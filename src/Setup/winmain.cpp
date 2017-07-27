// Setup.cpp : Defines the entry point for the application.
//

#include "stdafx.h"
#include "Setup.h"
#include "FxHelper.h"
#include "UpdateRunner.h"
#include "MachineInstaller.h"
#include "windows.h"
#include "fcntl.h"
#include "io.h"
#include <cstdio>

CAppModule* _Module;

typedef BOOL(WINAPI *SetDefaultDllDirectoriesFunction)(DWORD DirectoryFlags);

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                      _In_opt_ HINSTANCE hPrevInstance,
                      _In_ LPWSTR lpCmdLine,
                      _In_ int nCmdShow)
{
	// Attempt to mitigate http://textslashplain.com/2015/12/18/dll-hijacking-just-wont-die
	HMODULE hKernel32 = LoadLibrary(L"kernel32.dll");
	ATLASSERT(hKernel32 != NULL);

	SetDefaultDllDirectoriesFunction pfn = (SetDefaultDllDirectoriesFunction) GetProcAddress(hKernel32, "SetDefaultDllDirectories");
	if (pfn) { (*pfn)(LOAD_LIBRARY_SEARCH_SYSTEM32); }

	int exitCode = -1;
	CString cmdLine(lpCmdLine);

	if (cmdLine.Find(L"--checkInstall") >= 0) {
		// If we're already installed, exit as fast as possible
		if (!MachineInstaller::ShouldSilentInstall()) {
			return 0;
		}

		// Make sure update.exe gets silent
		wcscat(lpCmdLine, L" --silent");
	}

	HRESULT hr = ::CoInitialize(NULL);
	ATLASSERT(SUCCEEDED(hr));

	AtlInitCommonControls(ICC_COOL_CLASSES | ICC_BAR_CLASSES);
	_Module = new CAppModule();
	hr = _Module->Init(NULL, hInstance);

	bool isQuiet = (cmdLine.Find(L"-s") >= 0);
	bool weAreUACElevated = CUpdateRunner::AreWeUACElevated() == S_OK;
	bool attemptingToRerun = (cmdLine.Find(L"--rerunningWithoutUAC") >= 0);

	bool allUsersInstall = (cmdLine.Find(L"--allUsers") >= 0);

	if (allUsersInstall && !weAreUACElevated && !isQuiet)
	{
		// We will give a warning but NOT exit. Per BL-3404, it seems sometimes AreWeUACElevated
		// does not accurately indicate whether we will be able to install in program files.
		// So, we try to give the user a hint of what might be wrong in case it fails, but
		// go ahead and attempt the installation.

		// This bit of magic allows a windows application to attach to the console of its parent
		// (that is, to write to the DOS box from which we hope it was launched).
		// Currently (since about 3.9) it doesn't seem to be working, and I can't figure out why.
		// So I've gone back to the messagebox below.
		// Keeping this code for now since it just possibly might do something helpful in a network
		// install situation where the message box can't be seen.
		// If we can't attach to a parent console we just give up on sending this warning.
		// Some of the examples from which this code was adapted create a private console for the
		// program, but that isn't helpful in our most likely allUsers scenario when a domain
		// admin is installing on a remote machine.
		// If we really don't have admin privileges we will soon dialog saying we can't unpack the files.
		if (AttachConsole(ATTACH_PARENT_PROCESS)) {
			// A further part of the magic is to redirect unbuffered STDERR to the console.
			// We don't bother with stdout because we don't use it, but if we did it would
			// require its own redirection.
			HANDLE consoleHandleError = GetStdHandle(STD_ERROR_HANDLE);
			int fdError = _open_osfhandle((intptr_t)consoleHandleError, _O_TEXT);
			FILE * fpError = _fdopen(fdError, "w");
			*stderr = *fpError;
			setvbuf(stderr, NULL, _IONBF, 0);

			// Now we can actually write a message to stderr.
			fwprintf(stderr, L"\nWarning: all-users install requires the installer to be run with administrator privileges.\n");
			fflush(stderr);

			// Unfortunately, the dos box doesn't know to wait for this application to finish, so the
			// output appears following a command prompt. This attempts to type a return key to get a fresh prompt.
			INPUT ip;
			// Set up a generic keyboard event.
			ip.type = INPUT_KEYBOARD;
			ip.ki.wScan = 0; // hardware scan code for key
			ip.ki.time = 0;
			ip.ki.dwExtraInfo = 0;
			// Send the "Enter" key
			ip.ki.wVk = 0x0D; // virtual-key code for the "Enter" key
			ip.ki.dwFlags = 0; // 0 for key press
			SendInput(1, &ip, sizeof(INPUT));
			// Release the "Enter" key
			ip.ki.dwFlags = KEYEVENTF_KEYUP;  // KEYEVENTF_KEYUP for key release
			SendInput(1, &ip, sizeof(INPUT));
		}

		if (MessageBox(0L, L"Warning: you appear to be attempting an allUsers install without required administrator privileges.\nDo you want to try anyway?",
			L"Need to run as Administrator",
			MB_YESNO | MB_DEFBUTTON2 | MB_ICONWARNING) != IDYES)
		{
			exitCode = E_FAIL;
			goto out;
		}
	}

	if (weAreUACElevated && attemptingToRerun) {
		CUpdateRunner::DisplayErrorMessage(CString(L"Please re-run this installer as a normal user instead of \"Run as Administrator\"."), NULL);
		exitCode = E_FAIL;
		goto out;
	}

	if (!CFxHelper::CanInstallDotNet4_5()) {
		// Explain this as nicely as possible and give up.
		MessageBox(0L, L"Bloom cannot run on Windows XP or before; it requires a later version of Windows. You can still find Bloom 3.0, which runs on XP, on our website.", L"Incompatible Operating System", 0);
		exitCode = E_FAIL;
		goto out;
	}

	NetVersion requiredVersion = CFxHelper::GetRequiredDotNetVersion();

	if (!CFxHelper::IsDotNetInstalled(requiredVersion)) {
		hr = CFxHelper::InstallDotNetFramework(requiredVersion, isQuiet);
		if (FAILED(hr)) {
			exitCode = hr; // #yolo
			CUpdateRunner::DisplayErrorMessage(CString(L"Failed to install the .NET Framework, try installing the latest version manually"), NULL);
			goto out;
		}
	
		// S_FALSE isn't failure, but we still shouldn't try to install
		if (hr != S_OK) {
			exitCode = 0;
			goto out;
		}
	}

	// If we're UAC-elevated (and not doing all-users install), we shouldn't be because it will give us permissions
	// problems later. Just silently rerun ourselves.
	if (weAreUACElevated && !allUsersInstall) {
		wchar_t buf[4096];
		HMODULE hMod = GetModuleHandle(NULL);
		GetModuleFileNameW(hMod, buf, 4096);
		wcscat(lpCmdLine, L" --rerunningWithoutUAC");

		CUpdateRunner::ShellExecuteFromExplorer(buf, lpCmdLine);
		exitCode = 0;
		goto out;
	}

	exitCode = CUpdateRunner::ExtractUpdaterAndRun(lpCmdLine, false);

out:
	_Module->Term();
	return exitCode;
}
