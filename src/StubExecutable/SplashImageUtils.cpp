// SplashImageUtils.cpp : methods for displaying a splash image while starting
// the real program.  Many of these methods are adapted from those found at
// http://faithlife.codes/blog/2008/09/displaying_a_splash_screen_with_c_introduction/
// The original code is Copyright 2007-2008 Logos Bible Software and licensed
// under an MIT style license according  to
// http://faithlife.codes/blog/2008/09/logos_code_blog_license/.  The revisions
// and additions are Copyright 2017 SIL International and licensed under the same
// MIT style license.

#include "stdafx.h"
#include "Resource.h"

// forward declarations of local methods.
std::wstring GetSplashPath(std::wstring fullPath);
BOOL InitSplashWindows(HINSTANCE hInstance, LPCTSTR splashPath);
HWND CreateSplashWindow(HINSTANCE hInstance);
HBITMAP LoadSplashImage(LPCTSTR splashPath);
BOOL SetSplashImage(HWND hwndSplash, HBITMAP hbmpSplash);
ATOM RegisterWindowClass(HINSTANCE hInstance);

// create Splash window
HANDLE InitSplash(HINSTANCE hInstance, std::wstring fullPath)
{
	HANDLE hCloseSplashEvent = NULL;
	std::wstring splashPath = GetSplashPath(fullPath);
	if (splashPath.length() > 0)
	{
		// Create an event, making sure we're the first active process to create it.
		SetLastError(ERROR_SUCCESS);
		hCloseSplashEvent = CreateEvent(NULL, TRUE, FALSE, _T("CloseSquirrelSplashScreenEvent"));
		if (GetLastError() == ERROR_ALREADY_EXISTS)
			return NULL;
		::CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
		if (!InitSplashWindows(hInstance, splashPath.c_str()))
		{
			hCloseSplashEvent = NULL;
			::CoUninitialize();
		}
	}
	return hCloseSplashEvent;
}

// Replace the ".exe" at the end of the pathname with "Splash.png" and check
// whether the file exists.
std::wstring GetSplashPath(std::wstring fullPath)
{
	if (fullPath.length() > 4)
	{
		std::wstring path = fullPath;
		path = path.replace(path.length() - 4, 4, L"Splash.png");
		if (PathFileExists(path.c_str()))
			return path;
	}
	return std::wstring(L"");
}

BOOL InitSplashWindows(HINSTANCE hInstance, LPCTSTR splashPath)
{
	if (!RegisterWindowClass(hInstance))
		return FALSE;
	HWND hwndSplash = CreateSplashWindow(hInstance);
	if (!hwndSplash)
		return FALSE;
	HBITMAP hbmpSplash = LoadSplashImage(splashPath);
	if (!hbmpSplash)
		return FALSE;
	return SetSplashImage(hwndSplash, hbmpSplash);
}

// Creates a 32-bit DIB from the specified WIC bitmap.
HBITMAP CreateHBITMAP(IWICBitmapSource * ipBitmap)
{
	// initialize return value
	HBITMAP hbmp = NULL;

	// get image attributes and check for valid image
	UINT width = 0;
	UINT height = 0;
	if (FAILED(ipBitmap->GetSize(&width, &height)) || width == 0 || height == 0)
		return NULL;

	// prepare structure giving bitmap information (negative height indicates a top-down DIB)
	BITMAPINFO bminfo;
	::ZeroMemory(&bminfo, sizeof(bminfo));
	bminfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
	bminfo.bmiHeader.biWidth = width;
	bminfo.bmiHeader.biHeight = -((LONG)height);
	bminfo.bmiHeader.biPlanes = 1;
	bminfo.bmiHeader.biBitCount = 32;
	bminfo.bmiHeader.biCompression = BI_RGB;

	// create a DIB section that can hold the image
	void * pvImageBits = NULL;
	HDC hdcScreen = ::GetDC(NULL);
	if (hdcScreen)
	{
		hbmp = ::CreateDIBSection(hdcScreen, &bminfo, DIB_RGB_COLORS, &pvImageBits, NULL, 0);
		::ReleaseDC(NULL, hdcScreen);
	}
	if (hbmp == NULL)
		return NULL;

	// extract the image into the HBITMAP
	const UINT cbStride = width * 4;
	const UINT cbImage = cbStride * height;
	if (FAILED(ipBitmap->CopyPixels(NULL, cbStride, cbImage, static_cast<BYTE *>(pvImageBits))))
	{
		// couldn't extract image; delete HBITMAP
		DeleteObject(hbmp);
		hbmp = NULL;
	}
	return hbmp;
}

// Loads a PNG image from the specified stream (using Windows Imaging Component).
IWICBitmapSource * LoadBitmapFromStream(IStream * ipImageStream)
{
	// load WIC's PNG decoder
	IWICBitmapDecoder * ipDecoder = NULL;
	if (FAILED(::CoCreateInstance(CLSID_WICPngDecoder, NULL, CLSCTX_INPROC_SERVER, __uuidof(ipDecoder), reinterpret_cast<void**>(&ipDecoder))))
		return NULL;

	// initialize return value
	IWICBitmapSource * ipBitmap = NULL;
	// load the PNG
	if (SUCCEEDED(ipDecoder->Initialize(ipImageStream, WICDecodeMetadataCacheOnLoad)))
	{
		// check for the presence of the first frame in the bitmap
		UINT nFrameCount = 0;
		if (SUCCEEDED(ipDecoder->GetFrameCount(&nFrameCount)) || nFrameCount != 1)
		{
			// load the first frame (i.e., the image)
			IWICBitmapFrameDecode * ipFrame = NULL;
			if (SUCCEEDED(ipDecoder->GetFrame(0, &ipFrame)))
			{
				// convert the image to 32bpp BGRA format with pre-multiplied alpha
				//   (it may not be stored in that format natively in the PNG resource,
				//   but we need this format to create the DIB to use on-screen)
				::WICConvertBitmapSource(GUID_WICPixelFormat32bppPBGRA, ipFrame, &ipBitmap);
				ipFrame->Release();
			}
		}
	}
	ipDecoder->Release();
	return ipBitmap;
}

// Loads the PNG containing the splash image into a HBITMAP.
HBITMAP LoadSplashImage(LPCTSTR splashPath)
{
	HBITMAP hbmpSplash = NULL;
	// load the PNG image data into a stream
	IStream * ipImageStream = NULL;
	if (SUCCEEDED(SHCreateStreamOnFileEx(splashPath, STGM_READ, 0, FALSE, NULL, &ipImageStream)))
	{
		// load the bitmap with WIC
		IWICBitmapSource * ipBitmap = LoadBitmapFromStream(ipImageStream);
		if (ipBitmap != NULL)
		{
			// create a HBITMAP containing the image
			hbmpSplash = CreateHBITMAP(ipBitmap);
			ipBitmap->Release();
		}
		ipImageStream->Release();
	}
	return hbmpSplash;
}

// Window Class name
const TCHAR * c_szSplashClass = _T("SquirrelSplashWindow");

// Registers a window class for the splash and splash owner windows.
ATOM RegisterWindowClass(HINSTANCE hInstance)
{
	WNDCLASS wc = { 0 };
	wc.lpfnWndProc = DefWindowProc;
	wc.hInstance = hInstance;
	wc.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_STUBEXECUTABLE));
	wc.hCursor = LoadCursor(NULL, IDC_ARROW);
	wc.lpszClassName = c_szSplashClass;
	return ::RegisterClass(&wc);
}

// Creates the splash owner window and the splash window.
HWND CreateSplashWindow(HINSTANCE hInstance)
{
	HWND hwndOwner = ::CreateWindow(c_szSplashClass, NULL, WS_POPUP,
		0, 0, 0, 0, NULL, NULL, hInstance, NULL);
	if (!hwndOwner)
		return NULL;
	return ::CreateWindowEx(WS_EX_LAYERED, c_szSplashClass, NULL, WS_POPUP | WS_VISIBLE,
		0, 0, 0, 0, hwndOwner, NULL, hInstance, NULL);
}

// Calls UpdateLayeredWindow to set a bitmap (with alpha) as the content of the splash window.
BOOL SetSplashImage(HWND hwndSplash, HBITMAP hbmpSplash)
{
	// get the size of the bitmap
	BITMAP bm;
	if (!::GetObject(hbmpSplash, sizeof(bm), &bm))
		return FALSE;
	SIZE sizeSplash = { bm.bmWidth, bm.bmHeight };

	// get the primary monitor's info
	POINT ptZero = { 0 };
	HMONITOR hmonPrimary = ::MonitorFromPoint(ptZero, MONITOR_DEFAULTTOPRIMARY);
	if (!hmonPrimary)
		return FALSE;
	MONITORINFO monitorinfo = { 0 };
	monitorinfo.cbSize = sizeof(monitorinfo);
	if (!::GetMonitorInfo(hmonPrimary, &monitorinfo))
		return FALSE;

	// center the splash screen in the middle of the primary work area
	const RECT & rcWork = monitorinfo.rcWork;
	POINT ptOrigin;
	ptOrigin.x = rcWork.left + (rcWork.right - rcWork.left - sizeSplash.cx) / 2;
	ptOrigin.y = rcWork.top + (rcWork.bottom - rcWork.top - sizeSplash.cy) / 2;

	// create a memory DC holding the splash bitmap
	HDC hdcScreen = ::GetDC(NULL);
	if (!hdcScreen)
		return FALSE;
	HDC hdcMem = ::CreateCompatibleDC(hdcScreen);
	if (!hdcMem)
	{
		::ReleaseDC(NULL, hdcScreen);
		return FALSE;
	}
	HBITMAP hbmpOld = (HBITMAP)::SelectObject(hdcMem, hbmpSplash);
	if (!hbmpOld || hbmpOld == HGDI_ERROR)
	{
		::DeleteDC(hdcMem);
		::ReleaseDC(NULL, hdcScreen);
		return FALSE;
	}

	// use the source image's alpha channel for blending
	BLENDFUNCTION blend = { 0 };
	blend.BlendOp = AC_SRC_OVER;
	blend.SourceConstantAlpha = 255;
	blend.AlphaFormat = AC_SRC_ALPHA;

	// paint the window (in the right location) with the alpha-blended bitmap
	BOOL retval = ::UpdateLayeredWindow(hwndSplash, hdcScreen, &ptOrigin, &sizeSplash,
		hdcMem, &ptZero, RGB(0, 0, 0), &blend, ULW_ALPHA);

	// delete temporary objects
	::SelectObject(hdcMem, hbmpOld);
	::DeleteDC(hdcMem);
	::ReleaseDC(NULL, hdcScreen);
	return retval;
}

inline DWORD PumpMsgWaitForMultipleObjects(DWORD nCount, LPHANDLE pHandles, DWORD dwMilliseconds)
{
	// useful variables
	const DWORD dwStartTickCount = ::GetTickCount();

	// loop until done
	for (;;)
	{
		// calculate timeout
		const DWORD dwElapsed = ::GetTickCount() - dwStartTickCount;
		const DWORD dwTimeout = dwMilliseconds == INFINITE ? INFINITE :
			dwElapsed < dwMilliseconds ? dwMilliseconds - dwElapsed : 0;

		// wait for a handle to be signaled or a message
		const DWORD dwWaitResult = ::MsgWaitForMultipleObjects(nCount, pHandles, FALSE, dwTimeout, QS_ALLINPUT);
		if (dwWaitResult == WAIT_OBJECT_0 + nCount)
		{
			// pump messages
			MSG msg;
			while (::PeekMessage(&msg, NULL, 0, 0, PM_REMOVE) != FALSE)
			{
				// check for WM_QUIT
				if (msg.message == WM_QUIT)
				{
					// repost quit message and return
					PostQuitMessage((int)msg.wParam);
					return WAIT_OBJECT_0 + nCount;
				}

				// dispatch thread message
				::TranslateMessage(&msg);
				::DispatchMessage(&msg);
			}
		}
		else
		{
			// timeout on actual wait or any other object
			return dwWaitResult;
		}
	}
}

DWORD PumpMsgWaitingForEvent(HANDLE hProcess, HANDLE hCloseSplashEvent, DWORD dwMilliseconds)
{
	HANDLE aHandles[2] = { hProcess, hCloseSplashEvent };
	DWORD ret = PumpMsgWaitForMultipleObjects(2, &aHandles[0], 60 * 1000);
	::CoUninitialize();
	return ret;
}
