#pragma once
// public methods defined in SplashImageUtils.cpp
HANDLE InitSplash(HINSTANCE hInstance, std::wstring splashPath);
DWORD PumpMsgWaitingForEvent(HANDLE hProcess, HANDLE hCloseSplashEvent, DWORD dwMilliseconds);
