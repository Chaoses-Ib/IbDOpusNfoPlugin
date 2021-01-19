#pragma once

extern "C"
{
	__declspec(dllexport) BOOL DVP_Init();
	__declspec(dllexport) BOOL DVP_InitEx(LPDVPINITEXDATA pInitExData);
	__declspec(dllexport) void DVP_Uninit(void);
	__declspec(dllexport) BOOL DVP_Identify(LPVIEWERPLUGININFO lpVPInfo);
	__declspec(dllexport) BOOL DVP_IdentifyFile(HWND hWnd,LPTSTR lpszName,LPVIEWERPLUGINFILEINFO lpVPFileInfo,HANDLE hAbortEvent);
	__declspec(dllexport) BOOL DVP_IdentifyFileStream(HWND hWnd,LPSTREAM lpStream,LPTSTR lpszName,LPVIEWERPLUGINFILEINFO lpVPFileInfo,DWORD dwStreamFlags);
	__declspec(dllexport) HWND DVP_CreateViewer(HWND hWnd,LPRECT lpRc,DWORD dwFlags);
//	__declspec(dllexport) HWND DVP_Configure(HWND hWndParent,HWND hWndNotify,DWORD dwNotifyData);
	__declspec(dllexport) BOOL DVP_USBSafe(LPOPUSUSBSAFEDATA pUSBSafeData);
};
