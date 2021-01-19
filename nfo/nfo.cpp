// nfo.cpp : Defines the entry point for the DLL application.
//

#include "stdafx.h"
#include "resource.h"
#include "nfo.h"
#include "LeoHelpers.h"
#include "NfoViewerWindow.h"

// {0247DBD1-9679-44c3-8171-B586121E9D0D}
static const GUID GUIDPlugin_nfo =
{ 0x247dbd1, 0x9679, 0x44c3, { 0x81, 0x71, 0xb5, 0x86, 0x12, 0x1e, 0x9d, 0xd } };

static HMODULE s_hDllModule = NULL;
static CRITICAL_SECTION s_cs;
static int s_refCount = 0;
static ATOM s_atomViewer = 0;

void S_RegisterWindowClasses()
{
	if (s_atomViewer == 0)
	{
		s_atomViewer = CNfoViewerWindow::RegisterWindowClass(s_hDllModule);
	}
}

void S_FreeWindowClasses()
{
	if (s_atomViewer != 0)
	{
		CNfoViewerWindow::UnregisterWindowClass(s_hDllModule);
		s_atomViewer = 0;
	}
}

extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
	switch(dwReason)
	{
	case(DLL_PROCESS_ATTACH):
		// If DLL_PROCESS_ATTACH fails it should clean up everything it did
		// as DllMain won't be called again after such a failure.
		s_hDllModule = reinterpret_cast<HMODULE>(hInstance);
		InitializeCriticalSection(&s_cs);
	//	DisableThreadLibraryCalls(s_hDllModule); // Leo 03/Aug/2009: Removed. Doing this is wrong (and most likely the call failed) when we use the static CRT.
		break;
	case(DLL_PROCESS_DETACH):
		DeleteCriticalSection(&s_cs);
		break;
	case(DLL_THREAD_ATTACH):
	case(DLL_THREAD_DETACH):
		break;
	default:
		break;
	}

	return TRUE;
}

BOOL DVP_Init()
{
	// Only old versions of Opus will call DVP_Init when we also expose DVP_InitEx.
	// We use this as a clean, definite way to stop old versions loading us by mistake.

	return FALSE;
}

BOOL DVP_InitEx(LPDVPINITEXDATA pInitExData)
{
	// Note: DVP_InitEx won't even be called before Opus 9, which is the minimum supported version.

	// Note: We assume that Opus has initialized the common controls and COM for all
	// threads which call our code.

	// DVP_InitEx and DVP_Uninit must keep a refcount as Opus may init and uninit us multiple times.
	// (e.g. Click the Refresh button in the plugins list to trigger an additional Init/Uninit pair.)

	LeoHelpers::CriticalSectionScoper css(&s_cs);

	BOOL bInitSuccess = FALSE;

	if (0 < s_refCount++)
	{
		bInitSuccess = TRUE;
	}
	else
	{
		S_RegisterWindowClasses();

		if (s_atomViewer != 0)
		{
			DWORD64 dw64OpusVersion = pInitExData->dwOpusVerMajor;
			dw64OpusVersion <<= 32;
			dw64OpusVersion += pInitExData->dwOpusVerMinor;

			if (dw64OpusVersion >= MAKE64BITVERSIONNUMBER(9,1,3,0))
			{
				bInitSuccess = TRUE;
			}
		}

		if (!bInitSuccess)
		{
			DVP_Uninit(); // Opus does not call Uninit if Init fails. Clean ourselves up.
		}
	}

	return(bInitSuccess);
}

void DVP_Uninit(void)
{
	LeoHelpers::CriticalSectionScoper css(&s_cs);

	if (0 < s_refCount)
	{
		if (0 == --s_refCount)
		{
			S_FreeWindowClasses();
		}
	}
}

BOOL DVP_USBSafe(LPOPUSUSBSAFEDATA pUSBSafeData)
{
	return TRUE;
}

BOOL DVP_Identify(LPVIEWERPLUGININFO lpVPInfo)
{
	BOOL bResult = FALSE;

	lpVPInfo->dwFlags = DVPFIF_CanHandleStreams
					  | DVPFIF_ExtensionsOnly
					  | DVPFIF_NoThumbnails
					  | DVPFIF_NoFileInformation;
					  //| DVPFIF_CanConfigure;

	lpVPInfo->lpszHandleExts = _T(".nfo;.diz;.asc");
	lpVPInfo->dwlMinFileSize = 1;
	lpVPInfo->dwlMaxFileSize = 512000;
	lpVPInfo->dwlMinPreviewFileSize = lpVPInfo->dwlMinFileSize;
	lpVPInfo->dwlMaxPreviewFileSize = lpVPInfo->dwlMaxFileSize;
	lpVPInfo->uiMajorFileType = DVPMajorType_Text;
	lpVPInfo->idPlugin = GUIDPlugin_nfo;

	if (LeoHelpers::CopyVersionResourceToViewerPluginInfo(s_hDllModule, IDI_MAIN, STR_NFO_PLUGIN_NAME, STR_NFO_PLUGIN_DESCRIPTION, lpVPInfo))
	{
		// If initialisation fails after the call to CopyVersionResourceToViewerPluginInfo then the lpVPInfo->hIconSmall icon must be destroyed.
		bResult = TRUE;
	}

	return(bResult);
}

BOOL DVP_IdentifyFile(HWND hWnd,LPTSTR lpszName,LPVIEWERPLUGINFILEINFO lpVPFileInfo,HANDLE hAbortEvent)
{
	lpVPFileInfo->dwFlags = DVPFIF_CanReturnViewer;

	lpVPFileInfo->wMajorType = DVPMajorType_Text;
	lpVPFileInfo->wMinorType = 0;

	lpVPFileInfo->szImageSize.cx = 0;
	lpVPFileInfo->szImageSize.cy = 0;
	lpVPFileInfo->iNumBits = 0;

	if (NULL != lpVPFileInfo->lpszInfo && 0 != lpVPFileInfo->cchInfoMax)
	{
		lpVPFileInfo->lpszInfo[0] = _T('\0');
	}

	return(TRUE);
}

BOOL DVP_IdentifyFileStream(HWND hWnd,LPSTREAM lpStream,LPTSTR lpszName,LPVIEWERPLUGINFILEINFO lpVPFileInfo,DWORD dwStreamFlags)
{
	return(DVP_IdentifyFile(hWnd, lpszName, lpVPFileInfo, NULL));
}

//HWND DVP_Configure(HWND hWndParent,HWND hWndNotify,DWORD dwNotifyData)
//{
//	NNfoConfig::NNC_CongifData cd;
//	cd.hDllModule		= s_hDllModule;
//	cd.hWndNotify		= hWndNotify;
//	cd.dwNotifyData		= dwNotifyData;
//	cd.dw64OpusVersion	= getOpusVersionNumber();
//	cd.hWndViewer		= NULL;
//
//	return(CreateDialogParam(s_hDllModule, MAKEINTRESOURCE(IDD_CONFIG), hWndParent,
//								NNfoConfig::configDlgProc,
//								reinterpret_cast<LPARAM>(&cd)));
//}

HWND DVP_CreateViewer(HWND hWnd,LPRECT lpRc,DWORD dwFlags)
{
	return(CNfoViewerWindow::CreateViewerWindow(s_hDllModule, hWnd, lpRc, dwFlags));
}
