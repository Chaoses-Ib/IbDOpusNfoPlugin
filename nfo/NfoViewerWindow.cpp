#include "StdAfx.h"
#include "LeoHelpers.h"
#include "nfo.h"
#include "nfoviewerwindow.h"

CNfoViewerWindow::CNfoViewerWindow(HMODULE hDllModule, HWND hWnd, LPRECT lpRc, DWORD dwFlags)
:	m_hWnd(hWnd)
,	m_hInstance(hDllModule)
,	m_dwOpusViewerFlags(dwFlags)
,	m_szName(NULL)
,	m_dwNumBytes(0)
,	m_dwMaxLineWidth(0)
,	m_dwCapabilities(0)
,	m_hWndEdit(NULL)
,	m_bReportedFontError(false)
,	m_hFont(NULL)
//,	m_nFontSize(9)
#ifndef UNICODE
,	m_pCurrentLogFont(NULL)
#endif
{
	// initialize other stuff.
	reset(true);

#ifndef UNICODE
	const TCHAR *szFontName = _T("IBMPC");

	// Get IBMPC font sizes.
	HDC hDC = GetDC(hWnd);
	if (NULL != hDC)
	{
		LOGFONT logFont;
		ZeroMemory(&logFont, sizeof(logFont));
		_tcscpy_s(logFont.lfFaceName, LF_FACESIZE, szFontName);
		logFont.lfCharSet = OEM_CHARSET;
		logFont.lfPitchAndFamily = FIXED_PITCH|FF_DONTCARE;

		EnumFontFamiliesEx(hDC, &logFont, fontEnumCallback, reinterpret_cast<LPARAM>(this), 0);

		ReleaseDC(hWnd, hDC);
		hDC = NULL;
	}
#endif

	// Create edit control
	m_hWndEdit = CreateWindowEx(0, _T("EDIT"), _T(""),
#ifndef UNICODE
								ES_OEMCONVERT |
#endif
								ES_AUTOHSCROLL | ES_AUTOVSCROLL | ES_LEFT | ES_MULTILINE | ES_NOHIDESEL | ES_READONLY | ES_WANTRETURN | \
								WS_CHILD|WS_VISIBLE|WS_CLIPCHILDREN|WS_CLIPSIBLINGS|/*WS_HSCROLL|*/WS_VSCROLL,
								0, 0, lpRc->right-lpRc->left, lpRc->bottom-lpRc->top,
								m_hWnd, NULL, m_hInstance, NULL);

	// Subclass edit control
	if (NULL != m_hWndEdit)
	{
		assert(::IsWindowUnicode(m_hWndEdit));

		#pragma warning(suppress:4312) // spurious warning due to stupid Win32 SDK header definition.
		#pragma warning(suppress:4244) // spurious warning due to stupid Win32 SDK header definition.
		SetProp(m_hWndEdit, LeoHelpers::s_SubclassAtomPropHolder.GetAtomCast(), reinterpret_cast<WNDPROC>(SetWindowLongPtr(m_hWndEdit, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(editSubProc))));
	}

#ifndef UNICODE
	if (m_mapFontWidthToLogFont.empty())
	{
		reportFontError(szFontName);
	}
#endif

	updateFont(hWnd);

	updateCapabilities(NULL);
}

CNfoViewerWindow::~CNfoViewerWindow(void)
{
	if (NULL != m_hWndEdit)
	{
		#pragma warning(suppress:4244) // spurious warning due to stupid Win32 SDK header definition.
		SetWindowLongPtr(m_hWndEdit, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>( RemoveProp(m_hWndEdit, LeoHelpers::s_SubclassAtomPropHolder.GetAtomCast()) ));

		DestroyWindow(m_hWndEdit);
		m_hWndEdit = NULL;
	}

	if (NULL != m_hFont)
	{
		DeleteObject(m_hFont);
		m_hFont = NULL;
	}

	reset(true);
}

void CNfoViewerWindow::reset(bool bForgetAbortEvent)
{
	delete [] m_szName;
	m_szName = NULL;

	if (bForgetAbortEvent)
	{
		m_hAbortEvent = NULL;
	}
}

// static
ATOM CNfoViewerWindow::RegisterWindowClass(HMODULE hDllModule)
{
	WNDCLASS wc;

	// Register window class
	wc.style			= 0;
	wc.lpfnWndProc		= CNfoViewerWindow::wndProc;
	wc.cbClsExtra		= 0;
	wc.cbWndExtra		= 0;
	wc.hInstance		= hDllModule;
	wc.hIcon			= 0;
	wc.hCursor			= 0;
	wc.hbrBackground	= reinterpret_cast<HBRUSH>(COLOR_WINDOW+1);
	wc.lpszMenuName		= 0;
	wc.lpszClassName	= CNfoViewerWindow::getWindowClassName();

	return(RegisterClass(&wc));
}

// static
BOOL CNfoViewerWindow::UnregisterWindowClass(HMODULE hDllModule)
{
	return(UnregisterClass(CNfoViewerWindow::getWindowClassName(), hDllModule));
}

// static
HWND CNfoViewerWindow::CreateViewerWindow(HMODULE hDllModule, HWND hWndParent, LPRECT lpRc, DWORD dwFlags)
{
	RECT wrc;
	GetClientRect(hWndParent, &wrc);

	NVC_CreateParams cp;
	cp.lpRc = lpRc;
	cp.dwFlags = dwFlags;

	DWORD dwExStyle = 0; //WS_EX_NOPARENTNOTIFY; // WS_EX_NOPARENTNOTIFY is essential to avoid deadlock if on a separate thread.

//	const DWORD dwBorderFlag = DVPCVF_Border;
//	dwExStyle |= ((dwFlags&dwBorderFlag) ? WS_EX_CLIENTEDGE : 0);

	HWND hWnd = CreateWindowEx(	dwExStyle,
								getWindowClassName(),
								NULL,
								WS_CHILD|WS_VISIBLE|WS_CLIPCHILDREN|WS_CLIPSIBLINGS,
								wrc.left, wrc.top, wrc.right - wrc.left + 1, wrc.bottom - wrc.top + 1,
								hWndParent,
								NULL,
								hDllModule,
								&cp);
	return(hWnd);
}

// static
LRESULT WINAPI CNfoViewerWindow::wndProc(HWND hWnd,UINT uMsg,WPARAM wParam,LPARAM lParam)
{
	#pragma warning(suppress:4312) // spurious warning due to stupid Win32 SDK header definition.
	CNfoViewerWindow *pThis = reinterpret_cast<CNfoViewerWindow *>(GetWindowLongPtr(hWnd, GWLP_USERDATA));

	if (uMsg == WM_CREATE)
	{
		CREATESTRUCT *pCS = reinterpret_cast<CREATESTRUCT*>(lParam);
		NVC_CreateParams *pCP = (reinterpret_cast<NVC_CreateParams *>(pCS->lpCreateParams));
		pThis = new CNfoViewerWindow(pCS->hInstance, hWnd,
									 pCP == NULL ? NULL : pCP->lpRc,
									 pCP == NULL ? 0    : pCP->dwFlags);
		#pragma warning(suppress:4244) // spurious warning due to stupid Win32 SDK header definition.
		SetWindowLongPtr(hWnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(pThis));
		pThis->onCreate(hWnd);

		return 0;
	}
	else if (NULL != pThis)
	{
		switch (uMsg)
		{
		case WM_CLOSE:
			pThis->reset(true);
			DestroyWindow(hWnd);
			return 0;

		case WM_DESTROY:
			delete pThis;
			pThis = NULL;
			SetWindowLongPtr(hWnd, GWLP_USERDATA, NULL);
			return 0;

		case WM_CTLCOLORSTATIC:
			if (reinterpret_cast<HWND>(lParam) == pThis->m_hWndEdit)
			{
				return(reinterpret_cast<LRESULT>(GetSysColorBrush(COLOR_WINDOW)));
			}
			return NULL;

		case DVPLUGINMSG_GETCAPABILITIES:
			pThis->updateCapabilities(NULL);
			return pThis->m_dwCapabilities;

		case DVPLUGINMSG_SETABORTEVENT:
			pThis->m_hAbortEvent = reinterpret_cast<HANDLE>(lParam);
			return 0;

		case DVPLUGINMSG_GETIMAGEINFO:
			pThis->onDVPGetImageInfo(hWnd, reinterpret_cast<LPVIEWERPLUGINFILEINFO>(lParam));
			return TRUE;
			break;

		case DVPLUGINMSG_RESIZE:
			pThis->onSize(hWnd, false);
			return 0;

		case WM_SIZE:
			pThis->onSize(hWnd, wParam == SIZE_MINIMIZED);
			return 0;

		case DVPLUGINMSG_FULLSCREEN:
			if (wParam)
			{
				return TRUE; // We're going fullscreen; Don't allow it to happen.
			}
			return FALSE; // Allow change.

		//case DVPLUGINMSG_SETZOOM:
		//	{
		//		// Ignore initial zoom, for now at least.
		//	}
		//	return FALSE;

		//case DVPLUGINMSG_ZOOM:
		//	{

		//	}
		//	return TRUE;
		//case DVPLUGINMSG_GETZOOMFACTOR:
		//	{
		//		LONG lZoomFactor = 0;

		//		return(lZoomFactor);
		//	}
		//	return 0;

		//case DVPLUGINMSG_GETZOOMLIMITS:
		//	{
		//		// Return value indicates max/min zoom limits (HIWORD(max), LOWORD(min))
		//		DWORD dwZoomRange = 0;

		//		return(dwZoomRange);
		//	}
		//	return 0;

		case DVPLUGINMSG_TESTSELECTION:
			if (NULL != pThis->m_hWndEdit)
			{
				DWORD dwStart;
				DWORD dwEnd;
				SendMessage(pThis->m_hWndEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&dwStart), reinterpret_cast<LPARAM>(&dwEnd));
				return(dwStart < dwEnd ? TRUE : FALSE);
			}
			return FALSE;

		case DVPLUGINMSG_SELECTALL:
			if (NULL != pThis->m_hWndEdit)
			{
				SendMessage(pThis->m_hWndEdit, EM_SETSEL, 0, -1);
				return TRUE;
			}
			return FALSE;

		case DVPLUGINMSG_COPYSELECTION:
			if (NULL != pThis->m_hWndEdit)
			{
				DWORD dwStart;
				DWORD dwEnd;
				SendMessage(pThis->m_hWndEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&dwStart), reinterpret_cast<LPARAM>(&dwEnd));

				if (dwStart >= dwEnd)
				{
					//// Hide selection marker
					//SendMessage(pThis->m_hWndEdit,EM_HIDESELECTION,TRUE,0);
					// Turn off redraw
					SendMessage(pThis->m_hWndEdit,WM_SETREDRAW,FALSE,0);
					// Select All
					SendMessage(pThis->m_hWndEdit, EM_SETSEL, 0, -1);
				}

				SetCursor(LoadCursor(0,IDC_WAIT));
				SendMessage(pThis->m_hWndEdit, WM_COPY, 0, 0);
				SetCursor(LoadCursor(0,IDC_ARROW));

				if (dwStart >= dwEnd)
				{
					// Select None
					SendMessage(pThis->m_hWndEdit, EM_SETSEL, 0, 0);
					// Enable redraw
					SendMessage(pThis->m_hWndEdit,WM_SETREDRAW,TRUE,0);
					//// Show selection marker again
					//SendMessage(pThis->m_hWndEdit,EM_HIDESELECTION,FALSE,0);
				}

				return TRUE;
			}
			return FALSE;
			break;

		case DVPLUGINMSG_PRINT:
			return 0;

		case WM_ERASEBKGND:
			return 1;

		case WM_PAINT:
			return(pThis->onPaint(hWnd));

		case DVPLUGINMSG_LOAD:
			return(pThis->onDVPLoad(hWnd, reinterpret_cast<LPTSTR>(lParam)));

		case DVPLUGINMSG_LOADSTREAM:
			return(pThis->onDVPLoadStream(hWnd, reinterpret_cast<LPSTREAM>(lParam), reinterpret_cast<LPTSTR>(wParam)));

		case DVPLUGINMSG_CLEAR:
			pThis->onDVPClear(hWnd);
			return 0;

		case DVPLUGINMSG_REINITIALIZE:
			pThis->onDVPReinitialize(hWnd);
			return 0;

		//case WM_MOUSEACTIVATE:
		//	SetFocus(hWnd);
		//	return MA_ACTIVATE;

		case WM_SETFOCUS:
			if (NULL != pThis->m_hWndEdit)
			{
				SetFocus(pThis->m_hWndEdit);
			}
			break;

		case DVPLUGINMSG_TRANSLATEACCEL:
			return FALSE;

		case WM_MOUSEWHEEL:
			return SendMessage(GetParent(hWnd), uMsg, wParam, lParam);

		case DVPLUGINMSG_MOUSEWHEEL:
			if (NULL != pThis->m_hWndEdit)
			{
				return SendMessage(pThis->m_hWndEdit, DVPLUGINMSG_MOUSEWHEEL, wParam, lParam);
			}
			return FALSE;

//		case WM_SETCURSOR:
//			// If we're not covered by a control and able to get WM_SETCURSOR then we must be loading or in error.
//			SetCursor(LoadCursor(0,IDC_WAIT)); //IDC_ARROW
//			return 0;

		default:
			break;
		}
	}

	return(DefWindowProc(hWnd,uMsg,wParam,lParam));
}

// Sub-class for edit control
// static
LRESULT WINAPI CNfoViewerWindow::editSubProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	switch(uMsg)
	{
	case(WM_MOUSEWHEEL):
		// Intercept mousewheel message so that Opus gets a chance to handle it.
		return SendMessage(GetParent(GetParent(hWnd)), uMsg, wParam, lParam);

	case(DVPLUGINMSG_MOUSEWHEEL):
		// Cause the original Edit wndproc to receive a WM_MOUSEWHEEL message.
		uMsg = WM_MOUSEWHEEL;
		break;

	case(WM_KEYDOWN):
		// Intercept some keys and pass them back to Opus.
		switch(wParam)
		{
		case(VK_TAB):
			{
				NMKEY nmKey;
				ZeroMemory(&nmKey, sizeof(nmKey));
				nmKey.hdr.code = NM_KEYDOWN;
				nmKey.hdr.hwndFrom = GetParent(hWnd);
				nmKey.hdr.idFrom = 0;
				nmKey.nVKey = static_cast<UINT>(wParam);
				nmKey.uFlags = 0;
				SendMessage(GetParent(GetParent(hWnd)), WM_NOTIFY, 0, (LPARAM)&nmKey);
			}
			return 0;
		case('a'):
		case('A'):
			if (0x8000 & GetKeyState(VK_CONTROL))
			{
				SendMessage(GetParent(hWnd), DVPLUGINMSG_SELECTALL, 0, 0);
			}
			return 0;
		case('c'):
		case('C'):
			if (0x8000 & GetKeyState(VK_CONTROL))
			{
				SendMessage(GetParent(hWnd), DVPLUGINMSG_COPYSELECTION, 0, 0);
			}
			return 0;
		default:
			break;
		}
		break;

	case(WM_UNICHAR):
		// Make sure Unicode chars are converted to WM_CHAR messages rather than handled by the edit wndproc.
		return DefWindowProc(hWnd, uMsg, wParam, lParam);

	case(WM_CHAR):
		//return DefWindowProc(hWnd, uMsg, wParam, lParam);
		//switch(wParam)
		//{
		//case(VK_TAB):
		//	return 0;
		//case('a'):
		//case('A'):
		//	if (GetKeyState(VK_CONTROL))
		//	{
		//		SendMessage(GetParent(hWnd), DVPLUGINMSG_SELECTALL, 0, 0);
		//	}
		//	return 0;
		//default:
		//	break;
		//}
		break;

	case WM_SETFOCUS:
	case WM_KILLFOCUS:
		// If we get or lose focus, notify Opus.
		// Could be done without subclassing via EN_KILLFOCUS and EN_SETFOCUS, but since we're subclassing anyway...
		{
			DVPNMFOCUSCHANGE hdr;
			ZeroMemory(&hdr, sizeof(hdr));
			hdr.hdr.code = DVPN_FOCUSCHANGE;
			hdr.hdr.hwndFrom = GetParent(hWnd);
			hdr.hdr.idFrom = 0;
			hdr.fGotFocus = (uMsg == WM_SETFOCUS);
			SendMessage( GetParent(GetParent(hWnd)), WM_NOTIFY, 0, reinterpret_cast<LPARAM>(&hdr) );
		}
		break;

	case WM_CONTEXTMENU:
		PostMessage(GetParent(GetParent(hWnd)),WM_CONTEXTMENU,(WPARAM)hWnd,lParam);
		return 0;
		break;

	default:
		break;
	}

	// Default handling
	return CallWindowProc(reinterpret_cast<WNDPROC>( GetProp(hWnd, LeoHelpers::s_SubclassAtomPropHolder.GetAtomCast()) ), hWnd, uMsg, wParam, lParam);
}

void CNfoViewerWindow::onCreate(HWND hWnd)
{
}

LRESULT CNfoViewerWindow::onDVPLoad(HWND hWnd, LPTSTR szFilename)
{
	LRESULT lResult = FALSE;

	HANDLE hFile = CreateFile(szFilename, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL|FILE_FLAG_SEQUENTIAL_SCAN, NULL);

	if (INVALID_HANDLE_VALUE != hFile)
	{
		DWORD dwFileSizeLow = GetFileSize(hFile, NULL);

		if (INVALID_FILE_SIZE != dwFileSizeLow)
		{
			m_dwNumBytes = dwFileSizeLow;

			if (dwFileSizeLow != MAXDWORD)
			{
				dwFileSizeLow++; // For null-terminator. (If it's MAXDWORD then we'll overwrite the last char instead.)
			}

			HLOCAL hLocal = LocalAlloc(LMEM_MOVEABLE|LMEM_ZEROINIT, dwFileSizeLow);

			if (NULL != hLocal)
			{
				char *pLocalData = reinterpret_cast<char *>(LocalLock(hLocal));

				if (NULL != pLocalData)
				{
					DWORD dwFileSizeRead = 0;

					if (ReadFile(hFile, pLocalData, dwFileSizeLow-1, &dwFileSizeRead, NULL))
					{
						pLocalData[dwFileSizeLow-1] = '\0';

						LocalUnlock(hLocal);
						pLocalData = NULL;

						lResult = load(hWnd, szFilename, hLocal);

						hLocal = NULL; // It's now owned by the edit control, or was freed by the load function; don't free it again.
					}

					if (NULL != pLocalData)
					{
						LocalUnlock(hLocal);
						pLocalData = NULL;
					}
				}

				if (NULL != hLocal)
				{
					LocalFree(hLocal);
					hLocal = NULL;
				}
			}
		}

		CloseHandle(hFile);
	}

	return(lResult);
}

LRESULT CNfoViewerWindow::onDVPLoadStream(HWND hWnd, LPSTREAM pStream, LPTSTR szFilename)
{
	LRESULT lResult = FALSE;

	SIZE_T allocSize = 0;

	STATSTG statStg;
	ZeroMemory(&statStg, sizeof(statStg));

	if (SUCCEEDED(pStream->Stat(&statStg, STATFLAG_NONAME)))
	{
		allocSize = statStg.cbSize.LowPart;

		m_dwNumBytes = statStg.cbSize.LowPart;

		if (allocSize != MAXDWORD)
		{
			allocSize++; // For null-terminator. (If it's MAXDWORD then we'll overwrite the last char instead.)
		}

		HLOCAL hLocal = LocalAlloc(LMEM_MOVEABLE|LMEM_ZEROINIT, allocSize);

		if (NULL != hLocal)
		{
			char *pLocalData = reinterpret_cast<char *>(LocalLock(hLocal));

			if (NULL != pLocalData)
			{
				if (SUCCEEDED(pStream->Read(pLocalData, static_cast<ULONG>(allocSize-1), NULL)))
				{
					pLocalData[allocSize-1] = '\0';

					LocalUnlock(hLocal);
					pLocalData = NULL;

					lResult = load(hWnd, szFilename, hLocal);

					hLocal = NULL; // It's now owned by the edit control, or was freed by the load function; don't free it again.
				}

				if (NULL != pLocalData)
				{
					LocalUnlock(hLocal);
					pLocalData = NULL;
				}
			}

			if (NULL != hLocal)
			{
				LocalFree(hLocal);
				hLocal = NULL;
			}
		}
	}

	return(lResult);
}

LRESULT CNfoViewerWindow::load(HWND hWnd, const TCHAR *szName, HLOCAL hLocal)
{
	LRESULT lResult = FALSE;

	reset(false);
	updateCapabilities(hWnd);

	if (NULL != szName)
	{
		rsize_t len = _tcslen(szName) + 1;
		m_szName = new TCHAR[ len ];
		_tcscpy_s(m_szName, len, szName);
	}

	if (NULL != m_hWndEdit)
	{
		convertEOLs(hLocal);

		if (::IsWindowUnicode(m_hWndEdit))
		{
			convertToUnicode(hLocal);
		}

		HLOCAL hLocalOld = reinterpret_cast<HLOCAL>(SendMessage(m_hWndEdit, EM_GETHANDLE, 0, 0));

		if (0 != hLocalOld)
		{
			SendMessage(m_hWndEdit, EM_SETHANDLE, reinterpret_cast<WPARAM>(hLocal), 0);
			lResult = TRUE;

			LocalFree(hLocalOld);
		}
	}

	if (!lResult)
	{
		LocalFree(hLocal);
		reset(false); // If we failed then clean up any mess we made.
	}

	updateCapabilities(hWnd);

	return(lResult);
}

void CNfoViewerWindow::convertEOLs(HLOCAL &hLocal)
{
	bool bUseNew = false;
	HLOCAL hLocalNew = NULL;

	const char *pLocalDataIn = reinterpret_cast<const char *>(LocalLock(hLocal));

	if (NULL != pLocalDataIn)
	{
		bool bStuffToDo = false;
		DWORD dwBufferSize = 0;
		DWORD dwLineWidth = 0;

		m_dwMaxLineWidth = 0;

		{
			const char *pc = pLocalDataIn;
			while(pc[0] != '\0')
			{
				if (pc[0] != '\n' && pc[0] != '\r')
				{
					dwLineWidth++;
					dwBufferSize++;
					pc++;
				}
				else
				{
					if (m_dwMaxLineWidth < dwLineWidth)
					{
						m_dwMaxLineWidth = dwLineWidth;
					}

					dwLineWidth = 0;
					if (pc[1] != '\n' && pc[1] != '\r')
					{
						// EOL not in a pair. Will need an extra character.
						bStuffToDo = true;
						dwBufferSize++;
						dwBufferSize++;
						pc++;
					}
					else if (pc[0] == pc[1])
					{
						if (pc[0] == '\n')
						{
							// Two \n in a row. Treat as two lines.
							bStuffToDo = true;
							dwBufferSize++;
							dwBufferSize++;
							pc++;
							dwBufferSize++;
							dwBufferSize++;
							pc++;
						}
						else
						{
							// Two \r in a row. Skip the first one and handle the second one as if it's new.
							bStuffToDo = true;
							pc++;
						}
					}
					else
					{
						if (pc[0] != '\r' || pc[1] != '\n')
						{
							// EOLs in a pair but the wrong way around. Will need correcting.
							bStuffToDo = true;
						}

						dwBufferSize++;
						pc++;
						dwBufferSize++;
						pc++;
					}
				}
			}

			if (m_dwMaxLineWidth < dwLineWidth)
			{
				m_dwMaxLineWidth = dwLineWidth;
			}

			dwBufferSize++; // For null-termination.
		}

		if (bStuffToDo)
		{
			hLocalNew = LocalAlloc(LMEM_MOVEABLE|LMEM_ZEROINIT, dwBufferSize);

			if (NULL != hLocalNew)
			{
				char *pLocalNewData = reinterpret_cast<char *>(LocalLock(hLocalNew));

				if (NULL != pLocalNewData)
				{
					const char *pcIn = pLocalDataIn;
					char *pcOut = pLocalNewData;

					while(pcIn[0] != '\0')
					{
						if (pcIn[0] != '\n' && pcIn[0] != '\r')
						{
							*pcOut++ = *pcIn++;
						}
						else if (pcIn[1] != '\n' && pcIn[1] != '\r')
						{
							*pcOut++ = '\r';
							*pcOut++ = '\n';
							pcIn++;
						}
						else if (pcIn[0] == pcIn[1])
						{
							if (pcIn[0] == '\n')
							{
								// Two \n in a row. Treat as two lines.
								*pcOut++ = '\r';
								*pcOut++ = '\n';
								pcIn++;
								*pcOut++ = '\r';
								*pcOut++ = '\n';
								pcIn++;
							}
							else
							{
								// Two \r in a row. Skip the first one and handle the second one as if it's new.
								pcIn++;
							}
						}
						else
						{
							// Either \r\n or \n\r so output \r\n.
							*pcOut++ = '\r';
							*pcOut++ = '\n';
							pcIn++;
							pcIn++;
						}
					}
					*pcOut = '\0';

					bUseNew = true;

					LocalUnlock(hLocalNew);
					pLocalNewData = NULL;
				}
			}
		}

		LocalUnlock(hLocal);
		pLocalDataIn = NULL;
	}

	if (bUseNew)
	{
		LocalFree(hLocal);
		hLocal = hLocalNew;
	}
	else if (NULL != hLocalNew)
	{
		LocalFree(hLocalNew);
	}
}

void CNfoViewerWindow::convertToUnicode(HLOCAL &hLocal)
{
	bool bUseLocalUnicode = false;
	HLOCAL hLocalUnicode = NULL;

	const char *pLocalDataIn = reinterpret_cast<const char *>(LocalLock(hLocal));

	if (NULL != pLocalDataIn)
	{
		// Code page 437 is the IBM PC code page. http://en.wikipedia.org/wiki/Code_page_437
#ifdef UNICODE
		UINT uiCodePage = 437;
#else
		UINT uiCodePage = CP_ACP;
#endif

		int wcBufSize = MultiByteToWideChar(uiCodePage, 0, pLocalDataIn, -1, NULL, 0);

		hLocalUnicode = LocalAlloc(LMEM_MOVEABLE|LMEM_ZEROINIT, wcBufSize * sizeof(WCHAR));

		if (NULL != hLocalUnicode)
		{
			WCHAR *pLocalDataUnicode = reinterpret_cast<WCHAR *>(LocalLock(hLocalUnicode));

			if (NULL != pLocalDataUnicode)
			{
				// Hopefully this includes the '\0' in the copy and the count of characters,
				// else this will not work with empty strings.
				if (0 != MultiByteToWideChar(uiCodePage, 0, pLocalDataIn, -1, pLocalDataUnicode, wcBufSize))
				{
					bUseLocalUnicode = true;
				}

				LocalUnlock(hLocalUnicode);
				pLocalDataUnicode = NULL;
			}

			if (!bUseLocalUnicode)
			{
				LocalFree(hLocalUnicode);
				hLocalUnicode = NULL;
			}
		}

		LocalUnlock(hLocal);
		pLocalDataIn = NULL;
	}

	if (bUseLocalUnicode)
	{
		LocalFree(hLocal);
		hLocal = hLocalUnicode;
	}
	else if (NULL != hLocalUnicode)
	{
		LocalFree(hLocalUnicode);
	}
}


void CNfoViewerWindow::onDVPClear(HWND hWnd)
{
	reset(true);
	InvalidateRect(hWnd, NULL, FALSE);
	UpdateWindow(hWnd);

	NMHDR notHdr;
	notHdr.hwndFrom = hWnd;
	notHdr.idFrom = 0;
	notHdr.code = DVPN_CLEARED;

	SendMessage(GetParent(hWnd), WM_NOTIFY, 0, reinterpret_cast<LPARAM>(&notHdr));
}

void CNfoViewerWindow::onDVPReinitialize(HWND hWnd)
{
}

void CNfoViewerWindow::onDVPGetImageInfo(HWND hWnd, VIEWERPLUGINFILEINFO *pInfo)
{
	if (NULL != pInfo)
	{
        pInfo->dwFlags = DVPFIF_CanReturnViewer;

		pInfo->wMajorType = DVPMajorType_Text;
		pInfo->wMinorType = 0;

		pInfo->szImageSize.cx = 0;
		pInfo->szImageSize.cy = 0;
		pInfo->iNumBits = 0;

		if (NULL != pInfo->lpszInfo && 0 != pInfo->cchInfoMax)
		{
			LRESULT lNumLines = 1;

			if (NULL != m_hWndEdit)
			{
				lNumLines = SendMessage(m_hWndEdit, EM_GETLINECOUNT, 0, 0);
			}

			LeoHelpers::OpusStringLoader sl(&m_opusUtils);

			TCHAR *szInfoString = LeoHelpers::StringAllocAndFormat(sl.Get(STR_NFO_LINES_BYTES), lNumLines, m_dwNumBytes);

			LeoHelpers::StringCopy(pInfo->lpszInfo, szInfoString, pInfo->cchInfoMax);

			delete [] szInfoString;
		}
	}
}

void CNfoViewerWindow::onSize(HWND hWnd, bool bMinimized)
{
	if (!bMinimized)
	{
		if (NULL != m_hWndEdit)
		{
			RECT r;
			if (GetClientRect(hWnd, &r))
			{
				MoveWindow(m_hWndEdit, 0, 0, r.right, r.bottom, FALSE);
				updateFont(hWnd);
			}
		}
	}
}

LRESULT CNfoViewerWindow::onPaint(HWND hWnd)
{
	PAINTSTRUCT ps;

	HDC hDC = BeginPaint(hWnd, &ps);

	if (NULL != hDC)
	{
		EndPaint(hWnd, &ps);
	}

	return 0;
}

void CNfoViewerWindow::updateCapabilities(HWND hWnd)
{
	m_dwCapabilities = VPCAPABILITY_WANTMOUSEWHEEL
					 | VPCAPABILITY_WANTFOCUS
					 | VPCAPABILITY_CANTRACKFOCUS
					 | VPCAPABILITY_COPYSELECTION
					 | VPCAPABILITY_SELECTALL
					 | VPCAPABILITY_COPYALL
					 | VPCAPABILITY_NOFULLSCREEN;

	if (NULL != hWnd)
	{
		DVPNMCAPABILITIES notCaps;
		notCaps.hdr.hwndFrom = hWnd;
		notCaps.hdr.idFrom = 0;
		notCaps.hdr.code = DVPN_CAPABILITIES;
		notCaps.dwCapabilities = m_dwCapabilities;

		SendMessage(GetParent(hWnd), WM_NOTIFY, 0, reinterpret_cast<LPARAM>(&notCaps));
	}
}

void CNfoViewerWindow::updateFont(HWND hWnd)
{
	if (NULL != m_hWndEdit)
	{
		RECT r;
		GetClientRect(m_hWndEdit, &r);

		int maxCharWidth = (m_dwMaxLineWidth > 0 ? m_dwMaxLineWidth : 80);

		int nHeight = r.bottom / 10; // Try to show at least 10 lines.
		int nWidth = r.right / maxCharWidth;

		if (nHeight < 2) { nHeight = 2; }
		if (nWidth  < 1) { nWidth  = 1; }

		LOGFONT *pFontToCreate = NULL;

#ifdef UNICODE

		LONG fontHeight = nWidth * 2;
		LONG fontWidth  = nWidth;

		if (m_hFont == NULL
		||	m_currentLogFont.lfHeight != fontHeight
		||	m_currentLogFont.lfWidth  != fontWidth)
		{
			const TCHAR *szFontName = _T("Lucida Console");

			ZeroMemory(&m_currentLogFont, sizeof(m_currentLogFont));
			_tcscpy_s(m_currentLogFont.lfFaceName, LF_FACESIZE, szFontName);
			m_currentLogFont.lfPitchAndFamily = FIXED_PITCH|FF_DONTCARE;

			m_currentLogFont.lfHeight = fontHeight;
			m_currentLogFont.lfWidth  = fontWidth;

			m_currentLogFont.lfQuality = ANTIALIASED_QUALITY;

			pFontToCreate = &m_currentLogFont;
		}

#else
		if (!m_mapFontWidthToLogFont.empty())
		{
			// Find the first font that is too wide.
			std::map< LONG, LOGFONT >::iterator pLogFont = m_mapFontWidthToLogFont.upper_bound( nWidth );

			// Use a smaller font than the one we found (unless we got the smallest one).
			if (pLogFont != m_mapFontWidthToLogFont.begin())
			{
				--pLogFont;
			}

			// If the font is too tall, go for a smaller one...
			while (pLogFont->second.lfHeight > nHeight && pLogFont != m_mapFontWidthToLogFont.begin())
			{
				--pLogFont;
			}

			if (m_pCurrentLogFont != &(pLogFont->second))
			{
				m_pCurrentLogFont = &(pLogFont->second);

				pFontToCreate = m_pCurrentLogFont;
			}
		}
#endif
		if (pFontToCreate != NULL)
		{
			HFONT hFont = CreateFontIndirect(pFontToCreate);

			if (NULL == hFont)
			{
				reportFontError(pFontToCreate->lfFaceName);
			}
			else
			{
				SendMessage(m_hWndEdit, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);

				if (NULL != m_hFont)
				{
					DeleteObject(m_hFont);
				}

				m_hFont = hFont;
			}
		}
	}
}

#ifndef UNICODE
// static
int CALLBACK CNfoViewerWindow::fontEnumCallback(CONST LOGFONT *pLogFont, CONST TEXTMETRIC *pTextMetric, DWORD FontType, LPARAM lParam)
{
	CNfoViewerWindow *pThis = reinterpret_cast<CNfoViewerWindow *>(lParam);

	// Reject ugly and stupid fonts.
	if ((!(pLogFont->lfItalic || pLogFont->lfUnderline || pLogFont->lfStrikeOut)))
//	&&	pTextMetric->tmMaxCharWidth < pTextMetric->tmHeight
//	&&	pTextMetric->tmMaxCharWidth*7 > pTextMetric->tmHeight*4)
//	&&	0 != (RASTER_FONTTYPE & FontType))
	{
		if (pThis->m_mapFontWidthToLogFont.find( pTextMetric->tmMaxCharWidth ) != pThis->m_mapFontWidthToLogFont.end())
		{
			//MessageBeep(0);
		}
		else
		{
			LOGFONT *pMappedLogFont = &pThis->m_mapFontWidthToLogFont[ pTextMetric->tmMaxCharWidth ];

			*pMappedLogFont = *pLogFont;
			pMappedLogFont->lfOutPrecision = OUT_DEFAULT_PRECIS;
			pMappedLogFont->lfClipPrecision = CLIP_DEFAULT_PRECIS;
			pMappedLogFont->lfQuality = DEFAULT_QUALITY;
			pMappedLogFont->lfPitchAndFamily = FIXED_PITCH|FF_DONTCARE;
		}
	}

	return 1;
}
#endif

void CNfoViewerWindow::reportFontError(const TCHAR *szFontName)
{
	if (!m_bReportedFontError)
	{
		m_bReportedFontError = true;

		LeoHelpers::OpusStringLoader sl(&m_opusUtils);

		TCHAR *szErrMsg = LeoHelpers::StringAllocAndFormat(sl.Get(STR_NFO_FONT_ERROR), szFontName);
	
		if (szErrMsg != NULL)
		{
			MessageBox(m_hWnd, szErrMsg, sl.Get(STR_NFO_PLUGIN_DESCRIPTION), MB_OK|MB_ICONERROR);
			
			delete [] szErrMsg;
		}
	}
}
