#pragma once

class CNfoViewerWindow
{
public:
	struct NVC_CreateParams
	{
		LPRECT	lpRc;
		DWORD	dwFlags;
	};

protected:
	HWND m_hWnd;
	HINSTANCE m_hInstance;
	DWORD m_dwOpusViewerFlags;

	HWND m_hWndEdit;
	bool m_bReportedFontError;
	HFONT m_hFont;
	//int m_nFontSize;
	COLORREF m_FontColor;
	COLORREF m_BackgroundColor;
	HBRUSH m_BackgroundBrush;
#ifndef UNICODE
	LOGFONT *m_pCurrentLogFont;
	std::map< LONG, LOGFONT > m_mapFontWidthToLogFont;
#else
	LOGFONT m_currentLogFont;
#endif

	TCHAR *m_szName; // Name of document we're showing. Sometimes, but not always, a file path.
	DWORD m_dwNumBytes; // Size of the document we loaded (for information display only).
	DWORD m_dwMaxLineWidth;
	DWORD m_dwCapabilities;
	HANDLE m_hAbortEvent;

	DOpusPluginHelperUtil m_opusUtils;

public:
	CNfoViewerWindow(HMODULE hDllModule, HWND hWnd, LPRECT lpRc, DWORD dwFlags);
	~CNfoViewerWindow(void);

	static ATOM RegisterWindowClass(HMODULE hDllModule);
	static BOOL UnregisterWindowClass(HMODULE hDllModule);

	static HWND CreateViewerWindow(HMODULE hDllModule, HWND hWndParent, LPRECT lpRc, DWORD dwFlags);

protected:
	void reset(bool bForgetAbortEvent);

	inline static const TCHAR *getWindowClassName() { return(_T("dopusviewerplugin.nfo.viewer")); }
	static LRESULT WINAPI wndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	static LRESULT WINAPI editSubProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

	void onCreate(HWND hWnd);
	void onSize(HWND hWnd, bool bMinimized);
	LRESULT onDVPLoad(HWND hWnd, LPTSTR szFilename);
	LRESULT onDVPLoadStream(HWND hWnd, LPSTREAM pStream, LPTSTR szFilename);
	LRESULT load(HWND hWnd, const TCHAR *szName, HLOCAL hLocal);
	void onDVPClear(HWND hWnd);
	LRESULT onPaint(HWND hWnd);
	void onDVPReinitialize(HWND hWnd);
	void onDVPGetImageInfo(HWND hWnd, VIEWERPLUGINFILEINFO *pInfo);
	void onDVPPrint(HWND hWnd);
	void onDVPSelectAll(HWND hWnd);
	void onDVPCopySelection(HWND hWnd);
	void updateCapabilities(HWND hWnd);
	void updateFont(HWND hWnd);

#ifndef UNICODE
	static int CALLBACK fontEnumCallback(CONST LOGFONT *pLogFont, CONST TEXTMETRIC *pTextMetric, DWORD FontType, LPARAM lParam);
#endif
	void reportFontError(const TCHAR *szFontName);


	void convertEOLs(HLOCAL &hLocal);
	void convertToUnicode(HLOCAL &hLocal);
};
