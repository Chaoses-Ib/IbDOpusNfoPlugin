#pragma once

// Macros suck but I need constant expressions for switch statements. :-(

#define MAKE64BITVERSIONNUMBER(a,b,c,d) ((((DWORD64)(a))<<48) + (((DWORD64)(b))<<32) + (((DWORD64)(c))<<16) + (((DWORD64)(d))<<00))
#define MAKE_DWORD_NAME(a,b,c,d) (((DWORD)(unsigned char)(d)<<24)+((DWORD)(unsigned char)(c)<<16)+((DWORD)(unsigned char)(b)<<8)+((DWORD)(unsigned char)(a)))

#ifdef LEOHELP_DIALOG_DYNAMIC
#ifndef LEOHELP_WINDOW
#define LEOHELP_WINDOW
#endif
#endif

struct LeoHelpers
{
	// Returns false on failure. Use GetLastError().
	// hKeyParent -- hKey of parent under which szKeyPath should be created. Usually a hive like HKEY_CURRENT_USER
	// szKeyPath -- Sub key (can be a path) under hKeyHiveOrParent from which szValueName is read. Can be "" or NULL.
	// szValueName -- Name of value to query. Can be "" or NULL to read the unnamed default value.
	// sam -- Pass zero normally, or KEY_WOW64_32KEY or KEY_WOW64_64KEY if you need to access the alternative area. (KEY_QUERY_VALUE is specified automatically.)
	// pdwType -- Pointer to DWORD to receive type of data. NULL if unwanted.
	// plpData -- POINTER TO POINTER to receive data buffer. NULL if data unwanted.
	// pcbData -- Pointer to DWORD to receive size of data buffer. NULL if data unwanted.
	// You do NOT allocate a buffer, this function does it for you (like the stupid Registry API should have done, at least as an option).
	// If the function succeeds and plpData was not NULL you must delete[] *plpData when finished with it.
	// Like RegQueryValueEx, you may supply NULL plpData and non-NULL pcbData to get the required
	// buffer size (except for HKEY_PERFORMANCE_DATA), but again the function always allocates a buffer for you.
	// Although untested, reading HKEY_PERFORMANCE_DATA should work.
	// Pads whatever data comes out of the registry with 8 null bytes to avoid problems with badly stored data.
	static bool LeetRegQueryValue(HKEY hKeyParent, const TCHAR *szKeyPath, const TCHAR *szValueName, REGSAM sam,
						DWORD *pdwType, void **plpData, DWORD *pcbData);

	// Returns false on failure. Use GetLastError().
	// Uses LeetRegQueryValue() to get you a DWORD without having to worry about buffers, types and so on.
	static bool LeetRegQueryDWORDValue(HKEY hKeyParent, const TCHAR *szKeyPath, const TCHAR *szValueName, REGSAM sam, DWORD *pdwRes);

	// Returns false on failure. Use GetLastError().
	// Uses LeetRegQueryValue() to get you an int without having to worry about buffers, types and so on.
	static bool LeetRegQueryIntValue(HKEY hKeyParent, const TCHAR *szKeyPath, const TCHAR *szValueName, REGSAM sam, int *piRes);

	// Returns false on failure. Use GetLastError().
	// Uses LeetRegQueryValue() to get you a bool without having to worry about buffers, types and so on.
	static bool LeetRegQueryBoolValue(HKEY hKeyParent, const TCHAR *szKeyPath, const TCHAR *szValueName, REGSAM sam, bool *pBool);

	// Returns false on failure. Use GetLastError().
	// Uses LeetRegQueryValue() to get you a TCHAR string without having to worry about buffers, types and so on.
	static bool LeetRegQueryStringValue(HKEY hKeyParent, const TCHAR *szKeyPath, const TCHAR *szValueName, REGSAM sam, std::basic_string<TCHAR> *pString);

	// Returns false on failure. Use GetLastError().
	// Uses LeetRegQueryValue() to get you a vector of TCHAR strings without having to worry about buffers, types and so on.
	static bool LeetRegQueryMultiStringValue(HKEY hKeyParent, const TCHAR *szKeyPath, const TCHAR *szValueName, REGSAM sam, std::vector< std::basic_string<TCHAR> > *pVecStrings);

	// Returns false on failure. Use GetLastError().
	// Returns vector of child key names.
	// If szPrefixFilter is non-NULL then only key names starting with it will be included, and it will be removed from the start of them.
	static bool LeetRegEnumKey(HKEY hKeyParent, const TCHAR *szKeyPath, const TCHAR *szPrefixFilter, REGSAM sam, std::vector< std::basic_string<TCHAR> > *pVecStrings);

	// Returns false on failure. Use GetLastError().
	// Returns vector of value names.
	// KEY_QUERY_VALUE will automatically be specified as part of the access mask.
	// The sam argument allow you to pass KEY_WOW64_32KEY or KEY_WOW64_64KEY; pass zero otherwise.
	static bool LeetRegEnumValue(HKEY hKeyParent, const TCHAR *szKeyPath, REGSAM sam, std::vector< std::basic_string<TCHAR> > *pVecStrings);

	// Returns false on failure. Use GetLastError().
	// hKeyParent -- hKey of parent under which szKeyPath should be created. Usually a hive like HKEY_CURRENT_USER
	// szKeyPath -- Sub key (can be a path) under hKeyHiveOrParent in which szValueName is set. Can be "" or NULL.
	// szValueName -- Name of value to set. Can be "" or NULL to set the unnamed, default value.
	// dwType -- Type of data to set. (Same as dwType argument of RegSetValueEx())
	// lpData -- Pointer to data to set. (6/Feb/2001: Can be NULL if you just want to create the key but set no values.)
	// cbData -- Size of data pointed to by lpData, including null terminator(s) (Same as cbData arg of RegSetValueEx())
	static bool LeetRegSetValue(HKEY hKeyParent, const TCHAR *szKeyPath, const TCHAR *szValueName,
							DWORD dwType, const void *lpData, DWORD cbData);

	// Returns false on failure. Use GetLastError().
	// Uses LeetRegSetValue() to set a DWORD without having to worry about buffers, types and so on.
	static bool LeetRegSetDWORDValue(HKEY hKeyParent, const TCHAR *szKeyPath, const TCHAR *szValueName, DWORD dwData);

	// Returns false on failure. Use GetLastError().
	// Uses LeetRegSetValue() to set an int without having to worry about buffers, types and so on.
	static bool LeetRegSetIntValue(HKEY hKeyParent, const TCHAR *szKeyPath, const TCHAR *szValueName, int iData);

	// Returns false on failure. Use GetLastError().
	// Uses LeetRegSetValue() to set a bool without having to worry about buffers, types and so on.
	static bool LeetRegSetBoolValue(HKEY hKeyParent, const TCHAR *szKeyPath, const TCHAR *szValueName, bool bData);

	// Returns false on failure. Use GetLastError().
	// Uses LeetRegSetValue() to set a String without having to worry about buffers, types and so on.
	static bool LeetRegSetStringValue(HKEY hKeyParent, const TCHAR *szKeyPath, const TCHAR *szValueName, const TCHAR *szData);

	// Returns false on failure. Use GetLastError().
	static bool LeetRegSetMultiStringValue(HKEY hKeyParent, const TCHAR *szKeyPath, const TCHAR *szValueName, const std::vector< std::basic_string<TCHAR> > *pVecStrings);

	// Returns false on failure. Use GetLastError().
	static bool LeetRegCreateKey(HKEY hKeyParent, const TCHAR *szKeyPath);

	// Returns false on failure. Use GetLastError().
	// Will NOT delete a key that has sub-keys. (You can use SHDeleteKey for that, though.)
	static bool LeetRegDeleteKey(HKEY hKeyParent, const TCHAR *szKeyPath, const TCHAR *szKeyToDelete);

	// Returns false on failure. Use GetLastError().
	static bool LeetRegDeleteValue(HKEY hKeyParent, const TCHAR *szKeyPath, const TCHAR *szValueName);

	static const _locale_t LocaleC;

	static inline wchar_t WideCharUpper(wchar_t c) { return _towupper_l(c, LocaleC); }
	static inline wchar_t WideCharLower(wchar_t c) { return _towlower_l(c, LocaleC); }

	static inline int WideStrCmpUpper(const wchar_t *x, const wchar_t *y)
	{
		wchar_t ux, uy;
		while(*x && *y)
		{
			ux = WideCharUpper(*x++);
			uy = WideCharUpper(*y++);
			if (ux < uy) return -1;
			if (uy < ux) return 1;
		}
		if (*y) return -1;
		if (*x) return 1;
		return 0;
	}
	static int WideStrCmpUpper(const std::wstring &x, const std::wstring &y) { return WideStrCmpUpper(x.c_str(), y.c_str()); }
	static inline int WideStrCmpUpperLen(const wchar_t *x, const wchar_t *y, size_t n)
	{
		wchar_t ux, uy;
		while(*x && *y && n)
		{
			ux = WideCharUpper(*x++);
			uy = WideCharUpper(*y++);
			if (ux < uy) return -1;
			if (uy < ux) return 1;
			--n;
		}
		if (n == 0) return 0;
		if (*y) return -1;
		if (*x) return 1;
		return 0;
	}
	static int WideStrCmpUpperLen(const std::wstring &x, const std::wstring &y, size_t n) { return WideStrCmpUpperLen(x.c_str(), y.c_str(), n); }

	static inline char CharToUpper(unsigned char c)
	{
		if (__isascii(c) && _islower_l(c, LocaleC))
		{
			return _toupper_l(c, LocaleC);
		}
		return c;
	}

	static inline char CharToLower(unsigned char c)
	{
		if (__isascii(c) && _isupper_l(c, LocaleC))
		{
			return _tolower_l(c, LocaleC);
		}
		return c;
	}

	static int __cdecl WideStrBSearchCmp(void *pvContext, const void *pvLeft, const void *pvRight)
	{
		return wcscmp(*reinterpret_cast<const wchar_t * const *>(pvLeft), *reinterpret_cast<const wchar_t * const *>(pvRight));
	}

	static int __cdecl WideStrBSearchCmpUpper(void *pvContext, const void *pvLeft, const void *pvRight)
	{
		return WideStrCmpUpper(reinterpret_cast<const wchar_t *>(pvLeft), reinterpret_cast<const wchar_t *>(pvRight));
	}

	struct LessUpperCase
	{
		bool operator()(const std::wstring &x, const std::wstring &y) const
		{ return 0 > WideStrCmpUpper(x.c_str(), y.c_str()); }
	};

	static inline void ToUpper(std::string *pStrOutput, const char *szInput)
	{
		pStrOutput->clear();
		while(*szInput)
		{
			pStrOutput->push_back(LeoHelpers::CharToUpper(*szInput++));
		}
	}

	static inline void ToUpper(std::wstring *pstrOutput, const wchar_t *szInput)
	{
		size_t len = wcslen(szInput);

		pstrOutput->clear();
		pstrOutput->reserve(len);

		std::transform(szInput, szInput + len, std::back_inserter< std::wstring >(*pstrOutput), WideCharUpper);
	}

	static inline void ToLower(std::wstring *pstrOutput, const wchar_t *szInput)
	{
		size_t len = wcslen(szInput);

		pstrOutput->clear();
		pstrOutput->reserve(len);

		std::transform(szInput, szInput + len, std::back_inserter< std::wstring >(*pstrOutput), WideCharLower);
	}

	static inline void ToUpper(std::wstring *pstrOutput, const std::wstring &strInput)
	{
		pstrOutput->clear();
		pstrOutput->reserve(strInput.length());

		std::transform(strInput.begin(), strInput.end(), std::back_inserter< std::wstring >(*pstrOutput), WideCharUpper);
	}

	static inline void ToLower(std::wstring *pstrOutput, const std::wstring &strInput)
	{
		pstrOutput->clear();
		pstrOutput->reserve(strInput.length());

		std::transform(strInput.begin(), strInput.end(), std::back_inserter< std::wstring >(*pstrOutput), WideCharLower);
	}

	static inline void ToUpper(std::wstring *pstr)
	{
		std::transform(pstr->begin(), pstr->end(), pstr->begin(), WideCharUpper);
	}

	static inline void ToLower(std::wstring *pstr)
	{
		std::transform(pstr->begin(), pstr->end(), pstr->begin(), WideCharLower);
	}
/*
	static inline void ToUpper(wchar_t *sz)
	{
		while(*sz)
		{
			*sz = LeoHelpers::WideCharUpper(*sz);
			++sz;
		}
	}

	static inline void ToLower(wchar_t *sz)
	{
		while(*sz)
		{
			*sz = LeoHelpers::WideCharLower(*sz);
			++sz;
		}
	}
*/
	// Initialisation lists are about the only place it makes sense to use this.
	static inline const std::wstring ToUpperTransientString(const wchar_t *szIn)
	{
		std::wstring strOut;
		LeoHelpers::ToUpper(&strOut, szIn);
		return strOut;
	}

	// Initialisation lists are about the only place it makes sense to use this.
	static inline const std::wstring ToLowerTransientString(const wchar_t *szIn)
	{
		std::wstring strOut;
		LeoHelpers::ToLower(&strOut, szIn);
		return strOut;
	}

	static inline bool IsPrefix(const TCHAR *string, const TCHAR *prefix)
	{
		return 0 == _tcsncmp(string, prefix, _tcslen(prefix));
	}

	static inline bool IsPrefix(const std::wstring &string, const std::wstring &prefix)
	{
		return 0 == _tcsncmp(string.c_str(), prefix.c_str(), prefix.length());
	}

	static inline bool IsPrefixUpper(const TCHAR *string, const TCHAR *prefix)
	{
		return 0 == WideStrCmpUpperLen(string, prefix, _tcslen(prefix));
	}

	static inline bool IsPrefixUpper(const std::wstring &string, const std::wstring &prefix)
	{
		return 0 == WideStrCmpUpperLen(string, prefix, prefix.length());
	}

	static inline bool IsSuffix(const TCHAR *string, const TCHAR *suffix)
	{
		size_t lenString = _tcslen(string);
		size_t lenSuffix = _tcslen(suffix);

		if (lenString < lenSuffix) return false;

		return 0 == _tcscmp(string + (lenString - lenSuffix), suffix);
	}

	static inline bool IsSuffix(const std::wstring &string, const std::wstring &suffix)
	{
		size_t lenString = string.length();
		size_t lenSuffix = suffix.length();

		if (lenString < lenSuffix) return false;

		return 0 == _tcscmp(string.c_str() + (lenString - lenSuffix), suffix.c_str());
	}

	static inline bool IsSuffixUpper(const TCHAR *string, const TCHAR *suffix)
	{
		size_t lenString = _tcslen(string);
		size_t lenSuffix = _tcslen(suffix);

		if (lenString < lenSuffix) return false;

		return 0 == WideStrCmpUpper(string + (lenString - lenSuffix), suffix);
	}

	static inline bool IsSuffixUpper(const std::wstring &string, const std::wstring &suffix)
	{
		size_t lenString = string.length();
		size_t lenSuffix = suffix.length();

		if (lenString < lenSuffix) return false;

		return 0 == WideStrCmpUpper(string.c_str() + (lenString - lenSuffix), suffix.c_str());
	}

	static std::wstring StringTrim(const std::wstring &strIn, const wchar_t *szTrimChars, bool bTrimLeft, bool bTrimRight);

	static inline void StringCopy(TCHAR *dest, const TCHAR *source, size_t destSizeChars)
	{
		if (destSizeChars > 0)
		{
			_tcsncpy_s(dest, destSizeChars, source, _TRUNCATE);
		}
	}

	static inline void StringCopy(TCHAR *dest, size_t destSizeChars, const TCHAR *source)
	{
		if (destSizeChars > 0)
		{
			_tcsncpy_s(dest, destSizeChars, source, _TRUNCATE);
		}
	}

	// Unlike strncat, destSizeChars should be the total size of dest, not the remaining size.
	static inline void StringAppend(TCHAR *dest, const TCHAR *source, size_t destSizeChars)
	{
		const size_t destLen = _tcslen(dest);
		dest += destLen;
		StringCopy(dest, source, destSizeChars - destLen);
	}

	// Returns the index of the null terminator. -1 on error.
	static int StringFormat(TCHAR *buffer, size_t bufferSizeChars, const TCHAR *format, ...)
	{
		va_list args;
		va_start(args, format);
		int result = VStringFormat(buffer, bufferSizeChars, format, args);
		va_end(args);

		return(result);
	}

	// Returns the index of the null terminator. -1 on error.
	static int VStringFormat(TCHAR *buffer, size_t bufferSizeChars, const TCHAR *format, va_list args)
	{
		int result = 0;

		if (bufferSizeChars > 0)
		{
			result = _vsntprintf_s(buffer, bufferSizeChars, _TRUNCATE, format, args);
		}

		return(result);
	}

	// delete[] the result.
	// Returns NULL on failure.
	static TCHAR *StringAllocAndFormat(const TCHAR *format, ...)
	{
		va_list args;
		va_start(args, format);
		TCHAR *result = VStringAllocAndFormat(format, args);
		va_end(args);

		return(result);
	}

	// delete[] the result.
	// Returns NULL on failure.
	static TCHAR *VStringAllocAndFormat(const TCHAR *format, va_list args);

	// These are less efficient, since they copy the string an extra time, but safer and easier.
	// Returns boolean success.
	static bool StringFormatToString(std::basic_string< TCHAR > *pResult, const TCHAR *format, ...);
	// Returns boolean success.
	static bool VStringFormatToString(std::basic_string< TCHAR > *pResult, const TCHAR *format, va_list args);

	// delete[] the result.
	// Returns NULL on failure.
	// *pdwLengthChars will be set to the size of the returned buffer, including all the null terminators.
	static TCHAR *MultiStringFromVector(const std::vector< std::basic_string<TCHAR> > *pVecStrings, size_t *pLengthChars);

	// szInput -- String to parse.
	// delimiters -- Characters which delimit the string.
	// terminators -- Characters which end the string. Can be NULL.
	// results -- A string list to put the results into. Existing contents will be erased. Give NULL if you don't want any results.
	// bSkipEmpty -- Set true to skip any empty strings. (e.g. someone typed "blah,,blah")
	// bQuotes -- Set to true to handle double-quotes (") around items. A quoted empty string is NOT skipped even if bSkipEmpty is set. Terminators in quotes are ignored and included in the tokenized strings.
	// bCSVQuotes -- (Ignored if bQuotes false) -- Inside quoted strings, two double-quotes next to each other will be output as one double-quote and not signify the end of the string. A string like "Blah""blah" will become blah"blah.
	// Leading and trailing spaces are always stripped.
	// Returns false iff nothing could be extracted (e.g. szInput was NULL)
	static bool Tokenize(const TCHAR *szInput, const TCHAR *delimiters, const TCHAR *terminators, std::list< std::basic_string< TCHAR > > *pResults, bool bSkipEmpty, bool bQuotes, bool bCSVQuotes);

	// pp -- Pointer to buffer to tokenize. Will be advanced to the next position string. Set to NULL after the last token is extracted.
	// delimiters -- Characters which delimit the string.
	// terminators -- Characters which end the string. Can be NULL.
	// bSkipEmpty -- Set true to skip any empty strings. (e.g. someone typed "blah,,blah")
	// bQuotes -- Set to true to handle double-quotes (") around items. A quoted empty string is NOT skipped even if bSkipEmpty is set.
	// bCSVQuotes -- (Ignored if bQuotes false) -- Inside quoted strings, two double-quotes next to each other will be output as one double-quote and not signify the end of the string. A string like "Blah""blah" will become blah"blah.
	// buf -- Buffer to extract in to. Set to NULL not to extract.
	// bufSize -- Size of buf if it is non-NULL. Will truncate string if buf too small.
	// Leading and trailing spaces are always stripped.
	// Returns false iff nothing could be extracted (i.e. *pp was NULL)
	// pp and the contents of buf are only valid after calling if true is returns.
	// Lame function should allocate us a suitably sized buffer for each string... :-)
	// But hey, this isn't a public OS function and you can just give a buffer the size of the input...
	static bool Tokenize(const TCHAR **pp, const TCHAR *delimiters, const TCHAR *terminators, TCHAR *buf, size_t bufSize, bool bSkipEmpty, bool bQuotes, bool bCSVQuotes);

	// Returns NULL on failure (use GetLastError())
	// delete[] the result when finished with it.
	static inline WCHAR *LeetMBtoWC(const char *mbString)
	{
		WCHAR *result = NULL;
		LeoHelpers::MBtoWC(&result, mbString, -1, CP_ACP);
		return(result);
	}

	// Returns NULL on failure (use GetLastError())
	// delete[] the result when finished with it.
	static inline char *LeetWCtoMB(const WCHAR *wcString)
	{
		char *result = NULL;
		LeoHelpers::WCtoMB(&result, wcString, -1, CP_ACP);
		return(result);
	}

	static bool LeetMBtoWC(std::wstring *pOutString, const char *mbString)
	{
		 WCHAR *t = LeoHelpers::LeetMBtoWC(mbString);

		 if (!t)
		 {
			 pOutString->clear();
			 return false;
		 }

		 *pOutString = t;
		 delete[] t;

		 return true;
	}

	static bool LeetWCtoMB(std::string *pOutString, const WCHAR *wcString)
	{
		 char *t = LeoHelpers::LeetWCtoMB(wcString);

		 if (!t)
		 {
			 pOutString->clear();
			 return false;
		 }

		 *pOutString = t;
		 delete[] t;

		 return true;
	}

	// Returns 0 on failure (use GetLastError()), else the number of wide characters written.
	// delete [] *pwcStringResult when finished with it.
	// numInputBytes -1 for null-terminated string.
	static int MBtoWC(WCHAR **pwcStringResult, const char *mbStringInput, int numInputBytes, DWORD dwCodePage);

	// Returns 0 on failure (use GetLastError()), else the number of characters written.
	// delete [] *pmbStringResult when finished with it.
	// numInputWords -1 for null-terminated string.
	static int WCtoMB(char **pmbStringResult, const WCHAR *wcStringInput, int numInputWords, DWORD dwCodePage);

	static inline int MBtoWC(std::wstring *pstrStringResult, const char *mbStringInput, int numInputBytes, DWORD dwCodePage)
	{
		wchar_t *sz = NULL;

		int iResult = LeoHelpers::MBtoWC(&sz, mbStringInput, numInputBytes, dwCodePage);

		if (0 >= iResult)
		{
			pstrStringResult->clear();
		}
		else
		{
			*pstrStringResult = sz;
		}

		delete[] sz;

		return iResult;
	}

	static inline int WCtoMB(std::string *pstrStringResult, const WCHAR *wcStringInput, int numInputWords, DWORD dwCodePage)
	{
		char *sz = NULL;

		int iResult = LeoHelpers::WCtoMB(&sz, wcStringInput, numInputWords, dwCodePage);

		if (0 >= iResult)
		{
			pstrStringResult->clear();
		}
		else
		{
			*pstrStringResult = sz;
		}

		delete[] sz;

		return iResult;
	}

	static inline bool IsUtf16LeftSurrogate(wchar_t c)
	{
		// The (wchar_t) casts might help if wchar_t is signed. :)
		return (c >= ((wchar_t)0xD800) && c <= ((wchar_t)0xDBFF));
	}

	// Returns the new null position.
	static size_t TruncateString(wchar_t *szBuffer, size_t idxNull, size_t cchSpaceAfterNull, bool bTruncateWithElipsis);

	static bool MBtoWCInPlaceTruncate(const char *szSource, int cbInputBytes, wchar_t *szDestBuffer, int cchBufferSize, DWORD dwCodePage, bool bTruncateWithElipsis, bool *pbWasTruncated);

	static inline int MulDivRoundDown(int nNumber, int nNumerator, int nDenominator)
	{
		assert(nDenominator != 0);

		if (nDenominator == 0)
		{
			return -1; // Do what MulDiv does.
		}

		__int64 x = nNumber;
		x *= nNumerator;
		x /= nDenominator;

		assert(x <= INT_MAX && x >= INT_MIN);

		if (x > INT_MAX || x < INT_MIN)
		{
			return -1;
		}

		return static_cast<int>(x);
	}

#ifdef DOPUS_PLUGIN_HELPER
	
	class OpusStringLoader
	{
	private:
		DOpusPluginHelperUtil *m_pHelper;
		// Map is to pointers rather than strings so that it doesn't matter if the map nodes move around in memory.
		std::map< UINT, std::wstring * > m_mapStrings;

		// Disallow copy-constructor and assignment operator:
		OpusStringLoader(const OpusStringLoader &rhs);
		OpusStringLoader &operator=(const OpusStringLoader &rhs);
	public:
		OpusStringLoader(DOpusPluginHelperUtil *pHelper) : m_pHelper(pHelper)
		{
		}

		~OpusStringLoader()
		{
			for (std::map< UINT, std::wstring * >::iterator miter = m_mapStrings.begin(); miter != m_mapStrings.end(); ++miter)
			{
				std::wstring *pwstr = miter->second;
				delete pwstr;
			}
		}

		const wchar_t *Get(UINT id)
		{
			std::wstring *&pwstr = m_mapStrings[id];

			if (pwstr == NULL)
			{
				wchar_t szBuffer[1024];
				wchar_t *pb = szBuffer; // Just to make the compiler happy about __w64 pointer crap.

				SIZE_T res = m_pHelper->GetString(id, &pb, _countof(szBuffer));

				if (res == 0)
				{
					_CrtDbgBreak();
					return L"<missing string>";
				}

				if (res < _countof(szBuffer))
				{
					pwstr = new std::wstring(szBuffer);
				}
				else
				{
					wchar_t *szLocalAlloc = NULL;
					res = m_pHelper->GetString(id, &szLocalAlloc, 0);

					if (res == 0)
					{
						_CrtDbgBreak();
						return L"<missing string>";
					}

					pwstr = new std::wstring(szLocalAlloc);
					::LocalFree(szLocalAlloc);
					szLocalAlloc = NULL;

				}
			}

			return pwstr->c_str();
		}

		DOpusPluginHelperUtil *GetOpusPluginHelperUtil() const
		{
			return m_pHelper;
		}
	};

#endif // DOPUS_PLUGIN_HELPER

	class StringLoader
	{
	private:
		HINSTANCE m_hInstance;
		// Map is to pointers rather than strings so that it doesn't matter if the map nodes move around in memory.
		std::map< UINT, std::wstring * > m_mapStrings;
	public:
		StringLoader(HINSTANCE hInstance) : m_hInstance(hInstance)
		{
		}

		~StringLoader()
		{
			for (std::map< UINT, std::wstring * >::iterator miter = m_mapStrings.begin(); miter != m_mapStrings.end(); ++miter)
			{
				std::wstring *pwstr = miter->second;
				delete pwstr;
			}
		}

		const wchar_t *Get(UINT id)
		{
			std::wstring *&pwstr = m_mapStrings[id];

			if (pwstr == NULL)
			{
				wchar_t szBuffer[1024]; // Assume 1KB is enough for all our strings.

				int res = ::LoadString(m_hInstance, id, szBuffer, _countof(szBuffer)-1);

				if (res <= 0 || res >= _countof(szBuffer))
				{
					_CrtDbgBreak();
					return L"<missing string>";
				}

				szBuffer[res] = L'\0'; // Ensure the buffer is null-terminated. MSDN is unclear about what happens if the string is the same as the buffer length (i.e. no room for the null terminator) which is why we subtracted 1 from the size above.

				pwstr = new std::wstring(szBuffer);
			}

			return pwstr->c_str();
		}
	};

	static inline bool IsAscii(const wchar_t *sz)
	{
		while(*sz)
		{
			if (!iswascii(*sz))
			{
				return false;
			}
			++sz;
		}

		return true;
	}

	static inline bool IsAscii(unsigned char c) // This exists just to cast to unsigned char for the dangerously stupid CRT function.
	{
		return isascii(c) ? true : false; // Can't call ::isascii as isascii is a macro.
	}

	static inline bool IsDigit(unsigned char c) // This exists just to cast to unsigned char for the dangerously stupid CRT function.
	{
		return ::isdigit(c) ? true : false;
	}

	static bool GetShortFilePath(std::wstring *pstrResult, const wchar_t *szInput);

	static bool UrlFromPath(std::wstring *pstrResult, const wchar_t *szPath);

	static bool UrlFromPathShortIfPossible(std::wstring *pstrResult, const wchar_t *szPath, bool bOnlyTryShortIfNonAscii);

//	static void EncodeUri(std::wstring *pstrResult, const wchar_t *szInput, bool bAppend, bool bConvertColon, bool bProperUTF8Method);

	// A path with no slashes or the empty string returns an empty string.
	static std::basic_string< TCHAR > *GetParentPathString(std::basic_string< TCHAR > *pstrResult, const TCHAR *path);

	static std::basic_string< TCHAR > *AppendPathString(std::basic_string< TCHAR > *pstrPath, const TCHAR *szAppendage);

	static const TCHAR *GetLastPathPart(const TCHAR *szIn);

	static const TCHAR *GetExtensionPart(bool bIncludeDot, const TCHAR *szIn);

	static bool IsLegalPathChar(const wchar_t &c, bool bAllowPathSeps, bool bAllowDots)
	{
		switch(c)
		{
		default:
			return true;
		case L'\\':
		case L'/':
			return bAllowPathSeps;
		case L'.':
			return bAllowDots;
		case L'*':
		case L'?':
		case L'"':
		case L'<':
		case L'>':
		case L'|':
			return false;
		}
	}

	static std::wstring GenerateLegalFileName(const wchar_t * szInput, bool bAllowPathSeps, bool bAllowDots);
	static void GenerateLegalFileName(std::wstring *pstr, bool bAllowPathSeps, bool bAllowDots, std::wstring::size_type startPos);

	static std::wstring GenerateSafeName(const wchar_t *szInput);

	static bool GetTempPath(std::basic_string< TCHAR > *pstrTempPath);

	// If bDontReallyOpen is true then the function won't really create a file and always returns INVALID_HANDLE_VALUE, but will still set *pstrFilenameOut to a suitable filename.
	// On failure *pstrFilenameOut will be left empty so you can check that when bDontReallyOpen is true.
	// Of course, if you don't create the file then there is nothing to stop someone else (e.g. another caller of this function)
	// creating it unexpectedly, so the name should only be considered as an example (e.g. to see if we'll generate a path with any non-ASCII chars in it) and never really used.
	static HANDLE OpenTempFileNamePreserveExtension(const TCHAR *szTempPath, const TCHAR *szPrefix, const TCHAR *szFilenameIn, std::basic_string< TCHAR > *pstrFilenameOut, bool bDontReallyOpen);

	static bool CopyFile(const TCHAR *szSource, const TCHAR *szDest, HANDLE hAbortEvent, bool bClearReadOnlyAttrib, bool bFailIfExists);

	// Intended for reading from pipes.
	static bool ReadFileUntilDone(HANDLE hFile, LPVOID lpBuffer, DWORD dwNumberOfBytesToRead);

	// Intended for writing to pipes.
	static bool WriteFileUntilDone(HANDLE hFile, LPVOID lpBuffer, DWORD dwNumberOfBytesToWrite);

	static BOOL GetDlgItemRect(HWND hwndDlg, int nIDDlgItem, RECT *pRect);

	static BOOL OffsetDlgItem(HWND hwndDlg, int nIDDlgItem, LONG dx, LONG dy, LONG dw, LONG dh, BOOL bRepaint);

	static BOOL EnableDlgItem(HWND hwndDlg, int nIDDlgItem, BOOL bEnable);
	static BOOL ShowDlgItem(HWND hwndDlg, int nIDDlgItem, BOOL bShow);

	static BOOL GetDlgItemText(HWND hwndDlg, int nIDDlgItem, std::wstring *pStrResult);
	static BOOL GetListBoxItemText(HWND hwndListBox, LRESULT lrItemIndex, std::wstring *pStrResult);
	static BOOL SetupComboItem(HWND hWndParent, int iDlgItem, bool bClear, DWORD dwValue, DWORD dwSelectValue, const wchar_t *szTitle);
	static BOOL GetComboItem(HWND hWndParent, int iDlgItem, DWORD *pdwValue);
	static BOOL GetComboItemIntCast(HWND hWndParent, int iDlgItem, int *piValue);
	static BOOL SetupCheckBoxItem(HWND hWndParent, int iDlgItem, bool bCheck);
	static BOOL GetCheckBoxItem(HWND hWndParent, int iDlgItem, bool *pbValue);
	static BOOL SetupSpinItem(HWND hWndParent, int iDlgItem, int iMax, int iMin, int iValue);
	static BOOL SetSpinItem(HWND hWndParent, int iDlgItem, int iValue);
	static BOOL GetSpinItem(HWND hWndParent, int iDlgItem, int *piValue);
	static BOOL SetupEditItem(HWND hWndParent, int iDlgItem, const wchar_t *szText);
	static BOOL GetEditItem(HWND hWndParent, int iDlgItem, std::wstring *pStrResult);

//	static BOOL SubclassDlgItemToUserData(HWND hwndDlg, int nIDDlgItem, WNDPROC wndProc);

//	static BOOL GetDlgItemLong(HWND hwndDlg, int nIDDlgItem, int nIndex, LONG *plResult);
//	static BOOL SetDlgItemLong(HWND hwndDlg, int nIDDlgItem, int nIndex, LONG lValue);
	static BOOL GetDlgItemLongPtr(HWND hwndDlg, int nIDDlgItem, int nIndex, LONG_PTR *plResult);
	static BOOL SetDlgItemLongPtr(HWND hwndDlg, int nIDDlgItem, int nIndex, LONG_PTR lValue);
//	static BOOL ModifyDlgItemStyle(HWND hwndDlg, int nIDDlgItem, bool bStyleEx, LONG_PTR lAdd, LONG_PTR lRemove);

	static BOOL InvalidateDlgItem(HWND hwndDlg, int nIDDlgItem);

	static BOOL EnableDlgControl(HWND hwndControl, BOOL bEnable);

	static BOOL ScreenToClientRect(HWND hwnd, RECT *pRect);
	static BOOL ClientToScreenRect(HWND hwnd, RECT *pRect);

	static inline void MoveWindowRect(HWND hwnd, RECT *pRect, BOOL bRepaint)
	{
		::MoveWindow(hwnd, pRect->left, pRect->top, pRect->right-pRect->left, pRect->bottom-pRect->top, bRepaint);
	}

	static inline void AboveRect(RECT *pRectToMove, const RECT *pRectRelative, LONG lOffset)
	{
		LONG lWidth  = pRectToMove->right  - pRectToMove->left;
		LONG lHeight = pRectToMove->bottom - pRectToMove->top;

		pRectToMove->left   = pRectRelative->left;
		pRectToMove->right  = pRectToMove->left     + lWidth;
		pRectToMove->bottom = pRectRelative->top    - lOffset;
		pRectToMove->top    = pRectToMove->bottom   - lHeight;
	}

	static inline void BelowRect(RECT *pRectToMove, const RECT *pRectRelative, LONG lOffset)
	{
		LONG lWidth  = pRectToMove->right  - pRectToMove->left;
		LONG lHeight = pRectToMove->bottom - pRectToMove->top;

		pRectToMove->left   = pRectRelative->left;
		pRectToMove->right  = pRectToMove->left     + lWidth;
		pRectToMove->top    = pRectRelative->bottom + lOffset;
		pRectToMove->bottom = pRectToMove->top      + lHeight;
	}

	static inline void BelowRightRect(RECT *pRectToMove, const RECT *pRectRelative, LONG lOffset)
	{
		LONG lWidth  = pRectToMove->right  - pRectToMove->left;
		LONG lHeight = pRectToMove->bottom - pRectToMove->top;

		pRectToMove->right  = pRectRelative->right;
		pRectToMove->left   = pRectToMove->right    - lWidth;
		pRectToMove->top    = pRectRelative->bottom + lOffset;
		pRectToMove->bottom = pRectToMove->top      + lHeight;
	}

	static inline void RightRect(RECT *pRectToMove, const RECT *pRectRelative, LONG lOffset)
	{
		LONG lWidth  = pRectToMove->right  - pRectToMove->left;
		LONG lHeight = pRectToMove->bottom - pRectToMove->top;

		pRectToMove->left   = pRectRelative->right  + lOffset;
		pRectToMove->right  = pRectToMove->left     + lWidth;
		pRectToMove->top    = pRectRelative->top;
		pRectToMove->bottom = pRectToMove->top      + lHeight;
	}

	static inline void RightBelowRect(RECT *pRectToMove, const RECT *pRectRelative, LONG lOffset)
	{
		LONG lWidth  = pRectToMove->right  - pRectToMove->left;
		LONG lHeight = pRectToMove->bottom - pRectToMove->top;

		pRectToMove->left   = pRectRelative->right  + lOffset;
		pRectToMove->right  = pRectToMove->left     + lWidth;
		pRectToMove->bottom = pRectRelative->bottom;
		pRectToMove->top    = pRectToMove->bottom   - lHeight;
	}

	static inline void LeftRect(RECT *pRectToMove, const RECT *pRectRelative, LONG lOffset)
	{
		LONG lWidth  = pRectToMove->right  - pRectToMove->left;
		LONG lHeight = pRectToMove->bottom - pRectToMove->top;

		pRectToMove->right  = pRectRelative->left   - lOffset;
		pRectToMove->left   = pRectToMove->right    - lWidth;
		pRectToMove->top    = pRectRelative->top;
		pRectToMove->bottom = pRectToMove->top      + lHeight;
	}

	static inline void UnionRect(RECT *pRectDest, const RECT *pRectOther)
	{
		RECT rcTemp = *pRectDest;
		::UnionRect(pRectDest, &rcTemp, pRectOther);
	}

	// Give hWndRelative NULL to centre on the monitor work area with the mouse pointer.
	// The window is snapped to the monitor work area it appears on.
	static void CenterWindow(HWND hWndToMove, HWND hWndRelative);

	// Flags as for MonitorFromWindow.
	// If you don't want a rect pass in NULL.
	static bool GetAreasFromWindow(RECT *pRectFull, RECT *pRectWork, HWND hWnd, DWORD dwFlags);

	static BOOL GetWindowTextString(HWND hwnd, std::basic_string< TCHAR > *pStrResult);

	struct GlobalAtomHolder
	{
		ATOM a;
		GlobalAtomHolder(const wchar_t *szName) : a(::GlobalAddAtom(szName)) { }
		~GlobalAtomHolder() { if (a != 0) { ::GlobalDeleteAtom(a); a = 0; } }
		inline ATOM GetAtom() { return a; }
		inline const LPCWSTR GetAtomCast() { return MAKEINTRESOURCE(a); }
	};

	static GlobalAtomHolder s_SubclassAtomPropHolder;

	template< typename T > struct SubclassData
	{
	public:
		T *pData;
		HWND hWnd;
		WNDPROC fWndProcOrig;

		SubclassData(T *pDataIn, HWND hWndIn, WNDPROC fWndProcNewIn) : pData(pDataIn), hWnd(hWndIn), fWndProcOrig(0)
		{
			assert(hWnd != NULL);
			assert(IsWindowUnicode(hWnd));

			if (hWnd != NULL)
			{
				SetProp(hWnd, s_SubclassAtomPropHolder.GetAtomCast(), this);

				if (fWndProcNewIn)
				{
					#pragma warning(suppress:4244) // spurious warning due to stupid Win32 SDK header definition.
					#pragma warning(suppress:4312) // spurious warning due to stupid Win32 SDK header definition.
					fWndProcOrig = reinterpret_cast<WNDPROC>(SetWindowLongPtr(hWnd, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(fWndProcNewIn)));
				}
			}
		}

		~SubclassData()
		{
			if (hWnd != NULL)
			{
				RemoveProp(hWnd, s_SubclassAtomPropHolder.GetAtomCast());

				if (fWndProcOrig != NULL)
				{
					#pragma warning(suppress:4244) // spurious warning due to stupid Win32 SDK header definition.
					SetWindowLongPtr(hWnd, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(fWndProcOrig));

					fWndProcOrig = NULL;
				}

				hWnd = NULL;
			}
		}

		static inline SubclassData *GetSubclassData(HWND hWnd)
		{
			#pragma warning(suppress:4312) // spurious warning due to stupid Win32 SDK header definition.
			return reinterpret_cast<SubclassData *>(GetProp(hWnd, s_SubclassAtomPropHolder.GetAtomCast()));
		}
	private:
		// Prevent calls to copy constructor and assignment operator.
		SubclassData(const SubclassData &rhs);
		SubclassData &SubclassData::operator=(const SubclassData &rhs);
	};

	class DialogLayoutHelper
	{
	private:
		HWND m_hwndDlg;
		RECT m_rect1;
		RECT m_rect2;
		RECT m_rect3;
	public:
		DialogLayoutHelper(HWND hwndDlg)
		: m_hwndDlg (hwndDlg)
		{
			m_rect1.left   = 4; // Button X and related-controls X
			m_rect1.right  = 9; // Checkbox box width.
			m_rect1.top    = 4; // Button Y and related-controls Y
			m_rect1.bottom = 7; // Margin (X and Y)
			MapDialogRect(m_hwndDlg, &m_rect1);

			m_rect2.left   = 2; // Checkbox box-to-text gap width.
			m_rect2.right  = 0;
			m_rect2.top    = 3; // Label Assoc Gap Y
			m_rect2.bottom = 8; // Checkbox box height.
			MapDialogRect(m_hwndDlg, &m_rect2);

			m_rect3.left   = 50; // Standard button width.
			m_rect3.right  = 0;
			m_rect3.top    = 11; // First control in group-box.
			m_rect3.bottom = 14; // Edit control height. (Single line edit control.) Also button height.
			MapDialogRect(m_hwndDlg, &m_rect3);
		}

		// non-virtual
		~DialogLayoutHelper()
		{
		}

		LONG GetRelatedControlGapXPixels()    const { return m_rect1.left;   }
		LONG GetRelatedControlGapYPixels()    const { return m_rect1.top;    }
		LONG GetButtonWidthXPixels()          const { return m_rect3.left;   }
		LONG GetButtonGapXPixels()            const { return m_rect1.left;   }
		LONG GetButtonGapYPixels()            const { return m_rect1.top;    }
		LONG GetMarginPixels()                const { return m_rect1.bottom; }

		LONG GetLabelAssocGapYPixels()        const { return m_rect2.top;    }

		LONG GetGroupBoxFirstControlYPixels() const { return m_rect3.top; }

		// Resize static only pays attention to the input of pInOutRect if bSingleLine is false.
		bool ResizeStatic(  HWND hwndStatic,   RECT *pInOutRect, bool bResize, bool bRedraw, bool bSingleLine) const;
		bool ResizeCheckbox(HWND hwndCheckbox, RECT *pOutRect, bool bResize, bool bRedraw) const;
		bool ResizeEdit(    HWND hwndEdit,     RECT *pOutRect, bool bResize, bool bRedraw) const;
		bool ResizeButton(  HWND hwndButton,   RECT *pOutRect, bool bResize, bool bRedraw) const;

		void SetGroupBoxText(HWND hwndGroupBox, const wchar_t *szText, bool bNoPrefix) const;
	};

#ifdef DOPUS_PLUGIN_HELPER
#ifdef WC_BUTTON

	class WindowMaker
	{
	private:
		HMODULE m_hModule;
		HWND m_hWnd;
		HFONT m_hFont;
		DialogLayoutHelper *m_pLayout;
		OpusStringLoader *m_pStringLoader;
	public:
		WindowMaker(HMODULE hModule, HWND hWnd, HFONT hFont, DialogLayoutHelper *pLayout, OpusStringLoader *pStringLoader)
			: m_hModule(hModule)
			, m_hWnd(hWnd)
			, m_hFont(hFont)
			, m_pLayout(pLayout)
			, m_pStringLoader(pStringLoader)
		{
		}

		// non-virtual
		~WindowMaker()
		{
		}

		bool CreateCheck(HWND *pOutHwnd, RECT *pOutRect, ULONG uiCtrlId, UINT uiLabel, bool bResize, bool bRadio, bool bTriState, bool bAuto, bool bTabStop)
		{
			DWORD dwStyle = WS_CHILD | WS_VISIBLE;

			if (bTabStop)                dwStyle |= WS_TABSTOP;

			if (bRadio && bAuto)         dwStyle |= BS_AUTORADIOBUTTON;
			else if (bRadio)             dwStyle |= BS_RADIOBUTTON;
			else if (bTriState && bAuto) dwStyle |= BS_AUTO3STATE;
			else if (bTriState)          dwStyle |= BS_3STATE;
			else if (bAuto)              dwStyle |= BS_AUTOCHECKBOX;
			else                         dwStyle |= BS_CHECKBOX;

			HWND hWndResult = ::CreateWindowEx(WS_EX_NOPARENTNOTIFY, WC_BUTTON,
				uiLabel == 0 ? L"" : m_pStringLoader->Get(uiLabel), dwStyle, 0, 0, 0, 0,
				m_hWnd, (HMENU)UlongToPtr(uiCtrlId), m_hModule, NULL);

			if (hWndResult != NULL)
			{
				if (m_hFont != NULL)
				{
					::SendMessage(hWndResult, WM_SETFONT, reinterpret_cast< WPARAM >(m_hFont), FALSE);
				}

				m_pLayout->ResizeCheckbox(hWndResult, pOutRect, bResize, false);
			}

			if (pOutHwnd != NULL)
			{
				*pOutHwnd = hWndResult;
			}

			return (hWndResult != NULL);
		}

		bool CreateEdit(HWND *pOutHwnd, RECT *pOutRect, ULONG uiCtrlId, bool bResize, LONG lWidth, bool bAutoHScroll, bool bRight, bool bNumber, bool bReadOnly)
		{
			DWORD dwStyle = WS_CHILD | WS_VISIBLE | WS_TABSTOP;
			if (bAutoHScroll) dwStyle |= ES_AUTOHSCROLL;
			if (bRight)       dwStyle |= ES_RIGHT;
			if (bNumber)      dwStyle |= ES_NUMBER;
			if (bReadOnly)    dwStyle |= ES_READONLY;

			HWND hWndResult = ::CreateWindowEx(WS_EX_NOPARENTNOTIFY | WS_EX_CLIENTEDGE, WC_EDIT, L"", dwStyle,
				0, 0, lWidth, 0, m_hWnd, (HMENU)UlongToPtr(uiCtrlId), m_hModule, NULL);

			if (hWndResult != NULL)
			{
				if (m_hFont != NULL)
				{
					::SendMessage(hWndResult, WM_SETFONT, reinterpret_cast< WPARAM >(m_hFont), FALSE);
				}

				m_pLayout->ResizeEdit(hWndResult, pOutRect, bResize, false);
			}

			if (pOutHwnd != NULL)
			{
				*pOutHwnd = hWndResult;
			}

			return (hWndResult != NULL);
		}

		bool CreateSpin(HWND *pOutHwnd, ULONG uiCtrlId, bool bSetBuddyInt, bool bRight, bool bArrowKeys, bool bNoThousands)
		{
			DWORD dwStyle = WS_CHILD | WS_VISIBLE;
			if (bSetBuddyInt) dwStyle |= UDS_SETBUDDYINT;
			if (bRight)       dwStyle |= UDS_ALIGNRIGHT;
			if (bArrowKeys)   dwStyle |= UDS_ARROWKEYS;
			if (bNoThousands) dwStyle |= UDS_NOTHOUSANDS;

			HWND hWndResult = ::CreateWindowEx(WS_EX_NOPARENTNOTIFY, UPDOWN_CLASS, L"", dwStyle, 0, 0, 0, 0,
				m_hWnd, (HMENU)UlongToPtr(uiCtrlId), m_hModule, NULL);

			if (pOutHwnd != NULL)
			{
				*pOutHwnd = hWndResult;
			}

			return (hWndResult != NULL);
		}

		bool CreateGroup(HWND *pOutHwnd, ULONG uiCtrlId, UINT uiLabel)
		{
			HWND hWndResult = ::CreateWindowEx(WS_EX_NOPARENTNOTIFY, WC_BUTTON, m_pStringLoader->Get(uiLabel),
				WS_CLIPSIBLINGS | WS_CHILD | WS_VISIBLE | BS_GROUPBOX, 0, 0, 0, 0,
				m_hWnd, (HMENU)UlongToPtr(uiCtrlId), m_hModule, NULL);

			if (hWndResult != NULL)
			{
				if (m_hFont != NULL)
				{
					::SendMessage(hWndResult, WM_SETFONT, reinterpret_cast< WPARAM >(m_hFont), FALSE);
				}
			}

			if (pOutHwnd != NULL)
			{
				*pOutHwnd = hWndResult;
			}

			return (hWndResult != NULL);
		}

		bool CreateButton(HWND *pOutHwnd, RECT *pOutRect, ULONG uiCtrlId, UINT uiLabel, bool bResize, bool bDefPushButton)
		{
			DWORD dwStyle = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_TEXT;
			if (bDefPushButton) dwStyle |= BS_DEFPUSHBUTTON;
			else                dwStyle |= BS_PUSHBUTTON;

			HWND hWndResult = ::CreateWindowEx(WS_EX_NOPARENTNOTIFY, WC_BUTTON, m_pStringLoader->Get(uiLabel),
				dwStyle, 0, 0, 0, 0, m_hWnd, (HMENU)UlongToPtr(uiCtrlId), m_hModule, NULL);

			if (hWndResult != NULL)
			{
				if (m_hFont != NULL)
				{
					::SendMessage(hWndResult, WM_SETFONT, reinterpret_cast< WPARAM >(m_hFont), FALSE);
				}

				m_pLayout->ResizeButton(hWndResult, pOutRect, bResize, false);
			}

			if (pOutHwnd != NULL)
			{
				*pOutHwnd = hWndResult;
			}

			return (hWndResult != NULL);
		}

		bool CreateListBox(HWND *pOutHwnd, ULONG uiCtrlId, bool bSort, bool bExtendedSel)
		{
			DWORD dwStyle = WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_VSCROLL | LBS_NOINTEGRALHEIGHT | LBS_NOTIFY;
			if (bSort)        dwStyle |= LBS_SORT;
			if (bExtendedSel) dwStyle |= LBS_EXTENDEDSEL;

			HWND hWndResult = ::CreateWindowEx(WS_EX_NOPARENTNOTIFY | WS_EX_CLIENTEDGE, WC_LISTBOX, L"",
				dwStyle, 0, 0, 0, 0, m_hWnd, (HMENU)UlongToPtr(uiCtrlId), m_hModule, NULL);

			if (hWndResult != NULL)
			{
				if (m_hFont != NULL)
				{
					::SendMessage(hWndResult, WM_SETFONT, reinterpret_cast< WPARAM >(m_hFont), FALSE);
				}
			}

			if (pOutHwnd != NULL)
			{
				*pOutHwnd = hWndResult;
			}

			return (hWndResult != NULL);
		}

		bool CreateLabel(HWND *pOutHwnd, RECT *pOutRect, ULONG uiCtrlId, UINT uiLabel, bool bResize, bool bSingleLine, bool bVCenter)
		{
			DWORD dwStyle = WS_CHILD | WS_VISIBLE;
			if (bVCenter) dwStyle |= SS_CENTERIMAGE;

			HWND hWndResult = ::CreateWindowEx(WS_EX_NOPARENTNOTIFY, WC_STATIC, m_pStringLoader->Get(uiLabel),
				dwStyle, 0, 0, 0, 0, m_hWnd, (HMENU)UlongToPtr(uiCtrlId), m_hModule, NULL);

			if (hWndResult != NULL)
			{
				if (m_hFont != NULL)
				{
					::SendMessage(hWndResult, WM_SETFONT, reinterpret_cast< WPARAM >(m_hFont), FALSE);
				}

				m_pLayout->ResizeStatic(hWndResult, pOutRect, bResize, false, bSingleLine);
			}

			if (pOutHwnd != NULL)
			{
				*pOutHwnd = hWndResult;
			}

			return (hWndResult != NULL);
		}

		bool CreateBitmap(HWND *pOutHwnd, ULONG uiCtrlId, bool bNotify)
		{
			DWORD dwStyle = WS_CHILD | WS_VISIBLE | SS_BITMAP | SS_CENTERIMAGE | SS_REALSIZEIMAGE;
			if (bNotify) dwStyle |= SS_NOTIFY;

			HWND hWndResult = ::CreateWindowEx(WS_EX_NOPARENTNOTIFY, WC_STATIC, L"",
				dwStyle, 0, 0, 0, 0, m_hWnd, (HMENU)UlongToPtr(uiCtrlId), m_hModule, NULL);

			if (pOutHwnd != NULL)
			{
				*pOutHwnd = hWndResult;
			}

			return (hWndResult != NULL);
		}

		bool CreateCombo(HWND *pOutHwnd, RECT *pOutRect, ULONG uiCtrlId, bool bResize, bool bDropDownList, bool bDropDown)
		{
			DWORD dwStyle = WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_VSCROLL;
			if (bDropDownList)  dwStyle |= CBS_DROPDOWNLIST;
			else if (bDropDown) dwStyle |= CBS_DROPDOWN;
			else                dwStyle |= CBS_SIMPLE;

			HWND hWndResult = ::CreateWindowEx(WS_EX_NOPARENTNOTIFY, WC_COMBOBOX, L"",
				dwStyle, 0, 0, 0, 0, m_hWnd, (HMENU)UlongToPtr(uiCtrlId), m_hModule, NULL);

			if (hWndResult != NULL)
			{
				if (m_hFont != NULL)
				{
					::SendMessage(hWndResult, WM_SETFONT, reinterpret_cast< WPARAM >(m_hFont), FALSE);
				}

				m_pLayout->ResizeButton(hWndResult, pOutRect, bResize, false);
			}

			if (pOutHwnd != NULL)
			{
				*pOutHwnd = hWndResult;
			}

			return (hWndResult != NULL);
		}
	};

#endif
#endif

	static BOOL GradientFillIfColor(HDC hdc, PTRIVERTEX pVertices, ULONG nVertices, PVOID pMesh, ULONG nMeshElements, ULONG dwMode);

#ifndef LIM_SMALL
#ifndef LIM_LARGE
enum _LI_METRIC
{
   LIM_SMALL, // corresponds to SM_CXSMICON/SM_CYSMICON
   LIM_LARGE, // corresponds to SM_CXICON/SM_CYICON
};
#endif
#endif

	static HRESULT LoadIconMetricOrDefault(HINSTANCE hinst, PCWSTR pszName, int lims, HICON *phico, bool *pbDestroyIcon);

	// DeleteObject the result
	static HFONT GetMessageFont();

// From uxtheme.h

#ifndef THEMESIZE
typedef enum THEMESIZE
{
	TS_MIN,             // minimum size
	TS_TRUE,            // size without stretching
	TS_DRAW,            // size that theme mgr will use to draw part
};
#endif

#ifndef TABP_BODY
#define TABP_BODY 10
#endif

	// *********************************************************
	// This class is NOT THREAD SAFE
	// If you have multiple UI threads use some kind of locking.
	// *********************************************************
	class ThemeHelper  
	{
	public:
		ThemeHelper();
		virtual ~ThemeHelper();

		bool IsAvailable();

		bool CacheBackgroundTabThemeDialog(HWND hWnd);

		// Returns TRUE iff it was enabled.
		bool EnableThemeDialogTexture(HWND hWnd, DWORD dwFlags);

		// You have to call EnableThemeDialogTexture and have Windows erase the background in order to get
		// "transparent" controls if version 6 common-controls are enabled. On XP the tab dialog texture
		// is a limited height and runs out if the window is sized too tall. This works around that problem
		// by tiling the texture, flipping it over on every alternate line so that it remains smooth.
		// Some child controls may paint their own backgrounds based on what they think the parent window's
		// background should be, so some "transparent" controls may need subclassing to paint properly below
		// the point where the background texture runs out on XP.
		// On Vista the texture seems to tile correctly regardless of the dialog size without needing help.
		void EraseBackgroundTabThemeDialog(HWND hWnd, HDC hDC, const RECT &rcClient, const RECT &rcBackgroundWnd);

		HBRUSH GetTempBrush(HWND hWnd, HDC hDC, const RECT &rcClient, const RECT &rcBackgroundWnd);

//		BOOL IsThemeBackgroundPartiallyTransparentButton(HWND hWndControl, int iPartId, int iStateId);

		HRESULT DrawThemeParentBackground(HWND hWnd, HDC hDC, const RECT *prcDrawChildCoords);

	protected:
		typedef HANDLE  (STDAPICALLTYPE *OPENTHEMEDATA_PTR)(HWND hwnd, LPCWSTR pszClassList);
		typedef HRESULT (STDAPICALLTYPE *CLOSETHEMEDATA_PTR)(HANDLE hTheme);
		typedef HRESULT (STDAPICALLTYPE *DRAWTHEMEBACKGROUND_PTR)(HANDLE hTheme, HDC hdc, int iPartId, int iStateId, const RECT *pRect, OPTIONAL const RECT *pClipRect);
		typedef HRESULT (STDAPICALLTYPE *DRAWTHEMEPARENTBACKGROUND_PTR)(HWND hwnd, HDC hdc, __in_opt const RECT* prc);
		typedef HRESULT (STDAPICALLTYPE *GETTHEMEPARTSIZE_PTR)(HANDLE hTheme, HDC hdc, int iPartId, int iStateId, OPTIONAL RECT *prc, enum THEMESIZE eSize, OUT SIZE *psz);
		typedef HRESULT (STDAPICALLTYPE *ENABLETHEMEDIALOGTEXTURE_PTR)(HWND hwnd, DWORD dwFlags);
//		typedef BOOL    (STDAPICALLTYPE *ISTHEMEBACKGROUNDPARTIALLYTRANSPARENT_PTR(HTHEME hTheme, int iPartId, int iStateId);

		HMODULE m_hmodUxThemeDLL;

		OPENTHEMEDATA_PTR							m_fOpenThemeData;
		CLOSETHEMEDATA_PTR							m_fCloseThemeData;
		DRAWTHEMEBACKGROUND_PTR						m_fDrawThemeBackground;
		DRAWTHEMEPARENTBACKGROUND_PTR				m_fDrawThemeParentBackground;
		GETTHEMEPARTSIZE_PTR						m_fGetThemePartSize;
		ENABLETHEMEDIALOGTEXTURE_PTR				m_fEnableThemeDialogTexture;
//		ISTHEMEBACKGROUNDPARTIALLYTRANSPARENT_PTR	m_fIsThemeBackgroundPartiallyTransparent;

		SIZE m_sizePart;
		HBITMAP m_hbmpDIB;

		HDC m_hdcMemTemp;
		HBRUSH m_hBrushTemp;
	};

	// delete[] the result.
	// Only shrinking works at the moment. Enlarging is untested and lots of bugs were found in the shrinking code so don't trust it.
	static RGBQUAD *RGBQAllocateStretchPreserveAlpha(DWORD dwNewWidth, DWORD dwNewHeight, DWORD dwInWidth, DWORD dwInHeight, const RGBQUAD *lpSourceBits);

	static void InitRGBQBitmapHeader(BITMAPINFO *pBitmapInfo, int width, int height, bool topDown);

#ifdef VIEWERPLUGINFILEINFO

#ifndef DOPUS_PLUGIN_LEO_NO_COPYVERSION
	// If initialisation fails after the call to CopyVersionResourceToViewerPluginInfo then the lpVPInfo->hIconSmall icon must be destroyed.
	static bool CopyVersionResourceToViewerPluginInfo(HMODULE g_hDllModule, WORD iconId, UINT uiOpusMsgIdName, UINT uiOpusMsgIdDesc, LPVIEWERPLUGININFO lpVPInfo);
	static bool CopyVersionNumber(DWORD *pdwHigh, DWORD *pdwLow, TCHAR *szItemName, BYTE *pLang, const LPVOID pVerData);
	static bool CopyVersionString(TCHAR *szDestBuffer, UINT uiDestBufferSize, TCHAR *szItemName, BYTE *pLang, const LPVOID pVerData);
#endif

	static inline void WriteFileInfoInfoLine(VIEWERPLUGINFILEINFO *pInfo, const wchar_t *szTypeName)
	{
		if (NULL != pInfo
		&&	NULL != pInfo->lpszInfo
		&&	0    <  pInfo->cchInfoMax
		&&	NULL != szTypeName)
		{
			if (pInfo->szImageSize.cx > 0
			&&	pInfo->szImageSize.cy > 0
			&&	pInfo->iNumBits > 0)
			{
				_sntprintf_s(pInfo->lpszInfo, pInfo->cchInfoMax, _TRUNCATE, _T("%ld x %ld x %d %s"), pInfo->szImageSize.cx, pInfo->szImageSize.cy, pInfo->iNumBits, szTypeName);
			}
			else if (pInfo->szImageSize.cx > 0
				 &&	 pInfo->szImageSize.cy > 0)
			{
				_sntprintf_s(pInfo->lpszInfo, pInfo->cchInfoMax, _TRUNCATE, _T("%ld x %ld %s"), pInfo->szImageSize.cx, pInfo->szImageSize.cy, szTypeName);
			}
			else
			{
				StringCopy(pInfo->lpszInfo, pInfo->cchInfoMax, szTypeName);
			}
		}
	}
#endif // VIEWERPLUGINFILEINFO

	class CriticalSectionScoper
	{
	private:
		CRITICAL_SECTION *m_pCS;
	private:
		CriticalSectionScoper(const CriticalSectionScoper &rhs); // disallow
		CriticalSectionScoper &operator=(const CriticalSectionScoper &rhs); // disallow
	public:
		CriticalSectionScoper(CRITICAL_SECTION *pCS) : m_pCS(pCS) { EnterCriticalSection(m_pCS); }
		/*not virtual*/ ~CriticalSectionScoper()                  { LeaveCriticalSection(m_pCS); }
	};

	template< typename T > class CriticalObject
	{
	private:
		mutable CRITICAL_SECTION m_cs;
	private:
		CriticalObject(const CriticalObject< T > &rhs); // disallow
		CriticalObject &operator=(const CriticalObject< T > &rhs); // disallow
	public:
		CriticalObject(T t) : object(t)   { InitializeCriticalSection(&m_cs); }
		/*not virtual*/ ~CriticalObject() { DeleteCriticalSection(&m_cs); }
		void Lock() const                 { EnterCriticalSection(&m_cs); }
		void Unlock() const               { LeaveCriticalSection(&m_cs); }
		T object;
	};

	template< typename T > class CriticalObjectScoper
	{
	private:
		const CriticalObject< T > *m_pCO;
	private:
		CriticalObjectScoper(const CriticalObjectScoper< T > &rhs); // disallow
		CriticalObjectScoper &operator=(const CriticalObjectScoper< T > &rhs); // disallow
	public:
		CriticalObjectScoper(const CriticalObject< T > *pCO) : m_pCO(pCO) { m_pCO->Lock(); }
		/*not virtual*/ ~CriticalObjectScoper()                           { m_pCO->Unlock(); }
	};

	/*
	// Never pass ArrayScoper objects across process/DLL boundaries.
	// (It's all inline so the new/delete calls may hit different heaps.)
	class ArrayScoper
	{
	protected:
		void *m_pVoid;
	private:
		ArrayScoper(const ArrayScoper &rhs); // disallow
		ArrayScoper &operator=(const ArrayScoper &rhs); // disallow
	public:
		ArrayScoper(void *pVoid) : m_pVoid(pVoid) { }
		~ArrayScoper() { delete [] m_pVoid; m_pVoid = NULL; } // not virtual
	};
	*/

	// Never pass MultiArrayScoper objects across process/DLL boundaries.
	// (It's all inline so the new/delete calls may hit different heaps.)
	class MultiArrayScoper
	{
	private:
		std::vector< void * > m_voidPointers;
	private:
		MultiArrayScoper(const MultiArrayScoper &rhs); // disallow
		MultiArrayScoper &operator=(const MultiArrayScoper &rhs); // disallow
	public:
		MultiArrayScoper() { }
		/* not virtual */ ~MultiArrayScoper() { while(!m_voidPointers.empty()) { delete [] m_voidPointers.back(); m_voidPointers.pop_back(); } }
		void Add(void *pVoid) { m_voidPointers.push_back(pVoid); }
	};

	// Never pass AutoBuffer objects across process/DLL boundaries.
	// (It's all inline so the new/delete calls may hit different heaps.)
	template <typename T> class AutoBuffer
	{
	private:
		BYTE *m_pBuffer;
		size_t m_size;
	private:
		AutoBuffer(const AutoBuffer &rhs); // disallow
		AutoBuffer &operator=(const AutoBuffer &rhs); // disallow
	public:
		AutoBuffer() : m_pBuffer(0) , m_size(0) { }
		/* not virtual */ ~AutoBuffer() { Free(); } // Must not change Win32 LastError
		void Free() { delete[] m_pBuffer; m_pBuffer = 0; m_size = 0; } 		// Must not change Win32 LastError

		// Allocating a zero-length buffer is legal and results in a non-null buffer.
		// Make no assumptions about the content of the buffer after allocation.
		bool AllocateBytes(size_t s)
		{
			if (s == m_size && m_pBuffer != 0) { return true; }
			Free();
			m_pBuffer = new(std::nothrow) BYTE[s];
			if (m_pBuffer == 0) { m_size = 0; ::SetLastError(ERROR_NOT_ENOUGH_MEMORY); return false; }
			m_size = s;
			return true;
		}

		// If there's already a buffer big enough (or bigger) then do nothing, else (re-)allocate.
		bool AllocateMinimumBytes(size_t s)
		{
			if (s <= m_size && m_pBuffer != 0) { return true; }
			return AllocateBytes(s);
		}

		bool        AllocateElements(size_t c) {        return AllocateBytes(c * sizeof(T)); }
		bool AllocateMinimumElements(size_t c) { return AllocateMinimumBytes(c * sizeof(T)); }

		T *    GetBuffer()             { return reinterpret_cast<       T * >(m_pBuffer); }
		T *    GetBuffer()       const { return reinterpret_cast< const T * >(m_pBuffer); }
		size_t GetSizeBytes()    const { return m_size; }
		size_t GetSizeElements() const { return (m_size / sizeof(T)); }
		bool   IsAllocated()     const { return (m_pBuffer != 0); }
	};

	class HandleScoper
	{
	private:
		HANDLE m_h;
		bool m_bHasHandle;
	private:
		HandleScoper(const HandleScoper &rhs); // disallow
		HandleScoper &operator=(const HandleScoper &rhs); // disallow
	public:
		HandleScoper()         : m_h(INVALID_HANDLE_VALUE), m_bHasHandle(false) { }
		HandleScoper(HANDLE h) : m_h(h),                    m_bHasHandle(true)  { }
		/* not virtual */ ~HandleScoper() { Close(); }
		inline bool   HasHandle() const   { return m_bHasHandle; }
		inline void   Forget()            { m_h = INVALID_HANDLE_VALUE; m_bHasHandle = false; }
		inline void   Close()             { if (m_bHasHandle) { ::CloseHandle(m_h); Forget(); } }
		inline HANDLE Get() const         { assert(m_bHasHandle); return m_h; }
		inline void   Set(HANDLE h)       { Close(); m_h = h; m_bHasHandle = true; }
	};

	class LocalFreeScoper
	{
	private:
		HLOCAL m_h;
		bool m_bHasHandle;
	private:
		LocalFreeScoper(const LocalFreeScoper &rhs); // disallow
		LocalFreeScoper &operator=(const LocalFreeScoper &rhs); // disallow
	public:
		LocalFreeScoper()         : m_h(NULL), m_bHasHandle(false) { }
		LocalFreeScoper(HLOCAL h) : m_h(h),    m_bHasHandle(true)  { }
		/* not virtual */ ~LocalFreeScoper() { Free(); }
		inline bool HasHandle() const         { return m_bHasHandle; }
		inline void Forget()                  { m_h = NULL; m_bHasHandle = false; }
		inline void Free()                    { if (m_bHasHandle) { ::LocalFree(m_h); Forget(); } }
		inline HLOCAL Get() const             { assert(m_bHasHandle); return m_h; }
		inline void Set(HLOCAL h)             { Free(); m_h = h; m_bHasHandle = true; }
		inline bool AllocateFixed(SIZE_T uBytes)
		{
			Free();
			m_h = ::LocalAlloc(LMEM_FIXED, uBytes);
			if (m_h) { m_bHasHandle = true; }
			return m_bHasHandle;
		}
	};

	class FileAndStream
	{
	public:
		// hAbortEvent can be null.
		FileAndStream(const wchar_t *szFilePath, HANDLE hAbortEvent, bool bOpenTempCopy);

		// pStream can be NULL but only if you are never going to call methods which read data or get the steam or file path.
		FileAndStream(const wchar_t *szFileName, IStream *pStream, bool bNoRandomSeek, bool bSlow);

		FileAndStream(const wchar_t *szFileName, const BYTE *pData, UINT cbDataSize);

		~FileAndStream(); // Warning: Non-virtual destructor.

	private:
		// disallow
		FileAndStream(const FileAndStream &rhs);
		FileAndStream &operator=(const FileAndStream &rhs);

	public:

		// This allows you to set the OpenTempCopy flag after construction.
		// Any subsequent call to GetFilePath will result in the file being copied to a temp path if
		// it had not already been done. Any preceeding call to GetFilePath may, of course, have alredy
		// received the real filename. If the data is in a stream then calling this has no effect as
		// GetFilePath will always cause the data to be saved to a temp file regardless of the flag.
		void SetOpenTempCopy(bool bOpenTemp);

		// The file name may not match the name of the actual file. It is indicative of the original file's name
		// but the file path returned by GetFilePath may be a temp-file with a different name. GetFileName should
		// be used for display purposes and for any logic which depends on the original file's name.
		// Calling GetFileName will not trigger any files or streams to be created.
		bool GetFileName(std::basic_string< TCHAR > *pstrFileName);
		bool GetFileExtension(std::basic_string< TCHAR > *pstrFileExtension, bool bIncludeDot);

		// GetNominalFilePath gets a full path that will be similar to, but may not be the same as,
		// the file path returned by GetFilePath. You can use this to test for non-ASCII characters
		// in the file path without triggering a stream to be written to a file as you would if you
		// called GetFilePath.
		bool GetNominalFilePath(std::basic_string< TCHAR > *pstrNominalFilePath);
		bool GetNominalTempFilePath(std::basic_string< TCHAR > *pstrNominalTempFilePath);

		bool HasFilePath();
		bool HasStream();

		// These return the current seeking ability. Basically tells you whether or not calling GetStream will have to do an expensive extraction to create a seekable stream.
		bool HasFastRandomSeek() { return !m_bNoRandomSeekInStream && !m_bSlowStream; } // Does not imply there is a stream (but it'll be cheap to create if there isn't one).
		bool HasNoRandomSeekStream() { return m_bNoRandomSeekInStream; }
		bool HasSlowStream()         { return m_bSlowStream; }

		HANDLE GetAbortEvent() { return m_hAbortEvent; } // May return NULL if there isn't one.

		// These calls will cause a file or stream to be created if there isn't one already.
		// Call HasFilePath and HasStream to see what already exists if you can work with both and want
		// to avoid the overhead of conversion.

		bool GetFilePath(std::basic_string< TCHAR > *pstrFilePath);

		// The stream will not be AddRef'd by this call. It will exist as long as the CFileAndSteam does
		// so you should not normally need to AddRef it.
		// If a stream is returned it will always be fast and seekable.
		// Whenever GetStream is called the stream's position is reset to the start.
		bool GetStream(IStream **ppStream);

		// Call this to get the current IStream, if any.
		// Unlike the GetStream() method:
		// - The returned IStream *will* be AddRef'd and you must Release it.
		// - The returned IStream may not be seekable.
		// - If the object contains a filepath that hasn't already been converted into an IStream then that conversion will not happen and the method will fail instead.
		// - The returned IStream will not have its position reset to the start of the IStream (since it may not be seekable).
		bool GetStreamAddRefNoConversionNoSeekReset(IStream **ppStream, bool *pOutNoRandomSeek, bool *pOutSlow);

	private:
		bool PrepareFileForReading();
		bool WriteStreamToTempFile();

	private:
		std::wstring m_strFileName;

		std::wstring m_strFilePath;
		bool m_bDeleteFilePathOnClose;
		bool m_bOpenTempCopyFilePath;

		IStream *m_pStream;
		bool m_bNoRandomSeekInStream;
		bool m_bSlowStream;

		bool m_bStreamIsFile;

		HANDLE m_hAbortEvent;
	};

	static bool GetComInterfaceName(std::basic_string< TCHAR > *pstrOutName, REFIID iid);
	static HRESULT IDispGetNamedProperty(IDispatch *pDisp, const wchar_t *szName, VARIANT *pvOut);
	static HRESULT IDispPutNamedProperty(IDispatch *pDisp, const wchar_t *szName, VARIANT *pvIn);

#if 0
	class LowIntegrity
	{
	public:
		LowIntegrity(HANDLE hThread);
		/* not virtual */ ~LowIntegrity();

		bool Succeeded() { return m_bSuccess; }
	private:
		bool m_bSuccess;
		HANDLE m_hThread;
	};
#endif

	// This requires both a section and a key name. It won't get you a list of names like the real GetPrivateProfileString.
	// Blank values will never be returned and are considered errors.
	static bool GetPrivateProfileString(std::basic_string< TCHAR > *pstrResult, const TCHAR *szFileName, const TCHAR *szSectionName, const TCHAR *szKeyName);

#ifndef _DEBUG
	static inline void SetThreadName(LPCSTR szThreadName,  DWORD dwThreadId = -1) {}
	static inline void SetThreadName(LPCWSTR szThreadName, DWORD dwThreadId = -1) {}
#else
	// SetThreadName in Visual Studio Debugger.
	// See http://msdn2.microsoft.com/en-us/library/xcb2z8hs.aspx

	#pragma pack(push,8)
	typedef struct tagTHREADNAME_INFO
	{
		DWORD dwType; // Must be 0x1000.
		LPCSTR szName; // Pointer to name (in user addr space).
		DWORD dwThreadID; // Thread ID (-1=caller thread).
		DWORD dwFlags; // Reserved for future use, must be zero.
	} THREADNAME_INFO;
	#pragma pack(pop)

	static void SetThreadName(LPCSTR szThreadName, DWORD dwThreadId = -1)
	{
		Sleep(10);
		THREADNAME_INFO info;
		info.dwType = 0x1000;
		info.szName = szThreadName;
		info.dwThreadID = dwThreadId;
		info.dwFlags = 0;

		__try
		{
			::RaiseException(0x406D1388, 0, sizeof(info)/sizeof(ULONG_PTR), reinterpret_cast<ULONG_PTR*>(&info));
		}
		__except(EXCEPTION_CONTINUE_EXECUTION)
		{
		}
	}

	static void SetThreadName(LPCWSTR szThreadName, DWORD dwThreadId = -1)
	{
		char *szAsciiName = LeetWCtoMB(szThreadName);

		if (szAsciiName != NULL)
		{
			SetThreadName(szAsciiName);

			delete[] szAsciiName;
		}
	}
#endif

	class HousekeepingThread
	{
	public:
		HousekeepingThread(DWORD dwThreadIntervalMS);
		virtual ~HousekeepingThread();

		void SetThreadIntervalMS(DWORD dwThreadIntervalMS);
		DWORD GetThreadIntervalMS();

		// Start the thread if it is stopping or stopped.
		void Start();

		// Stop should never be called by two threads at once and in most cases should only ever be called implicitly by the destructor.
		void Stop();

	protected:
		// RequestStop should only be called by MainTask to prevent further iterations.
		void RequestStop();

		// Once the thread has started, MainTask is called every m_dwThreadIntervalMS until Stop or RequestStop has been called.
		virtual void MainTask() = 0;

	private:
		static unsigned __stdcall staticThread(void *pVoidThis);
		HousekeepingThread(const HousekeepingThread &rhs); // disallow
		HousekeepingThread &operator=(const HousekeepingThread &rhs); // disallow
	protected:
		CRITICAL_SECTION m_csThreadVariables;
	private:
		CRITICAL_SECTION m_csStart;
		DWORD m_dwThreadIntervalMS;
		HANDLE m_hThread;
		HANDLE m_hStopEvent;
	};


	static void OutputDebug(const TCHAR *szPrefix, const TCHAR *szClass, const TCHAR *szMethod, const TCHAR *szMessage)
	{
		std::basic_string< TCHAR > strMsg = szPrefix;

		if (szClass != NULL)
		{
			strMsg += szClass;

			if (szMethod != NULL)
			{
				strMsg += _T("::");
			}
			else
			{
				strMsg += _T(": ");
			}
		}

		if (szMethod != NULL)
		{
			strMsg += szMethod;
			strMsg += _T(": ");
		}

		if (szMessage != NULL)
		{
			strMsg += szMessage;
		}

		while(!strMsg.empty() && _istspace(strMsg[strMsg.length() - 1]))
		{
			strMsg.resize(strMsg.length() - 1);
		}

		strMsg += _T("\r\n");

		::OutputDebugString(strMsg.c_str());
	}

	static void VOutputDebugFormat(const TCHAR *szPrefix, const TCHAR *szClass, const TCHAR *szMethod, const TCHAR *szFormat, va_list args)
	{
		TCHAR *szMainMsg = LeoHelpers::VStringAllocAndFormat(szFormat, args);

		if (szMainMsg != NULL)
		{
			LeoHelpers::OutputDebug(szPrefix, szClass, szMethod, szMainMsg);

			delete[] szMainMsg;
		}
	}

	static void OutputDebugFormat(const TCHAR *szPrefix, const TCHAR *szClass, const TCHAR *szMethod, const TCHAR *szFormat, ...)
	{
		va_list args;
		va_start(args, szFormat);

		LeoHelpers::VOutputDebugFormat(szPrefix, szClass, szMethod, szFormat, args);

		va_end(args);
	}

	class BinaryData
	{
	protected:
		BYTE *m_pData;
		DWORD m_dwSize;
		bool  m_bUseLocalAlloc;

	public:
		BinaryData(bool bUseLocalAlloc)
		: m_pData(NULL)
		, m_dwSize(0)
		, m_bUseLocalAlloc(bUseLocalAlloc)
		{
		}

		~BinaryData()
		{
			Clear();
		}

		void Clear()
		{
			if (m_pData != NULL)
			{
				if (m_bUseLocalAlloc)
				{
					::LocalFree(m_pData);
				}
				else
				{
					delete[] m_pData;
				}
			}
			m_pData = NULL;
			m_dwSize = 0;
		}

		void Forget()
		{
			m_pData = NULL;
			m_dwSize = 0;
		}

		void SetAndOwn(BYTE *pData, DWORD dwSize, bool bUseLocalAlloc)
		{
			Clear();

			m_pData = pData;
			m_dwSize = dwSize;
			m_bUseLocalAlloc = bUseLocalAlloc;
		}

		bool Load(const TCHAR *filename)
		{
			bool bResult = false;

			Clear();

			HANDLE hFile = ::CreateFile(filename, FILE_READ_ATTRIBUTES|FILE_READ_DATA, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, NULL);

			if (INVALID_HANDLE_VALUE == hFile)
			{
			}
			else
			{
				DWORD dwFileSize = ::GetFileSize(hFile, NULL);

				if (INVALID_FILE_SIZE == dwFileSize)
				{
				}
				else if (0 == dwFileSize)
				{
				}
				else
				{
					if (m_bUseLocalAlloc)
					{
						m_pData = reinterpret_cast< BYTE *>( ::LocalAlloc(LMEM_FIXED, dwFileSize) );
					}
					else
					{
						m_pData = new(std::nothrow) BYTE[dwFileSize];
					}

					if (m_pData == NULL)
					{
					}
					else
					{
						if (0 == ::ReadFile(hFile, m_pData, dwFileSize, &m_dwSize, NULL) || m_dwSize != dwFileSize)
						{
						}
						else
						{
							bResult = true;
						}
					}
				}

				::CloseHandle(hFile);
			}

			if (!bResult)
			{
				Clear();
			}

			return bResult;
		}

		/*
		bool Save(const TCHAR *filename)
		{
			return Save(filename, m_pData, m_dwSize);
		}

		static bool Save(const TCHAR *filename, const BYTE *pData, DWORD dwSize)
		{
			bool bResult = false;

			HANDLE hFile = ::CreateFile(filename, FILE_WRITE_DATA, FILE_SHARE_READ, NULL, CREATE_NEW, FILE_FLAG_SEQUENTIAL_SCAN, NULL);

			if (hFile == INVALID_HANDLE_VALUE)
			{
				DWORD dwErr = GetLastError();

				if (dwErr == ERROR_ALREADY_EXISTS
				||	dwErr == ERROR_FILE_EXISTS)
				{
					_tcprintf(_T("Failed to save (already exists): %s\n"), filename);
				}
				else
				{
					_tcprintf(_T("Failed to save (error %ld): %s\n"), dwErr, filename);
				}
			}
			else
			{
				DWORD dwWritten = 0;

				if (!WriteFile(hFile, pData, dwSize, &dwWritten, NULL))
				{
					DWORD dwErr = GetLastError();
					_tcprintf(_T("Failed to write (error %ld): %s\n"), dwErr, filename);
				}
				else if (dwWritten != dwSize)
				{
					_tcprintf(_T("Failed to write (only partially written): %s\n"), filename);
				}
				else
				{
					bResult = true;
				}

				CloseHandle(hFile);

				if (!bResult)
				{
					DeleteFile(filename);
				}
			}

			return bResult;
		}
		*/

		BYTE *GetBuffer()
		{
			return m_pData;
		}

		DWORD GetSize()
		{
			return m_dwSize;
		}
	};

	static bool IsWindowsVersionGreaterOrEqual(DWORD dwMaj, DWORD dwMin, WORD wSPMaj, WORD wSPMin);

	static bool IsWindowsXPOrAbove()
	{
		// 5.0.0.0 would mean Windows Server 2003, Windows XP, or Windows 2000
		// 5.1.0.0 means Windows XP
		return IsWindowsVersionGreaterOrEqual(5,1,0,0);
	}

	static bool IsWindowsVistaOrAbove()
	{
		// 6.0.0.0 means Windows Vista or Windows Server 2008
		return IsWindowsVersionGreaterOrEqual(6,0,0,0);
	}

	static bool SetClipboard(HWND hwndClipboardOwner, const wchar_t *sz);
	static bool GetClipboard(HWND hwndClipboardOwner, std::wstring *pstrOut, bool bPrefixUTF16BOM, bool *pbOpenFailed);

	inline static DWORD SwapDWORD(DWORD dw)
	{
		return ((dw<<24)&0xFF000000)
			 + ((dw<<8 )&0x00FF0000)
			 + ((dw>>8 )&0x0000FF00)
			 + ((dw>>24)&0x000000FF);
	}

	inline static WORD SwapWORD(WORD w)
	{
		return ((w<<8 )&0xFF00)
			 + ((w>>8 )&0x00FF);
	}
};
