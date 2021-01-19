#include "stdafx.h"
#include "LeoHelpers.h"
#include <locale.h>

bool LeoHelpers::LeetRegQueryValue(HKEY hKeyParent, const TCHAR *szKeyPath, const TCHAR *szValueName, REGSAM sam,
									   DWORD *pdwType, void **plpData, DWORD *pcbData)
{
	// Documentation doesn't state that registry functions call SetLastError().
	// We need to preserve the error values they return through clean-up anyway.
	LONG regRes = ERROR_SUCCESS;

	HKEY hKeyToReadValueFrom;

	DWORD dwPcdDataStandin; // Used for getting the registry size if the user doesn't care about it.

	if (NULL == pcbData)
	{
		pcbData = &dwPcdDataStandin;
	}

	// szKeyPath may be NULL in which case another handle like hKeyParent is openned (and must be closed).
	regRes = RegOpenKeyEx(hKeyParent, szKeyPath, 0, sam | KEY_QUERY_VALUE, &hKeyToReadValueFrom);

	if (ERROR_SUCCESS == regRes)
	{
		if (NULL == plpData)
		{
			// They just want the type and/or the data size.
			regRes = RegQueryValueEx(hKeyToReadValueFrom, szValueName, 0, pdwType, NULL, pcbData);
		}
		else
		{
			// We add 8 bytes of zero padding to the end of any buffer we return. This is because
			// the registry does not check inputs to ensure they are the correct length for the
			// datatype and also does not ensure that strings are null-terminated and multi-strings are
			// double-null-terminated. Rather than add checks into all functions which call this one,
			// we pad with as many zero bytes as are needed to correct anything we might read.
			// 8 means a QUADDWORD that was zero length in the registry would still be safe.
			// 8 is also more than enough for double-null-terminator of a Unicode multi-string.
			const DWORD dwZeroPadding = 8;

			BYTE *allocatedBuffer = NULL;
			size_t AllocatedBufferSize = 0; // For HKEY_PERFORMANCE_DATA, really.

			// Get the required buffer size by giving NULL for data buffer.
			regRes = RegQueryValueEx(hKeyToReadValueFrom, szValueName, 0, pdwType, NULL, pcbData);

			if (ERROR_SUCCESS == regRes)
			{
				AllocatedBufferSize = *pcbData;
				allocatedBuffer = new BYTE[AllocatedBufferSize + dwZeroPadding];
			}
			else if (ERROR_MORE_DATA == regRes)
			{
				// Most likely tried to read HKEY_PERFORMANCE_DATA. *pcdData is undefined at this point.
				*pcbData = 1024; // Must be a number x such that (Int(x/2) > 0)
				AllocatedBufferSize = *pcbData;
				allocatedBuffer = new BYTE[AllocatedBufferSize + dwZeroPadding];
			}
			else
			{
				*pcbData = 0;
				// leave allocatedBuffer NULL so the while loop does not execute.
			}

			while (NULL != allocatedBuffer)
			{
				regRes = RegQueryValueEx(hKeyToReadValueFrom, szValueName, 0, pdwType, allocatedBuffer, pcbData);

				if (ERROR_SUCCESS == regRes)
				{
					// Zero out the padding we added.
					ZeroMemory(allocatedBuffer + AllocatedBufferSize, dwZeroPadding);

					*plpData = allocatedBuffer;
					allocatedBuffer = NULL;
					break;
				}
				else if (ERROR_MORE_DATA == regRes)
				{
					// Try a larger buffer.
					delete [] allocatedBuffer;
					*pcbData = static_cast<DWORD>(AllocatedBufferSize + (AllocatedBufferSize/2));
					AllocatedBufferSize = *pcbData;
					allocatedBuffer = new BYTE[AllocatedBufferSize + dwZeroPadding];
				}
				else
				{
					*pcbData = 0;
					delete [] allocatedBuffer;
					allocatedBuffer = NULL;
					break;
				}
			}
		}

		LONG closeRes = RegCloseKey(hKeyToReadValueFrom);
		if (ERROR_SUCCESS == regRes) { regRes = closeRes; }
	}

	if (ERROR_SUCCESS != regRes)
	{
		SetLastError(regRes);
	}

	return(ERROR_SUCCESS == regRes);
}

bool LeoHelpers::LeetRegQueryDWORDValue(HKEY hKeyParent, const TCHAR *szKeyPath, const TCHAR *szValueName, REGSAM sam, DWORD *pdwRes)
{
	DWORD dwErrToSet = ERROR_SUCCESS;

	DWORD dwType;
	void *pData;

	if (NULL != pdwRes)
	{
		*pdwRes = 0;
	}

	DWORD cbData = 0;

	if (! LeetRegQueryValue(hKeyParent, szKeyPath, szValueName, sam, &dwType, &pData, &cbData) )
	{
		dwErrToSet = GetLastError();
	}
	else
	{
		if (REG_DWORD != dwType)
		{
			dwErrToSet = ERROR_INVALID_DATA;
		}
		else
		{
			if (NULL != pdwRes)
			{
				*pdwRes = *((DWORD*)pData);
			}
		}

		delete [] pData;
	}

	if (ERROR_SUCCESS != dwErrToSet)
	{
		SetLastError(dwErrToSet);
	}

	return(ERROR_SUCCESS == dwErrToSet);
}

// Returns false on failure. Use GetLastError().
// Uses LeetRegQueryValue() to get you an int without having to worry about buffers, types and so on.
bool LeoHelpers::LeetRegQueryIntValue(HKEY hKeyParent, const TCHAR *szKeyPath, const TCHAR *szValueName, REGSAM sam, int *pInt)
{
	DWORD dwData = 0;

	if (!LeetRegQueryDWORDValue(hKeyParent, szKeyPath, szValueName, sam, &dwData))
	{
		return false;
	}

	assert(sizeof(int)==sizeof(DWORD));

	if (NULL != pInt && sizeof(int)==sizeof(DWORD))
	{
		*pInt = *reinterpret_cast<int*>(&dwData);
	}

	return true;
}

// Returns false on failure. Use GetLastError().
// Uses LeetRegQueryValue() to get you a bool without having to worry about buffers, types and so on.
bool LeoHelpers::LeetRegQueryBoolValue(HKEY hKeyParent, const TCHAR *szKeyPath, const TCHAR *szValueName, REGSAM sam, bool *pBool)
{
	DWORD dwData = 0;

	if (!LeetRegQueryDWORDValue(hKeyParent, szKeyPath, szValueName, sam, &dwData))
	{
		return false;
	}

	if (NULL != pBool)
	{
		*pBool = (dwData ? true : false);
	}

	return true;
}

// Returns false on failure. Use GetLastError().
// Uses LeetRegQueryValue() to get you a TCHAR string without having to worry about buffers, types and so on.
bool LeoHelpers::LeetRegQueryStringValue(HKEY hKeyParent, const TCHAR *szKeyPath, const TCHAR *szValueName, REGSAM sam, std::basic_string<TCHAR> *pString)
{
	DWORD dwErrToSet = ERROR_SUCCESS;

	DWORD dwType;
	void *pData;

	pString->erase();

	if (! LeetRegQueryValue(hKeyParent, szKeyPath, szValueName, sam, &dwType, &pData, NULL) )
	{
		dwErrToSet = GetLastError();
	}
	else
	{
		if (REG_SZ != dwType)
		{
			dwErrToSet = ERROR_INVALID_DATA;
		}
		else
		{
			(*pString) = reinterpret_cast<TCHAR*>(pData);
		}

		delete [] pData;
	}

	if (ERROR_SUCCESS != dwErrToSet)
	{
		SetLastError(dwErrToSet);
	}

	return(ERROR_SUCCESS == dwErrToSet);
}

// Returns false on failure. Use GetLastError().
// Uses LeetRegQueryValue() to get you a vector of TCHAR strings without having to worry about buffers, types and so on.
bool LeoHelpers::LeetRegQueryMultiStringValue(HKEY hKeyParent, const TCHAR *szKeyPath, const TCHAR *szValueName, REGSAM sam, std::vector< std::basic_string<TCHAR> > *pVecStrings)
{
	DWORD dwErrToSet = ERROR_SUCCESS;

	DWORD dwType;
	void *pData;

	pVecStrings->clear();

	if (! LeetRegQueryValue(hKeyParent, szKeyPath, szValueName, sam, &dwType, &pData, NULL) )
	{
		dwErrToSet = GetLastError();
	}
	else
	{
		if (REG_MULTI_SZ != dwType)
		{
			dwErrToSet = ERROR_INVALID_DATA;
		}
		else
		{
			const TCHAR *msz = reinterpret_cast<TCHAR*>(pData);

			for ( std::basic_string<TCHAR> strLine = msz; !strLine.empty(); strLine = ( msz = (msz + strLine.length() + 1) ) )
			{
				pVecStrings->push_back( strLine );
			}
		}

		delete [] pData;
	}

	if (ERROR_SUCCESS != dwErrToSet)
	{
		SetLastError(dwErrToSet);
	}

	return(ERROR_SUCCESS == dwErrToSet);
}

bool LeoHelpers::LeetRegEnumKey(HKEY hKeyParent, const TCHAR *szKeyPath, const TCHAR *szPrefixFilter, REGSAM sam, std::vector< std::basic_string<TCHAR> > *pVecStrings)
{
	// Documentation doesn't state that registry functions call SetLastError().
	// We need to preserve the error values they return through clean-up anyway.
	LONG regRes = ERROR_SUCCESS;

	pVecStrings->clear();

	HKEY hKeyToEnum;

	// szKeyPath may be NULL in which case another handle like hKeyParent is openned (and must be closed).
	regRes = RegOpenKeyEx(hKeyParent, szKeyPath, 0, sam | KEY_ENUMERATE_SUB_KEYS, &hKeyToEnum);

	if (ERROR_SUCCESS == regRes)
	{
		FILETIME ftLastWriteTime;	// Don't care but documentation doesn't say it can be null.

		// Max key-name   length is 255   chars according to MSDN. (Not sure if that includes the null.) Doesn't hurt to work with more.
		// Max value-name length is 16383 bytes according to MSDN. (Not sure if that includes the null.) Doesn't hurt to work with more.
		TCHAR szNameBuffer[1024];
		const DWORD dwNameBufferSize = (sizeof(szNameBuffer)/sizeof(szNameBuffer[0]));
		DWORD dwcName;

		size_t cPrefixFilter = (NULL != szPrefixFilter) ? _tcslen(szPrefixFilter) : 0;

		for(DWORD dwIndex = 0; true; dwIndex++)
		{
			dwcName = dwNameBufferSize;
			regRes = RegEnumKeyEx(hKeyToEnum, dwIndex, szNameBuffer, &dwcName, NULL, NULL, NULL, &ftLastWriteTime);

			if (ERROR_SUCCESS == regRes && dwcName >= dwNameBufferSize)
			{
				regRes = ERROR_INSUFFICIENT_BUFFER;
				break;
			}
			else if (ERROR_NO_MORE_ITEMS == regRes)
			{
				regRes = ERROR_SUCCESS;
				break;
			}
			else if (ERROR_SUCCESS != regRes)
			{
				break;
			}
			else
			{
				if (0 == cPrefixFilter || 0 == _tcsnicmp(szNameBuffer, szPrefixFilter, cPrefixFilter))
				{
					pVecStrings->push_back( szNameBuffer + cPrefixFilter );
				}
			}
		}

		LONG closeRes = RegCloseKey(hKeyToEnum);
		if (ERROR_SUCCESS == regRes) { regRes = closeRes; }
	}

	if (ERROR_SUCCESS != regRes)
	{
		SetLastError(regRes);
	}

	return(ERROR_SUCCESS == regRes);
}

bool LeoHelpers::LeetRegEnumValue(HKEY hKeyParent, const TCHAR *szKeyPath, REGSAM sam, std::vector< std::basic_string<TCHAR> > *pVecStrings)
{
	// Documentation doesn't state that registry functions call SetLastError().
	// We need to preserve the error values they return through clean-up anyway.
	LONG regRes = ERROR_SUCCESS;

	pVecStrings->clear();

	HKEY hKeyToEnum;

	// szKeyPath may be NULL in which case another handle like hKeyParent is openned (and must be closed).
	regRes = RegOpenKeyEx(hKeyParent, szKeyPath, 0, sam | KEY_QUERY_VALUE, &hKeyToEnum);

	if (ERROR_SUCCESS == regRes)
	{
		// Max key-name   length is 255   chars according to MSDN. (Not sure if that includes the null.) Doesn't hurt to work with more.
		// Max value-name length is 16383 chars according to MSDN. (Not sure if that includes the null.) Doesn't hurt to work with more.
		// ANSI version of RegEnumValue will go haywire if we give a size larger than 32767 chars/bytes.
		const DWORD dwNameBufferSize = 32767;
		TCHAR *szNameBuffer = new TCHAR[dwNameBufferSize];
		DWORD dwcName;

		for(DWORD dwIndex = 0; true; dwIndex++)
		{
			dwcName = dwNameBufferSize;
			regRes = RegEnumValue(hKeyToEnum, dwIndex, szNameBuffer, &dwcName, NULL, NULL, NULL, NULL);

			if (ERROR_SUCCESS == regRes && dwcName >= dwNameBufferSize)
			{
				regRes = ERROR_INSUFFICIENT_BUFFER;
				break;
			}
			else if (ERROR_NO_MORE_ITEMS == regRes)
			{
				regRes = ERROR_SUCCESS;
				break;
			}
			else if (ERROR_SUCCESS != regRes)
			{
				break;
			}
			else
			{
				pVecStrings->push_back( szNameBuffer );
			}
		}

		delete [] szNameBuffer;
		szNameBuffer = NULL;

		LONG closeRes = RegCloseKey(hKeyToEnum);
		if (ERROR_SUCCESS == regRes) { regRes = closeRes; }
	}

	if (ERROR_SUCCESS != regRes)
	{
		SetLastError(regRes);
	}

	return(ERROR_SUCCESS == regRes);
}

bool LeoHelpers::LeetRegSetValue(HKEY hKeyParent, const TCHAR *szKeyPath, const TCHAR *szValueName,
									 DWORD dwType, const void *lpData, DWORD cbData)
{
	// Documentation doesn't state that registry functions call SetLastError().
	// We need to preserve the error values they return through clean-up anyway.
	LONG regRes = ERROR_SUCCESS;

	bool bCloseKey = false;
	HKEY hKeyToCreateValueIn;

	if ((NULL == szKeyPath) || (0 == _tcsicmp(szKeyPath, _T(""))))
	{
		// Cannot give NULL as second argument to RegCreateKeyEx() as you can to RegOpenKeyEx()
		hKeyToCreateValueIn = hKeyParent;
	}
	else
	{
		DWORD disposition;

		// We don't need to create the key one level at a time, RegCreateKeyEx() does that.
		regRes = RegCreateKeyEx(hKeyParent, szKeyPath, 0, _T(""),
								REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, NULL,
								&hKeyToCreateValueIn, &disposition);

		if (ERROR_SUCCESS == regRes)
		{
			bCloseKey = true;
		}
	}

	if ((ERROR_SUCCESS == regRes) && (NULL != lpData))
	{
		regRes = RegSetValueEx(hKeyToCreateValueIn, szValueName, 0, dwType, reinterpret_cast<CONST BYTE *>(lpData), cbData);
	}

	if (bCloseKey)
	{
		LONG closeRes = RegCloseKey(hKeyToCreateValueIn);
		if (ERROR_SUCCESS == regRes) { regRes = closeRes; }

		bCloseKey = false;
	}

	if (ERROR_SUCCESS != regRes)
	{
		SetLastError(regRes);
	}

	return(ERROR_SUCCESS == regRes);
}

bool LeoHelpers::LeetRegSetDWORDValue(HKEY hKeyParent, const TCHAR *szKeyPath, const TCHAR *szValueName, DWORD dwData)
{
	return(LeetRegSetValue(hKeyParent, szKeyPath, szValueName, REG_DWORD, &dwData, sizeof(DWORD)));
}

bool LeoHelpers::LeetRegSetIntValue(HKEY hKeyParent, const TCHAR *szKeyPath, const TCHAR *szValueName, int iData)
{
	if (sizeof(int)==sizeof(DWORD))
	{
		DWORD dwData = *reinterpret_cast<DWORD*>(&iData);

		return(LeetRegSetValue(hKeyParent, szKeyPath, szValueName, REG_DWORD, &dwData, sizeof(DWORD)));
	}

	return false;
}

bool LeoHelpers::LeetRegSetBoolValue(HKEY hKeyParent, const TCHAR *szKeyPath, const TCHAR *szValueName, bool bData)
{
	DWORD dwData = bData ? 1 : 0;

	return(LeetRegSetValue(hKeyParent, szKeyPath, szValueName, REG_DWORD, &dwData, sizeof(DWORD)));
}

bool LeoHelpers::LeetRegSetStringValue(HKEY hKeyParent, const TCHAR *szKeyPath, const TCHAR *szValueName, const TCHAR *szData)
{
	return(LeetRegSetValue(hKeyParent, szKeyPath, szValueName, REG_SZ, szData, static_cast<DWORD>((_tcslen(szData)+1)*sizeof(szData[0]))));
}

bool LeoHelpers::LeetRegSetMultiStringValue(HKEY hKeyParent, const TCHAR *szKeyPath, const TCHAR *szValueName, const std::vector< std::basic_string<TCHAR> > *pVecStrings)
{
	bool bResult = false;

	size_t dwBufferSizeChars = 0;
	TCHAR *mszBuffer = MultiStringFromVector(pVecStrings, &dwBufferSizeChars);

	if (NULL != mszBuffer)
	{
		bResult = LeetRegSetValue(hKeyParent, szKeyPath, szValueName, REG_MULTI_SZ, mszBuffer, static_cast<DWORD>(dwBufferSizeChars * sizeof(TCHAR)));

		delete [] mszBuffer;
	}

	return(bResult);
}

bool LeoHelpers::LeetRegCreateKey(HKEY hKeyParent, const TCHAR *szKeyPath)
{
	// Documentation doesn't state that registry functions call SetLastError().
	// We need to preserve the error values they return through clean-up anyway.
	LONG regRes = ERROR_SUCCESS;

	HKEY hKeyResult = NULL;

	regRes = RegCreateKeyEx(hKeyParent, szKeyPath, 0, NULL, REG_OPTION_NON_VOLATILE, 0, NULL, &hKeyResult, NULL);

	if (ERROR_SUCCESS == regRes)
	{
		LONG closeRes = RegCloseKey(hKeyResult);
		if (ERROR_SUCCESS == regRes) { regRes = closeRes; }
	}

	if (ERROR_SUCCESS != regRes)
	{
		SetLastError(regRes);
	}

	return(ERROR_SUCCESS == regRes);
}

bool LeoHelpers::LeetRegDeleteKey(HKEY hKeyParent, const TCHAR *szKeyPath, const TCHAR *szKeyToDelete)
{
	// Documentation doesn't state that registry functions call SetLastError().
	// We need to preserve the error values they return through clean-up anyway.
	LONG regRes = ERROR_SUCCESS;

	HKEY hKeyToDeleteUnder;

	// szKeyPath may be NULL in which case another handle like hKeyParent is openned (and must be closed).
	regRes = RegOpenKeyEx(hKeyParent, szKeyPath, 0, DELETE, &hKeyToDeleteUnder);

	if (ERROR_SUCCESS == regRes)
	{
		regRes = RegDeleteKey(hKeyToDeleteUnder, szKeyToDelete);

		LONG closeRes = RegCloseKey(hKeyToDeleteUnder);
		if (ERROR_SUCCESS == regRes) { regRes = closeRes; }
	}

	if (ERROR_SUCCESS != regRes)
	{
		SetLastError(regRes);
	}

	return(ERROR_SUCCESS == regRes);
}

// Returns false on failure. Use GetLastError().
bool LeoHelpers::LeetRegDeleteValue(HKEY hKeyParent, const TCHAR *szKeyPath, const TCHAR *szValueName)
{
	// Documentation doesn't state that registry functions call SetLastError().
	// We need to preserve the error values they return through clean-up anyway.
	LONG regRes = ERROR_SUCCESS;

	HKEY hKeyToDeleteUnder;

	// szKeyPath may be NULL in which case another handle like hKeyParent is openned (and must be closed).
	regRes = RegOpenKeyEx(hKeyParent, szKeyPath, 0, KEY_SET_VALUE, &hKeyToDeleteUnder);

	if (ERROR_SUCCESS == regRes)
	{
		regRes = RegDeleteValue(hKeyToDeleteUnder, szValueName);

		LONG closeRes = RegCloseKey(hKeyToDeleteUnder);
		if (ERROR_SUCCESS == regRes) { regRes = closeRes; }
	}

	if (ERROR_SUCCESS != regRes)
	{
		SetLastError(regRes);
	}

	return(ERROR_SUCCESS == regRes);
}

// We always use the C locale. It's assumed that any case-insensitive string
// comparison, or upper/lower-casing of strings, is done to compare either
// file paths (or extensions), or internal data (e.g. CLSID strings), rather
// than anything which should be locale-dependent. We want consistency across
// machines and executions, not with the user's locale.
const _locale_t LeoHelpers::LocaleC = _create_locale(LC_ALL, "");

std::wstring LeoHelpers::StringTrim(const std::wstring &strIn, const wchar_t *szTrimChars, bool bTrimLeft, bool bTrimRight)
{
	std::wstring::size_type leftOffset  = 0;
	std::wstring::size_type rightOffset = strIn.length();

	if (bTrimLeft)
	{
		leftOffset = strIn.find_first_not_of(szTrimChars);
	}

	if (bTrimRight)
	{
		rightOffset = strIn.find_last_not_of(szTrimChars);

		if (rightOffset != std::wstring::npos)
		{
			++rightOffset;
		}
	}

	if (leftOffset == std::wstring::npos
	||	rightOffset == std::wstring::npos
	||	leftOffset >= rightOffset)
	{
		return L"";
	}

	return strIn.substr(leftOffset, rightOffset - leftOffset);
}

// delete[] the result.
// Returns NULL on failure.
TCHAR *LeoHelpers::VStringAllocAndFormat(const TCHAR *format, va_list args)
{
	// Leo 16/Jan/2009: Instead of defaulting to 64 bytes for the initial buffer size, make it 1.5* the format string length.
	size_t bufferSize = (_tcslen(format) * 3) / 2;

	if (bufferSize < 64)
	{
		bufferSize = 64;
	}

	TCHAR *szResult = new(std::nothrow) TCHAR[bufferSize];

	while(true)
	{
		if (NULL == szResult)
		{
			::SetLastError(ERROR_NOT_ENOUGH_MEMORY);
			break;
		}

		int formatRes = _vsntprintf_s(szResult, bufferSize, _TRUNCATE, format, args);

		// formatRes negative on error; (bufferSize == formatRes) should be redundant for _vsntprintf.

		if ((0 > formatRes) || (bufferSize == formatRes))
		{
			delete [] szResult;
			bufferSize *= 2;
			szResult = new(std::nothrow) TCHAR[bufferSize];
		}
		else
		{
			break;
		}
	}

	return(szResult);
}

// Returns boolean success.
bool LeoHelpers::StringFormatToString(std::basic_string< TCHAR > *pResult, const TCHAR *format, ...)
{
	va_list args;
	va_start(args, format);
	TCHAR *result = VStringAllocAndFormat(format, args);
	va_end(args);

	if (result == NULL)
	{
		pResult->clear();
		return false;
	}

	*pResult = result;
	delete[] result;

	return true;
}

// Returns boolean success.
bool LeoHelpers::VStringFormatToString(std::basic_string< TCHAR > *pResult, const TCHAR *format, va_list args)
{
	TCHAR *result = VStringAllocAndFormat(format, args);

	if (result == NULL)
	{
		pResult->clear();
		return false;
	}

	*pResult = result;
	delete[] result;

	return true;
}

TCHAR *LeoHelpers::MultiStringFromVector(const std::vector< std::basic_string<TCHAR> > *pVecStrings, size_t *pLengthChars)
{
	if (NULL != pLengthChars)
	{
		*pLengthChars = 0;
	}

	size_t bufferSize = 1; // Final null character.

	for (std::vector< std::basic_string<TCHAR> >::const_iterator pStr = pVecStrings->begin(); pStr != pVecStrings->end(); ++pStr)
	{
		if (!pStr->empty())
		{
			bufferSize += (pStr->length() + 1);
		}
	}

	TCHAR *mszBuffer = new(std::nothrow) TCHAR[ bufferSize ];

	if (NULL == mszBuffer)
	{
		::SetLastError(ERROR_NOT_ENOUGH_MEMORY);
		return NULL;
	}

	size_t bufferLeft = bufferSize;
	size_t partLen;

	TCHAR *szCurrent = mszBuffer;

	for (std::vector< std::basic_string<TCHAR> >::const_iterator pStr = pVecStrings->begin(); pStr != pVecStrings->end(); ++pStr)
	{
		if (!pStr->empty())
		{
			if (0 == _tcscpy_s(szCurrent, bufferLeft, pStr->c_str()))
			{
				partLen = (pStr->length() + 1);
				szCurrent += partLen;
				bufferLeft -= partLen;
			}
		}
	}

	szCurrent[0] = _T('\0'); // Final null character.

	if (NULL != pLengthChars)
	{
		*pLengthChars = bufferSize;
	}

	return mszBuffer;
}

// szInput -- String to parse.
// delimiters -- Characters which delimit the string.
// terminators -- Characters which end the string. Can be NULL.
// results -- A string list to put the results into. Existing contents will be erased. Give NULL if you don't want any results.
// bSkipEmpty -- Set true to skip any empty strings. (e.g. someone typed "blah,,blah")
// bQuotes -- Set to true to handle double-quotes (") around items. A quoted empty string is NOT skipped even if bSkipEmpty is set. Terminators in quotes are ignored and included in the tokenized strings.
// bCSVQuotes -- (Ignored if bQuotes false) -- Inside quoted strings, two double-quotes next to each other will be output as one double-quote and not signify the end of the string. A string like "Blah""blah" will become blah"blah.
// Leading and trailing spaces are always stripped.
// Returns false iff nothing could be extracted (e.g. szInput was NULL)
bool LeoHelpers::Tokenize(const TCHAR *szInput, const TCHAR *delimiters, const TCHAR *terminators, std::list< std::basic_string< TCHAR > > *pResults, bool bSkipEmpty, bool bQuotes, bool bCSVQuotes)
{
	if (NULL != pResults)
	{
		pResults->clear();
	}

	bool bResult = false;

	size_t inLen = _tcslen(szInput);
	TCHAR *szBuff = new TCHAR[inLen];

	while (LeoHelpers::Tokenize(&szInput, delimiters, terminators, szBuff, inLen, bSkipEmpty, bQuotes, bCSVQuotes))
	{
		bResult = true;

		if (NULL != pResults)
		{
			pResults->push_back(szBuff);
		}
	}

	delete[] szBuff;

	return bResult;
}

// pp -- Pointer to buffer to tokenize. Will be advanced to the next position string. Set to NULL after the last token is extracted.
// delimiters -- Characters which delimit the string.
// terminators -- Characters which end the string. Can be NULL.
// bSkipEmpty -- Set true to skip any empty strings. (e.g. someone typed "blah,,blah")
// bQuotes -- Set to true to handle double-quotes (") around items. A quoted empty string is NOT skipped even if bSkipEmpty is set. Terminators in quotes are ignored and included in the tokenized strings.
// bCSVQuotes -- (Ignored if bQuotes false) -- Inside quoted strings, two double-quotes next to each other will be output as one double-quote and not signify the end of the string. A string like "Blah""blah" will become blah"blah.
// buf -- Buffer to extract in to. Set to NULL not to extract.
// bufSize -- Size of buf if it is non-NULL. Will truncate string if buf too small.
// Leading and trailing spaces are always stripped.
// Returns false iff nothing could be extracted (i.e. *pp was NULL)
// pp and the contents of buf are only valid after calling if true is returns.
// Lame function should allocate us a suitably sized buffer for each string... :-)
// But hey, this isn't a public OS function and you can just give a buffer the size of the input...
bool LeoHelpers::Tokenize(const TCHAR **pp, const TCHAR *delimiters, const TCHAR *terminators, TCHAR *buf, size_t bufSize, bool bSkipEmpty, bool bQuotes, bool bCSVQuotes)
{
	bool bResult = false;

	if (NULL == terminators)
	{
		terminators = _T("");
	}

	if (NULL != *pp)
	{
		// Remove leading spaces.
		// Also, skip leading delimiters if we're to skip empties.
		// (Important that we skip the spaces here so that they're not included if it's the first string.)
		// (Note we must explicitly check for the terminator before checking the delimiters else we ask if
		// the terminator is in the delimiters string.)
		while ( _istspace(**pp) || ((_T('\0') != **pp) && (NULL == _tcschr(terminators, **pp)) && bSkipEmpty && (NULL != _tcschr(delimiters, **pp))) )
		{
			(*pp)++;
		}

		if (bSkipEmpty && ( (_T('\0') == **pp) || (NULL != _tcschr(terminators, **pp)) ))
		{
			// Last item is empty and we're skipping empties.
			*pp = NULL;
		}
		else
		{
			TCHAR *bufOrig = buf;

			bResult = true; // We got somethin'!

			bool bFindQuote = false;

			if (bQuotes && (_T('"') == **pp))
			{
				(*pp)++; // Skip over the starting quote.
				bFindQuote = true; // Copy string until the ending quote rather than a delimiter.
			}

			// There is some string left, extract what's next.
			// Ignore terminators and delmiters if we're looking for the end-quote.
			while (_T('\0') != **pp)
			{
				if (bFindQuote)
				{
					if (_T('"') == (*pp)[0])
					{
						if (bCSVQuotes && (_T('"') == (*pp)[1]))
						{
							(*pp)++; // It is a pair of quotes, skip one and copy the other.
						}
						else
						{
							break; // It is a closing quote.
						}
					}
				}
				else if ((NULL != _tcschr(delimiters,  **pp))
				||		 (NULL != _tcschr(terminators, **pp)))
				{
					// Not looking for quotes and found a delimiter or a terminator
					break;
				}

				if ((NULL != buf) && (bufSize > 1)) // Keep 1 character for the terminator.
				{
					*buf = **pp;
					bufSize--;
					buf++;
				}

				(*pp)++; // Must increment *pp here as copying to buf is optional and may also be trunctated.
			}

			// Handles the case where a zero-length buffer is passed in.
			if ((NULL != buf) && (bufSize > 0))
			{
				*buf = _T('\0'); // NULL-terminate it.

				// Remove trailing spaces if not quoted.
				if (! bFindQuote )
				{
					while (buf > bufOrig)
					{
						if (_istspace(*(--buf)))
						{
							*buf = _T('\0');
						}
						else
						{
							break;
						}
					}
				}
			}

			// If quoted, move over the closing quote and skip up to the next delimiter.
			if (bFindQuote && (_T('\0') != **pp) && (NULL == _tcschr(terminators, **pp)))
			{
				do
				{
					(*pp)++;
				} while ((_T('\0') != **pp) && (NULL == _tcschr(terminators, **pp)) && (NULL == _tcschr(delimiters, **pp)));
			}

			// Move over the delimiter (unless we're at the end of the string).
			if ((_T('\0') != **pp) && (NULL == _tcschr(terminators, **pp)) && (NULL != _tcschr(delimiters, **pp)))
			{
				(*pp)++;

				// Skip spaces before the next token. (Useful for when buf==NULL)
				while ((_T('\0') != **pp) && (NULL == _tcschr(terminators, **pp)) && _istspace(**pp))
				{
					(*pp)++;
				}
			}
			else
			{
				// We're returning the last thing on the line.
				*pp = NULL;
			}
		}
	}

	return(bResult);
}

/* Code to test how the MB/WC APIs behave:
	wchar_t szWide[16];

	char *aszTests[] =
	{
		"",
		"1",
		"12",
		"123",
		"1234",
		"12345",
		"123456",
		"1234567",
		"12345678",
		"123456789",
		"123456789A",
		"123456789AB",
		"123456789ABC"

	};

	for(int m = 0; m < 4; ++m)
	{
		if (m&1)
		{
			if (m&2)
			{
				printf("Length-specified strings (excluding null in length), and request required size:\n");
			}
			else
			{
				printf("Length-specified strings (excluding null in length):\n");
			}
		}
		else
		{
			if (m&2)
			{
				printf("Null terminated strings, and request required size:\n");
			}
			else
			{
				printf("Null terminated strings:\n");
			}
		}

		for(int t = 0; t < _countof(aszTests); ++t)
		{
			for(int i = 0; i < _countof(szWide); ++i)
			{
				szWide[i] = (i < 9 ? L'x' : L'!');
			}
			szWide[_countof(szWide)-1] = L'\0';

			// Change the +0 to a +1 to see how it behaves when the length is specified but does include the null. (Then it's the same as using -1 as the length.)
			int res = ::MultiByteToWideChar(CP_UTF8, 0, aszTests[t], (m&1) ? strlen(aszTests[t]) +  0 : -1, szWide, (m&2) ? 0 : 10);

			int nullPos;

			for (nullPos = 0; nullPos < _countof(szWide); ++nullPos)
			{
				if (szWide[nullPos] == L'\0')
				{
					break;
				}
			}

			if (nullPos == _countof(szWide) - 1)
			{
				printf("%12s - res = %2d, nullPos = NONE - %S\n", aszTests[t], res, szWide);
			}
			else if (nullPos > 9)
			{
				printf("%12s - res = %2d, nullPos = %4d (ILLEGAL) - %S\n", aszTests[t], res, nullPos, szWide);
			}
			else
			{
				printf("%12s - res = %2d, nullPos = %4d - %S\n", aszTests[t], res, nullPos, szWide);
			}
		}

		printf("\n");
	}
*/


// Returns 0 on failure (use GetLastError()), else the number of wide characters written including the null.
// delete [] *pwcStringResult when finished with it.
// numInputBytes -1 for null-terminated string.
int LeoHelpers::MBtoWC(WCHAR **pwcStringResult, const char *mbStringInput, int numInputBytes, DWORD dwCodePage)
{
	*pwcStringResult = NULL;

	if (numInputBytes == 0)
	{
		*pwcStringResult = new(std::nothrow) WCHAR[1];

		if (NULL == *pwcStringResult)
		{
			::SetLastError(ERROR_NOT_ENOUGH_MEMORY);
			return 0;
		}

		(*pwcStringResult)[0] = L'\0';
		return 1;
	}

	int wcBufSize = ::MultiByteToWideChar(dwCodePage, 0, mbStringInput, numInputBytes, NULL, 0);

	if (wcBufSize == 0)
	{
		return 0;
	}

	if (numInputBytes != -1 && mbStringInput[numInputBytes-1] != '\0')
	{
		// If the string length is specified, and it doesn't contain a null, then MultiByteToWideChar will not include space for the null.
		// We'll add and allocate room for that extra null, whether the caller requires it or not. If the caller doesn't then it's harmless.
		++wcBufSize;
	}

	*pwcStringResult = new(std::nothrow) WCHAR[wcBufSize];

	if (NULL == *pwcStringResult)
	{
		::SetLastError(ERROR_NOT_ENOUGH_MEMORY);
		return 0;
	}

	int convRes = ::MultiByteToWideChar(dwCodePage, 0, mbStringInput, numInputBytes, *pwcStringResult, wcBufSize);

	(*pwcStringResult)[wcBufSize-1] = L'\0'; // No matter what, the buffer is terminated.

	if (convRes <= 0 || convRes > wcBufSize)
	{
		assert(convRes==0); // Real WTF from the API if it isn't zero.

		delete[] *pwcStringResult;
		*pwcStringResult = NULL;

		return 0;
	}

	if (numInputBytes != -1 && mbStringInput[numInputBytes-1] != '\0' && convRes < wcBufSize)
	{
		++convRes; // Bump for the null that we wrote, which was not written by the API or included in its count.
	}

	// Finally, remove any left-surrogates from the end of the string.

	while(convRes > 1 && LeoHelpers::IsUtf16LeftSurrogate((*pwcStringResult)[convRes-2])) // -1 is the current null; -2 is the last character.
	{
		(*pwcStringResult)[convRes-2] = L'\0';
		--convRes;
	}

	// Ideally we'd ensure three are no partial multi-byte chars at the end of the string but I can't see a good way to do this
	// for the same dwCodePage argument which WideCharToMultiByte takes. CharPrevExA/CharNextExA are almost what we want, but not quite.

	return convRes;
}

// Returns 0 on failure (use GetLastError()), else the number of characters written.
// delete [] *pmbStringResult when finished with it.
// numInputWords -1 for null-terminated string.
int LeoHelpers::WCtoMB(char **pmbStringResult, const WCHAR *wcStringInput, int numInputWords, DWORD dwCodePage)
{
	*pmbStringResult = NULL;

	if (numInputWords == 0)
	{
		*pmbStringResult = new(std::nothrow) char[1];

		if (NULL == *pmbStringResult)
		{
			::SetLastError(ERROR_NOT_ENOUGH_MEMORY);
			return 0;
		}

		(*pmbStringResult)[0] = '\0';
		return 1;
	}

	int mbBufSize = ::WideCharToMultiByte(dwCodePage, 0, wcStringInput, numInputWords, NULL, 0, NULL, NULL);

	if (mbBufSize == 0)
	{
		return 0;
	}

	if (numInputWords != -1 && wcStringInput[numInputWords-1] != L'\0')
	{
		// If the string length is specified, and it doesn't contain a null, then WideCharToMultiByte will not include space for the null.
		// We'll add and allocate room for that extra null, whether the caller requires it or not. If the caller doesn't then it's harmless.
		++mbBufSize;
	}

	*pmbStringResult = new(std::nothrow) char[mbBufSize];

	if (NULL == *pmbStringResult)
	{
		::SetLastError(ERROR_NOT_ENOUGH_MEMORY);
		return 0;
	}

	int convRes = ::WideCharToMultiByte(dwCodePage, 0, wcStringInput, numInputWords, *pmbStringResult, mbBufSize, NULL, NULL);

	(*pmbStringResult)[mbBufSize-1] = '\0'; // No matter what, the buffer is terminated.

	if (convRes <= 0 || convRes > mbBufSize)
	{
		assert(convRes==0); // Real WTF from the API if it isn't zero.

		delete[] (*pmbStringResult);
		*pmbStringResult = NULL;

		return 0;
	}

	if (numInputWords != -1 && wcStringInput[numInputWords-1] != L'\0' && convRes < mbBufSize)
	{
		++convRes; // Bump for the null tht we wrote, which was not written by the API or included in its count.
	}

	return(convRes);
}

// Returns the new null position.
size_t LeoHelpers::TruncateString(wchar_t *szBuffer, size_t idxNull, size_t cchSpaceAfterNull, bool bTruncateWithElipsis)
{
	size_t ec = 0; // This is how many spaces for dots we want (but may not get if the buffer is too small).
	
	if (bTruncateWithElipsis && cchSpaceAfterNull < 3)
	{
		ec = 3 - cchSpaceAfterNull; // How much of the string we need to eat in order to fit three dots before the null.
	}

	cchSpaceAfterNull = 0; // nothing else should use it.

	while(idxNull > 0 && (ec > 0 || LeoHelpers::IsUtf16LeftSurrogate(szBuffer[idxNull-1])))
	{
		if (ec > 0)
		{
			--ec;
		}

		szBuffer[--idxNull] = L'\0';
	}

	assert(szBuffer[idxNull] == L'\0');

	if (bTruncateWithElipsis)
	{
		// Add up to three '.' in whatever space we've made.
		while(ec < 3)
		{
			szBuffer[  idxNull] = L'.';
			szBuffer[++idxNull] = L'\0';
			++ec;
		}
	}

	assert(szBuffer[idxNull] == L'\0');

	return idxNull;
}

bool LeoHelpers::MBtoWCInPlaceTruncate(const char *szSource, int cbInputBytes, wchar_t *szDestBuffer, int cchBufferSize, DWORD dwCodePage, bool bTruncateWithElipsis, bool *pbWasTruncated)
{
	if (pbWasTruncated != NULL) { *pbWasTruncated = false; }

	assert(cchBufferSize > 0);

	if (cchBufferSize < 1)
	{
		if (pbWasTruncated != NULL) { *pbWasTruncated = true; }
		return false;
	}

	if (cbInputBytes == 0)
	{
		szDestBuffer[0] = L'\0';
		return true;
	}

	int fres = ::MultiByteToWideChar(dwCodePage, 0, szSource, cbInputBytes, szDestBuffer, cchBufferSize);

	szDestBuffer[cchBufferSize-1] = L'\0'; // Whatever happened, it's null terminated.

	if (cbInputBytes != -1 && szSource[cbInputBytes-1] != '\0' && fres > 0 && fres <= cchBufferSize)
	{
		if (fres < cchBufferSize)
		{
			szDestBuffer[fres] = L'\0'; // We need to add a null to the string, and there is room for it.
			++fres; // Account for the fact we added it.
		}
		else
		{
			// The buffer has no null, and no room for the null. Turn this success into a failure so it is truncated.
			fres = 0;
			::SetLastError(ERROR_INSUFFICIENT_BUFFER);
		}
	}

	if (fres > 0 && fres <= cchBufferSize)
	{
		// Remove any left-surrogates from the end of the string.
		while(fres > 1 && LeoHelpers::IsUtf16LeftSurrogate(szDestBuffer[fres-2])) // -1 is the current null; -2 is the last character.
		{
			szDestBuffer[fres-2] = L'\0';
			--fres;
		}

		return true;
	}
	else if (fres == 0)
	{
		if (::GetLastError() != ERROR_INSUFFICIENT_BUFFER)
		{
			// Failed for some other reason. Who knows what's in the buffer. Nuke it, return failure, thank-you-drive-through.
			szDestBuffer[0] = L'\0';
			return false;
		}
		else
		{
			if (pbWasTruncated != NULL) { *pbWasTruncated = true; }

			// Assumption: MultiByteToWideChar filled the buffer as much as it could. (This isn't actually documented -- FFS! -- but is what happens.)
			// The buffer will not be null-terminated yet. We also need to remove any partial surrogates and, if requested, add the elpisis (if it'll fit).
			// Making room for the elipsis may create an orphaned surrogate as well.

			fres = cchBufferSize; // This is how many characters were written. We know this is at least 1. No null has been written yet.

			szDestBuffer[--fres] = L'\0'; // The buffer is now null terminated. fres may now be zero.

			LeoHelpers::TruncateString(szDestBuffer, fres, 0, bTruncateWithElipsis);

			return true;
		}
	}
	else
	{
		assert(false); // Getting here would be a real WTF on the API's part.
		szDestBuffer[0] = L'\0';
		return false;
	}
}

bool LeoHelpers::GetShortFilePath(std::wstring *pstrResult, const wchar_t *szInput)
{
	bool bResult = false;

	wchar_t szDummy[1]; // Just in case GetShortPath is stupid.

	DWORD dwRes = ::GetShortPathName(szInput, szDummy, 0);

	if (dwRes > 0)
	{
		wchar_t *szBuf = new(std::nothrow) wchar_t[(size_t)dwRes + 1];

		if (szBuf)
		{
			DWORD dwRes2 = ::GetShortPathName(szInput, szBuf, dwRes);

			if (dwRes2 <= dwRes)
			{
				*pstrResult = szBuf;
				bResult = true;
			}

			delete[] szBuf;
		}
	}

	return bResult;
}

#ifdef UrlCreateFromPath
bool LeoHelpers::UrlFromPath(std::wstring *pstrResult, const wchar_t *szPath)
{
	bool bResult = false;

	DWORD dwUrlBufferLen = 5120;
	wchar_t *szUrlBuffer = new wchar_t[dwUrlBufferLen];

	if (SUCCEEDED(UrlCreateFromPath(szPath, szUrlBuffer, &dwUrlBufferLen, NULL)))
	{
		*pstrResult = szUrlBuffer;
		bResult = true;
	}
	else
	{
		pstrResult->clear();
	}

	delete[] szUrlBuffer;

	return bResult;
}
#endif

#ifdef UrlCreateFromPath
bool LeoHelpers::UrlFromPathShortIfPossible(std::wstring *pstrResult, const wchar_t *szPath, bool bOnlyTryShortIfNonAscii)
{
	std::wstring strPathToConvert;

	if ((bOnlyTryShortIfNonAscii && LeoHelpers::IsAscii(szPath))
	||	!LeoHelpers::GetShortFilePath(&strPathToConvert, szPath))
	{
		strPathToConvert = szPath;
	}

	return LeoHelpers::UrlFromPath(pstrResult, strPathToConvert.c_str());
}
#endif

/*
void LeoHelpers::EncodeUri(std::wstring *pstrResult, const wchar_t *szInput, bool bAppend, bool bConvertColon, bool bProperUTF8Method)
{
	if (!bAppend)
	{
		pstrResult->clear();
	}

	// See http://en.wikipedia.org/wiki/Percent-encoding
	// We encode the "reserved" characters as well since if they are part of a filename then they are not
	// being used for their "reserved" purpose and thus must be encoded.
	// Note that there is another copy of this table below.
	const static char *szNoEscapeSortedAnsi = "-.0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ_abcdefghijklmnopqrstuvwxyz~";
	const static size_t nesLenAnsi = strlen(szNoEscapeSortedAnsi);

	const static wchar_t *szNoEscapeSortedWide = L"-.0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ_abcdefghijklmnopqrstuvwxyz~";
	const static size_t nesLenWide = wcslen(szNoEscapeSortedWide);

#ifdef _DEBUG
	assert(nesLenAnsi == nesLenWide);

	for (size_t i = 1; i < nesLenAnsi; ++i)
	{
		assert(szNoEscapeSortedAnsi[i - 1] < szNoEscapeSortedAnsi[i]);
	}

	for (size_t i = 1; i < nesLenWide; ++i)
	{
		assert(szNoEscapeSortedWide[i - 1] < szNoEscapeSortedWide[i]);
	}
#endif

	if (bProperUTF8Method)
	{
		char szNumBuf[8];

		char *szUtf8Input = NULL;

		if (0 == WCtoMB(&szUtf8Input, szInput, -1, CP_UTF8))
		{
			*pstrResult = szInput; // Fail.
		}
		else
		{
			std::string strAnsiResult;

			const size_t slen = strlen(szUtf8Input);

			for(size_t i = 0; i < slen; ++i)
			{
				const char c = szUtf8Input[i];

				if (c == '\\')
				{
					strAnsiResult.append(1, '/'); // Convert backslashes to forwardslashes.
				}
				else if (bConvertColon && c == ':')
				{
					strAnsiResult.append(1, '|');
				}
				else if (c == ':' || std::binary_search(szNoEscapeSortedAnsi, szNoEscapeSortedAnsi + nesLenAnsi, c))
				{
					strAnsiResult.append(1, c);
				}
				else
				{
					unsigned char uc = c;

					_snprintf_s(szNumBuf, _countof(szNumBuf), _TRUNCATE, "%%%02X", static_cast<int>(uc));
					strAnsiResult.append(szNumBuf);
				}
			}

			wchar_t *wszResult = NULL;

			if (0 == MBtoWC(&wszResult, strAnsiResult.c_str(), -1, CP_UTF8))
			{
				*pstrResult = szInput; // Fail.
			}
			else
			{
				pstrResult->append(wszResult);

				delete [] wszResult;
			}

			delete[] szUtf8Input;
		}
	}
	else
	{
		wchar_t szNumBuf[8];

		const size_t slen = wcslen(szInput);

		for(size_t i = 0; i < slen; ++i)
		{
			const wchar_t c = szInput[i];

			if (c == L'\\')
			{
				pstrResult->append(1, L'/'); // Convert backslashes to forwardslashes.
			}
			else if (bConvertColon && c == L':')
			{
				pstrResult->append(1, L'|');
			}
			else if (c == L':' || std::binary_search(szNoEscapeSortedWide, szNoEscapeSortedWide + nesLenWide, c))
			{
				pstrResult->append(1, c);
			}
			else
			{
				_snwprintf_s(szNumBuf, _countof(szNumBuf), _TRUNCATE, (c <= 0xFF) ? L"%%%02hX" : L"%%%04hX", c);
				pstrResult->append(szNumBuf);
			}
		}
	}
}
*/

std::basic_string< TCHAR > *LeoHelpers::GetParentPathString(std::basic_string< TCHAR > *pstrResult, const TCHAR *path)
{
	assert(NULL != pstrResult && NULL != path);

	if (NULL == pstrResult || NULL == path)
	{
		return NULL;
	}

	const TCHAR *lastBackslash  = _tcsrchr(path, _T('\\'));
	const TCHAR *lastFrontslash = _tcsrchr(path, _T('/') );
	const TCHAR *lastSlash = NULL;

	if ((NULL != lastBackslash) && (NULL != lastFrontslash))
	{
		if (lastBackslash > lastFrontslash)
		{
			lastSlash = lastBackslash;
		}
		else
		{
			lastSlash = lastFrontslash;
		}
	}
	else if (NULL != lastBackslash)
	{
		lastSlash = lastBackslash;
	}
	else if (NULL != lastFrontslash)
	{
		lastSlash = lastFrontslash;
	}

	if (NULL == lastSlash)
	{
		pstrResult->clear();
	}
	else
	{
		std::basic_string< TCHAR >::size_type resLen = (lastSlash - path);

		pstrResult->assign(path, resLen);
	}

	return pstrResult;
}

std::basic_string< TCHAR > *LeoHelpers::AppendPathString(std::basic_string< TCHAR > *pstrPath, const TCHAR *szAppendage)
{
	if (pstrPath == NULL)
	{
		return NULL;
	}

	if (szAppendage != NULL)
	{
		while (szAppendage[0] == _T('\\') || szAppendage[0] == _T('/'))
		{
			++szAppendage;
		}
	}

	if (szAppendage != NULL && szAppendage[0] != _T('\0'))
	{
		if (!pstrPath->empty())
		{
			TCHAR lastChar = (*pstrPath)[pstrPath->length() - 1];

			if (lastChar != _T('\\')
			&&	lastChar != _T('/'))
			{
				pstrPath->append(_T("\\"));
			}
		}

		pstrPath->append(szAppendage);
	}

	return pstrPath;
}

const TCHAR *LeoHelpers::GetLastPathPart(const TCHAR *szIn)
{
	if (NULL == szIn)
	{
		return(szIn);
	}

	size_t len = _tcslen(szIn);

	if (1 >= len)
	{
		return(szIn);
	}

	const TCHAR *szOut = szIn + (len - 1);

	// Ignore any slashes at the very end of the string.
	while (szIn < szOut && (*szOut == _T('\\') || *szOut == _T('/')))
	{
		szOut--;
	}

	// Stop at the next slash we find.
	while (szIn < szOut && (*szOut != _T('\\') && *szOut != _T('/')))
	{
		szOut--;
	}

	if (*szOut == _T('\\') || *szOut == _T('/'))
	{
		szOut++;
	}

	return(szOut);
}

const TCHAR *LeoHelpers::GetExtensionPart(bool bIncludeDot, const TCHAR *szIn)
{
	if (NULL == szIn)
	{
		return(szIn);
	}

	size_t len = _tcslen(szIn);

	if (1 >= len)
	{
		return(NULL);
	}

	const TCHAR *szOut = szIn + (len - 1);

	while (szIn < szOut && (*szOut != _T('.')))
	{
		szOut--;
	}

	if (*szOut != _T('.'))
	{
		return(NULL);
	}

	if (!bIncludeDot)
	{
		szOut++;
	}

	return(szOut);
}

bool LeoHelpers::GetTempPath(std::basic_string< TCHAR > *pstrTempPath)
{
	// If we ask for the size first then we have to worry about the temp dir
	// being redefined between calls. In theory we still have to worry about
	// that with the code below but only if the temp path is already
	// longer than MAX_PATH and it then gets redefined, between calls, and
	// the new definition is even longer than the old one. In which case we
	// will fail gracefully. Since a lot of other software will fail just because
	// the temp path is longer than MAX_PATH it doesn't seem worth looping to
	// handle the worst case scenario.

	DWORD dwTempBufferSize = MAX_PATH;
	TCHAR szTempPath1[MAX_PATH];

	DWORD dwTempRes = ::GetTempPath(dwTempBufferSize, szTempPath1);

	if (dwTempRes == 0)
	{
		return false;
	}
	else if (dwTempRes < dwTempBufferSize)
	{
		*pstrTempPath = szTempPath1;
		return true;
	}

	dwTempBufferSize = dwTempRes + 1;
	TCHAR *szTempPath2 = new TCHAR[dwTempBufferSize];

	dwTempRes = ::GetTempPath(dwTempBufferSize, szTempPath2);

	bool bResult = false;

	if (dwTempRes != 0 && dwTempRes < dwTempBufferSize)
	{
		*pstrTempPath = szTempPath2;
		bResult = true;
	}

	delete[] szTempPath2;

	return bResult;
}

// If bDontReallyOpen is true then the function won't really create a file and always returns INVALID_HANDLE_VALUE, but will still set *pstrFilenameOut to a name that is similar to what you'd get normally.
// On failure *pstrFilenameOut will be left empty so you can check that when bDontReallyOpen is true.
// If you set bDontReallyOpen then you absolutely must not create a file with the generated path because such a file may already exist and be used for other purposes. The generated path should only be used as a sample, for example if you want to test whether calling this function for real will result in a path that contains any non-ASCII characters.
HANDLE LeoHelpers::OpenTempFileNamePreserveExtension(const TCHAR *szTempPath, const TCHAR *szPrefix, const TCHAR *szFilenameIn, std::basic_string< TCHAR > *pstrFilenameOut, bool bDontReallyOpen)
{
	HANDLE hFileResult = INVALID_HANDLE_VALUE;

	if (NULL != pstrFilenameOut)
	{
		pstrFilenameOut->clear();
	}

	std::basic_string< TCHAR > strFileStart;
	
	if (szTempPath != NULL)
	{
		strFileStart = szTempPath;
	}
	else if (!LeoHelpers::GetTempPath(&strFileStart))
	{
		return hFileResult;
	}

	if (strFileStart.length() > 0 && '\\' != *(strFileStart.rbegin()) && '/' != *(strFileStart.rbegin()))
	{
		strFileStart += '\\';
	}

	strFileStart += szPrefix;

	std::basic_string< TCHAR > strExtension = _T(".tmp");

	if (NULL != szFilenameIn)
	{
		const TCHAR *szExt = _tcsrchr(szFilenameIn, _T('.'));

		if (NULL != szExt)
		{
			strExtension = szExt;
		}
	}

	SYSTEMTIME systime;
	GetSystemTime(&systime);

	int i1 = systime.wYear;
	int i2 = systime.wMonth;
	int i3 = systime.wDay;
	int i4 = systime.wHour;
	int i5 = systime.wMinute;
	int i6 = systime.wSecond;
	int i7a = systime.wMilliseconds;

	for (int i = 0; i < 100; i++)
	{
		int i7 = i7a + i;

		TCHAR *szFilename = StringAllocAndFormat(_T("%s%04d%02d%02d%02d%02d%02d%04d%s"),
													strFileStart.c_str(), i1,i2,i3,i4,i5,i6,i7, strExtension.c_str());

		if (NULL == szFilename)
		{
			break;
		}
		else
		{
			if (!bDontReallyOpen)
			{
				hFileResult = ::CreateFile(szFilename, GENERIC_WRITE, 0, 0, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, 0);
			}

			if (bDontReallyOpen || hFileResult != INVALID_HANDLE_VALUE)
			{
				if (NULL != pstrFilenameOut)
				{
					*pstrFilenameOut = szFilename;
				}
				delete [] szFilename;
				break;
			}
			delete [] szFilename;
		}
	}

	return(hFileResult);
}

std::wstring LeoHelpers::GenerateLegalFileName(const wchar_t * szInput, bool bAllowPathSeps, bool bAllowDots)
{
	std::wstring strResult;

	if (NULL != szInput)
	{
		while(*szInput)
		{
			if (LeoHelpers::IsLegalPathChar(*szInput, bAllowPathSeps, bAllowDots))
			{
				strResult += *szInput++;
			}
			else
			{
				strResult += L"_";
				szInput++;
			}
		}
	}

	return strResult;
}

void LeoHelpers::GenerateLegalFileName(std::wstring *pstr, bool bAllowPathSeps, bool bAllowDots, std::wstring::size_type startPos)
{
	std::wstring::iterator citer = pstr->begin();

	while(startPos > 0 && citer != pstr->end())
	{
		++citer;
		--startPos;
	}

	while(citer != pstr->end())
	{
		if (!LeoHelpers::IsLegalPathChar(*citer, bAllowPathSeps, bAllowDots))
		{
			*citer = L'_';
		}
		++citer;
	}
}

std::wstring LeoHelpers::GenerateSafeName(const wchar_t *szInput)
{
	std::wstring strResult;

	if (NULL != szInput)
	{
		while(*szInput)
		{
			if (iswalpha(*szInput) || iswdigit(*szInput))
			{
				strResult += *szInput++;
			}
			else
			{
				strResult += L"_";
				szInput++;
			}
		}
	}

	return strResult;
}

bool LeoHelpers::CopyFile(const TCHAR *szSource, const TCHAR *szDest, HANDLE hAbortEvent, bool bClearReadOnlyAttrib, bool bFailIfExists)
{
	class CopyFileDummyInner
	{
	public:
		static DWORD CALLBACK CopyProgressRoutine(
			LARGE_INTEGER TotalFileSize,
			LARGE_INTEGER TotalBytesTransferred,
			LARGE_INTEGER StreamSize,
			LARGE_INTEGER StreamBytesTransferred,
			DWORD dwStreamNumber,
			DWORD dwCallbackReason,
			HANDLE hSourceFile,
			HANDLE hDestinationFile,
			LPVOID lpData)
		{
			if (lpData == NULL)
			{
				return PROGRESS_QUIET;
			}

			if (WAIT_TIMEOUT != ::WaitForSingleObject(static_cast<HANDLE>(lpData), 0))
			{
				return PROGRESS_CANCEL;
			}

			return PROGRESS_CONTINUE;
		}
	};

	BOOL bCancel = FALSE; // Never used.
	DWORD dwCopyFlags = (bFailIfExists ? COPY_FILE_FAIL_IF_EXISTS : 0);

	if (0 == ::CopyFileEx(szSource, szDest, CopyFileDummyInner::CopyProgressRoutine, hAbortEvent, &bCancel, dwCopyFlags))
	{
		return false;
	}

	if (bClearReadOnlyAttrib)
	{
		DWORD dwAttribs = ::GetFileAttributes(szDest);

		if (dwAttribs != INVALID_FILE_ATTRIBUTES
		&&	(dwAttribs&FILE_ATTRIBUTE_READONLY))
		{
			dwAttribs -= FILE_ATTRIBUTE_READONLY;
			::SetFileAttributes(szDest, dwAttribs);
		}
	}

	return true;
}

// Intended for reading from pipes.
bool LeoHelpers::ReadFileUntilDone(HANDLE hFile, LPVOID lpBuffer, DWORD dwNumberOfBytesToRead)
{
	BYTE *pBuffer = reinterpret_cast<BYTE *>(lpBuffer);

	DWORD dwBytesRead;

	while(dwNumberOfBytesToRead > 0)
	{
		dwBytesRead = 0;

		if (!ReadFile(hFile, pBuffer, dwNumberOfBytesToRead, &dwBytesRead, NULL)
		||	dwBytesRead == 0
		||	dwBytesRead > dwNumberOfBytesToRead)
		{
			return false;
		}

		dwNumberOfBytesToRead -= dwBytesRead;
		pBuffer += dwBytesRead;
	}

	return true;
}

// Intended for writing to pipes.
bool LeoHelpers::WriteFileUntilDone(HANDLE hFile, LPVOID lpBuffer, DWORD dwNumberOfBytesToWrite)
{
	BYTE *pBuffer = reinterpret_cast<BYTE *>(lpBuffer);

	DWORD dwBytesWritten;

	while(dwNumberOfBytesToWrite > 0)
	{
		dwBytesWritten = 0;

		if (!WriteFile(hFile, pBuffer, dwNumberOfBytesToWrite, &dwBytesWritten, NULL)
		||	dwBytesWritten == 0
		||	dwBytesWritten > dwNumberOfBytesToWrite)
		{
			return false;
		}

		dwNumberOfBytesToWrite -= dwBytesWritten;
		pBuffer += dwBytesWritten;
	}

	return true;
}

BOOL LeoHelpers::GetDlgItemRect(HWND hwndDlg, int nIDDlgItem, RECT *pRect)
{
	BOOL bResult = FALSE;

	HWND hwnd = GetDlgItem(hwndDlg, nIDDlgItem);

	if (NULL != hwnd)
	{
		bResult = GetWindowRect(hwnd, pRect);
	}

	return(bResult);
}

BOOL LeoHelpers::OffsetDlgItem(HWND hwndDlg, int nIDDlgItem, LONG dx, LONG dy, LONG dw, LONG dh, BOOL bRepaint)
{
	BOOL bResult = FALSE;
	HWND hwnd = hwndDlg; // We'll adjust the dialog itself if nIDDlgItem is zero.

	if (0 != nIDDlgItem && NULL != hwndDlg)
	{
		hwnd = GetDlgItem(hwndDlg, nIDDlgItem);
	}

	if (NULL != hwnd)
	{
		RECT r;
		GetWindowRect(hwnd, &r);

		if (0 == nIDDlgItem || ScreenToClientRect(hwndDlg, &r))
		{
			bResult = MoveWindow(hwnd, r.left + dx, r.top + dy, (r.right - r.left) + dw, (r.bottom - r.top) + dh, bRepaint);
		}
	}

	return(bResult);
}

BOOL LeoHelpers::EnableDlgItem(HWND hwndDlg, int nIDDlgItem, BOOL bEnable)
{
	BOOL bResult = FALSE;

	HWND hwnd = GetDlgItem(hwndDlg, nIDDlgItem);

	if (NULL != hwnd)
	{
		if (!bEnable && ::GetFocus() == hwnd)
		{
			// We're about to disable the control with focus. Give focus to the next control.
			// Set focus to next control. (Do not use SetFocus in dialogs. Use WM_NEXTDLGCTL instead.)
			SendMessage(hwndDlg, WM_NEXTDLGCTL, 0, FALSE);
		}

		bResult = ::EnableWindow(hwnd, bEnable);
	}

	return(bResult);
}

BOOL LeoHelpers::ShowDlgItem(HWND hwndDlg, int nIDDlgItem, BOOL bShow)
{
	BOOL bResult = FALSE;

	HWND hwnd = GetDlgItem(hwndDlg, nIDDlgItem);

	if (NULL != hwnd)
	{
		if (!bShow && ::GetFocus() == hwnd)
		{
			// We're about to hide the control with focus. Give focus to the next control.
			// Set focus to next control. (Do not use SetFocus in dialogs. Use WM_NEXTDLGCTL instead.)
			SendMessage(hwndDlg, WM_NEXTDLGCTL, 0, FALSE);
		}

		bResult = ::SetWindowPos(hwnd, NULL, 0, 0, 0, 0,
			SWP_NOACTIVATE|SWP_NOMOVE|SWP_NOOWNERZORDER|SWP_NOSIZE|SWP_NOZORDER|(bShow?SWP_SHOWWINDOW:SWP_HIDEWINDOW));
	}

	return(bResult);
}

BOOL LeoHelpers::GetDlgItemText(HWND hwndDlg, int nIDDlgItem, std::wstring *pStrResult)
{
	BOOL bResult = FALSE;

	pStrResult->erase();

	HWND hwnd = ::GetDlgItem(hwndDlg, nIDDlgItem);

	if (NULL != hwnd)
	{
		bResult = LeoHelpers::GetWindowTextString(hwnd, pStrResult);
	}

	return(bResult);
}

BOOL LeoHelpers::GetListBoxItemText(HWND hwndListBox, LRESULT lrItemIndex, std::wstring *pStrResult)
{
	BOOL bResult = FALSE;

	pStrResult->clear();

	LRESULT itemTextLength = SendMessage(hwndListBox, LB_GETTEXTLEN, lrItemIndex, 0);
	if (0 == itemTextLength)
	{
		bResult = TRUE;
	}
	else if (0 < itemTextLength)
	{
		TCHAR *szItemText = new TCHAR[ itemTextLength + 1 ];

		if (0 < SendMessage(hwndListBox, LB_GETTEXT, lrItemIndex, reinterpret_cast<LPARAM>(szItemText)))
		{
			*pStrResult = szItemText;
			bResult = TRUE;
		}

		delete [] szItemText;
	}

	return(bResult);
}

BOOL LeoHelpers::SetupComboItem(HWND hWndParent, int iDlgItem, bool bClear, DWORD dwValue, DWORD dwSelectValue, const TCHAR *szTitle)
{
	BOOL bResult = FALSE;

	HWND hWndCombo = GetDlgItem(hWndParent, iDlgItem);

	if (hWndCombo != NULL)
	{
		if (bClear)
		{
			SendMessage(hWndCombo, CB_RESETCONTENT, 0, 0);
		}

		LRESULT idx = SendMessage(hWndCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(szTitle));
		if (0 <= idx)
		{
		//	DWORD dwValue = *reinterpret_cast<DWORD *>(&iValue);

			SendMessage(hWndCombo, CB_SETITEMDATA, idx, dwValue);

			if (dwSelectValue == dwValue)
			{
				SendMessage(hWndCombo, CB_SETCURSEL, idx, 0);
			}

			bResult = TRUE;
		}
	}

	return(bResult);
}

BOOL LeoHelpers::GetComboItem(HWND hWndParent, int iDlgItem, DWORD *pdwValue)
{
	BOOL bResult = FALSE;

	HWND hWndCombo = GetDlgItem(hWndParent, iDlgItem);

	if (hWndCombo != NULL)
	{
		LRESULT curSel = SendMessage(hWndCombo, CB_GETCURSEL, 0, 0);

		if (CB_ERR != curSel)
		{
			*pdwValue = static_cast<DWORD>(SendMessage(hWndCombo, CB_GETITEMDATA, curSel, 0));

		//	iValue = *reinterpret_cast<int *>(&dwValue);

			bResult = TRUE;
		}
	}

	return(bResult);
}

BOOL LeoHelpers::GetComboItemIntCast(HWND hWndParent, int iDlgItem, int *piValue)
{
	DWORD dwValue = 0;
	if (!LeoHelpers::GetComboItem(hWndParent, iDlgItem, &dwValue))
	{
		return FALSE;
	}

	*piValue = *reinterpret_cast< int * >(&dwValue);
	return TRUE;
}

BOOL LeoHelpers::SetupCheckBoxItem(HWND hWndParent, int iDlgItem, bool bCheck)
{
	return(::CheckDlgButton(hWndParent, iDlgItem, bCheck ? BST_CHECKED : BST_UNCHECKED));
}

BOOL LeoHelpers::GetCheckBoxItem(HWND hWndParent, int iDlgItem, bool *pbValue)
{
	UINT res = ::IsDlgButtonChecked(hWndParent, iDlgItem);

	*pbValue = (BST_CHECKED == res);

	return(res != 0 ? TRUE : FALSE);
}

#ifdef UDM_SETRANGE32
BOOL LeoHelpers::SetupSpinItem(HWND hWndParent, int iDlgItem, int iMax, int iMin, int iValue)
{
	BOOL bResult = FALSE;

	HWND hWndSpin = GetDlgItem(hWndParent, iDlgItem);

	if (hWndSpin != NULL)
	{
		SendMessage(hWndSpin, UDM_SETRANGE32, iMin, iMax);
		SendMessage(hWndSpin, UDM_SETPOS32, 0, iValue);
		// Recipient of WM_VSCROLL shouldn't rely on the 16-bit scroll amount.
		SendMessage(hWndParent, WM_VSCROLL, MAKELONG(0, iValue), reinterpret_cast<LPARAM>(hWndSpin));

		bResult = TRUE;
	}

	return(bResult);
}

BOOL LeoHelpers::SetSpinItem(HWND hWndParent, int iDlgItem, int iValue)
{
	BOOL bResult = FALSE;

	HWND hWndSpin = GetDlgItem(hWndParent, iDlgItem);

	if (hWndSpin != NULL)
	{
		SendMessage(hWndSpin, UDM_SETPOS32, 0, iValue);
		// Recipient of WM_VSCROLL shouldn't rely on the 16-bit scroll amount.
		SendMessage(hWndParent, WM_VSCROLL, MAKELONG(0, iValue), reinterpret_cast<LPARAM>(hWndSpin));

		bResult = TRUE;
	}

	return(bResult);
}

BOOL LeoHelpers::GetSpinItem(HWND hWndParent, int iDlgItem, int *piValue)
{
	BOOL bResult = FALSE;

	HWND hWndSpin = GetDlgItem(hWndParent, iDlgItem);

	if (hWndSpin != NULL)
	{
		BOOL bError = TRUE;

		*piValue = static_cast<int>(SendMessage(hWndSpin, UDM_GETPOS32, 0, reinterpret_cast<LPARAM>(&bError)));

		if (!bError)
		{
			bResult = TRUE;
		}
	}

	return(bResult);
}
#endif

BOOL LeoHelpers::SetupEditItem(HWND hWndParent, int iDlgItem, const TCHAR *szText)
{
	BOOL bResult = FALSE;

	HWND hWndEdit = GetDlgItem(hWndParent, iDlgItem);

	if (hWndEdit != NULL)
	{
		bResult = SetWindowText(hWndEdit, szText);
	}

	return(bResult);
}

BOOL LeoHelpers::GetEditItem(HWND hWndParent, int iDlgItem, std::wstring *pStrResult)
{
	return(LeoHelpers::GetDlgItemText(hWndParent, iDlgItem, pStrResult));
}

/*
GWLP_USERDATA shouldn't be used for this. See the SubclassData class for a better method.
BOOL LeoHelpers::SubclassDlgItemToUserData(HWND hwndDlg, int nIDDlgItem, WNDPROC wndProc)
{
	BOOL bResult = FALSE;

	HWND hwnd = GetDlgItem(hwndDlg, nIDDlgItem);

	if (NULL != hwnd)
	{
		#pragma warning(suppress:4244) // spurious warning due to stupid Win32 SDK header definition.
		SetWindowLongPtr(hwnd, GWLP_USERDATA, SetWindowLongPtr(hwnd, GWLP_WNDPROC, reinterpret_cast<LPARAM>(wndProc)));
		bResult = TRUE;
	}

	return(bResult);
}
*/

/*
BOOL LeoHelpers::GetDlgItemLong(HWND hwndDlg, int nIDDlgItem, int nIndex, LONG *plResult)
{
	BOOL bResult = FALSE;

	HWND hwnd = GetDlgItem(hwndDlg, nIDDlgItem);

	if (NULL != hwnd)
	{
		SetLastError(0);

		LONG lStyleValue = GetWindowLong(hwnd, nIndex);

		if (0 != lStyleValue || 0 == GetLastError())
		{
			*plResult = lStyleValue;
			bResult = TRUE;
		}
	}

	return(bResult);
}

BOOL LeoHelpers::SetDlgItemLong(HWND hwndDlg, int nIDDlgItem, int nIndex, LONG lValue)
{
	BOOL bResult = FALSE;

	HWND hwnd = GetDlgItem(hwndDlg, nIDDlgItem);

	if (NULL != hwnd)
	{
		SetLastError(0);

		LONG lOldStyleValue = SetWindowLong(hwnd, nIndex, lValue);

		if (0 != lOldStyleValue || 0 == GetLastError())
		{
			bResult = TRUE;
		}
	}

	return(bResult);
}
*/

BOOL LeoHelpers::GetDlgItemLongPtr(HWND hwndDlg, int nIDDlgItem, int nIndex, LONG_PTR *plResult)
{
	BOOL bResult = FALSE;

	HWND hwnd = GetDlgItem(hwndDlg, nIDDlgItem);

	if (NULL != hwnd)
	{
		SetLastError(0);

		LONG_PTR lStyleValue = GetWindowLongPtr(hwnd, nIndex);

		if (0 != lStyleValue || 0 == GetLastError())
		{
			*plResult = lStyleValue;
			bResult = TRUE;
		}
	}

	return(bResult);
}

BOOL LeoHelpers::SetDlgItemLongPtr(HWND hwndDlg, int nIDDlgItem, int nIndex, LONG_PTR lValue)
{
	BOOL bResult = FALSE;

	HWND hwnd = GetDlgItem(hwndDlg, nIDDlgItem);

	if (NULL != hwnd)
	{
		SetLastError(0);

		#pragma warning(suppress:4244) // spurious warning due to stupid Win32 SDK header definition.
		LONG_PTR lOldStyleValue = SetWindowLongPtr(hwnd, nIndex, lValue);

		if (0 != lOldStyleValue || 0 == GetLastError())
		{
			bResult = TRUE;
		}
	}

	return(bResult);
}

/*
BOOL LeoHelpers::ModifyDlgItemStyle(HWND hwndDlg, int nIDDlgItem, bool bStyleEx, LONG_PTR lAdd, LONG_PTR lRemove)
{
	BOOL bResult = FALSE;

	HWND hwnd = GetDlgItem(hwndDlg, nIDDlgItem);

	if (NULL != hwnd)
	{
		int nStyleKey = bStyleEx ? GWL_EXSTYLE : GWL_STYLE;

		SetLastError(0);

		LONG_PTR lStyleValue = GetWindowLongPtr(hwnd, nStyleKey);

		if (0 != lStyleValue || 0 == GetLastError())
		{
			lStyleValue &= (~lRemove);
			lStyleValue |= lAdd;

			lStyleValue = SetWindowLongPtr(hwnd, nStyleKey, lStyleValue);

			if (0 != lStyleValue || 0 == GetLastError())
			{
				bResult = TRUE;
			}
		}
	}

	return(bResult);
}
*/

BOOL LeoHelpers::InvalidateDlgItem(HWND hwndDlg, int nIDDlgItem)
{
	BOOL bResult = FALSE;

	HWND hwnd = GetDlgItem(hwndDlg, nIDDlgItem);

	if (NULL != hwnd)
	{
		bResult = InvalidateRect(hwnd, NULL, FALSE);
	}

	return(bResult);
}

BOOL LeoHelpers::EnableDlgControl(HWND hwndControl, BOOL bEnable)
{
	BOOL bResult = FALSE;

	if (NULL != hwndControl)
	{
		if (!bEnable && ::GetFocus() == hwndControl)
		{
			// We're about to disable the control with focus. Give focus to the next control.
			// Set focus to next control. (Do not use SetFocus in dialogs. Use WM_NEXTDLGCTL instead.)
			SendMessage(GetParent(hwndControl), WM_NEXTDLGCTL, 0, FALSE);
		}

		bResult = ::EnableWindow(hwndControl, bEnable);
	}

	return(bResult);
}

BOOL LeoHelpers::ScreenToClientRect(HWND hwnd, RECT *pRect)
{
	BOOL bResult = FALSE;

	POINT ptTL;
	POINT ptBR;

	ptTL.x = pRect->left;
	ptTL.y = pRect->top;
	ptBR.x = pRect->right;
	ptBR.y = pRect->bottom;

	if (ScreenToClient(hwnd, &ptTL) && ScreenToClient(hwnd, &ptBR))
	{
		pRect->left   = ptTL.x;
		pRect->top    = ptTL.y;
		pRect->right  = ptBR.x;
		pRect->bottom = ptBR.y;

		bResult = TRUE;
	}

	return(bResult);
}

BOOL LeoHelpers::ClientToScreenRect(HWND hwnd, RECT *pRect)
{
	BOOL bResult = FALSE;

	POINT ptTL;
	POINT ptBR;

	ptTL.x = pRect->left;
	ptTL.y = pRect->top;
	ptBR.x = pRect->right;
	ptBR.y = pRect->bottom;

	if (ClientToScreen(hwnd, &ptTL) && ClientToScreen(hwnd, &ptBR))
	{
		pRect->left   = ptTL.x;
		pRect->top    = ptTL.y;
		pRect->right  = ptBR.x;
		pRect->bottom = ptBR.y;

		bResult = TRUE;
	}

	return(bResult);
}

void LeoHelpers::CenterWindow(HWND hWndToMove, HWND hWndRelative)
{
	bool bResult = true;

	RECT rectSnap;

	RECT rectRelative;

	if (NULL != hWndRelative)
	{
		::GetWindowRect(hWndRelative, &rectRelative);

		if (!LeoHelpers::GetAreasFromWindow(NULL, &rectSnap, hWndRelative, MONITOR_DEFAULTTONEAREST))
		{
			bResult = false;
		}
	}
	else
	{
		// Get work area under the mouse pointer.

		POINT pointCursor;
		MONITORINFO monitorInfo;
		ZeroMemory(&monitorInfo, sizeof(monitorInfo));
		monitorInfo.cbSize = sizeof(monitorInfo);
		HMONITOR hMonitor;

		if (GetCursorPos(&pointCursor)
		&&	(NULL != (hMonitor = ::MonitorFromPoint(pointCursor, MONITOR_DEFAULTTOPRIMARY)))
		&&	::GetMonitorInfo(hMonitor, &monitorInfo))
		{
			rectRelative = monitorInfo.rcWork;
			CopyRect(&rectSnap, &rectRelative);
		}
		else if (::GetWindowRect(::GetDesktopWindow(), &rectRelative))
		{
			CopyRect(&rectSnap, &rectRelative);
		}
		else
		{
			bResult = false;
		}
	}

	RECT rectToMove;

	if (bResult && ::GetWindowRect(hWndToMove, &rectToMove))
	{
		int relativeWidth  = (rectRelative.right  - rectRelative.left);
		int relativeHeight = (rectRelative.bottom - rectRelative.top);
		int toMoveWidth    = (rectToMove.right    - rectToMove.left);
		int toMoveHeight   = (rectToMove.bottom   - rectToMove.top);

		int x = rectRelative.left + ( (relativeWidth  - toMoveWidth ) / 2 );
		int y = rectRelative.top  + ( (relativeHeight - toMoveHeight) / 2 );

		if (x < rectSnap.left)
		{
			x = rectSnap.left;
		}

		if ((x + toMoveWidth) > rectSnap.right)
		{
			x = rectSnap.right - toMoveWidth;
		}

		if (y < rectSnap.top)
		{
			y = rectSnap.top;
		}

		if ((y + toMoveHeight) > rectSnap.bottom)
		{
			y = rectSnap.bottom - toMoveHeight;
		}

		::SetWindowPos(hWndToMove, NULL, x, y, 0, 0, SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE);
	}
}

bool LeoHelpers::GetAreasFromWindow(RECT *pRectFull, RECT *pRectWork, HWND hWnd, DWORD dwFlags)
{
	if (NULL != pRectFull) { ZeroMemory(pRectFull, sizeof(RECT)); }
	if (NULL != pRectWork) { ZeroMemory(pRectWork, sizeof(RECT)); }

	bool bResult = false;

	bool bGetDefault = true;

	HMONITOR hMonitor = ::MonitorFromWindow(hWnd, dwFlags);

	if ((NULL == hMonitor) && (MONITOR_DEFAULTTONULL == dwFlags))
	{
		bGetDefault = false;
	}
	else if (NULL != hMonitor)
	{
		MONITORINFO mi;
		ZeroMemory(&mi, sizeof(mi));
		mi.cbSize = sizeof(mi);

		if (::GetMonitorInfo(hMonitor, &mi)
		&&	(NULL == pRectFull || CopyRect(pRectFull, &mi.rcMonitor))
		&&	(NULL == pRectWork || CopyRect(pRectWork, &mi.rcWork   )))
		{
			bResult = true;
		}
	}

	if (bGetDefault && (!bResult))
	{
		bResult = true;

		if (bResult && NULL != pRectFull)
		{
			pRectFull->left   = 0;
			pRectFull->top    = 0;
			pRectFull->right  = ::GetSystemMetrics(SM_CXFULLSCREEN);
			pRectFull->bottom = ::GetSystemMetrics(SM_CYFULLSCREEN);

			if (0 == pRectFull->right || 0 == pRectFull->bottom)
			{
				bResult = false;
			}
		}

		if (bResult && NULL != pRectWork)
		{
			if (!::SystemParametersInfo(SPI_GETWORKAREA, 0, pRectWork, 0))
			{
				bResult = false;
			}
		}
	}

	return(bResult);
}

BOOL LeoHelpers::GetWindowTextString(HWND hwnd, std::basic_string< TCHAR > *pStrResult)
{
	BOOL bResult = FALSE;

	pStrResult->erase();

	int iLength = GetWindowTextLength(hwnd);

	if (0 == iLength)
	{
		bResult = TRUE;
	}
	else if (0 < iLength)
	{
		iLength++; // Not clear if GetWindowTextLength includes the null but GetWindowText does.

		TCHAR *szBuffer = new TCHAR[ iLength ];

		if (0 < GetWindowText(hwnd, szBuffer, iLength))
		{
			*pStrResult = szBuffer;

			bResult = TRUE;
		}

		delete [] szBuffer;
	}

	return bResult;
}

LeoHelpers::GlobalAtomHolder LeoHelpers::s_SubclassAtomPropHolder(L"{C73ECBF0-8E75-4421-B1B1-EB3D8EB0C19F}");

bool LeoHelpers::DialogLayoutHelper::ResizeStatic(HWND hwndStatic, RECT *pInOutRect, bool bResize, bool bRedraw, bool bSingleLine) const
{
	bool bResult = false;

	RECT r = {0};

	std::basic_string< TCHAR > strWindowText;

	if (LeoHelpers::GetWindowTextString(hwndStatic, &strWindowText))
	{
		HDC hDC = ::GetDC(hwndStatic);

		if (hDC != NULL)
		{
			HFONT hFont = reinterpret_cast< HFONT >(::SendMessage(hwndStatic, WM_GETFONT, 0, 0));
			HGDIOBJ hOldFont = NULL;

			if (hFont != NULL)
			{
				hOldFont = ::SelectObject(hDC, hFont);
			}

			UINT uFormat = DT_CALCRECT | DT_NOCLIP | DT_EXPANDTABS; // I think DT_EXPANDTABS is always in place with statics. Not sure though.
			
			if (bSingleLine)
			{
				uFormat |= DT_SINGLELINE;
			}
			else if (pInOutRect != NULL)
			{
				uFormat |= DT_WORDBREAK;
				r = *pInOutRect;
			}

			if (SS_NOPREFIX == (::GetWindowLongPtr(hwndStatic, GWL_STYLE) & SS_NOPREFIX))
			{
				uFormat |= DT_NOPREFIX;
			}

			if (SS_ENDELLIPSIS == (::GetWindowLongPtr(hwndStatic, GWL_STYLE) & SS_ENDELLIPSIS))
			{
				uFormat |= DT_END_ELLIPSIS;
			}

			if (SS_WORDELLIPSIS == (::GetWindowLongPtr(hwndStatic, GWL_STYLE) & SS_WORDELLIPSIS))
			{
				uFormat |= DT_WORD_ELLIPSIS;
			}

			if (SS_PATHELLIPSIS == (::GetWindowLongPtr(hwndStatic, GWL_STYLE) & SS_PATHELLIPSIS))
			{
				uFormat |= DT_PATH_ELLIPSIS;
			}

			if (hFont == NULL)
			{
				uFormat |= DT_INTERNAL; // Use system font.
			}

			if (0 != ::DrawText(hDC, strWindowText.c_str(), static_cast<int>(strWindowText.length()), &r, uFormat))
			{
				bResult = true;
			}

			if (hFont != NULL)
			{
				::SelectObject(hDC, hOldFont);
			}

			::ReleaseDC(hwndStatic, hDC);
		}
	}

	if (bResult)
	{
		if (bResize)
		{
			::SetWindowPos(hwndStatic, NULL, 0, 0, r.right - r.left, r.bottom - r.top,
				SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOOWNERZORDER | SWP_NOZORDER | (bRedraw ? 0 : SWP_NOREDRAW));
		}

		if (pInOutRect != NULL)
		{
			*pInOutRect = r;
		}
	}

	return bResult;
}

bool LeoHelpers::DialogLayoutHelper::ResizeCheckbox(HWND hwndCheckbox, RECT *pOutRect, bool bResize, bool bRedraw) const
{
	bool bResult = false;

	RECT r = {0};

	std::basic_string< TCHAR > strWindowText;

	if (LeoHelpers::GetWindowTextString(hwndCheckbox, &strWindowText))
	{
		HDC hDC = ::GetDC(hwndCheckbox);

		if (hDC != NULL)
		{
			HFONT hFont = reinterpret_cast< HFONT >(::SendMessage(hwndCheckbox, WM_GETFONT, 0, 0));
			HGDIOBJ hOldFont = NULL;

			if (hFont != NULL)
			{
				hOldFont = ::SelectObject(hDC, hFont);
			}

			UINT uFormat = DT_CALCRECT | DT_NOCLIP | DT_SINGLELINE;

		//	if (SS_NOPREFIX == (::GetWindowLongPtr(hwndStatic, GWL_STYLE) & SS_NOPREFIX))
		//	{
		//		uFormat |= DT_NOPREFIX;
		//	}

			if (hFont == NULL)
			{
				uFormat |= DT_INTERNAL; // Use system font.
			}

			if (0 != ::DrawText(hDC, strWindowText.c_str(), static_cast<int>(strWindowText.length()), &r, uFormat))
			{
				r.right += (m_rect1.right + m_rect2.left);

				if (r.bottom < m_rect2.bottom)
				{
					r.bottom = m_rect2.bottom;
				}

				bResult = true;
			}

			if (hFont != NULL)
			{
				::SelectObject(hDC, hOldFont);
			}

			::ReleaseDC(hwndCheckbox, hDC);
		}
	}

	if (bResult)
	{
		if (bResize)
		{
			::SetWindowPos(hwndCheckbox, NULL, 0, 0, r.right - r.left, r.bottom - r.top,
				SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOOWNERZORDER | SWP_NOZORDER | (bRedraw ? 0 : SWP_NOREDRAW));
		}

		if (pOutRect != NULL)
		{
			*pOutRect = r;
		}
	}

	return bResult;
}

bool LeoHelpers::DialogLayoutHelper::ResizeEdit(HWND hwndEdit, RECT *pOutRect, bool bResize, bool bRedraw) const
{
	bool bResult = false;

	RECT r = {0};

	r.bottom = m_rect3.bottom;
	r.right  = m_rect3.bottom; // Arbitrary width.

	bResult = true;

	if (bResult)
	{
		if (bResize)
		{
			::SetWindowPos(hwndEdit, NULL, 0, 0, r.right - r.left, r.bottom - r.top,
				SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOOWNERZORDER | SWP_NOZORDER | (bRedraw ? 0 : SWP_NOREDRAW));
		}

		if (pOutRect != NULL)
		{
			*pOutRect = r;
		}
	}

	return bResult;
}

bool LeoHelpers::DialogLayoutHelper::ResizeButton(HWND hwndButton, RECT *pOutRect, bool bResize, bool bRedraw) const
{
	// This is also called on ComboBox controls.

	bool bResult = false;

	RECT r = {0};

	std::basic_string< TCHAR > strWindowText;

	if (LeoHelpers::GetWindowTextString(hwndButton, &strWindowText))
	{
		HDC hDC = ::GetDC(hwndButton);

		if (hDC != NULL)
		{
			HFONT hFont = reinterpret_cast< HFONT >(::SendMessage(hwndButton, WM_GETFONT, 0, 0));
			HGDIOBJ hOldFont = NULL;

			if (hFont != NULL)
			{
				hOldFont = ::SelectObject(hDC, hFont);
			}

			UINT uFormat = DT_CALCRECT | DT_NOCLIP | DT_SINGLELINE;

		//	if (SS_NOPREFIX == (::GetWindowLongPtr(hwndStatic, GWL_STYLE) & SS_NOPREFIX))
		//	{
		//		uFormat |= DT_NOPREFIX;
		//	}

			if (hFont == NULL)
			{
				uFormat |= DT_INTERNAL; // Use system font.
			}

			if (0 != ::DrawText(hDC, strWindowText.c_str(), static_cast<int>(strWindowText.length()), &r, uFormat))
			{
				r.right += (m_rect1.left * 4); // Use 2 * button-gap as the margin. Note that the UX Guide doesn't specify an official margin for button text.

				if ((r.right - r.left) < m_rect3.left)
				{
					r.right = r.left + m_rect3.left; // Minimum standard button width.
				}

				if (r.bottom < m_rect3.bottom)
				{
					r.bottom = m_rect3.bottom; // Minimum standard button height.
				}

				bResult = true;
			}

			if (hFont != NULL)
			{
				::SelectObject(hDC, hOldFont);
			}

			::ReleaseDC(hwndButton, hDC);
		}
	}

	if (bResult)
	{
		if (bResize)
		{
			::SetWindowPos(hwndButton, NULL, 0, 0, r.right - r.left, r.bottom - r.top,
				SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOOWNERZORDER | SWP_NOZORDER | (bRedraw ? 0 : SWP_NOREDRAW));
		}

		if (pOutRect != NULL)
		{
			*pOutRect = r;
		}
	}

	return bResult;

}

void LeoHelpers::DialogLayoutHelper::SetGroupBoxText(HWND hwndGroupBox, const wchar_t *szText, bool bNoPrefix) const
{
	bool bResult = false;

	RECT r = {0};

	wchar_t *szTextCopy = NULL;
	
	if (szText != NULL && szText[0] != '\0')
	{
		size_t slen = 0;
		
		for (const wchar_t *szIn = szText; *szIn != '\0'; ++szIn)
		{
			++slen;
			if (bNoPrefix && *szIn == L'&')
			{
				++slen;
			}
		}

		szTextCopy = new(std::nothrow) wchar_t[slen + 1 + 4]; // + 1 for null, + 4 for DT_MODIFYSTRING

		if (szTextCopy != NULL)
		{
			wchar_t *szOut = szTextCopy;
			for (const wchar_t *szIn = szText; *szIn != '\0'; ++szIn)
			{
				*szOut++ = *szIn;
				if (bNoPrefix && *szIn == L'&')
				{
					*szOut++ = L'&';
				}
			}
			*szOut = '\0';
			szOut = NULL;

			RECT r;

			if (::GetClientRect(hwndGroupBox, &r))
			{
				r.right -= 2 * GetMarginPixels();

				HDC hDC = ::GetDC(hwndGroupBox);

				if (hDC != NULL)
				{
					HFONT hFont = reinterpret_cast< HFONT >(::SendMessage(hwndGroupBox, WM_GETFONT, 0, 0));
					HGDIOBJ hOldFont = NULL;

					if (hFont != NULL)
					{
						hOldFont = ::SelectObject(hDC, hFont);
					}

					UINT uFormat = DT_CALCRECT | DT_NOCLIP | DT_SINGLELINE | DT_MODIFYSTRING | DT_WORD_ELLIPSIS;

					if (hFont == NULL)
					{
						uFormat |= DT_INTERNAL; // Use system font.
					}

					if (0 == ::DrawText(hDC, szTextCopy, static_cast<int>(slen), &r, uFormat))
					{
						delete[] szTextCopy;
						szTextCopy = NULL;
					}

					if (hFont != NULL)
					{
						::SelectObject(hDC, hOldFont);
					}

					::ReleaseDC(hwndGroupBox, hDC);
				}
			}
		}
	}

	if (szTextCopy != NULL)
	{
		::SetWindowText(hwndGroupBox, szTextCopy);
		delete[] szTextCopy;
	}
	else
	{
		::SetWindowText(hwndGroupBox, szText);
	}
}

#ifdef LEOHELP_GDI

BOOL LeoHelpers::GradientFillIfColor(HDC hdc, PTRIVERTEX pVertices, ULONG nVertices, PVOID pMesh, ULONG nMeshElements, ULONG dwMode)
{
	BOOL bResult = FALSE;

	// Although the bitmap in the DC is 32-bit, GetDeviceCaps still reports the device's colours.
	int iNumColors = GetDeviceCaps(hdc, NUMCOLORS); // Will be -1 if more than 256 colours.

	if (iNumColors >= 0 && iNumColors <= 256) // Only attempt a GradientFill if display is more than 256 colours.
	{
		::SetLastError(ERROR_NOT_SUPPORTED);
	}
	else
	{
		if (::GradientFill(hdc, pVertices, nVertices, pMesh, nMeshElements, dwMode))
		{
			bResult = TRUE;
		}
	}

	return(bResult);
}

#endif

HRESULT LeoHelpers::LoadIconMetricOrDefault(HINSTANCE hinst, PCWSTR pszName, int lims, HICON *phico, bool *pbDestroyIcon)
{
	HRESULT hr = E_FAIL;

	*pbDestroyIcon = false;

	HMODULE hModComCtl32 = ::LoadLibrary(L"comctl32.dll");

	if (hModComCtl32 != NULL)
	{
		typedef HRESULT (WINAPI *LOADICONMETRIC_PTR)(HINSTANCE hinst, PCWSTR pszName, int lims, __out HICON *phico);

		LOADICONMETRIC_PTR fLoadIconMetric = reinterpret_cast<LOADICONMETRIC_PTR>(::GetProcAddress(hModComCtl32, "LoadIconMetric"));

		if (fLoadIconMetric != NULL)
		{
			hr = fLoadIconMetric(hinst, pszName, lims, phico);
			*pbDestroyIcon = true; // Presumably the result of this is never shared, even if the icon wasn't resized, although MSDN doesn't mention anything about clean-up.
		}

		::FreeLibrary(hModComCtl32);
	}

	if (S_OK != hr)
	{
		int cx = 0;
		int cy = 0;

		if (LIM_SMALL == lims)
		{
			cx = ::GetSystemMetrics(SM_CXSMICON);
			cy = ::GetSystemMetrics(SM_CYSMICON);
		}
		else if (LIM_LARGE == lims)
		{
			cx = ::GetSystemMetrics(SM_CXICON);
			cy = ::GetSystemMetrics(SM_CYICON);
		}

		if (cx != 0 && cy != 0)
		{
			HICON hTemp = reinterpret_cast<HICON>(::LoadImage(hinst, pszName, IMAGE_ICON, cx, cy, 0));

			if (hTemp != NULL)
			{
				hr = S_OK;
				*phico = hTemp;
				*pbDestroyIcon = true;
			}
			else
			{
				hTemp = reinterpret_cast<HICON>(::LoadImage(hinst, pszName, IMAGE_ICON, cx, cy, LR_SHARED));

				if (hTemp != NULL)
				{
					hr = S_OK;
					*phico = hTemp;
					*pbDestroyIcon = false;
				}
			}
		}
	}

	if (S_OK != hr)
	{
		*phico = NULL;
		*pbDestroyIcon = false;
	}

	return hr;
}

// DeleteObject the result
HFONT LeoHelpers::GetMessageFont()
{
	HFONT hFont = NULL;

	NONCLIENTMETRICS ncm = {0};
	ncm.cbSize = sizeof(NONCLIENTMETRICS);

	if (::SystemParametersInfo(SPI_GETNONCLIENTMETRICS, sizeof(NONCLIENTMETRICS), &ncm, 0))
	{
		hFont = ::CreateFontIndirect(&ncm.lfMessageFont);
	}

	return hFont;
}

LeoHelpers::ThemeHelper::ThemeHelper()
{
	m_hmodUxThemeDLL = LoadLibrary(_T("uxtheme.dll"));

	if (NULL == m_hmodUxThemeDLL
	||	NULL == (m_fOpenThemeData						  = reinterpret_cast<OPENTHEMEDATA_PTR            >(GetProcAddress(m_hmodUxThemeDLL, "OpenThemeData")))
	||	NULL == (m_fCloseThemeData						  = reinterpret_cast<CLOSETHEMEDATA_PTR           >(GetProcAddress(m_hmodUxThemeDLL, "CloseThemeData")))
	||	NULL == (m_fDrawThemeBackground					  = reinterpret_cast<DRAWTHEMEBACKGROUND_PTR      >(GetProcAddress(m_hmodUxThemeDLL, "DrawThemeBackground")))
	||	NULL == (m_fDrawThemeParentBackground			  = reinterpret_cast<DRAWTHEMEPARENTBACKGROUND_PTR>(GetProcAddress(m_hmodUxThemeDLL, "DrawThemeParentBackground")))
//	||	NULL == (m_fIsThemeBackgroundPartiallyTransparent = reinterpret_cast<ISTHEMEBACKGROUNDPARTIALLYTRANSPARENT_PTR>(GetProcAddress(m_hmodUxThemeDLL, "IsThemeBackgroundPartiallyTransparent")))
	||	NULL == (m_fGetThemePartSize					  = reinterpret_cast<GETTHEMEPARTSIZE_PTR         >(GetProcAddress(m_hmodUxThemeDLL, "GetThemePartSize")))
	||	NULL == (m_fEnableThemeDialogTexture			  = reinterpret_cast<ENABLETHEMEDIALOGTEXTURE_PTR >(GetProcAddress(m_hmodUxThemeDLL, "EnableThemeDialogTexture"))))
	{
		if (NULL != m_hmodUxThemeDLL)
		{
			FreeLibrary(m_hmodUxThemeDLL);
			m_hmodUxThemeDLL = NULL;
		}
		
		m_fOpenThemeData = NULL;
		m_fCloseThemeData = NULL;
		m_fDrawThemeBackground = NULL;
		m_fDrawThemeParentBackground = NULL;
		m_fGetThemePartSize = NULL;
		m_fEnableThemeDialogTexture = NULL;
	}

	m_sizePart = SIZE();
	m_hbmpDIB = NULL;
	m_hdcMemTemp = NULL;
	m_hBrushTemp = NULL;
}

LeoHelpers::ThemeHelper::~ThemeHelper()
{
	if (m_hBrushTemp != NULL)
	{
		DeleteObject(m_hBrushTemp);
		m_hBrushTemp = NULL;
	}

	if (m_hdcMemTemp != NULL)
	{
		DeleteDC(m_hdcMemTemp);
		m_hdcMemTemp = NULL;
	}

	if (NULL != m_hbmpDIB)
	{
		DeleteObject(m_hbmpDIB);
		m_hbmpDIB = NULL;
	}

	if (NULL != m_hmodUxThemeDLL)
	{
		FreeLibrary(m_hmodUxThemeDLL);
		m_hmodUxThemeDLL = NULL;
	}
}

bool LeoHelpers::ThemeHelper::EnableThemeDialogTexture(HWND hWnd, DWORD dwFlags)
{
	bool bResult = false;

	if (IsAvailable())
	{
		if (SUCCEEDED(m_fEnableThemeDialogTexture(hWnd, dwFlags)))
		{
			bResult = true;
		}
	}

	return(bResult);
}

bool LeoHelpers::ThemeHelper::IsAvailable()
{
	return(NULL != m_hmodUxThemeDLL);
}


#if 0
// Create a test pattern that makes it easy to spot when we're doing something wrong.
bool LeoHelpers::ThemeHelper::CacheBackgroundTabThemeDialog(HWND hWnd)
{
	bool bResult = false;

	if (NULL != m_hbmpDIB)
	{
		DeleteObject(m_hbmpDIB);
		m_hbmpDIB = NULL;
	}

	if (IsAvailable())
	{
		HDC hdcMem = CreateCompatibleDC(NULL);

		if (NULL != hdcMem)
		{
			m_sizePart.cx = 128;
			m_sizePart.cy = 128;

			BITMAPINFO bmi;
			ZeroMemory(&bmi, sizeof(bmi));
			bmi.bmiHeader.biSize = sizeof(bmi.bmiHeader);
			bmi.bmiHeader.biWidth = m_sizePart.cx;
			bmi.bmiHeader.biHeight = m_sizePart.cy * 2;
			bmi.bmiHeader.biPlanes = 1;
			bmi.bmiHeader.biBitCount = 24;
			bmi.bmiHeader.biCompression = BI_RGB;

			void *pDontCare = NULL;

			m_hbmpDIB = CreateDIBSection(hdcMem, &bmi, DIB_RGB_COLORS, &pDontCare, NULL, 0);

			if (NULL != m_hbmpDIB)
			{
				HGDIOBJ hbmpOld = SelectObject(hdcMem, m_hbmpDIB);

				if (NULL != hbmpOld)
				{
					static int iBrushType = HS_BDIAGONAL;

					HBRUSH hBrushNew = CreateHatchBrush(iBrushType, RGB(200,200,200));

					++iBrushType;
					iBrushType %= 6;

					if (hBrushNew != NULL)
					{
					//	HGDIOBJ hBrushOld = SelectObject(hdcMem, hBrushNew);

						RECT rcFill;
						::SetRect(&rcFill, 0, 0, bmi.bmiHeader.biWidth, bmi.bmiHeader.biHeight);

						if (FillRect(hdcMem, &rcFill, hBrushNew))
						{
							bResult = true;
						}

					//	SelectObject(hdcMem, hBrushOld);
					//	DeleteObject(hBrushNew);
					}

					SelectObject(hdcMem, hbmpOld);
				}
			}

			DeleteDC(hdcMem);
		}
	}

	if (!bResult)
	{
		if (NULL != m_hbmpDIB)
		{
			DeleteObject(m_hbmpDIB);
			m_hbmpDIB = NULL;
		}
	}

	return bResult;
}
#else
bool LeoHelpers::ThemeHelper::CacheBackgroundTabThemeDialog(HWND hWnd)
{
	bool bResult = false;

	if (NULL != m_hbmpDIB)
	{
		DeleteObject(m_hbmpDIB);
		m_hbmpDIB = NULL;
	}

	if (IsAvailable())
	{
		HANDLE hTheme = m_fOpenThemeData(hWnd, L"Tab");

		if (NULL != hTheme)
		{
			HDC hdcMem = CreateCompatibleDC(NULL);

			if (NULL != hdcMem)
			{
				if (SUCCEEDED(m_fGetThemePartSize(hTheme, hdcMem, TABP_BODY, 0, NULL, TS_TRUE, &m_sizePart)))
				{
					BITMAPINFO bmi;
					ZeroMemory(&bmi, sizeof(bmi));
					bmi.bmiHeader.biSize = sizeof(bmi.bmiHeader);
					bmi.bmiHeader.biWidth = m_sizePart.cx;
					bmi.bmiHeader.biHeight = m_sizePart.cy * 2;
					bmi.bmiHeader.biPlanes = 1;
					bmi.bmiHeader.biBitCount = 24;
					bmi.bmiHeader.biCompression = BI_RGB;

					void *pDontCare = NULL;

					m_hbmpDIB = CreateDIBSection(hdcMem, &bmi, DIB_RGB_COLORS, &pDontCare, NULL, 0);

					if (NULL != m_hbmpDIB)
					{
						HGDIOBJ hbmpOld = SelectObject(hdcMem, m_hbmpDIB);

						if (NULL != hbmpOld)
						{
							// Get an upside-down copy of the bitmap so that we can draw over every second
							// line of right-way-up bitmaps, so that the stupid thing tiles.

							RECT rectBackgroud;
							rectBackgroud.left	 = 0;
							rectBackgroud.top	 = 0;
							rectBackgroud.right	 = m_sizePart.cx;
							rectBackgroud.bottom = m_sizePart.cy;

							if (SUCCEEDED(m_fDrawThemeBackground(hTheme, hdcMem, TABP_BODY, 0, &rectBackgroud, NULL)))
							{
								StretchBlt(hdcMem, 0, m_sizePart.cy * 2 - 1, m_sizePart.cx, -m_sizePart.cy,
								           hdcMem, 0, 0,                     m_sizePart.cx,  m_sizePart.cy, SRCCOPY);

								bResult = true;
							}

							SelectObject(hdcMem, hbmpOld);
						}
					}
				}

				DeleteDC(hdcMem);
			}

			m_fCloseThemeData(hTheme);
		}
	}

	if (!bResult)
	{
		if (NULL != m_hbmpDIB)
		{
			DeleteObject(m_hbmpDIB);
			m_hbmpDIB = NULL;
		}
	}

	return bResult;
}
#endif

void LeoHelpers::ThemeHelper::EraseBackgroundTabThemeDialog(HWND hWnd, HDC hDC, const RECT &rcClient, const RECT &rcBackgroundWnd) 
{
	if (NULL != m_hbmpDIB || CacheBackgroundTabThemeDialog(hWnd))
	{
		HBITMAP hbmpOld = NULL;

		HDC hdcMem = CreateCompatibleDC(hDC);

		if (NULL != hdcMem)
		{
			HBRUSH hBrushNew = CreatePatternBrush(m_hbmpDIB);

			if (NULL != hBrushNew)
			{
				POINT ptOldOrigin={0};

				// SetBrushOrgEx sets the origin for the next brush that gets selected.
				if (SetBrushOrgEx(hDC, rcBackgroundWnd.left, rcBackgroundWnd.top, &ptOldOrigin))
				{
					HGDIOBJ hBrushOld = SelectObject(hDC, hBrushNew);

					if (hBrushOld != NULL)
					{
						PatBlt(hDC, rcClient.left, rcClient.top, rcClient.right - rcClient.left, rcClient.bottom - rcClient.top, PATCOPY);
					}

					SetBrushOrgEx(hDC, ptOldOrigin.x, ptOldOrigin.y, NULL);

					if (hBrushOld != NULL)
					{
						SelectObject(hDC, hBrushOld);
					}
				}

				DeleteObject(hBrushNew);
			}

			DeleteDC(hdcMem);
		}
	}
	else
	{
		HBRUSH hBrushNotOurs = NULL;

		for (HWND hWndForBackground = hWnd; hWndForBackground != NULL; hWndForBackground = GetParent(hWndForBackground))
		{
			#pragma warning(suppress:4312) // spurious warning due to stupid Win32 SDK header definition.
			hBrushNotOurs = reinterpret_cast<HBRUSH>(::GetClassLongPtr(hWndForBackground, GCLP_HBRBACKGROUND));

			if (hBrushNotOurs != NULL
			||	!(GetWindowLongPtr(hWndForBackground, GWL_STYLE) & WS_CHILD))
			{
				break;
			}
		}

		if (hBrushNotOurs)
		{
			hBrushNotOurs = GetSysColorBrush(COLOR_3DFACE);
		}

		POINT ptOldOrigin={0};

		// SetBrushOrgEx sets the origin for the next brush that gets selected.
		if (hBrushNotOurs != NULL && SetBrushOrgEx(hDC, rcBackgroundWnd.left, rcBackgroundWnd.top, &ptOldOrigin))
		{
			HGDIOBJ hBrushOld = SelectObject(hDC, hBrushNotOurs);

			if (hBrushOld != NULL)
			{
				PatBlt(hDC, rcClient.left, rcClient.top, rcClient.right - rcClient.left, rcClient.bottom - rcClient.top, PATCOPY);
			}

			SetBrushOrgEx(hDC, ptOldOrigin.x, ptOldOrigin.y, NULL);

			if (hBrushOld != NULL)
			{
				SelectObject(hDC, hBrushOld);
			}
		}
	}
}

HBRUSH LeoHelpers::ThemeHelper::GetTempBrush(HWND hWnd, HDC hDC, const RECT &rcClient, const RECT &rcBackgroundWnd)
{
	if (m_hBrushTemp != NULL)
	{
		DeleteObject(m_hBrushTemp);
		m_hBrushTemp = NULL;
	}

	if (m_hdcMemTemp != NULL)
	{
		DeleteDC(m_hdcMemTemp);
		m_hdcMemTemp = NULL;
	}

	if (NULL != m_hbmpDIB || CacheBackgroundTabThemeDialog(hWnd))
	{
		HDC hdcMem = CreateCompatibleDC(hDC);

		if (hdcMem != NULL)
		{
			HBRUSH hBrushNew = CreatePatternBrush(m_hbmpDIB);

			if (NULL != hBrushNew)
			{
				// SetBrushOrgEx sets the origin for the next brush that gets selected.
				if (SetBrushOrgEx(hDC, rcBackgroundWnd.left, rcBackgroundWnd.top, NULL))
				{
					m_hBrushTemp = hBrushNew;
					m_hdcMemTemp = hdcMem;
					hBrushNew = NULL;
					hdcMem = NULL;
				}

				if (hBrushNew != NULL)
				{
					DeleteObject(hBrushNew);
				}
			}

			if (hdcMem != NULL)
			{
				DeleteDC(hdcMem);
			}
		}
	}
	else
	{
		#pragma warning(suppress:4312) // spurious warning due to stupid Win32 SDK header definition.
		HBRUSH hBrushNotOurs = reinterpret_cast<HBRUSH>(::GetClassLongPtr(hWnd, GCLP_HBRBACKGROUND));

		if (hBrushNotOurs == NULL)
		{
			hBrushNotOurs = GetSysColorBrush(COLOR_3DFACE);
		}

		// SetBrushOrgEx sets the origin for the next brush that gets selected.
		if (hBrushNotOurs != NULL && SetBrushOrgEx(hDC, rcBackgroundWnd.left, rcBackgroundWnd.top, NULL))
		{
			return hBrushNotOurs;
		}
	}

	return m_hBrushTemp;
}

/*
BOOL LeoHelpers::ThemeHelper::IsThemeBackgroundPartiallyTransparentButton(HWND hWndControl, int iPartId, int iStateId)
{
	if (!IsAvailable())
	{
		return FALSE;
	}

	const wchar_t szClassList = L"Button";

	if ((GetWindowLongPtr(hWndControl, GWL_STYLE) & BS_DEFPUSHBUTTON))
	{
		szClassList = L"OkButton;Button";
	}

	HTHEME hTheme = m_fOpenThemeData(hWndControl, szClassList);

	if (NULL == hTheme)
	{
		return FALSE;
	}

	BOOL bResult = m_fIsThemeBackgroundPartiallyTransparent(hTheme, iPartId, iStateId);

	m_fCloseThemeData(hTheme);

}
*/

HRESULT LeoHelpers::ThemeHelper::DrawThemeParentBackground(HWND hWnd, HDC hDC, const RECT *prcDrawChildCoords)
{
	if (!IsAvailable())
	{
		return E_NOTIMPL;
	}

	return m_fDrawThemeParentBackground(hWnd, hDC, prcDrawChildCoords);
}

enum { GISCALE_PIXBYTES		= 4 }; // 4 bytes per pixel; preserves alpha channel.

int *createCoeffInt(int nLen, int nNewLen, bool bShrink)
{
	int nSum=0,nSum2;
	int nNorm  = bShrink ? (nNewLen<<12)/nLen : 0x1000;
	int	nDenom = bShrink ? nLen               : nNewLen;

	int *pResult = new(std::nothrow) int[2*((size_t)nLen+1)];

	if (NULL != pResult)
	{
		ZeroMemory(pResult, 2*((size_t)nLen+1)*sizeof(int));

		int *pCoeff = pResult;

		for (int i = 0; i < nLen; i++, pCoeff += 2)
		{
			nSum2 = nSum + nNewLen;

			if (nSum2 > nLen)
			{
				pCoeff[0] = ((nLen-nSum)<<12)/nDenom;
				pCoeff[1] = ((nSum2-nLen)<<12)/nDenom;
				nSum2 -= nLen;
			}
			else
			{
				*pCoeff = nNorm;
				if (nSum2 == nLen)
				{
					pCoeff[1] =- 1;
					nSum2=0;
				}
			}

			nSum = nSum2;
		}

		pResult[ 2*nLen + 0 ] = 0x1000;
		pResult[ 2*nLen + 1 ] = 1;
	}

	return(pResult);
}

bool stretchPreserveAlpha_Shrink(const BYTE *pInBuff,DWORD dwWidth,DWORD dwHeight,BYTE *pOutBuff,DWORD dwNewWidth,DWORD dwNewHeight)
{
	bool bResult = false;

	const BYTE* pLine=pInBuff;
	const BYTE *pInPix;
	BYTE *pOutPix;
	BYTE *pOutLine=pOutBuff;
	// If GISCALE_PIXBYTES is ever not four then the stuff in comments may be needed to long-word align the rows.
	DWORD dwInLn=(GISCALE_PIXBYTES*dwWidth);     // +3)&~3;
	DWORD dwOutLn=(GISCALE_PIXBYTES*dwNewWidth); // +3)&~3;
	BOOL bCrossRow,bCrossCol;
	DWORD x,y,i,ii;
	int *pRowCoeff=0,*pColCoeff=0,*pXCoeff,*pYCoeff;
	DWORD dwBuffLn=GISCALE_PIXBYTES*(dwNewWidth)*sizeof(DWORD),dwTmp;
	DWORD* pdwBuff=0,*pdwCurrLn,*pdwCurrPix,*pdwNextLn,*pdwNextPix;

	pRowCoeff=createCoeffInt(dwWidth,dwNewWidth,true);
	pColCoeff=createCoeffInt(dwHeight,dwNewHeight,true);
	pdwBuff = new(std::nothrow) DWORD[GISCALE_PIXBYTES*2*size_t(dwNewWidth)];

	if (pRowCoeff && pColCoeff && pdwBuff)
	{
		bResult = true;

		pYCoeff=pColCoeff;
		pdwCurrLn=pdwBuff;
		pdwNextLn=pdwBuff+GISCALE_PIXBYTES*size_t(dwNewWidth);

		ZeroMemory(pdwBuff,(size_t)2*dwBuffLn);

		y = 0;
		while (y<dwNewHeight)
		{
			pInPix=pLine;
			pLine+=dwInLn;

			pdwCurrPix=pdwCurrLn;
			pdwNextPix=pdwNextLn;

			x = 0;
			pXCoeff = pRowCoeff;
			bCrossRow = (y < (dwNewHeight - 1) && pYCoeff[1] > 0);
			while(x < dwNewWidth)
			{
				dwTmp = *pXCoeff * *pYCoeff;
				for(i = 0; i < GISCALE_PIXBYTES; i++)
					pdwCurrPix[i] += dwTmp * pInPix[i];
				bCrossCol = (x < (dwNewWidth - 1) && pXCoeff[1] > 0);

				if(bCrossCol)
				{
					dwTmp = pXCoeff[1] * *pYCoeff;
					for(i = 0, ii = GISCALE_PIXBYTES; i < GISCALE_PIXBYTES; i++, ii++)
						pdwCurrPix[ii] += dwTmp * pInPix[i];
				}
				if(bCrossRow)
				{
					dwTmp = *pXCoeff * pYCoeff[1];
					for(i = 0; i < GISCALE_PIXBYTES; i++)
						pdwNextPix[i] += dwTmp * pInPix[i];
					if(bCrossCol)
					{
						dwTmp = pXCoeff[1] * pYCoeff[1];
						for(i = 0, ii = GISCALE_PIXBYTES; i < GISCALE_PIXBYTES; i++, ii++)
							pdwNextPix[ii] += dwTmp * pInPix[i];
					}
				}
				if(pXCoeff[1])
				{
					x++;
					pdwCurrPix += GISCALE_PIXBYTES;
					pdwNextPix += GISCALE_PIXBYTES;
				}
				pXCoeff += 2;
				pInPix += GISCALE_PIXBYTES;
			}
			if(pYCoeff[1])
			{
				// set result line
				pdwCurrPix = pdwCurrLn;
				pOutPix = pOutLine;
				for(i = GISCALE_PIXBYTES * dwNewWidth; i > 0; i--, pdwCurrPix++, pOutPix++)
					*pOutPix = ((LPBYTE)pdwCurrPix)[3];

				// prepare line buffers
				pdwCurrPix = pdwNextLn;
				pdwNextLn = pdwCurrLn;
				pdwCurrLn = pdwCurrPix;

				if (y < (dwNewHeight - 1))
				{
					ZeroMemory(pdwNextLn, dwBuffLn);
				}

				y++;
				pOutLine += dwOutLn;
			}
			pYCoeff += 2;
		}
	}

	delete [] pRowCoeff;
	delete [] pColCoeff;
	delete [] pdwBuff;

	return(bResult);
}

bool stretchPreserveAlpha_Enlarge(const BYTE *pInBuff,DWORD dwWidth,DWORD dwHeight,BYTE *pOutBuff,DWORD dwNewWidth,DWORD dwNewHeight)
{
	return false;

	// Note: Enlarge function shouldn't be needed by any current code. It has not been tested. Since I found several bugs in the shrink function I wouldn't trust this without testing...

	// Opus has version of this that I think I've fixed.
/*
	bool bResult = false;

	LPBYTE pLine=pInBuff,pPix=pLine,pPixOld,pUpPix,pUpPixOld,pOutLine=pOutBuff,pOutPix;
	DWORD dw3 = 3;
	DWORD dwInLn=(GISCALE_PIXBYTES*wWidth+3)&~dw3,dwOutLn=(GISCALE_PIXBYTES*wNewWidth+3)&~dw3;
	BOOL bCrossRow,bCrossCol;
	int x,y,i,*pRowCoeff=0,*pColCoeff=0,*pXCoeff,*pYCoeff;
	DWORD dwTmp,dwPtTmp[GISCALE_PIXBYTES];

	pRowCoeff=createCoeffInt(wNewWidth,wWidth,false);
	pColCoeff=createCoeffInt(wNewHeight,wHeight,false);

	if (pRowCoeff && pColCoeff)
	{
		bResult = true;

		pYCoeff=pColCoeff;

		y = 0;
		while(y < wHeight)
		{
			bCrossRow = pYCoeff[1] > 0;
			x = 0;
			pXCoeff = pRowCoeff;
			pOutPix = pOutLine;
			pOutLine += dwOutLn;
			pUpPix = pLine;
			if(pYCoeff[1])
			{
				y++;
				pLine += dwInLn;
				pPix = pLine;
			}

			while(x < wWidth)
			{
				bCrossCol = pXCoeff[1] > 0;
				pUpPixOld = pUpPix;
				pPixOld = pPix;
				if(pXCoeff[1])
				{
					x++;
					pUpPix += GISCALE_PIXBYTES;
					pPix += GISCALE_PIXBYTES;
				}

				dwTmp = *pXCoeff * *pYCoeff;

				for(i = 0; i < GISCALE_PIXBYTES; i++)
					dwPtTmp[i] = dwTmp * pUpPixOld[i];

				if(bCrossCol)
				{
					dwTmp = pXCoeff[1] * *pYCoeff;
					for(i = 0; i < GISCALE_PIXBYTES; i++)
						dwPtTmp[i] += dwTmp * pUpPix[i];
				}

				if(bCrossRow)
				{
					dwTmp = *pXCoeff * pYCoeff[1];
					for(i = 0; i < GISCALE_PIXBYTES; i++)
						dwPtTmp[i] += dwTmp * pPixOld[i];
					if(bCrossCol)
					{
						dwTmp = pXCoeff[1] * pYCoeff[1];
						for(i = 0; i < GISCALE_PIXBYTES; i++)
							dwPtTmp[i] += dwTmp * pPix[i];
					}
				}

				for(i = 0; i < GISCALE_PIXBYTES; i++, pOutPix++)
					*pOutPix = ((LPBYTE)(dwPtTmp + i))[3];

				pXCoeff += 2;
			}
			pYCoeff += 2;
		}
	}

	delete [] pRowCoeff;
	delete [] pColCoeff;

	return(bResult);
*/
}

// delete[] the result.
// Only shrinking works at the moment. Enlarging is untested and lots of bugs were found in the shrinking code so don't trust it.
RGBQUAD *LeoHelpers::RGBQAllocateStretchPreserveAlpha(DWORD dwNewWidth, DWORD dwNewHeight, DWORD dwInWidth, DWORD dwInHeight, const RGBQUAD *lpSourceBits)
{
	RGBQUAD *lpNewBits = 0;
	const RGBQUAD *lpInBits  = lpSourceBits;

	// Note that there are two separate algorithms, one for expanding an image and one
	// for shrinking an image. If we want to do both, we have to expand first and then shrink :)

	// Expand image?
	if (dwNewWidth>dwInWidth || dwNewHeight>dwInHeight)
	{
		// Get expanded size and allocate a new bitmap for it
		DWORD dww=max(dwNewWidth,dwInWidth);
		DWORD dwh=max(dwNewHeight,dwInHeight);
		lpNewBits = new(std::nothrow) RGBQUAD[ (size_t)dww*dwh ];

		if (!lpNewBits)
		{
			return 0;
		}

		// Stretch into the bitmap
		stretchPreserveAlpha_Enlarge(reinterpret_cast<const BYTE*>(lpInBits),dwInWidth,dwInHeight,
									 reinterpret_cast<BYTE*>(lpNewBits),dww,dwh);

		// Use this as the new 'input' bitmap if we also need to shrink
		lpInBits=lpNewBits;
		dwInWidth=dww;
		dwInHeight=dwh;
	}

	// Shrink image?
	if (dwNewWidth<dwInWidth || dwNewHeight<dwInHeight)
	{
		// Get shrunk size and allocate a new bitmap for it
		DWORD dww=min(dwNewWidth,dwInWidth);
		DWORD dwh=min(dwNewHeight,dwInHeight);

		RGBQUAD *lpShrinkBits = new(std::nothrow) RGBQUAD[ (size_t)dww*dwh ];

		if (!lpShrinkBits)
		{
			delete [] lpNewBits;
			return 0;
		}

		// Shrink into the bitmap
		stretchPreserveAlpha_Shrink(reinterpret_cast<const BYTE*>(lpInBits),dwInWidth,dwInHeight,
									reinterpret_cast<BYTE*>(lpShrinkBits),dww,dwh);

		// Free previous bitmap if we expanded
		delete [] lpNewBits;

		// Use this as the new bitmap
		lpNewBits=lpShrinkBits;
	}

	return lpNewBits;
}

void LeoHelpers::InitRGBQBitmapHeader(BITMAPINFO *pBitmapInfo, int width, int height, bool topDown)
{
	//ZeroMemory(pBitmapInfo, sizeof(BITMAPINFO));

	ZeroMemory(&pBitmapInfo->bmiHeader, sizeof(BITMAPINFOHEADER));

	pBitmapInfo->bmiHeader.biSize			= sizeof(BITMAPINFOHEADER);
	pBitmapInfo->bmiHeader.biWidth			= width;
	pBitmapInfo->bmiHeader.biHeight			= (topDown ? (-height) : height);
	pBitmapInfo->bmiHeader.biPlanes			= 1;
	pBitmapInfo->bmiHeader.biBitCount		= 32;
	pBitmapInfo->bmiHeader.biCompression	= BI_RGB;
	pBitmapInfo->bmiHeader.biSizeImage		= 0;
	pBitmapInfo->bmiHeader.biXPelsPerMeter	= 0;
	pBitmapInfo->bmiHeader.biYPelsPerMeter	= 0;
	pBitmapInfo->bmiHeader.biClrUsed		= 0;
	pBitmapInfo->bmiHeader.biClrImportant	= 0;
/*
	pBitmapInfo->bmiColors[0].rgbRed        = 0;
	pBitmapInfo->bmiColors[0].rgbGreen      = 0;
	pBitmapInfo->bmiColors[0].rgbBlue       = 0;
	pBitmapInfo->bmiColors[0].rgbReserved   = 0;
*/
}

#ifdef DOPUS_PLUGIN_HELPER
#ifndef DOPUS_PLUGIN_LEO_NO_COPYVERSION

bool LeoHelpers::CopyVersionResourceToViewerPluginInfo(HMODULE g_hDllModule, WORD iconId, UINT uiOpusMsgIdName, UINT uiOpusMsgIdDesc, LPVIEWERPLUGININFO lpVPInfo)
{
	bool bResult = false;

	TCHAR szModName[_MAX_PATH + 1];
	DWORD dwVerSize;
	DWORD dwUnused;

	if (NULL != g_hDllModule
	&&	GetModuleFileName(g_hDllModule, szModName, _MAX_PATH)
	&&	(0 != (dwVerSize = GetFileVersionInfoSize(szModName, &dwUnused))))
	{
		BYTE *pVerData = new BYTE[dwVerSize];

		if (GetFileVersionInfo(szModName, NULL, dwVerSize, pVerData))
		{
			LPVOID pBuffer;
			UINT bufLen;

			// Find out language of version resource
			if (VerQueryValue(pVerData, _T("\\VarFileInfo\\Translation"), &pBuffer, &bufLen)
			&&	4 == bufLen)
			{
				BYTE *pLang = reinterpret_cast<BYTE*>(pBuffer);

				if (LeoHelpers::CopyVersionNumber(&lpVPInfo->dwVersionHigh,		&lpVPInfo->dwVersionLow,		_T("FileVersion"),		pLang, pVerData)
				&&	LeoHelpers::CopyVersionString(lpVPInfo->lpszCopyright,		lpVPInfo->cchCopyrightMax,		_T("LegalCopyright"),	pLang, pVerData)
				&&	LeoHelpers::CopyVersionString(lpVPInfo->lpszName,			lpVPInfo->cchNameMax,			_T("InternalName"),		pLang, pVerData)
				&&	LeoHelpers::CopyVersionString(lpVPInfo->lpszDescription,	lpVPInfo->cchDescriptionMax,	_T("FileDescription"),	pLang, pVerData)
				&&	LeoHelpers::CopyVersionString(lpVPInfo->lpszURL,			lpVPInfo->cchURLMax,			_T("CompanyName"),		pLang, pVerData))
				{
					// If we can, use a translated version of the plugin name.
					// This may fail when the GetString support function is not available, in which case we stick with the English strings.

					DOpusPluginHelperUtil *pHelper = NULL;

					if (uiOpusMsgIdName != 0 && lpVPInfo->lpszName != NULL && lpVPInfo->cchNameMax != 0)
					{
						if (!pHelper) { pHelper = new DOpusPluginHelperUtil; }

						pHelper->GetString(uiOpusMsgIdName, &lpVPInfo->lpszName, lpVPInfo->cchNameMax);
					}

					if (uiOpusMsgIdDesc != 0 && lpVPInfo->lpszDescription != NULL && lpVPInfo->cchDescriptionMax != 0)
					{
						if (!pHelper) { pHelper = new DOpusPluginHelperUtil; }

						pHelper->GetString(uiOpusMsgIdDesc, &lpVPInfo->lpszDescription, lpVPInfo->cchDescriptionMax);
					}

#ifndef UNICODE
					LeoHelpers::StringAppend(lpVPInfo->lpszDescription, " (ANSI)", lpVPInfo->cchDescriptionMax);
#endif
					// If anything fails after the next line then the icon must be destroyed.
					if (0 != iconId && lpVPInfo->cbSize >= VIEWERPLUGININFO_V4_SIZE)
					{
						lpVPInfo->hIconSmall = reinterpret_cast<HICON>(LoadImage(g_hDllModule, MAKEINTRESOURCE(iconId), IMAGE_ICON, 0, 0, LR_DEFAULTCOLOR));
					}

					delete pHelper;

					bResult = true;
				}
			}
		}

		delete [] pVerData;
	}
	
	return bResult;
}

bool LeoHelpers::CopyVersionNumber(DWORD *pdwHigh, DWORD *pdwLow, TCHAR *szItemName, BYTE *pLang, const LPVOID pVerData)
{
	bool bResult = false;

	TCHAR szVerPath[128];
	_stprintf_s(szVerPath, 128, _T("\\StringFileInfo\\%02x%02x%02x%02x\\%s"), pLang[1], pLang[0], pLang[3], pLang[2], szItemName);

	LPVOID pBuffer;
	UINT bufLen;
	int v0,v1,v2,v3;

	if (VerQueryValue(pVerData, szVerPath, &pBuffer, &bufLen)
	&&	4 == _stscanf_s(reinterpret_cast<TCHAR *>(pBuffer), _T("%d, %d, %d, %d"), &v0, &v1, &v2, &v3))
	{
		*pdwHigh = ((v0<<16) + v1);
		*pdwLow  = ((v2<<16) + v3);

		bResult = true;
	}

	return(bResult);
}

bool LeoHelpers::CopyVersionString(TCHAR *szDestBuffer, UINT uiDestBufferSize, TCHAR *szItemName, BYTE *pLang, const LPVOID pVerData)
{
	bool bResult = false;

	if (NULL == szDestBuffer)
	{
		bResult = true;
	}
	else if (2 < uiDestBufferSize)
	{
		TCHAR szVerPath[128];
		_stprintf_s(szVerPath, 128, _T("\\StringFileInfo\\%02x%02x%02x%02x\\%s"), pLang[1], pLang[0], pLang[3], pLang[2], szItemName);

		LPVOID pBuffer;
		UINT bufLen;

		if (VerQueryValue(pVerData, szVerPath, &pBuffer, &bufLen))
		{
			_tcsncpy_s(szDestBuffer, uiDestBufferSize, reinterpret_cast<TCHAR *>(pBuffer), _TRUNCATE);

			bResult = true;
		}
	}

	return(bResult);
}

#endif // DOPUS_PLUGIN_LEO_NO_COPYVERSION
#endif // DOPUS_PLUGIN_HELPER

LeoHelpers::FileAndStream::FileAndStream(const wchar_t *szFilePath, HANDLE hAbortEvent, bool bOpenTempCopy)
: m_strFileName(LeoHelpers::GetLastPathPart(szFilePath))
, m_strFilePath(szFilePath)
, m_bDeleteFilePathOnClose(false)
, m_bOpenTempCopyFilePath(bOpenTempCopy)
, m_pStream(NULL)
, m_bNoRandomSeekInStream(false)
, m_bSlowStream(false)
, m_hAbortEvent(hAbortEvent)
, m_bStreamIsFile(false)
{
}

// pStream can be NULL but only if you are never going to call methods which read data or get the steam or file path.
LeoHelpers::FileAndStream::FileAndStream(const wchar_t *szFileName, IStream *pStream, bool bNoRandomSeek, bool bSlow)
: m_strFileName(LeoHelpers::GetLastPathPart(szFileName))
, m_bDeleteFilePathOnClose(false)
, m_bOpenTempCopyFilePath(false)
, m_pStream(pStream)
, m_bNoRandomSeekInStream(bNoRandomSeek)
, m_bSlowStream(bSlow)
, m_hAbortEvent(NULL)
, m_bStreamIsFile(false)
{
	if (m_pStream != NULL)
	{
		m_pStream->AddRef();
	}
}

LeoHelpers::FileAndStream::FileAndStream(const wchar_t *szFileName, const BYTE *pData, UINT cbDataSize)
: m_strFileName(LeoHelpers::GetLastPathPart(szFileName))
, m_bDeleteFilePathOnClose(false)
, m_bOpenTempCopyFilePath(false)
, m_pStream(NULL)
, m_bNoRandomSeekInStream(false)
, m_bSlowStream(false)
, m_hAbortEvent(NULL)
, m_bStreamIsFile(false)
{
	HGLOBAL hGlobal = ::GlobalAlloc(GMEM_MOVEABLE, cbDataSize);

	if (hGlobal != NULL)
	{
		void *pBuffer = ::GlobalLock(hGlobal);

		if (pBuffer != NULL)
		{
			::memcpy_s(pBuffer, cbDataSize, pData, cbDataSize);

			::GlobalUnlock(hGlobal);

			if (S_OK != ::CreateStreamOnHGlobal(hGlobal, TRUE, &m_pStream))
			{
				m_pStream = NULL;
			}
		}
	}
}

LeoHelpers::FileAndStream::~FileAndStream()
{
	if (m_pStream != NULL)
	{
		m_pStream->Release();
		m_pStream = NULL;
	}

	if (m_bDeleteFilePathOnClose && !m_strFilePath.empty())
	{
		for(int i = 0; i < 5; ++i)
		{
			if (DeleteFile(m_strFilePath.c_str()))
			{
				break;
			}

			Sleep(1000);
		}
	}

	m_bDeleteFilePathOnClose = false;
	m_bOpenTempCopyFilePath = false;
	m_bNoRandomSeekInStream = false;
	m_bSlowStream = false;
	m_bStreamIsFile = false;
	m_strFilePath.clear();
	m_strFileName.clear();
	m_hAbortEvent = NULL;
}

// This allows you to set the OpenTempCopy flag after construction.
// Any subsequent call to GetFilePath will result in the file being copied to a temp path if
// it had not already been done. Any preceeding call to GetFilePath may, of course, have alredy
// received the real filename. If the data is in a stream then calling this has no effect as
// GetFilePath will always cause the data to be saved to a temp file regardless of the flag.
void LeoHelpers::FileAndStream::SetOpenTempCopy(bool bOpenTemp)
{
	m_bOpenTempCopyFilePath = bOpenTemp;
}

// The file name may not match the name of the actual file. It is indicative of the original file's name
// but the file path returned by GetFilePath may be a temp-file with a different name. GetFileName should
// be used for display purposes and for any logic which depends on the original file's name.
// Calling GetFileName will not trigger any files or streams to be created.

bool LeoHelpers::FileAndStream::GetFileName(std::basic_string< TCHAR > *pstrFileName)
{
	if (m_strFileName.empty())
	{
		return false;
	}

	*pstrFileName = m_strFileName;
	return true;
}

bool LeoHelpers::FileAndStream::GetFileExtension(std::basic_string< TCHAR > *pstrFileExtension, bool bIncludeDot)
{
	if (m_strFileName.empty())
	{
		return false;
	}

	const TCHAR *szExt = LeoHelpers::GetExtensionPart(bIncludeDot, m_strFileName.c_str());

	if (szExt == NULL)
	{
		return false;
	}

	*pstrFileExtension = szExt;
	return true;
}

bool LeoHelpers::FileAndStream::HasFilePath()
{
	return !m_strFilePath.empty();
}

bool LeoHelpers::FileAndStream::HasStream()
{
	return m_pStream != NULL;
}

bool LeoHelpers::FileAndStream::GetNominalFilePath(std::basic_string< TCHAR > *pstrNominalFilePath)
{
	if (pstrNominalFilePath == NULL)
	{
		return false;
	}

	pstrNominalFilePath->clear();

	if (m_strFileName.empty())
	{
		return false;
	}

	if ((m_bOpenTempCopyFilePath && !m_bDeleteFilePathOnClose)
	||	(m_strFilePath.empty() && m_pStream != NULL))
	{
		return GetNominalTempFilePath(pstrNominalFilePath);
	}

	if (!m_strFilePath.empty())
	{
		*pstrNominalFilePath = m_strFilePath;
		return true;
	}

	return false;
}

bool LeoHelpers::FileAndStream::GetNominalTempFilePath(std::basic_string< TCHAR > *pstrNominalTempFilePath)
{
	if (pstrNominalTempFilePath == NULL)
	{
		return false;
	}

	pstrNominalTempFilePath->clear();

	if (m_strFileName.empty())
	{
		return false;
	}

	// Work out a path that will be similar to the real temp file path if the file ever gets written to disk.
	// We tell OpenTempFileNamePreserveExtension to not really create the file.
	OpenTempFileNamePreserveExtension(NULL, _T("dnvt"), m_strFileName.c_str(), pstrNominalTempFilePath, true);

	return !pstrNominalTempFilePath->empty();
}

// These calls will cause a file or stream to be created if there isn't one already.
// Call HasFilePath and HasStream to see what already exists if you can work with both and want
// to avoid the overhead of conversion.

#ifdef SHCreateStreamOnFile
bool LeoHelpers::FileAndStream::GetFilePath(std::basic_string< TCHAR > *pstrFilePath)
{
	bool bResult = false;

	if (PrepareFileForReading()
	||	WriteStreamToTempFile())
	{
		*pstrFilePath = m_strFilePath;

		bResult = true;
	}
/*
	if (m_bStreamIsFile && m_pStream != NULL)
	{
		m_pStream->Release();
		m_pStream = NULL;

		m_bNoRandomSeekInStream = false;
		m_bSlowStream = false;

		m_bStreamIsFile = false;
	}
*/
	return bResult;
}
#endif

// The stream will not be AddRef'd by this call. It will exist as long as the CFileAndSteam does
// so you should not normally need to AddRef it.
// If a stream is returned it will always be fast and seekable.
// Whenever GetStream is called the stream's position is reset to the start.

#ifdef SHCreateStreamOnFile
bool LeoHelpers::FileAndStream::GetStream(IStream **ppStream)
{
	bool bResult = false;

	if (m_pStream != NULL)
	{
		if (m_bNoRandomSeekInStream || 	m_bSlowStream)
		{
			if (WriteStreamToTempFile() && m_pStream != NULL)
			{
				*ppStream = m_pStream;
				bResult = true;
			}
		}
		else
		{
			*ppStream = m_pStream;
			bResult = true;
		}
	 }
	 else if (PrepareFileForReading())
	 {
		m_bNoRandomSeekInStream = false;
		m_bSlowStream = false;

		if (S_OK != SHCreateStreamOnFile(m_strFilePath.c_str(), STGM_READ|STGM_SHARE_DENY_NONE, &m_pStream)
		||	m_pStream == NULL)
		{
			m_pStream = NULL;
		}
		else
		{
			m_bStreamIsFile = true;

			*ppStream = m_pStream;
			bResult = true;
		}
	}

	if (bResult)
	{
		LARGE_INTEGER liZero;
		liZero.QuadPart = 0;

		if (S_OK != m_pStream->Seek(liZero, STREAM_SEEK_SET, NULL))
		{
			bResult = false;
		}
	 }
	else
	{
		*ppStream = NULL;
	}

	return bResult;
}
#endif

// Call this to get the current IStream, if any.
// Unlike the GetStream() method:
// - The returned IStream *will* be AddRef'd and you must Release it.
// - The returned IStream may not be seekable.
// - If the object contains a filepath that hasn't already been converted into an IStream then that conversion will not happen and the method will fail instead.
// - The returned IStream will not have its position reset to the start of the IStream (since it may not be seekable).

bool LeoHelpers::FileAndStream::GetStreamAddRefNoConversionNoSeekReset(IStream **ppStream, bool *pOutNoRandomSeek, bool *pOutSlow)
{
	*ppStream = m_pStream;

	if (*ppStream)
	{
		(*ppStream)->AddRef();
		*pOutNoRandomSeek = m_bNoRandomSeekInStream;
		*pOutSlow = m_bSlowStream;
		return true;
	}

	*pOutNoRandomSeek = false;
	return false;
}

bool LeoHelpers::FileAndStream::PrepareFileForReading()
{
	bool bResult = false;

	if (!m_strFileName.empty()
	&&	!m_strFilePath.empty())
	{
		if (!m_bOpenTempCopyFilePath || m_bDeleteFilePathOnClose)
		{
			bResult = true;
		}
		else
		{
			// Copy the file to a temporary file and refer to that from now on.

			std::basic_string< TCHAR > strTempFilePath;

			HANDLE hTempFile = OpenTempFileNamePreserveExtension(NULL, _T("dnvt"), m_strFileName.c_str(), &strTempFilePath, false);

			if (hTempFile != INVALID_HANDLE_VALUE)
			{
				CloseHandle(hTempFile);

				if (!LeoHelpers::CopyFile(m_strFilePath.c_str(), strTempFilePath.c_str(), m_hAbortEvent, true, false))
				{
					DeleteFile(strTempFilePath.c_str());
				}
				else
				{
					m_bOpenTempCopyFilePath = false;
					m_bDeleteFilePathOnClose = true;
					m_strFilePath = strTempFilePath;

					bResult = true;
				}
			}
		}
	}

	return bResult;
}

#ifdef SHCreateStreamOnFile
bool LeoHelpers::FileAndStream::WriteStreamToTempFile()
{
	bool bResult = false;

	if (m_pStream != NULL && !m_strFileName.empty())
	{
		std::basic_string< TCHAR > strTempFilePath;

		HANDLE hTempFile = OpenTempFileNamePreserveExtension(NULL, _T("dnvt"), m_strFileName.c_str(), &strTempFilePath, false);

		if (hTempFile != INVALID_HANDLE_VALUE)
		{
			bool bWriteError = false;
			DWORD dwSize;
			BYTE bBuf[8192];

			LARGE_INTEGER liZero;
			ULARGE_INTEGER uliPrevious;

			liZero.QuadPart = 0;
			uliPrevious.QuadPart = 0;

			if (!m_bNoRandomSeekInStream)
			{
				m_pStream->Seek(liZero, STREAM_SEEK_CUR, &uliPrevious);
				m_pStream->Seek(liZero, STREAM_SEEK_SET, NULL);
			}

			while (S_OK == m_pStream->Read(bBuf, sizeof(bBuf), &dwSize))
			{
				if (0 == dwSize)
				{
					// Some IStream implementations return S_OK and set dwSize to zero to indicate
					// the end of the stream. (Some use S_FALSE instead.)
					break;
				}

				if (0 < dwSize)
				{
					DWORD dwTemp = 0;

					if (0 == WriteFile(hTempFile,bBuf,dwSize,&dwTemp,0))
					{
						bWriteError = true;
						break;
					}
				}
			}

			if (!m_bNoRandomSeekInStream)
			{
				// When using STREAM_SEEK_SET the first argument to Seek is treated as unsigned.

				LARGE_INTEGER liPrevious;
				liPrevious.LowPart  = uliPrevious.LowPart;
				liPrevious.HighPart = uliPrevious.HighPart;

				m_pStream->Seek(liPrevious, STREAM_SEEK_SET, NULL);
			}

			// Close temporary file
			CloseHandle(hTempFile);

			if (bWriteError)
			{
				DeleteFile(strTempFilePath.c_str());
			}
			else
			{
				m_bOpenTempCopyFilePath = false;
				m_bDeleteFilePathOnClose = true;
				m_strFilePath = strTempFilePath;

				m_pStream->Release();
				m_bStreamIsFile = false;
				m_bNoRandomSeekInStream = false;
				m_bSlowStream = false;

				if (S_OK != SHCreateStreamOnFile(m_strFilePath.c_str(), STGM_READ|STGM_SHARE_DENY_NONE, &m_pStream))
				{
					m_pStream = NULL;
				}
				else
				{
					m_bStreamIsFile = true;
					bResult = true;
				}
			}
		}
	}

	return bResult;
}
#endif

bool LeoHelpers::GetComInterfaceName(std::basic_string< TCHAR > *pstrOutName, REFIID iid)
{
#ifndef UNICODE
#pragma error("This function hasn't been written/tested for ANSI builds.")
#endif

	if (IID_IBindHost == iid)
	{
		*pstrOutName = L"(Interface) IBindHost";
		return true;
	}

	// This is more or less a replica of ATL's AtlDumpIID function using the registry code in here instead of ATL's.

	bool bResult = false;

	LPOLESTR wszGuid = NULL;

	if (S_OK == ::StringFromCLSID(iid, &wszGuid))
	{
		const TCHAR *szGuid = wszGuid;

		std::basic_string< TCHAR > strNameFromReg;
		std::basic_string< TCHAR > strRegPath = _T("Interface");
		LeoHelpers::AppendPathString(&strRegPath, szGuid);

#ifdef _WIN64
		if ((LeoHelpers::LeetRegQueryStringValue(HKEY_CLASSES_ROOT, strRegPath.c_str(), NULL, 0,               &strNameFromReg)
		||	 LeoHelpers::LeetRegQueryStringValue(HKEY_CLASSES_ROOT, strRegPath.c_str(), NULL, KEY_WOW64_32KEY, &strNameFromReg))
		&&	!strNameFromReg.empty())
#else
		if (LeoHelpers::LeetRegQueryStringValue(HKEY_CLASSES_ROOT, strRegPath.c_str(), NULL, 0, &strNameFromReg)
		&&	!strNameFromReg.empty())
#endif
		{
			*pstrOutName = _T("(Interface) ");
			pstrOutName->append(strNameFromReg);
			bResult = true;
		}
		else
		{
			strNameFromReg.clear();
			strRegPath = _T("CLSID");
			LeoHelpers::AppendPathString(&strRegPath, szGuid);

#ifdef _WIN64
			if ((LeoHelpers::LeetRegQueryStringValue(HKEY_CLASSES_ROOT, strRegPath.c_str(), NULL, 0,               &strNameFromReg)
			||	 LeoHelpers::LeetRegQueryStringValue(HKEY_CLASSES_ROOT, strRegPath.c_str(), NULL, KEY_WOW64_32KEY, &strNameFromReg))
			&&	!strNameFromReg.empty())
#else
			if (LeoHelpers::LeetRegQueryStringValue(HKEY_CLASSES_ROOT, strRegPath.c_str(), NULL, 0, &strNameFromReg)
			&&	!strNameFromReg.empty())
#endif
			{
				*pstrOutName = _T("(CLSID) ");
				pstrOutName->append(strNameFromReg);
				bResult = true;
			}
			else
			{
				*pstrOutName = szGuid;
				bResult = true;
			}
		}

		::CoTaskMemFree(wszGuid);
	}

	return bResult;
}

HRESULT LeoHelpers::IDispGetNamedProperty(IDispatch *pDisp, const wchar_t *szName, VARIANT *pvOut)
{
	HRESULT hr = S_OK;

	DISPID dispId = 0;
	DISPPARAMS dispParams = {0};

	wchar_t *szNameNonConst = const_cast<wchar_t *>(szName);

	// LOCALE_INVARIANT
	// LOCALE_SYSTEM_DEFAULT
	LCID lcid = MAKELCID(MAKELANGID(LANG_ENGLISH, SUBLANG_ENGLISH_US), SORT_DEFAULT);

	if (S_OK == (hr = pDisp->GetIDsOfNames(IID_NULL, &szNameNonConst, 1, lcid, &dispId)))
	{
		hr = pDisp->Invoke(dispId, IID_NULL, lcid, DISPATCH_PROPERTYGET, &dispParams, pvOut, NULL, NULL);
	}

	return hr;
}

HRESULT LeoHelpers::IDispPutNamedProperty(IDispatch *pDisp, const wchar_t *szName, VARIANT *pvIn)
{
	HRESULT hr = S_OK;

	DISPID dispId = 0;

	DISPID dispIdPropertyPut = DISPID_PROPERTYPUT;

	DISPPARAMS dispParams = {0};
	dispParams.rgdispidNamedArgs = &dispIdPropertyPut;
	dispParams.rgvarg = pvIn;
	dispParams.cArgs = 1;
	dispParams.cNamedArgs = 1;

	wchar_t *szNameNonConst = const_cast<wchar_t *>(szName);

	// LOCALE_INVARIANT
	// LOCALE_SYSTEM_DEFAULT
	LCID lcid = MAKELCID(MAKELANGID(LANG_ENGLISH, SUBLANG_ENGLISH_US), SORT_DEFAULT);

	if (S_OK == (hr = pDisp->GetIDsOfNames(IID_NULL, &szNameNonConst, 1, lcid, &dispId)))
	{
		hr = pDisp->Invoke(dispId, IID_NULL, lcid, DISPATCH_PROPERTYPUT, &dispParams, NULL, NULL, NULL);
	}

	return hr;
}

#if 0
LeoHelpers::LowIntegrity::LowIntegrity(HANDLE hThread)
: m_hThread(hThread)
, m_bSuccess(false)
{
	if (m_hThread == NULL)
	{
		m_hThread = ::GetCurrentThread(); // Does not need closing.
	}

	if (::ImpersonateSelf(SecurityImpersonation))
	{
		HANDLE hThreadToken;

		if (::OpenThreadToken(m_hThread, MAXIMUM_ALLOWED, TRUE, &hThreadToken))
		{
			LeoHelpers::CHandleScoper hsThreadToken(hThreadToken);
			hThreadToken = INVALID_HANDLE_VALUE;

			// Low integrity SID
			const wchar_t *szIntegritySid = L"S-1-16-4096"; // 16 = SECURITY_MANDATORY_LABEL_AUTHORITY, 4096 = SECURITY_MANDATORY_LOW_RID
			PSID pIntegritySid = NULL;

			if (::ConvertStringSidToSid(szIntegritySid, &pIntegritySid))
			{
				LeoHelpers::LocalFreeScoper lfsSid(pIntegritySid);

				TOKEN_MANDATORY_LABEL til = {0};
				til.Label.Attributes = SE_GROUP_INTEGRITY | SE_GROUP_INTEGRITY_ENABLED;
				til.Label.Sid        = pIntegritySid;

				if (::SetTokenInformation(hsThreadToken.Get(), TokenIntegrityLevel, &til, sizeof(til) + GetLengthSid(pIntegritySid)))
				{
					if (::SetThreadToken(&m_hThread, hsThreadToken.Get()))
					{
						m_bSuccess = true;
					}
				}
			}

			if (!m_bSuccess)
			{
				::RevertToSelf();
			}
		}
	}
}

LeoHelpers::LowIntegrity::~LowIntegrity()
{
	if (m_bSuccess)
	{
		::SetThreadToken(&m_hThread, NULL);
		::RevertToSelf();
		m_bSuccess = false;
	}
}
#endif

// This requires both a section and a key name. It won't get you a list of names like the real GetPrivateProfileString.
// Blank values will never be returned and are considered errors.
bool LeoHelpers::GetPrivateProfileString(std::basic_string< TCHAR > *pstrResult, const TCHAR *szFileName, const TCHAR *szSectionName, const TCHAR *szKeyName)
{
	if (szFileName    == NULL || szFileName[0]    == _T('\0')
	||	szSectionName == NULL || szSectionName[0] == _T('\0')
	||	szKeyName     == NULL || szKeyName[0]     == _T('\0'))
	{
		return false;
	}

	bool bResult = false;

	DWORD dwBufferSize = 512;
	TCHAR *szBuffer = new(std::nothrow) TCHAR[dwBufferSize];

	while(true)
	{
		DWORD dwRes = ::GetPrivateProfileString(szSectionName, szKeyName, _T(""), szBuffer, dwBufferSize, szFileName);

		if (dwRes == (dwBufferSize - 1))
		{
			delete [] szBuffer;
			dwBufferSize *= 2;
			szBuffer = new(std::nothrow) TCHAR[dwBufferSize];
		}
		else
		{
			if (dwRes != 0 && szBuffer)
			{
				*pstrResult = szBuffer;
				bResult = true;
			}

			break;
		}
	}

	delete [] szBuffer;

	return bResult;
}

#ifdef _INC_PROCESS

LeoHelpers::HousekeepingThread::HousekeepingThread(DWORD dwThreadIntervalMS)
: m_hThread(NULL)
, m_hStopEvent( CreateEvent(NULL, TRUE, FALSE, NULL) )
, m_dwThreadIntervalMS(dwThreadIntervalMS)
{
	InitializeCriticalSection(&m_csStart);
	InitializeCriticalSection(&m_csThreadVariables);
}

// virtual
LeoHelpers::HousekeepingThread::~HousekeepingThread()
{
	Stop();
	if (NULL != m_hStopEvent)
	{
		CloseHandle(m_hStopEvent);
		m_hStopEvent = NULL;
	}

	DeleteCriticalSection(&m_csStart);
	DeleteCriticalSection(&m_csThreadVariables);
}

void LeoHelpers::HousekeepingThread::SetThreadIntervalMS(DWORD dwThreadIntervalMS)
{
	LeoHelpers::CriticalSectionScoper css(&m_csThreadVariables);
	m_dwThreadIntervalMS = dwThreadIntervalMS;
}

DWORD LeoHelpers::HousekeepingThread::GetThreadIntervalMS()
{
	LeoHelpers::CriticalSectionScoper css(&m_csThreadVariables);
	return m_dwThreadIntervalMS;
}

// Start the thread if it is stopping or stopped.
void LeoHelpers::HousekeepingThread::Start()
{
	// Critical section deals with case where two threads are calling Start at once.
	LeoHelpers::CriticalSectionScoper css(&m_csStart);

	if (NULL != m_hThread && WAIT_OBJECT_0 == WaitForSingleObject(m_hStopEvent, 0))
	{
		// The thread is stopping. Wait for it to finish.
		Stop();
	}

	if (NULL == m_hThread)
	{
		// The thread is not running. Start a new one.

		unsigned int uiIgnored;

		m_hThread = reinterpret_cast<HANDLE>(_beginthreadex(NULL, 0, staticThread, this, 0, &uiIgnored));
	}
}

// Stop should never be called by two threads at once and in most cases should only ever be called implicitly by the destructor.
void LeoHelpers::HousekeepingThread::Stop()
{
	if (NULL != m_hThread)
	{
		if (NULL != m_hStopEvent)
		{
			SetEvent(m_hStopEvent);
			WaitForSingleObject(m_hThread, INFINITE);
			ResetEvent(m_hStopEvent);
		}
		CloseHandle(m_hThread);
		m_hThread = NULL; // Thist must be the only way for m_hThread to become null again.
	}
}

// RequestStop should only be called by MainTask to prevent further iterations.
void LeoHelpers::HousekeepingThread::RequestStop()
{
	if (NULL != m_hStopEvent)
	{
		SetEvent(m_hStopEvent);
	}
}

// static
unsigned __stdcall LeoHelpers::HousekeepingThread::staticThread(void *pVoidThis)
{
	LeoHelpers::SetThreadName("Plugin Housekeeping");

//	OutputDebugString(L"Housekeeping start");

	HousekeepingThread *pThis = reinterpret_cast<HousekeepingThread *>(pVoidThis);
	
	if (NULL == pThis->m_hStopEvent) { return 0; }
	
	while(WAIT_TIMEOUT == WaitForSingleObject(pThis->m_hStopEvent, pThis->GetThreadIntervalMS()))
	{
//		OutputDebugString(L"Housekeeping wake");

		pThis->MainTask();

//		OutputDebugString(L"Housekeeping sleep");
	}

//	OutputDebugString(L"Housekeeping end");

	return 0;
}

#endif // _INC_PROCESS

bool LeoHelpers::IsWindowsVersionGreaterOrEqual(DWORD dwMaj, DWORD dwMin, WORD wSPMaj, WORD wSPMin)
{
	OSVERSIONINFOEX osVersionInfoEx = {0};
	osVersionInfoEx.dwOSVersionInfoSize = sizeof(osVersionInfoEx);

	osVersionInfoEx.dwMajorVersion    = dwMaj;  // Windows Server 2003, Windows XP, or Windows 2000
	osVersionInfoEx.dwMinorVersion    = dwMin;  // Windows XP
	osVersionInfoEx.wServicePackMajor = wSPMaj; // Usually don't care about service pack but docs say we...
	osVersionInfoEx.wServicePackMinor = wSPMin; // ...must test it if we test the major version.

	const DWORD dwTypeMask = VER_MAJORVERSION | VER_MINORVERSION | VER_SERVICEPACKMAJOR | VER_SERVICEPACKMINOR;

	DWORDLONG dwlConditionMask = 0;
	dwlConditionMask = ::VerSetConditionMask(dwlConditionMask, VER_MAJORVERSION,     VER_GREATER_EQUAL);
	dwlConditionMask = ::VerSetConditionMask(dwlConditionMask, VER_MINORVERSION,     VER_GREATER_EQUAL);
	dwlConditionMask = ::VerSetConditionMask(dwlConditionMask, VER_SERVICEPACKMAJOR, VER_GREATER_EQUAL);
	dwlConditionMask = ::VerSetConditionMask(dwlConditionMask, VER_SERVICEPACKMINOR, VER_GREATER_EQUAL);

	if (::VerifyVersionInfo(&osVersionInfoEx, dwTypeMask, dwlConditionMask))
	{
		return true;
	}

	return false;
}

bool LeoHelpers::SetClipboard(HWND hwndClipboardOwner, const wchar_t *sz)
{
	bool bResult = false;

	size_t cchBufLen = wcslen(sz) + 1;

	size_t cbBufSize = cchBufLen * sizeof(wchar_t);

	HGLOBAL hGlobalClipMem = ::GlobalAlloc(GMEM_MOVEABLE, cbBufSize);

	if (hGlobalClipMem != NULL)
	{
		wchar_t *pGlobalClipBuffer = reinterpret_cast<wchar_t *>( ::GlobalLock(hGlobalClipMem) );

		if (NULL != pGlobalClipBuffer)
		{
			LeoHelpers::StringCopy(pGlobalClipBuffer, sz, cchBufLen);

			::GlobalUnlock(hGlobalClipMem);

			if (::OpenClipboard(hwndClipboardOwner))
			{
				if (::EmptyClipboard())
				{
					if (NULL != ::SetClipboardData(CF_UNICODETEXT, hGlobalClipMem))
					{
						hGlobalClipMem = NULL; // Clipboard owns it now.
						bResult = true;
					}
				}

				if (!::CloseClipboard())
				{
					bResult = false;
				}
			}
		}

		if (hGlobalClipMem != NULL)
		{
			::GlobalFree(hGlobalClipMem);
			hGlobalClipMem = NULL;
		}
	}

	return bResult;
}

bool LeoHelpers::GetClipboard(HWND hwndClipboardOwner, std::wstring *pstrOut, bool bPrefixUTF16BOM, bool *pbOpenFailed)
{
	bool bResult = false;

	if (pbOpenFailed)
	{
		*pbOpenFailed = false;
	}

	if (!::OpenClipboard(hwndClipboardOwner))
	{
		if (pbOpenFailed)
		{
			*pbOpenFailed = true;
		}
	}
	else
	{
		if (::IsClipboardFormatAvailable(CF_UNICODETEXT))
		{
			HANDLE hGlobalClipMem = ::GetClipboardData(CF_UNICODETEXT);

			if (hGlobalClipMem != NULL)
			{
				const wchar_t *szClipboard = reinterpret_cast<const wchar_t *>( ::GlobalLock(hGlobalClipMem) );

				if (szClipboard)
				{
					const wchar_t szIntelBom[2] = { 0xFEFF, 0 };

					try // We really don't want an exception from std::wstring to stop us closing the clipboard.
					{
						if (bPrefixUTF16BOM && szClipboard[0] != szIntelBom[0])
						{
							*pstrOut += szIntelBom;
						}

						*pstrOut += szClipboard;
						bResult = true;
					}
					catch(...)
					{
					}

					::GlobalUnlock(hGlobalClipMem);
				}
			}
		}

		::CloseClipboard();
	}

	if (!bResult)
	{
		pstrOut->clear();
	}

	return bResult;
}
