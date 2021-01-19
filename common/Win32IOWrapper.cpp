#include "stdafx.h"
#include "LeoHelpers.h"
#include "Win32IOWrapper.h"

// static
Win32IOWrapper *Win32IOWrapper::CreateFromFAS(LeoHelpers::FileAndStream *pFAS, bool bNeedRandomSeek)
{
	if (pFAS == NULL)
	{
		return NULL;
	}

	if (pFAS->HasFilePath())
	{
		std::wstring strFilePath;

		if (pFAS->GetFilePath(&strFilePath) && !strFilePath.empty())
		{
			Win32IOFileWrapper *pFile = new Win32IOFileWrapper(strFilePath.c_str(), pFAS->GetAbortEvent());

			if (pFile->IsHandleValid())
			{
				return pFile;
			}

			delete pFile;
			return NULL;
		}
	}

	if (pFAS->HasStream())
	{
		if (bNeedRandomSeek)
		{
			IStream *pStreamNoRef = 0;
			
			if (pFAS->GetStream(&pStreamNoRef) && pStreamNoRef != NULL)
			{
				return new Win32IOStreamWrapper(pStreamNoRef, true);
			}
		}
		else
		{
			IStream *pStreamWithRef = 0;
			bool bNoRandomSeek = false;
			bool bSlow = false;

			if (pFAS->GetStreamAddRefNoConversionNoSeekReset(&pStreamWithRef, &bNoRandomSeek, &bSlow) && pStreamWithRef != NULL)
			{
				Win32IOWrapper *pWrapper = new Win32IOStreamWrapper(pStreamWithRef, !bNoRandomSeek && !bSlow);

				pStreamWithRef->Release();

				return pWrapper;
			}
		}
	}

	return NULL;
}

Win32IOWrapper::Win32IOWrapper()
: m_bThrowExceptions(false)
{
}

Win32IOWrapper::~Win32IOWrapper()
{
}

bool Win32IOWrapper::GetThrowExceptions()
{
	return m_bThrowExceptions;
}

void Win32IOWrapper::SetThrowExceptions(bool bThrow)
{
	m_bThrowExceptions = bThrow;
}

/*
int Win32IOWrapper::getc()
{
	char c;
	if (1 != read(&c, sizeof(char)))
	{
		return EOF;
	}
	return static_cast<unsigned char>(c);
}

char *Win32IOWrapper::gets(char *string, int n)
{
	char *result = string;

	if (0 >= n)
	{
		if (m_bThrowExceptions) { throw "gets given empty buffer"; }
		result = NULL;
	}
	else
	{
		char c;
		char *p = string;

		while(true)
		{
			if (1 == n
			||	1 != read(&c, sizeof(char)))
			{
				// If read() failed then an exception has already been thrown if they're on.
				if (m_bThrowExceptions) { throw "gets buffer overflow"; }
				result = NULL;
				break;
			}

			if (c == '\r' || c == '\n')
			{
				*p++ = '\n'; n--;
				break;
			}

			*p++ = c; n--;
		}

		*p = '\0';
	}

	return result;
}

int Win32IOWrapper::getInt()
{
	std::string tempString;
	long seekCorrection = 0;

	while(true)
	{
		bool readError = false;
		char c = '\0';

		try
		{
			if (1 != read(&c, sizeof(char)))
			{
				readError = true;
			}
		}
		catch(...)
		{
			readError = true;
		}

		if (readError)
		{
			break;
		}

		if (isdigit(c))
		{
			seekCorrection = 0;
			tempString += c;
		}
		else if (tempString.empty() && (c == '-' || c == '+'))
		{
			seekCorrection -= 1;
			tempString += c;
		}
		else
		{
			seekCorrection -= 1;
			seek(seekCorrection, SEEK_CUR);
			break;
		}
	}

	if (tempString.empty())
	{
		if (m_bThrowExceptions)
		{
			throw "Could not read integer from input";
		}
		else
		{
			return 0;
		}
	}
	else
	{
		return atoi(tempString.c_str());
	}
}

float Win32IOWrapper::getFloat()
{
	std::string tempString;
	long seekCorrection = 0;

	char lastChar = '\0';
	bool reachedExponent = false;
	bool reachedDecmial = false;

	// [sign] [digits] [.digits] [ {d | D | e | E }[sign]digits]

	while(true)
	{
		bool readError = false;
		char c = '\0';

		try
		{
			if (1 != read(&c, sizeof(char)))
			{
				readError = true;
			}
		}
		catch(...)
		{
			readError = true;
		}

		if (readError)
		{
			break;
		}

		if (isdigit(c))
		{
			seekCorrection = 0;
			tempString += c;
			lastChar = c;
		}
		else if (c == '.' && !reachedDecmial && !reachedExponent)
		{
			seekCorrection -= 1;
			tempString += c;
			lastChar = c;
			reachedDecmial = true;
		}
		else if ((c == 'd' || c == 'D' || c == 'e' || c == 'E') && !reachedExponent && isdigit(lastChar))
		{
			seekCorrection -= 1;
			tempString += c;
			lastChar = c;
			reachedExponent = true;
		}
		else if ((c == '-' || c == '+') && (lastChar == '\0' || lastChar == 'd' || lastChar == 'D' || lastChar == 'e' || lastChar == 'E'))
		{
			seekCorrection -= -1;
			tempString += c;
			lastChar = c;
		}
		else
		{
			seekCorrection -= 1;
			seek(seekCorrection, SEEK_CUR);
			break;
		}
	}

	if (tempString.empty())
	{
		if (m_bThrowExceptions)
		{
			throw "Could not read float from input";
		}
		else
		{
			return 0.0f;
		}
	}
	else
	{
		return static_cast<float>(atof(tempString.c_str()));
	}
}
*/

Win32IOStreamWrapper::Win32IOStreamWrapper(IStream *pStream, bool bFastRandomSeek)
: m_pStream(pStream), m_bFastRandomSeek(bFastRandomSeek)
{
	m_pStream->AddRef();
}

Win32IOStreamWrapper::~Win32IOStreamWrapper()
{
	if (m_pStream)
	{
		m_pStream->Release();
		m_pStream = NULL;
	}
}

size_t Win32IOStreamWrapper::read(void *ptr, size_t byteCount)
{
	if (byteCount <= ULONG_MAX)
	{
		ULONG ulBytesToRead = static_cast<ULONG>(byteCount);

		ULONG ulBytesRead = 0;

		if (S_OK == m_pStream->Read(ptr, ulBytesToRead, &ulBytesRead))
		{
			errno = 0;
			return ulBytesRead;
		}
	}

	errno = 1;
	if (m_bThrowExceptions) { throw "read failed"; }
	return 0;
}

int Win32IOStreamWrapper::seek(long offset, int whence)
{
	// Even when they can't random-seek, Opus streams can always seek forwards
	// (though it may be slow due to having to decompress/download everything on the way).
	if (m_bFastRandomSeek || (offset >= 0 && whence == SEEK_CUR))
	{
		LARGE_INTEGER liOffset;
		liOffset.QuadPart = offset;

		DWORD dwOrigin;

		switch(whence)
		{
		default:
		case SEEK_SET:	dwOrigin = STREAM_SEEK_SET;	break;
		case SEEK_CUR:	dwOrigin = STREAM_SEEK_CUR;	break;
		case SEEK_END:	dwOrigin = STREAM_SEEK_END;	break;
		}

		if (S_OK == m_pStream->Seek(liOffset, dwOrigin, NULL))
		{
			return 0; // success
		}
	}

	if (m_bThrowExceptions) { throw "seek failed"; }
	return -1; // error
}

long Win32IOStreamWrapper::tell()
{
	// Even when they can't random-seek, Opus streams can always seek forwards
	// (though it may be slow due to having to decompress/download everything on the way).
//	if (m_bFastRandomSeek)
	{
		LARGE_INTEGER liSeek;
		liSeek.QuadPart = 0;

		ULARGE_INTEGER uliPos;
		uliPos.QuadPart = 0;

		if (S_OK == m_pStream->Seek(liSeek, STREAM_SEEK_CUR, &uliPos))
		{
			if (uliPos.HighPart == 0 && static_cast<long>(uliPos.LowPart) >= 0)
			{
				return uliPos.LowPart; // success
			}
		}
	}

	if (m_bThrowExceptions) { throw "tell failed"; }
	return -1; // error
}

bool Win32IOStreamWrapper::tell64(unsigned __int64 *pOut)
{
	// Even when they can't random-seek, Opus streams can always seek forwards
	// (though it may be slow due to having to decompress/download everything on the way).
//	if (m_bFastRandomSeek)
	{
		LARGE_INTEGER liSeek;
		liSeek.QuadPart = 0;

		ULARGE_INTEGER uliPos;
		uliPos.QuadPart = 0;

		if (S_OK == m_pStream->Seek(liSeek, STREAM_SEEK_CUR, &uliPos))
		{
			*pOut = uliPos.QuadPart;
			return true;
		}
	}

	*pOut = -1;
	if (m_bThrowExceptions) { throw "tell64 failed"; }
	return false; // error
}

bool Win32IOStreamWrapper::size64(unsigned __int64 *pOut)
{
	STATSTG statstg = {0};

	if (S_OK == m_pStream->Stat(&statstg, STATFLAG_NONAME))
	{
		*pOut = statstg.cbSize.QuadPart;
		return true;
	}

	*pOut = -1;
	if (m_bThrowExceptions) { throw "size64 failed"; }
	return false; // error
}

bool Win32IOStreamWrapper::CanFastRandomSeek()
{
	return m_bFastRandomSeek;
}




Win32IOFileWrapper::Win32IOFileWrapper(const wchar_t *szName, HANDLE hAbortEvent)
: m_hAbortEvent(hAbortEvent)
{
	m_hFile = ::CreateFile(szName, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING,
								FILE_ATTRIBUTE_NORMAL|FILE_FLAG_SEQUENTIAL_SCAN, NULL);

#ifdef NJR_FILEWRAPDEBUG
	if (INVALID_HANDLE_VALUE == m_hFile
	||	NULL != strstr(szName, "_debug")
	||	NULL == strstr(szName, ".bmp"))
	{
		m_hDebugFile = INVALID_HANDLE_VALUE;
	}
	else
	{
		std::string strDebugName = szName;
		strDebugName += "_debug";

		m_hDebugFile = ::CreateFile(strDebugName.c_str(), GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);

		if (INVALID_HANDLE_VALUE != m_hDebugFile)
		{
			BYTE zeros[1024];
			ZeroMemory(zeros, 1024);

			DWORD dwSize = ::GetFileSize(m_hFile, NULL);
			DWORD dwDontCare = 0;

			while (dwSize > 0)
			{
				DWORD dwToWrite = 1024;
				if (dwToWrite > dwSize)
				{
					dwToWrite = dwSize;
				}
				::WriteFile(m_hDebugFile, zeros, dwToWrite, &dwDontCare, NULL);
				dwSize -= dwToWrite;
			}

			::SetFilePointer(m_hDebugFile, 0, NULL, SEEK_SET);
		}
	}
#endif
}

Win32IOFileWrapper::~Win32IOFileWrapper()
{
	if (INVALID_HANDLE_VALUE != m_hFile)
	{
		CloseHandle(m_hFile);
		m_hFile = INVALID_HANDLE_VALUE;
	}

#ifdef NJR_FILEWRAPDEBUG
	if (INVALID_HANDLE_VALUE != m_hDebugFile)
	{
		CloseHandle(m_hDebugFile);
		m_hDebugFile = INVALID_HANDLE_VALUE;
	}
#endif
}

bool Win32IOFileWrapper::IsHandleValid()
{
	return (INVALID_HANDLE_VALUE != m_hFile);
}

size_t Win32IOFileWrapper::read(void *ptr, size_t byteCount)
{
	if (INVALID_HANDLE_VALUE != m_hFile && byteCount <= ULONG_MAX)
	{
		DWORD dwBytesToRead = static_cast<DWORD>(byteCount);
		DWORD dwBytesRead = 0;

		if (NULL != m_hAbortEvent && WAIT_OBJECT_0 == WaitForSingleObject(m_hAbortEvent, 0))
		{
			if (m_bThrowExceptions) { throw "abort event"; }
			return 0; // Indicate a read error.
		}

		if (::ReadFile(m_hFile, ptr, dwBytesToRead, &dwBytesRead, NULL))
		{
#ifdef NJR_FILEWRAPDEBUG
			if (INVALID_HANDLE_VALUE != m_hDebugFile)
			{
				DWORD dwDontCare = 0;
				::WriteFile(m_hDebugFile, ptr, dwBytesRead, &dwDontCare, NULL);
			}
#endif
			errno = 0;
			return dwBytesRead; // If dwBytesRead==0 then we are signalling the end of the input.
		}
	}

#ifdef NJR_FILEWRAPDEBUG
	MessageBeep(-1);
#endif

	errno = 1;
	if (m_bThrowExceptions) { throw "read failed"; }
	return 0; // Indicate a read error.
}

int Win32IOFileWrapper::seek(long offset, int whence)
{
	if (INVALID_HANDLE_VALUE != m_hFile)
	{
		if (NULL != m_hAbortEvent && WAIT_OBJECT_0 == WaitForSingleObject(m_hAbortEvent, 0))
		{
			if (m_bThrowExceptions) { throw "abort event"; }
			return -1; // error
		}

		DWORD dwMoveMethod;

		switch(whence)
		{
		default:
		case SEEK_SET:	dwMoveMethod = FILE_BEGIN;		break;
		case SEEK_CUR:	dwMoveMethod = FILE_CURRENT;	break;
		case SEEK_END:	dwMoveMethod = FILE_END;		break;
		}

		DWORD dwResult = ::SetFilePointer(m_hFile, offset, NULL, dwMoveMethod);

		if (INVALID_SET_FILE_POINTER != dwResult || NO_ERROR == GetLastError())
		{
#ifdef NJR_FILEWRAPDEBUG
			if (INVALID_HANDLE_VALUE != m_hDebugFile)
			{
				::SetFilePointer(m_hDebugFile, offset, NULL, dwMoveMethod);
			}
#endif
			return 0; // success
		}
	}

#ifdef NJR_FILEWRAPDEBUG
	MessageBeep(-1);
#endif

	if (m_bThrowExceptions) { throw "seek failed"; }
	return -1; // error
}

long Win32IOFileWrapper::tell()
{
	if (INVALID_HANDLE_VALUE != m_hFile)
	{
		if (NULL != m_hAbortEvent && WAIT_OBJECT_0 == WaitForSingleObject(m_hAbortEvent, 0))
		{
			if (m_bThrowExceptions) { throw "abort event"; }
			return -1; // Error.
		}

		DWORD dwResult = ::SetFilePointer(m_hFile, 0, NULL, FILE_CURRENT);

		if (INVALID_SET_FILE_POINTER != dwResult && static_cast<long>(dwResult) >= 0)
		{
			return dwResult; // Success
		}
	}

#ifdef NJR_FILEWRAPDEBUG
	MessageBeep(-1);
#endif

	if (m_bThrowExceptions) { throw "tell failed"; }
	return -1; // Error.
}

bool Win32IOFileWrapper::tell64(unsigned __int64 *pOut)
{
	if (INVALID_HANDLE_VALUE != m_hFile)
	{
		if (NULL != m_hAbortEvent && WAIT_OBJECT_0 == WaitForSingleObject(m_hAbortEvent, 0))
		{
			if (m_bThrowExceptions) { throw "abort event"; }
			return false; // Error.
		}

		LARGE_INTEGER liDistance;
		liDistance.QuadPart = 0;

		LARGE_INTEGER liPosition;
		liPosition.QuadPart = 0;

		// It's not clear why Win32 uses *signed* 64-bit values for file sizes. Maybe it can have negative size anti-files?
		if (::SetFilePointerEx(m_hFile, liDistance, &liPosition, FILE_CURRENT)
		&&	liPosition.QuadPart >= 0)
		{
			*pOut = liPosition.QuadPart;
			return true;
		}
	}

#ifdef NJR_FILEWRAPDEBUG
	MessageBeep(-1);
#endif

	*pOut = -1;
	if (m_bThrowExceptions) { throw "tell64 failed"; }
	return false; // Error.
}

bool Win32IOFileWrapper::size64(unsigned __int64 *pOut)
{
	if (INVALID_HANDLE_VALUE != m_hFile)
	{
		if (NULL != m_hAbortEvent && WAIT_OBJECT_0 == WaitForSingleObject(m_hAbortEvent, 0))
		{
			if (m_bThrowExceptions) { throw "abort event"; }
			return false; // Error.
		}

		LARGE_INTEGER liSize;
		liSize.QuadPart = 0;

		// It's not clear why Win32 uses *signed* 64-bit values for file sizes. Maybe it can have negative size anti-files?
		if (::GetFileSizeEx(m_hFile, &liSize)
		&&	liSize.QuadPart >= 0)
		{
			*pOut = liSize.QuadPart;
			return true;
		}
	}

#ifdef NJR_FILEWRAPDEBUG
	MessageBeep(-1);
#endif

	*pOut = -1;
	if (m_bThrowExceptions) { throw "size64 failed"; }
	return false; // Error.
}

bool Win32IOFileWrapper::CanFastRandomSeek()
{
	return true;
}

/*
Win32IOJpegSource works but isn't needed anymore.

void Win32IOJpegSource::impl_init_source(j_decompress_ptr cinfo)
{
	Win32IOJpegSource *pThis = reinterpret_cast<Win32IOJpegSource *>(cinfo->src);

	pThis->m_bStartOfFile = true;
}

boolean Win32IOJpegSource::impl_fill_input_buffer(j_decompress_ptr cinfo)
{
	Win32IOJpegSource *pThis = reinterpret_cast<Win32IOJpegSource *>(cinfo->src);

	size_t nbytes;

	try
	{
		nbytes = pThis->m_pFile->read(pThis->m_buffer, sizeof(JOCTET), WIJS_BUFFER_SIZE);
	}
	catch(...)
	{
		nbytes = 0;
	}

	if (nbytes <= 0)
	{
		if (pThis->m_bStartOfFile)
		{
			ERREXIT(cinfo, JERR_INPUT_EMPTY);
		}
		WARNMS(cinfo,JWRN_JPEG_EOF);
		pThis->m_buffer[0]=(JOCTET)0xFF;
		pThis->m_buffer[1]=(JOCTET)JPEG_EOI;
		nbytes=2;
	}
	else
	{
		_swab(reinterpret_cast<char *>(pThis->m_buffer), reinterpret_cast<char *>(pThis->m_buffer), nbytes);
	}

	pThis->next_input_byte = pThis->m_buffer;
	pThis->bytes_in_buffer = nbytes;
	pThis->m_bStartOfFile = false;

	return TRUE;
}

void Win32IOJpegSource::impl_skip_input_data(j_decompress_ptr cinfo, long num_bytes)
{
	Win32IOJpegSource *pThis = reinterpret_cast<Win32IOJpegSource *>(cinfo->src);

	if (num_bytes > 0)
	{
		while (static_cast<size_t>(num_bytes) > pThis->bytes_in_buffer)
		{
			num_bytes -= static_cast<size_t>(pThis->bytes_in_buffer);
			impl_fill_input_buffer(cinfo);
		}
		pThis->next_input_byte += num_bytes;
		pThis->bytes_in_buffer -= num_bytes;
	}
}

void Win32IOJpegSource::impl_term_source(j_decompress_ptr cinfo)
{
}

Win32IOJpegSource::Win32IOJpegSource(Win32IOWrapper *pFile)
: m_pFile(pFile)
, m_bStartOfFile(true)
{
	m_buffer = new JOCTET[ WIJS_BUFFER_SIZE ];

	next_input_byte = NULL;
	bytes_in_buffer = 0;

	init_source = impl_init_source;
	fill_input_buffer = impl_fill_input_buffer;
	skip_input_data = impl_skip_input_data;
	resync_to_restart = jpeg_resync_to_restart;
	term_source = impl_term_source;
}

Win32IOJpegSource::~Win32IOJpegSource()
{
	delete [] m_buffer;
}
*/
