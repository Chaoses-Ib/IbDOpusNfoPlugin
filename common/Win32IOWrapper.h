#pragma once

// Define this to write a debug file that shows you which bits of the input were read.
// It's more useful for debugging code which uses the wrapper rather than for debugging the wrapper itself.
//#define NJR_FILEWRAPDEBUG

class Win32IOWrapper
{
protected:
	bool m_bThrowExceptions;

public:
	static Win32IOWrapper *CreateFromFAS(LeoHelpers::FileAndStream *pFAS, bool bNeedSeek);

	Win32IOWrapper();
	virtual ~Win32IOWrapper();

private:
	// disallow
	Win32IOWrapper(const Win32IOWrapper &rhs);
	Win32IOWrapper &operator=(const Win32IOWrapper &rhs);

public:
	virtual size_t read(void *ptr, size_t byteCount) = 0;
	virtual int    seek(long offset, int whence) = 0;
	virtual long   tell() = 0;
	virtual bool   tell64(unsigned __int64 *pOut) = 0;
	virtual bool   size64(unsigned __int64 *pOut) = 0;

	virtual int seekOrReadForward(long offset)
	{
		if (offset < 0)
		{
			if (m_bThrowExceptions) { throw "read failed"; }
			return -1;
		}

		// Even when they can't random-seek, Opus streams can always seek forwards
		// (though it may be slow due to having to decompress/download everything on the way).
		return seek(offset, SEEK_CUR);

	/*
		if (CanFastRandomSeek())
		{
			return seek(offset, SEEK_CUR);
		}
		if () 
		{
			return 0;
		}
		else
		{
			BYTE buf[1024];

			while(offset > _countof(buf))
			{
				if (_countof(buf) != read(buf, _countof(buf)))
				{
					if (m_bThrowExceptions) { throw "read failed"; }
					return -1;
				}
				offset -= _countof(buf);
			}

			if (offset > 0)
			{
				if (offset != read(buf, offset))
				{
					if (m_bThrowExceptions) { throw "read failed"; }
					return -1;
				}
			}

			return 0;
		}
	*/
	}

	template< typename T > bool readValue(T *pt)
	{
		if (sizeof(T) != read(pt, sizeof(T)))
		{
			errno = 1;
			if (m_bThrowExceptions) { throw "readValue failed"; }
			return false;
		}
		return true;
	}

	template< typename T > bool readValueSwap(T *pt)
	{
		T t;

		assert(sizeof(T)==4);

		if (sizeof(T) != 4
		||	sizeof(T) != read(&t, sizeof(T)))
		{
			errno = 1;
			if (m_bThrowExceptions) { throw "readValueSwap failed"; }
			return false;
		}

		// Swap endian-ness.
		*pt = ((t<<24)&0xFF000000)
			+ ((t<<8 )&0x00FF0000)
			+ ((t>>8 )&0x0000FF00)
			+ ((t>>24)&0x000000FF);

		return true;
	}


//	virtual int    getc();
//	virtual char * gets(char *string, int n);
//	virtual int    getInt();
//	virtual float  getFloat();

	bool GetThrowExceptions();
	void SetThrowExceptions(bool bThrow);
	virtual bool CanFastRandomSeek() = 0;
};

class Win32IOStreamWrapper : public Win32IOWrapper
{
protected:
	IStream *m_pStream;
	bool m_bFastRandomSeek;

public:
	Win32IOStreamWrapper(IStream *pStream, bool bFastRandomSeek);
	virtual ~Win32IOStreamWrapper();

	virtual size_t read(void *ptr, size_t byteCount);
	virtual int    seek(long offset, int whence);
	virtual long   tell();
	virtual bool   tell64(unsigned __int64 *pOut);
	virtual bool   size64(unsigned __int64 *pOut);

	virtual bool CanFastRandomSeek();
};

class Win32IOFileWrapper : public Win32IOWrapper
{
protected:
	HANDLE m_hFile;
	HANDLE m_hAbortEvent;

#ifdef NJR_FILEWRAPDEBUG
	HANDLE m_hDebugFile;
#endif

public:
	Win32IOFileWrapper(const wchar_t *szName, HANDLE hAbortEvent);
	virtual ~Win32IOFileWrapper();

	virtual size_t read(void *ptr, size_t byteCount);
	virtual int    seek(long offset, int whence);
	virtual long   tell();
	virtual bool   tell64(unsigned __int64 *pOut);
	virtual bool   size64(unsigned __int64 *pOut);

	virtual bool CanFastRandomSeek();

	bool IsHandleValid();
};

/*
Win32IOJpegSource works but isn't needed anymore.

class Win32IOJpegSource : public jpeg_source_mgr
{
protected:
	enum { WIJS_BUFFER_SIZE = 32768 };

	Win32IOWrapper *m_pFile;
	bool m_bStartOfFile;
	JOCTET *m_buffer;

protected:
	static void impl_init_source(j_decompress_ptr cinfo);
	static boolean impl_fill_input_buffer(j_decompress_ptr cinfo);
	static void impl_skip_input_data(j_decompress_ptr cinfo, long num_bytes);
	static void impl_term_source(j_decompress_ptr cinfo);

public:
	Win32IOJpegSource(Win32IOWrapper *pFile);
	virtual ~Win32IOJpegSource();
};
*/