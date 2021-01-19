// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently, but
// are changed infrequently
//

#pragma once

#define WINVER         0x0500	// Win2k and above
#define _WIN32_WINNT   0x0500	// Win2k and above
#define _WIN32_WINDOWS 0x0410	// Win98 and above
#define _WIN32_IE      0x0600	// IE 6.0 and above

//#define WIN32_LEAN_AND_MEAN		// Exclude rarely-used stuff from Windows headers

// Windows Header Files:
#include <windows.h>
#include <commctrl.h>
#include <tchar.h>

#include <string>
#include <algorithm>
#include <vector>
#include <list>
#include <map>

#include <assert.h>

#ifndef WM_MOUSEWHEEL
#define WM_MOUSEWHEEL                   0x020A
#endif

#ifndef WM_UNICHAR
#define WM_UNICHAR                      0x0109
#endif

#include "..\common\viewer plugins.h"
#include "..\common\plugin support.h"
#include "..\common\messages.hpp"
#include "resource.h"
