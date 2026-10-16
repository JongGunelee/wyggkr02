// src/base/stdafx.h
#ifndef __BASE_STDAFX_H__
#define __BASE_STDAFX_H__

#pragma once

#ifndef _CRT_SECURE_NO_WARNINGS
#define _CRT_SECURE_NO_WARNINGS 1
#endif

// Windows headers
#ifndef VC_EXTRALEAN
#define VC_EXTRALEAN
#endif

#ifdef __cplusplus
#include <windows.h>
#include <tchar.h>
#include <process.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <objbase.h>
#include <stdint.h>
#include <string>

#ifdef __cplusplus
typedef std::basic_string<TCHAR> tstring;
namespace std {
    using tstring = ::tstring;
}
#endif

#ifndef PTRDIFF_MAX
#define PTRDIFF_MAX INTPTR_MAX
#endif

// Standard C++
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <list>
#include <vector>
#include <deque>
#include <map>
#include <set>

// xpr (Must be before def.h to provide OS config)
#include "xpr.h"

// base utility macros (COM_RELEASE, etc.)
#include "def.h"
#else
#include <windows.h>
#include <tchar.h>
#endif

#endif // __BASE_STDAFX_H__
