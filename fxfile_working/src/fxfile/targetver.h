//
// Copyright (c) 2001-2012 Leon Lee author. All rights reserved.
//
//   homepage: http://www.flychk.com
//   e-mail:   mailto:flychk@flychk.com
//
// Use of this source code is governed by a GPLv3 license that can be
// found in the LICENSE file.

#ifndef __FXFILE_TARGET_VER_H__
#define __FXFILE_TARGET_VER_H__ 1
#pragma once

// The following macros define the minimum required platform.  The minimum required platform
// is the earliest version of Windows, Internet Explorer etc. that has the necessary features to run 
// your application.  The macros work by enabling all features available on platform versions up to and 
// including the version specified.

// [Win11 Optimization] Updated to target Windows 10/11 (0x0A00)
// Previous values targeted Windows XP (0x0501) / Windows 98 (0x0410)
// which caused compatibility issues with modern Windows API calls.

#ifndef WINVER
#define WINVER 0x0A00           // Windows 10 / Windows 11
#endif

#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0A00     // Windows 10 / Windows 11
#endif

// Note: _WIN32_WINDOWS is for Win9x-based OS and is no longer relevant.
// Removed the deprecated _WIN32_WINDOWS definition.

#ifndef _WIN32_IE
#define _WIN32_IE 0x0A00        // Internet Explorer 10+
#endif

// [Win11 Optimization] Added NTDDI_VERSION for Windows 10+ API availability
#ifndef NTDDI_VERSION
#define NTDDI_VERSION 0x0A000000  // NTDDI_WIN10
#endif

#endif // __FXFILE_TARGET_VER_H__
