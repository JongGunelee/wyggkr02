//
// Copyright (c) 2012 Leon Lee author. All rights reserved.
//
// Use of this source code is governed by a GPLv3 license that can be
// found in the LICENSE file.

#ifndef __XPR_SYSTEM_H__
#define __XPR_SYSTEM_H__ 1
#pragma once

#include "xpr_types.h"
#include "xpr_dlsym.h"

namespace xpr
{
enum
{
    kOsVerWinUnknown  = 0,

    kOsVerWin9x       = 10,
    kOsVerWin95       = 20,
    kOsVerWin95Osr2   = 21,
    kOsVerWin98       = 30,
    kOsVerWin98SE     = 30,
    kOsVerWinMe       = 40,

    kOsVerWinNT       = 100,
    kOsVerWinNT35     = 110,
    kOsVerWinNT351    = 111,
    kOsVerWinNT4      = 120,
    kOsVerWinNT4Sp2   = 121,
    kOsVerWinNT4Sp3   = 122,
    kOsVerWinNT4Sp4   = 123,
    kOsVerWinNT4Sp5   = 124,
    kOsVerWinNT4Sp6   = 125,
    kOsVerWin2000     = 130,
    kOsVerWin2000Sp1  = 131,
    kOsVerWin2000Sp2  = 132,
    kOsVerWin2000Sp3  = 133,
    kOsVerWin2000Sp4  = 134,
    kOsVerWinXP       = 140,
    kOsVerWinXPSp1    = 141,
    kOsVerWinXPSp2    = 142,
    kOsVerWinXPSp3    = 143,
    kOsVerWinXPx64    = 145,
    kOsVerWin2003     = 150,
    kOsVerWin2003R2   = 151,
    kOsVerWinVista    = 160,
    kOsVerWinVistaSp1 = 161,
    kOsVerWinVistaSp2 = 162,
    kOsVerWinVistaSp3 = 163,
    kOsVerWin2008     = 170,
    kOsVerWin2008R2   = 171,
    kOsVerWin7        = 180,
    kOsVerWin7Sp1     = 181,
    kOsVerWin7Sp2     = 182,
    kOsVerWin8        = 190,
    kOsVerWin2012     = 200,

    // [Win11 Optimization] Added Windows 8.1, 10, 11 version constants
    kOsVerWin8_1      = 210,
    kOsVerWin2012R2   = 211,
    kOsVerWin10       = 220,
    kOsVerWin10_1511  = 221,  // November Update (TH2)
    kOsVerWin10_1607  = 222,  // Anniversary Update (RS1)
    kOsVerWin10_1703  = 223,  // Creators Update (RS2)
    kOsVerWin10_1709  = 224,  // Fall Creators Update (RS3)
    kOsVerWin10_1803  = 225,  // April 2018 Update (RS4)
    kOsVerWin10_1809  = 226,  // October 2018 Update (RS5)
    kOsVerWin10_1903  = 227,  // May 2019 Update (19H1)
    kOsVerWin10_1909  = 228,  // November 2019 Update
    kOsVerWin10_2004  = 229,  // May 2020 Update
    kOsVerWin10_21H1  = 230,
    kOsVerWin10_21H2  = 231,
    kOsVerWin10_22H2  = 232,
    kOsVerWin11       = 240,
    kOsVerWin11_22H2  = 241,
    kOsVerWin11_23H2  = 242,
    kOsVerWin11_24H2  = 243,
    kOsVerWin2022     = 250,

    kOsVerWinHigher,
};

struct SystemInfo
{
    xpr_uint_t mOsVer;
};

XPR_DL_API xpr_rcode_t initSystemInfo(void);

XPR_DL_API SystemInfo *getSystemInfo(void);
XPR_DL_API xpr_uint_t  getOsVer(void);

} // namespace xpr

#endif // __XPR_SYSTEM_H__
