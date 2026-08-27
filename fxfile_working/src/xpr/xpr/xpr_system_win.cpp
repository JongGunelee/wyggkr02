//
// Copyright (c) 2012 Leon Lee author. All rights reserved.
//
// Use of this source code is governed by a GPLv3 license that can be
// found in the LICENSE file.

#include "xpr_system.h"
#include "xpr_rcode.h"
#include "xpr_char.h"

namespace xpr
{
namespace
{
SystemInfo gSystemInfo;
} // namespace anonymous

XPR_INLINE xpr_rcode_t initOsVer(xpr_uint_t &aOsVer);

XPR_DL_API xpr_rcode_t initSystemInfo(void)
{
    xpr_rcode_t sRcode = initOsVer(gSystemInfo.mOsVer);
    if (sRcode != XPR_RCODE_SUCCESS)
        return sRcode;

    return XPR_RCODE_SUCCESS;
}

// [Win11 Optimization] Complete rewrite of OS version detection.
// GetVersionEx is deprecated since Win 8.1 and lies about OS version.
// RtlGetVersion from ntdll.dll always returns the true version.
typedef LONG (WINAPI *PRtlGetVersion)(OSVERSIONINFOEXW*);

XPR_INLINE xpr_rcode_t initOsVer(xpr_uint_t &aOsVer)
{
    OSVERSIONINFOEXW sOsVerInfo = {0,};
    sOsVerInfo.dwOSVersionInfoSize = sizeof(OSVERSIONINFOEXW);
    
    xpr_bool_t sResult = XPR_FALSE;
    
    // Try RtlGetVersion first (always returns true OS version)
    HMODULE hNtDll = GetModuleHandle(_T("ntdll.dll"));
    if (hNtDll != NULL)
    {
        PRtlGetVersion pRtlGetVersion = (PRtlGetVersion)
            GetProcAddress(hNtDll, "RtlGetVersion");
        if (pRtlGetVersion != NULL)
        {
            if (pRtlGetVersion(&sOsVerInfo) == 0) // STATUS_SUCCESS
            {
                sResult = XPR_TRUE;
            }
        }
    }
    
    // Fallback to GetVersionEx
    if (XPR_IS_FALSE(sResult))
    {
        #pragma warning(push)
        #pragma warning(disable: 4996)
        sResult = ::GetVersionExW((LPOSVERSIONINFOW)&sOsVerInfo);
        #pragma warning(pop)
        if (XPR_IS_FALSE(sResult))
            return XPR_RCODE_GET_OS_ERROR();
    }

    xpr_uint_t sWinVer = kOsVerWinUnknown;
    xpr_bool_t sWorkstation = (sOsVerInfo.wProductType == VER_NT_WORKSTATION) ? XPR_TRUE : XPR_FALSE;

    if (sOsVerInfo.dwMajorVersion == 5)
    {
        if (sOsVerInfo.dwMinorVersion == 0)
            sWinVer = kOsVerWin2000;
        else if (sOsVerInfo.dwMinorVersion == 1)
            sWinVer = kOsVerWinXP;
        else if (sOsVerInfo.dwMinorVersion == 2)
        {
            SYSTEM_INFO si = {0};
            ::GetSystemInfo(&si);
            if (sWorkstation == XPR_TRUE && si.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64)
                sWinVer = kOsVerWinXPx64;
            else
                sWinVer = kOsVerWin2003;
        }
    }
    else if (sOsVerInfo.dwMajorVersion == 6)
    {
        if (sOsVerInfo.dwMinorVersion == 0)
        {
            sWinVer = (sWorkstation == XPR_TRUE) ? kOsVerWinVista : kOsVerWin2008;
        }
        else if (sOsVerInfo.dwMinorVersion == 1)
        {
            sWinVer = (sWorkstation == XPR_TRUE) ? kOsVerWin7 : kOsVerWin2008R2;
        }
        else if (sOsVerInfo.dwMinorVersion == 2)
        {
            sWinVer = (sWorkstation == XPR_TRUE) ? kOsVerWin8 : kOsVerWin2012;
        }
        else if (sOsVerInfo.dwMinorVersion == 3)
        {
            sWinVer = (sWorkstation == XPR_TRUE) ? kOsVerWin8_1 : kOsVerWin2012R2;
        }
        else
        {
            sWinVer = kOsVerWinHigher;
        }
    }
    else if (sOsVerInfo.dwMajorVersion == 10)
    {
        if (sWorkstation == XPR_TRUE)
        {
            // Windows 11 starts at build 22000
            if (sOsVerInfo.dwBuildNumber >= 22000)
            {
                if (sOsVerInfo.dwBuildNumber >= 26100)
                    sWinVer = kOsVerWin11_24H2;
                else if (sOsVerInfo.dwBuildNumber >= 22631)
                    sWinVer = kOsVerWin11_23H2;
                else if (sOsVerInfo.dwBuildNumber >= 22621)
                    sWinVer = kOsVerWin11_22H2;
                else
                    sWinVer = kOsVerWin11;
            }
            else
            {
                // Windows 10 builds
                if (sOsVerInfo.dwBuildNumber >= 19045)
                    sWinVer = kOsVerWin10_22H2;
                else if (sOsVerInfo.dwBuildNumber >= 19044)
                    sWinVer = kOsVerWin10_21H2;
                else if (sOsVerInfo.dwBuildNumber >= 19043)
                    sWinVer = kOsVerWin10_21H1;
                else
                    sWinVer = kOsVerWin10;
            }
        }
        else
        {
            // Server versions
            if (sOsVerInfo.dwBuildNumber >= 20348)
                sWinVer = kOsVerWin2022;
            else
                sWinVer = kOsVerWinHigher;
        }
    }
    else if (sOsVerInfo.dwMajorVersion > 10)
    {
        sWinVer = kOsVerWinHigher;
    }

    aOsVer = sWinVer;

    return XPR_RCODE_SUCCESS;
}

XPR_DL_API SystemInfo *getSystemInfo(void)
{
    return &gSystemInfo;
}

XPR_DL_API xpr_uint_t getOsVer(void)
{
    return gSystemInfo.mOsVer;
}
} // namespace xpr
