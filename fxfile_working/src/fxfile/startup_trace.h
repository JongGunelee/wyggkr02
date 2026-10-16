//
// Lightweight opt-in startup tracing for release diagnostics.
// Set FXFILE_STARTUP_TRACE=1 and capture OutputDebugString messages with a
// debugger.  With the variable unset this is only one environment lookup at
// each coarse startup checkpoint and creates no files.
//

#ifndef __FXFILE_STARTUP_TRACE_H__
#define __FXFILE_STARTUP_TRACE_H__ 1
#pragma once

namespace fxfile
{
inline void TraceStartup(const xpr_tchar_t *aStage)
{
    xpr_tchar_t sEnabled[2] = {0};
    if (::GetEnvironmentVariable(XPR_STRING_LITERAL("FXFILE_STARTUP_TRACE"), sEnabled, 2) == 0)
        return;

    static ULONGLONG sStart = ::GetTickCount64();
    ULONGLONG sElapsed = ::GetTickCount64() - sStart;

    xpr_tchar_t sMessage[256] = {0};
    _sntprintf_s(sMessage,
                 _countof(sMessage),
                 _TRUNCATE,
                 XPR_STRING_LITERAL("FXFILE_STARTUP\t%lu\t%I64u\t%s\r\n"),
                 ::GetCurrentProcessId(),
                 sElapsed,
                 aStage);
    ::OutputDebugString(sMessage);
}
} // namespace fxfile

#endif // __FXFILE_STARTUP_TRACE_H__
