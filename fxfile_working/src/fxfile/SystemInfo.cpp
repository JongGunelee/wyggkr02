// Written by Zoltan Csizmadia, zoltan_csizmadia@yahoo.com
// For companies(Austin,TX): If you would like to get my resume, send an email.
//
// The source is free, but if you want to use it, mention my name and e-mail address
//
//////////////////////////////////////////////////////////////////////////////////////
//
// SystemInfo.cpp v1.1
//
// History:
// 
// Date      Version     Description
// --------------------------------------------------------------------------------
// 10/16/00	 1.0	     Initial version
// 11/09/00  1.1         NT4 doesn't like if we bother the System process fix :)
//                       SystemInfoUtils::GetDeviceFileName() fix (subst drives added)
//                       NT Major version added to INtDLL class
//
//////////////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include <process.h>
#include "SystemInfo.h"

//#ifndef WINNT
//#error You need Windows NT to use this source code. Define WINNT!
//#endif

//////////////////////////////////////////////////////////////////////////////////////
//
// SystemInfoUtils
//
//////////////////////////////////////////////////////////////////////////////////////

// From wide char string to CString
void SystemInfoUtils::LPCWSTR2CString( LPCWSTR strW, CString& str )
{
#ifdef UNICODE
	// if it is already UNICODE, no problem
	str = strW;
#else
	str = _T("");

	TCHAR* actChar = (TCHAR*)strW;

	if ( actChar == _T('\0') )
		return;

	size_t len = wcslen(strW) + 1;
	TCHAR* pBuffer = new TCHAR[ len ];
	TCHAR* pNewStr = pBuffer;

	while ( len-- )
	{
		*(pNewStr++) = *actChar;
		actChar += 2;
	}

	str = pBuffer;

	delete [] pBuffer;
#endif
}

// From wide char string to unicode
void SystemInfoUtils::Unicode2CString( UNICODE_STRING* strU, CString& str )
{
	if ( *(DWORD*)strU != 0 )
		LPCWSTR2CString( (LPCWSTR)strU->Buffer, str );
	else
		str = _T("");
}

// From device file name to DOS filename
BOOL SystemInfoUtils::GetFsFileName( LPCTSTR lpDeviceFileName, CString& fsFileName )
{
	BOOL rc = FALSE;

	TCHAR lpDeviceName[0x1000];
	TCHAR lpDrive[3] = _T("A:");

	// Iterating through the drive letters
	for ( TCHAR actDrive = _T('A'); actDrive <= _T('Z'); actDrive++ )
	{
		lpDrive[0] = actDrive;

		// Query the device for the drive letter
		if ( QueryDosDevice( lpDrive, lpDeviceName, 0x1000 ) != 0 )
		{
			// Network drive?
			if ( _tcsnicmp( _T("\\Device\\LanmanRedirector\\"), lpDeviceName, 25 ) == 0 )
			{
				//Mapped network drive 

				TCHAR cDriveLetter;
				DWORD dwParam;

				TCHAR lpSharedName[0x1000];

				if ( _stscanf(  lpDeviceName, 
								_T("\\Device\\LanmanRedirector\\;%c:%d\\%s"), 
								&cDriveLetter, 
								&dwParam, 
								lpSharedName ) != 3 )
						continue;

				_tcscpy( lpDeviceName, _T("\\Device\\LanmanRedirector\\") );
				_tcscat( lpDeviceName, lpSharedName );
			}
			
			// Is this the drive letter we are looking for?
			if ( _tcsnicmp( lpDeviceName, lpDeviceFileName, _tcslen( lpDeviceName ) ) == 0 )
			{
				fsFileName = lpDrive;
				fsFileName += (LPCTSTR)( lpDeviceFileName + _tcslen( lpDeviceName ) );

				rc = TRUE;

				break;
			}
		}
	}

	return rc;
}

// From DOS file name to device file name
BOOL SystemInfoUtils::GetDeviceFileName( LPCTSTR lpFsFileName, CString& deviceFileName )
{
	BOOL rc = FALSE;
	TCHAR lpDrive[3];

	// Get the drive letter 
	// unfortunetaly it works only with DOS file names
	_tcsncpy( lpDrive, lpFsFileName, 2 );
	lpDrive[2] = _T('\0');

	TCHAR lpDeviceName[0x1000];

	// Query the device for the drive letter
	if ( QueryDosDevice( lpDrive, lpDeviceName, 0x1000 ) != 0 )
	{
		// Subst drive?
		if ( _tcsnicmp( _T("\\??\\"), lpDeviceName, 4 ) == 0 )
		{
			deviceFileName = lpDeviceName + 4;
			deviceFileName += lpFsFileName + 2;

			return TRUE;
		}
		else
		// Network drive?
		if ( _tcsnicmp( _T("\\Device\\LanmanRedirector\\"), lpDeviceName, 25 ) == 0 )
		{
			//Mapped network drive 

			TCHAR cDriveLetter;
			DWORD dwParam;

			TCHAR lpSharedName[0x1000];

			if ( _stscanf(  lpDeviceName, 
							_T("\\Device\\LanmanRedirector\\;%c:%d\\%s"), 
							&cDriveLetter, 
							&dwParam, 
							lpSharedName ) != 3 )
					return FALSE;

			_tcscpy( lpDeviceName, _T("\\Device\\LanmanRedirector\\") );
			_tcscat( lpDeviceName, lpSharedName );
		}

		_tcscat( lpDeviceName, lpFsFileName + 2 );

		deviceFileName = lpDeviceName;

		rc = TRUE;
	}

	return rc;
}

// [Win11 Optimization] Replaced deprecated GetVersionEx with RtlGetVersion.
// GetVersionEx has been deprecated since Windows 8.1 and returns incorrect
// version info (lies about OS version) on Windows 10/11.
// RtlGetVersion (from ntdll.dll) always returns the true OS version.
typedef LONG (WINAPI *PRtlGetVersion)(OSVERSIONINFOEXW*);

DWORD SystemInfoUtils::GetNTMajorVersion()
{
   static DWORD s_dwMajorVersion = 0;
   if (s_dwMajorVersion != 0)
      return s_dwMajorVersion;

   // Try RtlGetVersion first (preferred method for Win10/11)
   HMODULE hNtDll = GetModuleHandle(_T("ntdll.dll"));
   if (hNtDll != NULL)
   {
      PRtlGetVersion pRtlGetVersion = (PRtlGetVersion)
         GetProcAddress(hNtDll, "RtlGetVersion");
      
      if (pRtlGetVersion != NULL)
      {
         OSVERSIONINFOEXW osvi;
         ZeroMemory(&osvi, sizeof(osvi));
         osvi.dwOSVersionInfoSize = sizeof(osvi);
         
         if (pRtlGetVersion(&osvi) == 0) // STATUS_SUCCESS
         {
            s_dwMajorVersion = osvi.dwMajorVersion;
            return s_dwMajorVersion;
         }
      }
   }

   // Fallback: use GetVersionEx (will return 6.2 on Win8.1+
   // unless manifested, but better than nothing)
   OSVERSIONINFOEX osvi;
   ZeroMemory(&osvi, sizeof(OSVERSIONINFOEX));
   osvi.dwOSVersionInfoSize = sizeof(OSVERSIONINFOEX);

   #pragma warning(push)
   #pragma warning(disable: 4996) // Suppress deprecation warning
   if (GetVersionEx((OSVERSIONINFO *)&osvi))
   {
      s_dwMajorVersion = osvi.dwMajorVersion;
      return s_dwMajorVersion;
   }
   
   osvi.dwOSVersionInfoSize = sizeof(OSVERSIONINFO);
   if (GetVersionEx((OSVERSIONINFO *)&osvi))
   {
      s_dwMajorVersion = osvi.dwMajorVersion;
      return s_dwMajorVersion;
   }
   #pragma warning(pop)

   s_dwMajorVersion = 10; // Default to Windows 10+ for modern systems
   return s_dwMajorVersion;
}

//////////////////////////////////////////////////////////////////////////////////////
//
// INtDll
//
//////////////////////////////////////////////////////////////////////////////////////
INtDll::PNtQuerySystemInformation INtDll::NtQuerySystemInformation = NULL;
INtDll::PNtQueryObject INtDll::NtQueryObject = NULL;
INtDll::PNtQueryInformationThread	INtDll::NtQueryInformationThread = NULL;
INtDll::PNtQueryInformationFile	INtDll::NtQueryInformationFile = NULL;
INtDll::PNtQueryInformationProcess INtDll::NtQueryInformationProcess = NULL;
DWORD INtDll::dwNTMajorVersion = SystemInfoUtils::GetNTMajorVersion();

BOOL INtDll::NtDllStatus = INtDll::Init();

BOOL INtDll::Init()
{
	// Get the NtDll function pointers
	NtQuerySystemInformation = (PNtQuerySystemInformation)
					GetProcAddress( GetModuleHandle( _T( "ntdll.dll" ) ),
                    "NtQuerySystemInformation");

	NtQueryObject = (PNtQueryObject)
					GetProcAddress(	GetModuleHandle( _T( "ntdll.dll" ) ),
                    "NtQueryObject");

	NtQueryInformationThread = (PNtQueryInformationThread)
					GetProcAddress(	GetModuleHandle( _T( "ntdll.dll" ) ),
                    "NtQueryInformationThread");

	NtQueryInformationFile = (PNtQueryInformationFile)
					GetProcAddress(	GetModuleHandle( _T( "ntdll.dll" ) ),
                    "NtQueryInformationFile");

	NtQueryInformationProcess = (PNtQueryInformationProcess)
					GetProcAddress(	GetModuleHandle( _T( "ntdll.dll" ) ),
                    "NtQueryInformationProcess");

	return  NtQuerySystemInformation	!= NULL &&
			NtQueryObject				!= NULL &&
			NtQueryInformationThread	!= NULL &&
			NtQueryInformationFile		!= NULL &&
			NtQueryInformationProcess	!= NULL;
}

//////////////////////////////////////////////////////////////////////////////////////
//
// SystemProcessInformation
//
//////////////////////////////////////////////////////////////////////////////////////

SystemProcessInformation::SystemProcessInformation( BOOL bRefresh )
{
	// [Win11 Optimization] Removed fixed address (0x100000) from VirtualAlloc.
	// Fixed addresses fail on 64-bit Windows 11 due to ASLR (Address Space
	// Layout Randomization). Using NULL lets the OS choose a safe address.
	// Also using MEM_RESERVE|MEM_COMMIT for proper memory management.
	m_pBuffer = (UCHAR*)VirtualAlloc (NULL,
						BufferSize, 
						MEM_RESERVE | MEM_COMMIT,
						PAGE_READWRITE);

	if ( bRefresh )
		Refresh();
}

SystemProcessInformation::~SystemProcessInformation()
{
	VirtualFree( m_pBuffer, 0, MEM_RELEASE );
}

BOOL SystemProcessInformation::Refresh()
{
	m_ProcessInfos.RemoveAll();
	m_pCurrentProcessInfo = NULL;

	if ( !NtDllStatus || m_pBuffer == NULL )
		return FALSE;
	
	// query the process information
	if ( INtDll::NtQuerySystemInformation( 5, m_pBuffer, BufferSize, NULL ) != 0 )
		return FALSE;

	DWORD currentProcessID = GetCurrentProcessId(); //Current Process ID

	SYSTEM_PROCESS_INFORMATION* pSysProcess = (SYSTEM_PROCESS_INFORMATION*)m_pBuffer;
	do 
	{
		// fill the process information map
		m_ProcessInfos.SetAt( pSysProcess->dUniqueProcessId, pSysProcess );

		// we found this process
		if ( pSysProcess->dUniqueProcessId == currentProcessID )
			m_pCurrentProcessInfo = pSysProcess;

		// get the next process information block
		if ( pSysProcess->dNext != 0 )
			pSysProcess = (SYSTEM_PROCESS_INFORMATION*)((UCHAR*)pSysProcess + pSysProcess->dNext);
		else
			pSysProcess = NULL;

	} while ( pSysProcess != NULL );

	return TRUE;
}

//////////////////////////////////////////////////////////////////////////////////////
//
// SystemThreadInformation
//
//////////////////////////////////////////////////////////////////////////////////////

SystemThreadInformation::SystemThreadInformation( DWORD pID, BOOL bRefresh )
{
	m_processId = pID;

	if ( bRefresh )
		Refresh();
}

BOOL SystemThreadInformation::Refresh()
{
	// Get the Thread objects ( set the filter to "Thread" )
	SystemHandleInformation hi( m_processId );
	BOOL rc = hi.SetFilter( _T("Thread"), TRUE );

	m_ThreadInfos.RemoveAll();

	if ( !rc )
		return FALSE;

	THREAD_INFORMATION ti;

	// Iterating through the found Thread objects
	for ( POSITION pos = hi.m_HandleInfos.GetHeadPosition(); pos != NULL; )
	{
		SystemHandleInformation::SYSTEM_HANDLE& h = hi.m_HandleInfos.GetNext(pos);
		
		ti.ProcessId = h.ProcessID;
		ti.ThreadHandle = (HANDLE)h.HandleNumber;
		
		// This is one of the threads we are lokking for
		if ( SystemHandleInformation::GetThreadId( ti.ThreadHandle, ti.ThreadId, ti.ProcessId ) )
			m_ThreadInfos.AddTail( ti );
	}

	return TRUE;
}

//////////////////////////////////////////////////////////////////////////////////////
//
// SystemHandleInformation
//
//////////////////////////////////////////////////////////////////////////////////////

SystemHandleInformation::SystemHandleInformation( DWORD pID, BOOL bRefresh, LPCTSTR lpTypeFilter )
{
	m_processId = pID;
	
	// Set the filter
	SetFilter( lpTypeFilter, bRefresh );
}

SystemHandleInformation::~SystemHandleInformation()
{
}

BOOL SystemHandleInformation::SetFilter( LPCTSTR lpTypeFilter, BOOL bRefresh )
{
	// Set the filter ( default = all objects )
	m_strTypeFilter = lpTypeFilter == NULL ? _T("") : lpTypeFilter;

	return bRefresh ? Refresh() : TRUE;
}

const CString& SystemHandleInformation::GetFilter()
{
	return m_strTypeFilter;
}

BOOL SystemHandleInformation::IsSupportedHandle( SYSTEM_HANDLE& handle )
{
	//Here you can filter the handles you don't want in the Handle list

	// Windows 2000 supports everything :)
	if ( dwNTMajorVersion >= 5 )
		return TRUE;

	//NT4 System process doesn't like if we bother his internal security :)
	if ( handle.ProcessID == 2 && handle.HandleType == 16 )
		return FALSE;

	return TRUE;
}

BOOL SystemHandleInformation::Refresh()
{
	DWORD size = 0x2000;
	DWORD needed = 0;
	DWORD i = 0;
	BOOL  ret = TRUE;
	CString strType;

	m_HandleInfos.RemoveAll();

	if ( !INtDll::NtDllStatus )
		return FALSE;

	// Allocate the memory for the buffer
	SYSTEM_HANDLE_INFORMATION* pSysHandleInformation = (SYSTEM_HANDLE_INFORMATION*)
				VirtualAlloc( NULL, size, MEM_COMMIT, PAGE_READWRITE );

	if ( pSysHandleInformation == NULL )
		return FALSE;

	// Query the needed buffer size for the objects ( system wide )
	if ( INtDll::NtQuerySystemInformation( 16, pSysHandleInformation, size, &needed ) != 0 )
	{
		if ( needed == 0 )
		{
			ret = FALSE;
			goto cleanup;
		}

		// The size was not enough
		VirtualFree( pSysHandleInformation, 0, MEM_RELEASE );

		pSysHandleInformation = (SYSTEM_HANDLE_INFORMATION*)
				VirtualAlloc( NULL, size = needed + 256, MEM_COMMIT, PAGE_READWRITE );
	}
	
	if ( pSysHandleInformation == NULL )
		return FALSE;

	// Query the objects ( system wide )
	if ( INtDll::NtQuerySystemInformation( 16, pSysHandleInformation, size, NULL ) != 0 )
	{
		ret = FALSE;
		goto cleanup;
	}
	
	// Iterating through the objects
	for ( i = 0; i < pSysHandleInformation->Count; i++ )
	{
		if ( !IsSupportedHandle( pSysHandleInformation->Handles[i] ) )
			continue;
		
		// ProcessId filtering check
		if ( pSysHandleInformation->Handles[i].ProcessID == m_processId || m_processId == (DWORD)-1 )
		{
			BOOL bAdd = FALSE;
			
			if ( m_strTypeFilter == _T("") )
				bAdd = TRUE;
			else
			{
				// Type filtering
				GetTypeToken( (HANDLE)pSysHandleInformation->Handles[i].HandleNumber, strType, pSysHandleInformation->Handles[i].ProcessID  );

				bAdd = strType == m_strTypeFilter;
			}

			// That's it. We found one.
			if ( bAdd )
			{	
				pSysHandleInformation->Handles[i].HandleType = (WORD)(pSysHandleInformation->Handles[i].HandleType % 256);
				
				m_HandleInfos.AddTail( pSysHandleInformation->Handles[i] );

			}
		}
	}

cleanup:
	
	if ( pSysHandleInformation != NULL )
		VirtualFree( pSysHandleInformation, 0, MEM_RELEASE );

	return ret;
}

HANDLE SystemHandleInformation::OpenProcess( DWORD processId )
{
	// Open the process for handle duplication
	return ::OpenProcess( PROCESS_DUP_HANDLE, TRUE, processId );
}

HANDLE SystemHandleInformation::DuplicateHandle( HANDLE hProcess, HANDLE hRemote )
{
	HANDLE hDup = NULL;

	// Duplicate the remote handle for our process
	::DuplicateHandle( hProcess, hRemote,	GetCurrentProcess(), &hDup,	0, FALSE, DUPLICATE_SAME_ACCESS );

	return hDup;
}

//Information functions
BOOL SystemHandleInformation::GetTypeToken( HANDLE h, CString& str, DWORD processId )
{
	ULONG size = 0x2000;
	UCHAR* lpBuffer = NULL;
	BOOL ret = FALSE;

	HANDLE handle = NULL;
	HANDLE hRemoteProcess = NULL;
	BOOL remote = processId != GetCurrentProcessId();
	
	if ( !NtDllStatus )
		return FALSE;

	if ( remote )
	{
		// Open the remote process
		hRemoteProcess = OpenProcess( processId );
		
		if ( hRemoteProcess == NULL )
			return FALSE;

		// Duplicate the handle
		handle = DuplicateHandle( hRemoteProcess, h );
	}
	else
		handle = h;

	// Query the info size
	INtDll::NtQueryObject( handle, 2, NULL, 0, &size );

	lpBuffer = new UCHAR[size];

	// Query the info size ( type )
	if ( INtDll::NtQueryObject( handle, 2, lpBuffer, size, NULL ) == 0 )
	{
		str = _T("");
		SystemInfoUtils::LPCWSTR2CString( (LPCWSTR)(lpBuffer+0x60), str );

		ret = TRUE;
	}

	if ( remote )
	{
		if ( hRemoteProcess != NULL )
			CloseHandle( hRemoteProcess );

		if ( handle != NULL )
			CloseHandle( handle );
	}
	
	if ( lpBuffer != NULL )
		delete [] lpBuffer;

	return ret;
}

BOOL SystemHandleInformation::GetType( HANDLE h, WORD& type, DWORD processId )
{
	CString strType;

	type = OB_TYPE_UNKNOWN;

	if ( !GetTypeToken( h, strType, processId ) )
		return FALSE;

	return GetTypeFromTypeToken( strType, type );
}

BOOL SystemHandleInformation::GetTypeFromTypeToken( LPCTSTR typeToken, WORD& type )
{
	const WORD count = 27;
	CString constStrTypes[count] = { 
		_T(""), _T(""), _T("Directory"), _T("SymbolicLink"), _T("Token"),
		_T("Process"), _T("Thread"), _T("Unknown7"), _T("Event"), _T("EventPair"), _T("Mutant"), 
		_T("Unknown11"), _T("Semaphore"), _T("Timer"), _T("Profile"), _T("WindowStation"),
		_T("Desktop"), _T("Section"), _T("Key"), _T("Port"), _T("WaitablePort"), 
		_T("Unknown21"), _T("Unknown22"), _T("Unknown23"), _T("Unknown24"),
		_T("IoCompletion"), _T("File") };

	type = OB_TYPE_UNKNOWN;

	for ( WORD i = 1; i < count; i++ )
		if ( constStrTypes[i] == typeToken )
		{
			type = i;
			return TRUE;
		}
		
	return FALSE;
}

BOOL SystemHandleInformation::GetName( HANDLE handle, CString& str, DWORD processId )
{
	WORD type = 0;

	if ( !GetType( handle, type, processId  ) )
		return FALSE;

	return GetNameByType( handle, type, str, processId );
}

BOOL SystemHandleInformation::GetNameByType( HANDLE h, WORD type, CString& str, DWORD processId )
{
	ULONG size = 0x2000;
	UCHAR* lpBuffer = NULL;
	BOOL ret = FALSE;

	HANDLE handle;
	HANDLE hRemoteProcess = NULL;
	BOOL remote = processId != GetCurrentProcessId();
	DWORD dwId = 0;
	
	if ( !NtDllStatus )
		return FALSE;

	if ( remote )
	{
		hRemoteProcess = OpenProcess( processId );
		
		if ( hRemoteProcess == NULL )
			return FALSE;

		handle = DuplicateHandle( hRemoteProcess, h );
	}
	else
		handle = h;

	// let's be happy, handle is in our process space, so query the infos :)
	switch( type )
	{
	case OB_TYPE_PROCESS:
		GetProcessId( handle, dwId );
		
		str.Format( _T("PID: 0x%X"), dwId );
			
		ret = TRUE;
		goto cleanup;
		break;

	case OB_TYPE_THREAD:
		GetThreadId( handle, dwId );

		str.Format( _T("TID: 0x%X") , dwId );
				
		ret = TRUE;
		goto cleanup;
		break;

	case OB_TYPE_FILE:
		ret = GetFileName( handle, str );

		// access denied :(
		if ( ret && str == _T("") )
			goto cleanup;
		break;

	};

	INtDll::NtQueryObject ( handle, 1, NULL, 0, &size );

	// let's try to use the default
	if ( size == 0 )
		size = 0x2000;

	lpBuffer = new UCHAR[size];

	if ( INtDll::NtQueryObject( handle, 1, lpBuffer, size, NULL ) == 0 )
	{
		SystemInfoUtils::Unicode2CString( (UNICODE_STRING*)lpBuffer, str );
		ret = TRUE;
	}
	
cleanup:

	if ( remote )
	{
		if ( hRemoteProcess != NULL )
			CloseHandle( hRemoteProcess );

		if ( handle != NULL )
			CloseHandle( handle );
	}

	if ( lpBuffer != NULL )
		delete [] lpBuffer;
	
	return ret;
}

//Thread related functions
BOOL SystemHandleInformation::GetThreadId( HANDLE h, DWORD& threadID, DWORD processId )
{
	SystemThreadInformation::BASIC_THREAD_INFORMATION ti;
	HANDLE handle;
	HANDLE hRemoteProcess = NULL;
	BOOL remote = processId != GetCurrentProcessId();
	
	if ( !NtDllStatus )
		return FALSE;

	if ( remote )
	{
		// Open process
		hRemoteProcess = OpenProcess( processId );
		
		if ( hRemoteProcess == NULL )
			return FALSE;

		// Duplicate handle
		handle = DuplicateHandle( hRemoteProcess, h );
	}
	else
		handle = h;
	
	// Get the thread information
	if ( INtDll::NtQueryInformationThread( handle, 0, &ti, sizeof(ti), NULL ) == 0 )
		threadID = ti.ThreadId;

	if ( remote )
	{
		if ( hRemoteProcess != NULL )
			CloseHandle( hRemoteProcess );

		if ( handle != NULL )
			CloseHandle( handle );
	}

	return TRUE;
}

//Process related functions
BOOL SystemHandleInformation::GetProcessPath( HANDLE h, CString& strPath, DWORD remoteProcessId )
{
	h; strPath; remoteProcessId;

	strPath.Format( _T("%d"), remoteProcessId );

	return TRUE;
}

BOOL SystemHandleInformation::GetProcessId( HANDLE h, DWORD& processId, DWORD remoteProcessId )
{
	BOOL ret = FALSE;
	HANDLE handle;
	HANDLE hRemoteProcess = NULL;
	BOOL remote = remoteProcessId != GetCurrentProcessId();
	SystemProcessInformation::PROCESS_BASIC_INFORMATION pi;
	
	ZeroMemory( &pi, sizeof(pi) );
	processId = 0;
	
	if ( !NtDllStatus )
		return FALSE;

	if ( remote )
	{
		// Open process
		hRemoteProcess = OpenProcess( remoteProcessId );
		
		if ( hRemoteProcess == NULL )
			return FALSE;

		// Duplicate handle
		handle = DuplicateHandle( hRemoteProcess, h );
	}
	else
		handle = h;

	// Get the process information
	if ( INtDll::NtQueryInformationProcess( handle, 0, &pi, sizeof(pi), NULL) == 0 )
	{
		processId = pi.UniqueProcessId;
		ret = TRUE;
	}

	if ( remote )
	{
		if ( hRemoteProcess != NULL )
			CloseHandle( hRemoteProcess );

		if ( handle != NULL )
			CloseHandle( handle );
	}

	return ret;
}

//File related functions
void SystemHandleInformation::ReleaseFileNameThreadParam( GetFileNameThreadParam* p )
{
	if ( p == NULL || InterlockedDecrement( &p->refs ) != 0 )
		return;

	if ( p->ownFile && p->hFile != NULL )
		CloseHandle( p->hFile );
	if ( p->doneEvent != NULL )
		CloseHandle( p->doneEvent );
	delete p;
}

unsigned __stdcall SystemHandleInformation::GetFileNameThread( PVOID pParam )
{
	// A filesystem filter can block this native query.  The caller owns the
	// timeout/cancellation policy and the shared request remains ref-counted.
	GetFileNameThreadParam* p = (GetFileNameThreadParam*)pParam;

	typedef struct _FX_IO_STATUS_BLOCK
	{
		union
		{
			LONG  Status;
			PVOID Pointer;
		};
		ULONG_PTR Information;
	} FX_IO_STATUS_BLOCK;

	typedef struct _FX_FILE_NAME_INFORMATION
	{
		ULONG FileNameLength;
		WCHAR FileName[1];
	} FX_FILE_NAME_INFORMATION;

	UCHAR lpBuffer[0x1000] = { 0 };
	FX_IO_STATUS_BLOCK iob = { 0 };
	
	p->rc = INtDll::NtQueryInformationFile( p->hFile, &iob, lpBuffer, sizeof(lpBuffer), 9 );

	if ( p->rc == 0 )
	{
		const FX_FILE_NAME_INFORMATION* fileNameInfo =
			reinterpret_cast<const FX_FILE_NAME_INFORMATION*>( lpBuffer );
		const ULONG headerBytes = FIELD_OFFSET( FX_FILE_NAME_INFORMATION, FileName );
		const ULONG maximumNameBytes = sizeof(lpBuffer) - headerBytes;

		// FILE_NAME_INFORMATION is a byte-counted UTF-16 string and is not
		// guaranteed to be NUL terminated.  Validate it before conversion.
		if ( fileNameInfo->FileNameLength <= maximumNameBytes &&
			(fileNameInfo->FileNameLength % sizeof(WCHAR)) == 0 )
		{
			const size_t characterCount =
				fileNameInfo->FileNameLength / sizeof(WCHAR);
			WCHAR safeName[(0x1000 / sizeof(WCHAR)) + 1];
			memcpy( safeName, fileNameInfo->FileName,
				fileNameInfo->FileNameLength );
			safeName[characterCount] = L'\0';
			SystemInfoUtils::LPCWSTR2CString( safeName, p->name );
		}
		else
		{
			p->rc = ERROR_INVALID_DATA;
		}
	}

	SetEvent( p->doneEvent );
	ReleaseFileNameThreadParam( p );
	return 0;
}

BOOL SystemHandleInformation::GetFileName( HANDLE h, CString& str, DWORD processId )
{
	BOOL ret= FALSE;
	HANDLE hThread = NULL;
	GetFileNameThreadParam* tp = NULL;
	HANDLE handle = NULL;
	HANDLE hRemoteProcess = NULL;
	DWORD waitResult = WAIT_FAILED;
	BOOL remote = processId != GetCurrentProcessId();
	
	if ( !NtDllStatus )
		return FALSE;

	if ( remote )
	{
		// Open process
		hRemoteProcess = OpenProcess( processId );
		
		if ( hRemoteProcess == NULL )
			return FALSE;

		// Duplicate handle
		handle = DuplicateHandle( hRemoteProcess, h );
	}
	else
	{
		// The worker may outlive this call after a timeout.  It must never use
		// the caller's handle, which can be closed and recycled immediately.
		if ( !::DuplicateHandle( GetCurrentProcess(), h, GetCurrentProcess(),
			&handle, 0, FALSE, DUPLICATE_SAME_ACCESS ) )
			return FALSE;
	}

	if ( handle == NULL )
		goto cleanup;

	tp = new GetFileNameThreadParam;
	if ( tp == NULL )
		goto cleanup;
	tp->hFile = handle;
	tp->ownFile = TRUE;
	tp->doneEvent = CreateEvent( NULL, TRUE, FALSE, NULL );
	tp->rc = (ULONG)-1;
	tp->refs = 2; // caller + worker
	if ( tp->doneEvent == NULL )
	{
		tp->refs = 1;
		ReleaseFileNameThreadParam( tp );
		tp = NULL;
		goto cleanup;
	}
	handle = NULL; // ownership moved to the shared request

	// Let's start the thread to get the file name
	hThread = (HANDLE)_beginthreadex( NULL, 0, GetFileNameThread, tp, 0, NULL );

	if ( hThread == NULL )
	{
		// No worker took its reference.
		ReleaseFileNameThreadParam( tp );
		ReleaseFileNameThreadParam( tp );
		tp = NULL;
		ret = FALSE;
		goto cleanup;
	}

	// NtQueryInformationFile can block in a filesystem filter.  Never terminate
	// the thread while it owns CRT/heap state; request cancellation and let the
	// ref-counted request clean itself up when the kernel call eventually exits.
	waitResult = WaitForSingleObject( tp->doneEvent, 500 );
	if ( waitResult != WAIT_OBJECT_0 )
	{
		CancelSynchronousIo( hThread );
		str = _T("");
		ret = FALSE;
	}
	else
	{
		ret = ( tp->rc == 0 );
		if ( ret )
			str = tp->name;
	}

	CloseHandle( hThread );
	hThread = NULL;
	ReleaseFileNameThreadParam( tp );
	tp = NULL;

cleanup:

	if ( hRemoteProcess != NULL )
		CloseHandle( hRemoteProcess );

	if ( handle != NULL )
		CloseHandle( handle );
		
	return ret;
}

//////////////////////////////////////////////////////////////////////////////////////
//
// SystemModuleInformation
//
//////////////////////////////////////////////////////////////////////////////////////

SystemModuleInformation::SystemModuleInformation( DWORD pID, BOOL bRefresh )
{
	m_processId = pID;

	if ( bRefresh )
		Refresh();
}

void SystemModuleInformation::GetModuleListForProcess( DWORD processID )
{
	DWORD i = 0;
	DWORD cbNeeded = 0;
	HMODULE* hModules = NULL;
	MODULE_INFO moduleInfo;

	// Open process to read to query the module list
	HANDLE hProcess = OpenProcess( PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, processID );

	if ( hProcess == NULL )
		goto cleanup;

	//Get the number of modules
	if ( !(*m_EnumProcessModules)( hProcess, NULL, 0, &cbNeeded ) )
		goto cleanup;

	hModules = new HMODULE[ cbNeeded / sizeof( HMODULE ) ];

	//Get module handles
    if ( !(*m_EnumProcessModules)( hProcess, hModules, cbNeeded, &cbNeeded ) )
		goto cleanup;
	
	for ( i = 0; i < cbNeeded / sizeof( HMODULE ); i++ )
	{
		moduleInfo.ProcessId = processID;
		moduleInfo.Handle = hModules[i];
		
		//Get module full paths
		if ( (*m_GetModuleFileNameEx)( hProcess, hModules[i], moduleInfo.FullPath, _MAX_PATH ) )
			m_ModuleInfos.AddTail( moduleInfo );
	}

cleanup:
	if ( hModules != NULL )
		delete [] hModules;

	if ( hProcess != NULL )
		CloseHandle( hProcess );
}

BOOL SystemModuleInformation::Refresh()
{
	BOOL rc = FALSE;
	m_EnumProcessModules = NULL;
	m_GetModuleFileNameEx = NULL;

	m_ModuleInfos.RemoveAll();
	
	//Load Psapi.dll
	HINSTANCE hDll = LoadLibrary(_T("PSAPI.DLL"));
 
	if ( hDll == NULL )
	{
		rc = FALSE;
		goto cleanup;
	}

	//Get Psapi.dll functions
	m_EnumProcessModules = (PEnumProcessModules)GetProcAddress( hDll, "EnumProcessModules" );

	m_GetModuleFileNameEx = (PGetModuleFileNameEx)GetProcAddress( hDll, 
#ifdef UNICODE
								"GetModuleFileNameExW" );
#else
								"GetModuleFileNameExA" );
#endif

	if ( m_GetModuleFileNameEx == NULL || m_EnumProcessModules == NULL )
	{
		rc = FALSE;
		goto cleanup;
	}
	
	// Everey process or just a particular one
	if ( m_processId != -1 )
		// For a particular one
		GetModuleListForProcess( m_processId );
	else
	{
		// Get teh process list
		DWORD pID;
		SystemProcessInformation::SYSTEM_PROCESS_INFORMATION* p = NULL;
		SystemProcessInformation pi( TRUE );
		
		if ( pi.m_ProcessInfos.GetStartPosition() == NULL )
		{
			rc = FALSE;
			goto cleanup;
		}

		// Iterating through the processes and get the module list
		for ( POSITION pos = pi.m_ProcessInfos.GetStartPosition(); pos != NULL; )
		{
			pi.m_ProcessInfos.GetNextAssoc( pos, pID, p ); 
			GetModuleListForProcess( pID );
		}
	}
	
	rc = TRUE;

cleanup:

	//Free psapi.dll
	if ( hDll != NULL )
		FreeLibrary( hDll );

	return rc;
}

//////////////////////////////////////////////////////////////////////////////////////
//
// SystemWindowInformation
//
//////////////////////////////////////////////////////////////////////////////////////

SystemWindowInformation::SystemWindowInformation( DWORD pID, BOOL bRefresh )
{
	m_processId = pID;

	if ( bRefresh )
		Refresh();
}

BOOL SystemWindowInformation::Refresh()
{
	m_WindowInfos.RemoveAll();

	// Enumerating the windows
	EnumWindows( EnumerateWindows, (LPARAM)this );
	
	return TRUE;
}

BOOL CALLBACK SystemWindowInformation::EnumerateWindows( HWND hwnd, LPARAM lParam )
{
	SystemWindowInformation* _this = (SystemWindowInformation*)lParam;
	WINDOW_INFO wi;

	wi.hWnd = hwnd;
	GetWindowThreadProcessId(hwnd, &wi.ProcessId ) ;

	// Filtering by process ID
	if ( _this->m_processId == -1 || _this->m_processId == wi.ProcessId )
	{
		GetWindowText( hwnd, wi.Caption, MaxCaptionSize );

		// That is we are looking for
		if ( GetLastError() == 0 )
			_this->m_WindowInfos.AddTail( wi );
	}

	return TRUE;
};
