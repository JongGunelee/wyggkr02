//
// Copyright (c) 2001-2013 Leon Lee author. All rights reserved.
//
//   homepage: http://www.flychk.com
//   e-mail:   mailto:flychk@flychk.com
//
// Use of this source code is governed by a GPLv3 license that can be
// found in the LICENSE file.

#include "stdafx.h"
#include "win_app.h"

#include "language_table.h"
#include "language_pack.h"
#include "string_table.h"
#include "format_string_table.h"

#include "file_op_undo.h"     // for Undo
#include "drive_shcn.h"
#include "program_opts.h"
#include "file_filter.h"
#include "size_format.h"
#include "winapi_ex.h"

#include "main_frame.h"
#include "explorer_view.h"
#include "option.h"
#include "option_manager.h"
#include "conf_dir.h"
#include "app_ver.h"
#include "command_string_table.h"
#include "shell_registry.h"
#include "launcher_manager.h"
#include "upchecker_manager.h"
#include "singleton_manager.h"
#include "single_process.h"
#include "bookmark.h"
#include "startup_trace.h"

#include "gfl/libgfl.h"

#include "cfg/cfg_main_dlg.h"

#include "cmd/tip_dlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

namespace fxfile
{
namespace
{
#if (XPR_CFG_COMPILER_MSVC && XPR_CFG_COMPILER_VER < 1700)
// reference: https://svn.boost.org/trac/boost/ticket/5582
void freeTypeinfoMemory(void)
{
   __type_info_node *sNode     = __type_info_root_node.next;
   __type_info_node *sTempNode = &__type_info_root_node;

   for (; sNode != XPR_NULL; sNode = sTempNode)
   {
      sTempNode = sNode->next;
      delete sNode->memPtr;
      delete sNode;
   }
}
#endif // XPR_CFG_COMPILER_MSVC && XPR_CFG_COMPILER_VER < 1700

const xpr_tchar_t kPreviewSection[] = XPR_STRING_LITERAL("Settings");
const xpr_tchar_t kPreviewEntry[]   = XPR_STRING_LITERAL("PreviewPages");
} // namespace anonymous

WinApp gApp;
Option *gOpt;

WinApp::WinApp(void)
    : mInitialized(XPR_FALSE)
    , mLanguageTable(XPR_NULL), mStringTable(XPR_NULL), mFormatStringTable(XPR_NULL)
{
#if defined(XPR_CFG_BUILD_RELEASE)

    xpr_tchar_t sAppVer[0xff] = {0};
    getAppVer(sAppVer);

#ifdef XPR_CFG_UNICODE
    _tcscat(sAppVer, XPR_STRING_LITERAL(" Unicode"));
#else
    _tcscat(sAppVer, XPR_STRING_LITERAL(""));
#endif

    fxfile_crash_init();
    fxfile_crash_setAppName(FXFILE_PROGRAM_NAME);
    fxfile_crash_setAppVer(sAppVer);
    fxfile_crash_setDevelopInfo(XPR_STRING_LITERAL("flychk@flychk.com"), XPR_STRING_LITERAL("http://www.flychk.com"));

#endif
}

WinApp::~WinApp(void)
{
#if (XPR_CFG_COMPILER_MSVC && XPR_CFG_COMPILER_VER < 1700)
    // clean RTTI memory leak
    freeTypeinfoMemory();
#endif // XPR_CFG_COMPILER_MSVC && XPR_CFG_COMPILER_VER < 1700
}

BEGIN_MESSAGE_MAP(WinApp, super)
    ON_COMMAND(ID_FILE_NEW, OnFileNew)
    ON_COMMAND(ID_FILE_OPEN, super::OnFileOpen)
END_MESSAGE_MAP()

xpr_bool_t WinApp::InitInstance(void)
{
    TraceStartup(XPR_STRING_LITERAL("WinApp.InitInstance.begin"));

    // Initialize OLE libraries
    if (AfxOleInit() == XPR_FALSE)
    {
        AfxMessageBox(IDP_OLE_INIT_FAILED);
        return XPR_FALSE;
    }

    // Task 065: Configure COleMessageFilter to prevent UI thread deadlocks
    // when Windows Shell extensions / COM Surrogate (dllhost.exe) take time.
    COleMessageFilter *sMsgFilter = AfxOleGetMessageFilter();
    if (XPR_IS_NOT_NULL(sMsgFilter))
    {
        sMsgFilter->EnableBusyDialog(FALSE);
        sMsgFilter->EnableNotRespondingDialog(FALSE);
        sMsgFilter->SetMessagePendingDelay(5000);
        sMsgFilter->SetRetryReply(0);
    }

    AfxEnableControlContainer();

    //
    // initialize XPR
    //
    xpr_bool_t sResult = xpr::initialize();
    XPR_ASSERT(sResult == XPR_TRUE);
    TraceStartup(XPR_STRING_LITERAL("WinApp.core_initialized"));

    // Html Help File Name
    xpr_tchar_t *sDot = (xpr_tchar_t *)_tcsrchr(m_pszHelpFilePath, '.');
    if (sDot != XPR_NULL)
        _tcscpy(sDot, XPR_STRING_LITERAL(".chm"));

    // Program file name verification routine
#if defined(XPR_CFG_BUILD_RELEASE)
    // [x64 Improvement] Disable name check for build flexibility
    /*
    if (_tcsicmp(m_pszExeName, FXFILE_PROGRAM_NAME) != 0)
    {
        xpr::string sMsg;
        sMsg  = XPR_STRING_LITERAL("This program is \'fxfile\' file manager.\n");
        sMsg += XPR_STRING_LITERAL("It is contained that a routine verfiy for this program file name.\n");
        sMsg += XPR_STRING_LITERAL("The file name must be \'fxfile.exe\'.");
        sMsg += XPR_STRING_LITERAL("Please, rename to \'fxfile.exe\' and then execute this program again.\n");
        sMsg += XPR_STRING_LITERAL("\n");
        sMsg += XPR_STRING_LITERAL("Please, contact to homepage: http://flychk.com or e-mail: flychk@flychk.com for any question.");
        sMsg += XPR_STRING_LITERAL("\n");
        MessageBox(XPR_NULL, sMsg.c_str(), FXFILE_PROGRAM_NAME, MB_OK | MB_ICONSTOP);
        return XPR_FALSE;
    }
    */
#endif // XPR_CFG_BUILD_RELEASE


    // shell change notification
    ShellChangeNotify &sShellChangeNotify = ShellChangeNotify::instance();
    sShellChangeNotify.create();
    sShellChangeNotify.start();

    DriveShcn &sDriveShcn = DriveShcn::instance();
    sDriveShcn.create();
    sDriveShcn.start();

    // initailze gfl Library
    gflLibraryInit();
    gflEnableLZW(GFL_TRUE);

    // [Win11 Optimization] Expanded common controls initialization.
    // ICC_WIN95_CLASSES alone is insufficient for Windows 10/11.
    // Added modern control classes for full compatibility.
    INITCOMMONCONTROLSEX sInitCommonContrlEx;
    sInitCommonContrlEx.dwSize = sizeof(sInitCommonContrlEx);
    sInitCommonContrlEx.dwICC = ICC_WIN95_CLASSES
                              | ICC_BAR_CLASSES
                              | ICC_TAB_CLASSES
                              | ICC_LISTVIEW_CLASSES
                              | ICC_TREEVIEW_CLASSES
                              | ICC_COOL_CLASSES
                              | ICC_USEREX_CLASSES
                              | ICC_STANDARD_CLASSES
                              | ICC_LINK_CLASS;
    InitCommonControlsEx(&sInitCommonContrlEx);

    // [Win11 Optimization] Enable Per-Monitor DPI awareness for sharp
    // rendering on Windows 10/11 high-DPI displays.
    // This prevents blurry UI on 4K monitors and multi-monitor setups.
    {
        typedef BOOL (WINAPI *PSetProcessDpiAwarenessContext)(DPI_AWARENESS_CONTEXT);
        HMODULE hUser32 = GetModuleHandle(_T("user32.dll"));
        if (hUser32 != NULL)
        {
            PSetProcessDpiAwarenessContext pSetDpiContext = 
                (PSetProcessDpiAwarenessContext)GetProcAddress(hUser32, "SetProcessDpiAwarenessContext");
            if (pSetDpiContext != NULL)
            {
                pSetDpiContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
            }
        }
    }
    TraceStartup(XPR_STRING_LITERAL("WinApp.common_controls_dpi"));

    // set registry key of this program
#ifdef XPR_CFG_BUILD_DEBUG
    SetRegistryKey(XPR_STRING_LITERAL("fxfile_dbg"));
#else
    SetRegistryKey(XPR_STRING_LITERAL("fxfile"));
#endif

    // parse command line arguments
    ProgramOpts &sProgramOpts = SingletonManager::get<ProgramOpts>();
    sProgramOpts.parse();

    if (sProgramOpts.isShowUsage() == XPR_TRUE)
    {
        ProgramOpts::showUsage();
        return XPR_FALSE;
    }

    if (sProgramOpts.isShowVersion() == XPR_TRUE)
    {
        ProgramOpts::showVersion();
    }

    xpr_bool_t  sResetConf     = sProgramOpts.isResetConf();
    xpr::string sConfDirToLoad = sProgramOpts.getConfDir();

    // load configuration directory
    ConfDir &sConfDir = ConfDir::instance();

    if (sConfDirToLoad.empty() == false)
    {
        sConfDir.setConfDir(sConfDirToLoad.c_str(), XPR_TRUE);
    }
    else
    {
        sConfDir.load();
    }

    // load options from file or load default option if configuration file does not exist
    OptionManager &sOptionManager = OptionManager::instance();

    xpr_bool_t sInitCfg = XPR_FALSE;

    if (XPR_IS_TRUE(sResetConf))
    {
        sOptionManager.initDefault();
        sInitCfg = XPR_TRUE;
    }
    else
    {
        sOptionManager.load(sInitCfg);
    }

    gOpt = sOptionManager.getOption();

    gOpt->setObserver(dynamic_cast<OptionObserver *>(this));
    TraceStartup(XPR_STRING_LITERAL("WinApp.configuration_loaded"));

    // The bookmark toolbar is populated while MainFrame creates its rebar.
    // Load the portable bookmark file before LoadFrame; otherwise the
    // BookmarkMgr is still empty, the band height is calculated as zero and
    // the saved bookmark bar appears as a thin blank line.
    BookmarkMgr::instance().load();
    TraceStartup(XPR_STRING_LITERAL("WinApp.bookmarks_loaded"));

    // [x64 Improvement] Robust language loading logic (Fixed Position)
    if (loadLanguageTable() == XPR_FALSE)
    {
        xpr_tchar_t sDirDbg[XPR_MAX_PATH + 1] = {0};
        GetModuleDir(sDirDbg, XPR_MAX_PATH);
        xpr_tchar_t sMsg[1024];
        _stprintf(sMsg, XPR_STRING_LITERAL("Critical Error: 'Languages' folder not found!\nBase Dir: %s\nPlease ensure 'Languages' folder exists next to fxfile.exe."), sDirDbg);
        MessageBox(XPR_NULL, sMsg, FXFILE_PROGRAM_NAME, MB_OK | MB_ICONSTOP);
        return XPR_FALSE;
    }

    xpr_bool_t sLanLoaded = loadLanguage(gOpt->mConfig.mLanguage);

    if (sLanLoaded == XPR_FALSE && mLanguageTable->getLanguageCount() > 0)
    {
        const fxfile::base::LanguagePack::Desc *sFirstLang = mLanguageTable->getLanguageDesc((xpr_size_t)0);
        if (sFirstLang != XPR_NULL)
        {
            sLanLoaded = loadLanguage(sFirstLang->mLanguage.c_str());
            if (sLanLoaded == XPR_TRUE)
            {
                _tcscpy_s(gOpt->mConfig.mLanguage, sFirstLang->mLanguage.c_str());
            }
        }
    }

    if (sLanLoaded == XPR_FALSE)
    {
        xpr_tchar_t sDirDbg[XPR_MAX_PATH + 1] = {0};
        GetModuleDir(sDirDbg, XPR_MAX_PATH);
        xpr_tchar_t sMsg[2048];
        _stprintf(sMsg, XPR_STRING_LITERAL("Error: Language Pack (Korean.xml) load FAILED!\n\n1. Search Dir: %s\\Languages\\\n2. Found Count: %Iu\n\nPlease check if Korean.xml exists in the above path."),
                  sDirDbg, mLanguageTable->getLanguageCount());
        MessageBox(XPR_NULL, sMsg, FXFILE_PROGRAM_NAME, MB_OK | MB_ICONSTOP);
        return XPR_FALSE;
    }
    TraceStartup(XPR_STRING_LITERAL("WinApp.language_loaded"));

    CommandStringTable::instance().load();
    TraceStartup(XPR_STRING_LITERAL("WinApp.command_strings_loaded"));

    // check it if single process by option
    if (XPR_IS_TRUE(gOpt->mConfig.mSingleProcess))
    {
        if (SingleProcess::check() == XPR_FALSE)
        {
            SingleProcess::postMsg();

            return XPR_FALSE;
        }

        SingleProcess::lock();
    }
    TraceStartup(XPR_STRING_LITERAL("WinApp.single_process_ready"));

    // load recent executed file list
    LoadStdProfileSettings(10);  // Load standard INI file options (including MRU)
    TraceStartup(XPR_STRING_LITERAL("WinApp.profile_loaded"));

    // load main frame
    MainFrame *sMainFrame = new MainFrame;
    if (XPR_IS_NULL(sMainFrame))
        return XPR_FALSE;

    m_pMainWnd = sMainFrame;
    TraceStartup(XPR_STRING_LITERAL("WinApp.before_LoadFrame"));
    if (sMainFrame->LoadFrame(IDR_MAINFRAME, WS_OVERLAPPEDWINDOW | FWS_ADDTOTITLE, XPR_NULL, XPR_NULL) == XPR_FALSE)
        return XPR_FALSE;
    TraceStartup(XPR_STRING_LITERAL("WinApp.after_LoadFrame"));

    // The one and only window has been initialized, so show and update it.
    sMainFrame->SetForegroundWindow();
    sMainFrame->ShowWindow(m_nCmdShow);
    sMainFrame->UpdateWindow();
    TraceStartup(XPR_STRING_LITERAL("WinApp.frame_shown"));

    // The skeleton is already on screen. Start the hidden-pane batch now,
    // ahead of lower-priority bookmark/icon messages already in the queue.
    sMainFrame->completeDeferredStartupViews();

    // popup preference dialog if configuration file does not exist on loading time.
    if (XPR_IS_TRUE(sInitCfg))
    {
        cfg::CfgMainDlg sDlg;
        sDlg.DoModal();
    }

    // popup tip of today dialog by option
    if (XPR_IS_TRUE(gOpt->mMain.mTipOfTheToday))
    {
        cmd::TipDlg sDlg;
        sDlg.DoModal();
    }

    mInitialized = XPR_TRUE;

    // show main frame
    sMainFrame->ShowWindow(SW_SHOW);

    return XPR_TRUE;
}

xpr_sint_t WinApp::ExitInstance(void) 
{
    // delete undo directory
    FileOpUndo::deleteUndoDir();

    // Exit gfl Library
    gflLibraryExit(); 

    // save main option
    if (OptionManager::isInstance() == XPR_TRUE)
        OptionManager::instance().saveMainOption();

    // destroy drive shell change notify
    DriveShcn &sDriveShcn = DriveShcn::instance();
    sDriveShcn.stop();
    sDriveShcn.destroy();

    // destroy shell change notify
    ShellChangeNotify &sShellChangeNotify = ShellChangeNotify::instance();
    sShellChangeNotify.stop();
    sShellChangeNotify.destroy();

    XPR_SAFE_DELETE(mLanguageTable);

    // unlock if single process is locked
    SingleProcess::unlock();

    // clean singleton manager
    SingletonManager::clean();

    //
    // finalize XPR
    //
    xpr::finalize();

    return super::ExitInstance();
}

void WinApp::LoadStdProfileSettings(xpr_uint_t aMaxMRU)
{
    ASSERT_VALID(this);

    // 0 by default means not set
    m_nNumPreviewPages = GetProfileInt(kPreviewSection, kPreviewEntry, 0);
}

void WinApp::OnFileNew(void)
{
    super::OnFileNew();
}

xpr_bool_t WinApp::loadLanguageTable(void)
{
    mLanguageTable = new fxfile::base::LanguageTable;
    if (XPR_IS_NULL(mLanguageTable))
        return XPR_FALSE;

    // [x64 Fix] Directly use GetModuleFileNameW for reliable Unicode/Korean path support
    wchar_t sModulePath[XPR_MAX_PATH + 1] = {0};
    if (::GetModuleFileNameW(NULL, sModulePath, XPR_MAX_PATH) == 0)
        return XPR_FALSE;

    // Remove file name fxfile.exe to get directory
    wchar_t *sLastSlash = wcsrchr(sModulePath, L'\\');
    if (sLastSlash != NULL)
        *sLastSlash = L'\0';

    // Append \Languages (Using Wide literal to be safe)
    wcscat_s(sModulePath, XPR_MAX_PATH, L"\\Languages");

    mLanguageTable->setDir((const xpr_tchar_t *)sModulePath);

    if (mLanguageTable->scan(gOpt->mConfig.mLanguage) == XPR_FALSE)
    {
        // TODO load embedded english language pack
        return XPR_FALSE;
    }

    return XPR_TRUE;
}

xpr_bool_t WinApp::loadLanguage(const xpr_tchar_t *aLanguage)
{
    if (XPR_IS_NULL(aLanguage))
        return XPR_FALSE;

    mStringTable = XPR_NULL;
    mFormatStringTable = XPR_NULL;

    xpr_time_t sTime1 = xpr::timer_ms();

    if (mLanguageTable->loadLanguage(aLanguage) == XPR_FALSE)
        return XPR_FALSE;

    xpr_time_t sTime2 = xpr::timer_ms();

    XPR_TRACE(XPR_STRING_LITERAL("Loading time of language file = %I64dms\n"), sTime2 - sTime1);

    mStringTable = mLanguageTable->getStringTable();
    mFormatStringTable = mLanguageTable->getFormatStringTable();

    FileFilter::setString(
        gApp.loadString(XPR_STRING_LITERAL("popup.cfg.body.appearance.filter.default_filter_name.folder")),
        gApp.loadString(XPR_STRING_LITERAL("popup.cfg.body.appearance.filter.default_filter_name.general_file")),
        gApp.loadString(XPR_STRING_LITERAL("popup.cfg.body.appearance.filter.default_filter_name.executable_file")),
        gApp.loadString(XPR_STRING_LITERAL("popup.cfg.body.appearance.filter.default_filter_name.compressed_file")),
        gApp.loadString(XPR_STRING_LITERAL("popup.cfg.body.appearance.filter.default_filter_name.document_file")),
        gApp.loadString(XPR_STRING_LITERAL("popup.cfg.body.appearance.filter.default_filter_name.image_file")),
        gApp.loadString(XPR_STRING_LITERAL("popup.cfg.body.appearance.filter.default_filter_name.sound_file")),
        gApp.loadString(XPR_STRING_LITERAL("popup.cfg.body.appearance.filter.default_filter_name.movie_file")),
        gApp.loadString(XPR_STRING_LITERAL("popup.cfg.body.appearance.filter.default_filter_name.web_file")),
        gApp.loadString(XPR_STRING_LITERAL("popup.cfg.body.appearance.filter.default_filter_name.programming_file")),
        gApp.loadString(XPR_STRING_LITERAL("popup.cfg.body.appearance.filter.default_filter_name.temporary_file")));

    SizeFormat::setText(
        gApp.loadString(XPR_STRING_LITERAL("popup.cfg.body.appearance.size_format.unit.none")),
        gApp.loadString(XPR_STRING_LITERAL("popup.cfg.body.appearance.size_format.unit.automatic")),
        gApp.loadString(XPR_STRING_LITERAL("popup.cfg.body.appearance.size_format.unit.default")),
        gApp.loadString(XPR_STRING_LITERAL("popup.cfg.body.appearance.size_format.unit.custom")),
        gApp.loadString(XPR_STRING_LITERAL("common.size.byte")));

    return XPR_TRUE;
}

const fxfile::base::LanguageTable *WinApp::getLanguageTable(void) const
{
    return mLanguageTable;
}

const xpr_tchar_t *WinApp::loadString(const xpr_tchar_t *aId, xpr_bool_t aNullAvailable)
{
    if (XPR_IS_NULL(mStringTable))
        return aId;

    return mStringTable->loadString(aId, aNullAvailable);
}

const xpr_tchar_t *WinApp::loadString(const xpr::string &aId, xpr_bool_t aNullAvailable)
{
    const xpr_tchar_t *sId = aId.c_str();

    if (XPR_IS_NULL(mStringTable))
        return sId;

    return mStringTable->loadString(sId, aNullAvailable);
}

const xpr_tchar_t *WinApp::loadFormatString(const xpr_tchar_t *aId, const xpr_tchar_t *aReplaceFormatSpecifier)
{
    if (XPR_IS_NULL(mFormatStringTable))
        return aId;

    return mFormatStringTable->loadString(aId, aReplaceFormatSpecifier);
}

const xpr_tchar_t *WinApp::loadFormatString(const xpr::string &aId, const xpr_tchar_t *aReplaceFormatSpecifier)
{
    const xpr_tchar_t *sId = aId.c_str();

    if (XPR_IS_NULL(mFormatStringTable))
        return sId;

    return mFormatStringTable->loadString(sId, aReplaceFormatSpecifier);
}

void WinApp::onChangedConfig(Option &aOption)
{
    // unregister shell registry for old language
    ShellRegistry::unregisterShell();

    // popup warning message about language change
    const fxfile::base::LanguagePack::Desc *sLoadedLanguagePackDesc = mLanguageTable->getLanguageDesc();
    if (sLoadedLanguagePackDesc->mLanguage.compare_case(aOption.mConfig.mLanguage) != 0)
    {
        const xpr_tchar_t *sMsg = gApp.loadString(XPR_STRING_LITERAL("popup.cfg.msg.apply_language_on_next_loading_time"));
        AfxMessageBox(sMsg, MB_OK | MB_ICONWARNING);
    }

    // load language on runtime
    //loadLanguage(aOption.mConfig.mLanguage);

    // set single process
    if (XPR_IS_TRUE(aOption.mConfig.mSingleProcess))
    {
        SingleProcess::lock();
    }
    else
    {
        SingleProcess::unlock();
    }

    // shell registry
    if (XPR_IS_TRUE(aOption.mConfig.mRegShellContextMenu))
    {
        ShellRegistry::registerShell();
    }
    else
    {
        ShellRegistry::unregisterShell();
    }

    // fxfile-launcher
    if (XPR_IS_TRUE(aOption.mConfig.mLauncher))
    {
        LauncherManager::startupProcess(aOption.mConfig.mLauncherGlobalHotKey, aOption.mConfig.mLauncherTray);
    }
    else
    {
        LauncherManager::shutdownProcess(aOption.mConfig.mLauncherGlobalHotKey, aOption.mConfig.mLauncherTray);
    }

    if (XPR_IS_TRUE(aOption.mConfig.mLauncherWinStartup))
    {
        LauncherManager::registerWinStartup();
    }
    else
    {
        LauncherManager::unregisterWinStartup();
    }

    // fxfile-upchecker
    UpcheckerManager::writeConfFile(aOption.mConfig);

    if (XPR_IS_TRUE(aOption.mConfig.mUpdateCheckEnable))
    {
        UpcheckerManager::startupProcess();
        UpcheckerManager::registerWinStartup();
    }
    else
    {
        UpcheckerManager::shutdownProcess();
        UpcheckerManager::unregisterWinStartup();
    }

    // notify changed options to main frame
    MainFrame *sMainFrame = (MainFrame *)GetMainWnd();

    sMainFrame->setChangedOption(aOption);
}

void WinApp::saveAllOptions(void)
{
    // save main frame options
    MainFrame *sMainFrame = (MainFrame *)GetMainWnd();
    sMainFrame->saveAllOptions();
}
} // namespace fxfile
