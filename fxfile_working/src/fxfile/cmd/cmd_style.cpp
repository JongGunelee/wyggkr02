//
// Copyright (c) 2012 Leon Lee author. All rights reserved.
//
//   homepage: http://www.flychk.com
//   e-mail:   mailto:flychk@flychk.com
//
// Use of this source code is governed by a GPLv3 license that can be
// found in the LICENSE file.

#include "stdafx.h"
#include "cmd_style.h"

#include "resource.h"
#include "explorer_ctrl.h"
#include "option.h"
#include "option_manager.h"
#include "main_frame.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

namespace fxfile
{
namespace cmd
{
xpr_sint_t ViewStyleCommand::canExecute(CommandContext &aContext)
{
    FXFILE_COMMAND_DECLARE_CTRL;

    xpr_bool_t sState = 0;

    if (sExplorerCtrl != XPR_NULL)
    {
        xpr_bool_t sChecked = XPR_FALSE;
        xpr_sint_t sViewStyle = sExplorerCtrl->getViewStyle();

        sState |= StateEnable;

        switch (sCommandId)
        {
        case ID_VIEW_STYLE_EXTRA_LARGE_ICONS: sChecked = (sViewStyle == VIEW_STYLE_EXTRA_LARGE_ICONS); break;
        case ID_VIEW_STYLE_LARGE_ICONS:       sChecked = (sViewStyle == VIEW_STYLE_LARGE_ICONS);       break;
        case ID_VIEW_STYLE_MEDIUM_ICONS:      sChecked = (sViewStyle == VIEW_STYLE_MEDIUM_ICONS);      break;
        case ID_VIEW_STYLE_SMALL_ICONS:       sChecked = (sViewStyle == VIEW_STYLE_SMALL_ICONS);       break;
        case ID_VIEW_STYLE_LIST:              sChecked = (sViewStyle == VIEW_STYLE_LIST);              break;
        case ID_VIEW_STYLE_DETAILS:           sChecked = (sViewStyle == VIEW_STYLE_DETAILS);           break;
        case ID_VIEW_STYLE_THUMBNAIL:         sChecked = (sViewStyle == VIEW_STYLE_THUMBNAIL);         break;
        default:                              sChecked = XPR_FALSE;                                    break;
        }

        if (XPR_IS_TRUE(sChecked))
        {
            sState |= StateRadio;
        }
    }

    return sState;
}

void ViewStyleCommand::execute(CommandContext &aContext)
{
    FXFILE_COMMAND_DECLARE_CTRL;

    if (sExplorerCtrl != XPR_NULL)
    {
        xpr_sint_t sViewStyle = -1;

        switch (sCommandId)
        {
        case ID_VIEW_STYLE_EXTRA_LARGE_ICONS: sViewStyle = VIEW_STYLE_EXTRA_LARGE_ICONS; break;
        case ID_VIEW_STYLE_LARGE_ICONS:       sViewStyle = VIEW_STYLE_LARGE_ICONS;       break;
        case ID_VIEW_STYLE_MEDIUM_ICONS:      sViewStyle = VIEW_STYLE_MEDIUM_ICONS;      break;
        case ID_VIEW_STYLE_SMALL_ICONS:       sViewStyle = VIEW_STYLE_SMALL_ICONS;       break;
        case ID_VIEW_STYLE_LIST:              sViewStyle = VIEW_STYLE_LIST;              break;
        case ID_VIEW_STYLE_DETAILS:           sViewStyle = VIEW_STYLE_DETAILS;           break;
        case ID_VIEW_STYLE_THUMBNAIL:         sViewStyle = VIEW_STYLE_THUMBNAIL;         break;
        }

        if (sViewStyle != -1)
        {
            xpr_bool_t sRefesh = (sViewStyle == VIEW_STYLE_THUMBNAIL) ? XPR_TRUE : XPR_FALSE;
            sExplorerCtrl->setViewStyle(sViewStyle, sRefesh);
        }
    }
}

xpr_sint_t ViewStyleToolBarCommand::canExecute(CommandContext &aContext)
{
    return StateEnable;
}

void ViewStyleToolBarCommand::execute(CommandContext &aContext)
{
}

xpr_sint_t UIScaleCommand::canExecute(CommandContext &aContext)
{
    xpr_sint_t sState = StateEnable;
    xpr_sint_t sCurrentPercent = gOpt->mConfig.mUIScalePercent;
    if (sCurrentPercent <= 0) sCurrentPercent = 100;

    xpr_bool_t sChecked = XPR_FALSE;
    switch (aContext.getCommandId())
    {
    case ID_VIEW_UI_SCALE_25:  sChecked = (sCurrentPercent == 25);  break;
    case ID_VIEW_UI_SCALE_50:  sChecked = (sCurrentPercent == 50);  break;
    case ID_VIEW_UI_SCALE_75:  sChecked = (sCurrentPercent == 75);  break;
    case ID_VIEW_UI_SCALE_100: sChecked = (sCurrentPercent == 100); break;
    case ID_VIEW_UI_SCALE_125: sChecked = (sCurrentPercent == 125); break;
    case ID_VIEW_UI_SCALE_150: sChecked = (sCurrentPercent == 150); break;
    case ID_VIEW_UI_SCALE_175: sChecked = (sCurrentPercent == 175); break;
    case ID_VIEW_UI_SCALE_200: sChecked = (sCurrentPercent == 200); break;
    }

    if (sChecked == XPR_TRUE)
        sState |= StateRadio;

    return sState;
}

void UIScaleCommand::execute(CommandContext &aContext)
{
    FXFILE_COMMAND_DECLARE_CTRL;

    xpr_sint_t sPercent = 100;
    switch (aContext.getCommandId())
    {
    case ID_VIEW_UI_SCALE_25:  sPercent = 25;  break;
    case ID_VIEW_UI_SCALE_50:  sPercent = 50;  break;
    case ID_VIEW_UI_SCALE_75:  sPercent = 75;  break;
    case ID_VIEW_UI_SCALE_100: sPercent = 100; break;
    case ID_VIEW_UI_SCALE_125: sPercent = 125; break;
    case ID_VIEW_UI_SCALE_150: sPercent = 150; break;
    case ID_VIEW_UI_SCALE_175: sPercent = 175; break;
    case ID_VIEW_UI_SCALE_200: sPercent = 200; break;
    }

    if (gOpt->mConfig.mUIScalePercent != sPercent)
    {
        gOpt->mConfig.mUIScalePercent = sPercent;
        OptionManager::instance().saveConfigOption();
        gOpt->notifyConfig();
        if (sMainFrame != XPR_NULL)
        {
            sMainFrame->applyUIScale();
        }
    }
}
} // namespace cmd
} // namespace fxfile
