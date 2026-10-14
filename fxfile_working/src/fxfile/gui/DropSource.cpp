//
// Copyright (c) 2001-2012 Leon Lee author. All rights reserved.
//
//   homepage: http://www.flychk.com
//   e-mail:   mailto:flychk@flychk.com
//
// Use of this source code is governed by a GPLv3 license that can be
// found in the LICENSE file.

#include "stdafx.h"
#include "DropSource.h"
#include "file_operation_confirmation.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

STDMETHODIMP DropSource::QueryInterface(REFIID iid, void FAR* FAR* ppv) 
{
    if (iid == IID_IUnknown || iid == IID_IDropSource)
    {
        *ppv = this;
        ++mRefCount;
        return NOERROR;
    }
    *ppv = XPR_NULL;

    return E_NOINTERFACE;
}

STDMETHODIMP_(xpr_ulong_t) DropSource::AddRef(void)
{
    return ++mRefCount;
}

STDMETHODIMP_(xpr_ulong_t) DropSource::Release(void)
{
    if (--mRefCount == 0)
    {
        delete this;
        return 0;
    }
    return mRefCount;
}  

DropSource::DropSource(IDataObject* aDataObject, HWND aOwner)
{
    mRefCount = 1;
    mOwner = aOwner;
    mEffect = DROPEFFECT_NONE;
    fxfile::FileOperationConfirmation::dataObjectPaths(aDataObject, mSources);
}

STDMETHODIMP DropSource::QueryContinueDrag(xpr_bool_t fEscapePressed, DWORD grfKeyState)
{
    HRESULT hr;
    if (fEscapePressed)// || (grfKeyState & MK_RBUTTON))
        hr = DRAGDROP_S_CANCEL;
    else if (!(grfKeyState & MK_LBUTTON) && !(grfKeyState & MK_RBUTTON))
    {
        POINT point; ::GetCursorPos(&point);
        HWND target = ::WindowFromPoint(point);
        DWORD process = 0; ::GetWindowThreadProcessId(target, &process);
        if (!target || mEffect == DROPEFFECT_NONE) return DRAGDROP_S_CANCEL;
        // An in-process target confirms after resolving its actual path.
        if (process == ::GetCurrentProcessId()) return DRAGDROP_S_DROP;
        wchar_t title[256] = {};
        ::GetWindowTextW(::GetAncestor(target, GA_ROOT), title, _countof(title));
        std::wstring destination = L"외부 프로그램: ";
        destination += *title ? title : L"이름을 확인할 수 없는 외부 창";
        destination += L"\r\n실제 도착 경로와 최종 작업은 외부 프로그램이 결정합니다.";
        UINT action = mEffect == DROPEFFECT_MOVE ? FO_MOVE :
            mEffect == DROPEFFECT_COPY ? FO_COPY :
            mEffect == DROPEFFECT_LINK ? fxfile::FileOperationConfirmation::Link :
            fxfile::FileOperationConfirmation::ExternalDrop;
        const DWORD effect = mEffect;
        bool accepted = fxfile::FileOperationConfirmation::confirm(mOwner, action, mSources,
            {destination});
        // Consent is for the original drop point, never for the dialog button
        // now beneath the pointer or a replacement external window.
        if (accepted) {
            ::SetCursorPos(point.x, point.y);
            accepted = ::IsWindow(target) && ::WindowFromPoint(point) == target && mEffect == effect;
        }
        hr = accepted ? DRAGDROP_S_DROP : DRAGDROP_S_CANCEL;
    }
    else
        hr = NOERROR;
    return hr;
}

STDMETHODIMP DropSource::GiveFeedback(DWORD dwEffect)
{
    mEffect = dwEffect & (DROPEFFECT_COPY | DROPEFFECT_MOVE | DROPEFFECT_LINK);
    return DRAGDROP_S_USEDEFAULTCURSORS;
}
