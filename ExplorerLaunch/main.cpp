#define STRICT
#include <windows.h>
#include <shldisp.h>
#include <shlobj.h>
#include <exdisp.h>
#include <atlbase.h>
#include <stdlib.h>

// https://devblogs.microsoft.com/oldnewthing/20131118-00/?p=2643
// can be used to break process trees.

void FindDesktopFolderView(REFIID riid, void** ppv)
{
    CComPtr<IShellWindows> spShellWindows;
    spShellWindows.CoCreateInstance(CLSID_ShellWindows);

    CComVariant vtLoc(CSIDL_DESKTOP);
    CComVariant vtEmpty;
    long lhwnd;
    CComPtr<IDispatch> spdisp;
    spShellWindows->FindWindowSW(
        &vtLoc, &vtEmpty,
        SWC_DESKTOP, &lhwnd, SWFO_NEEDDISPATCH, &spdisp);

    CComPtr<IShellBrowser> spBrowser;
    CComQIPtr<IServiceProvider>(spdisp)->
        QueryService(SID_STopLevelBrowser,
            IID_PPV_ARGS(&spBrowser));

    CComPtr<IShellView> spView;
    spBrowser->QueryActiveShellView(&spView);

    spView->QueryInterface(riid, ppv);
}

// FindDesktopFolderView incorporated by reference
void GetDesktopAutomationObject(REFIID riid, void** ppv)
{
    CComPtr<IShellView> spsv;
    FindDesktopFolderView(IID_PPV_ARGS(&spsv));
    CComPtr<IDispatch> spdispView;
    spsv->GetItemObject(SVGIO_BACKGROUND, IID_PPV_ARGS(&spdispView));
    spdispView->QueryInterface(riid, ppv);
}

void ShellExecuteFromExplorer(
    PCWSTR pszFile,
    PCWSTR pszParameters = nullptr,
    PCWSTR pszDirectory = nullptr,
    PCWSTR pszOperation = nullptr,
    int nShowCmd = SW_SHOWNORMAL)
{
    CComPtr<IShellFolderViewDual> spFolderView;
    GetDesktopAutomationObject(IID_PPV_ARGS(&spFolderView));
    CComPtr<IDispatch> spdispShell;
    spFolderView->get_Application(&spdispShell);
    CComQIPtr<IShellDispatch2>(spdispShell)
        ->ShellExecute(CComBSTR(pszFile),
            CComVariant(pszParameters ? pszParameters : L""),
            CComVariant(pszDirectory ? pszDirectory : L""),
            CComVariant(pszOperation ? pszOperation : L""),
            CComVariant(nShowCmd));
}

int __cdecl wmain(int argc, wchar_t** argv)
{
    if (argc < 2) return 0;
    CoInitialize(nullptr);
    ShellExecuteFromExplorer(
        argv[1],
        argc >= 3 ? argv[2] : L"",
        argc >= 4 ? argv[3] : L"",
        argc >= 5 ? argv[4] : L"",
        argc >= 6 ? _wtoi(argv[5]) : SW_SHOWNORMAL);
    return 0;
}
