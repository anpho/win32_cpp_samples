#include <iostream>
#include <Windows.h>
#include <atlbase.h>
#include <ShlObj.h>
#include <atlstr.h>
#include <atlpath.h>

int main()
{
	CoInitialize(NULL);
	ShellExecute(NULL, L"open", L"c:\\users\\default", NULL, NULL, SW_SHOWDEFAULT);
	Sleep(2000);

	ATL::CString folder_to_select(L"C:\\Users\\default\\Downloads");
	CComPtr<IShellWindows> shellwindow;
	HRESULT h = CoCreateInstance(CLSID_ShellWindows, NULL, CLSCTX_ALL, IID_PPV_ARGS(&shellwindow));
	if (FAILED(h)) { return 0; }
	CComPtr<IDispatch> pdisp;
	long count;
	shellwindow->get_Count(&count);
	for (long i = 0; i < count; i++) {
		CComVariant va(i, VT_I4);
		if (SUCCEEDED(shellwindow->Item(va, &pdisp)) && pdisp) {
			CComPtr<IWebBrowserApp> pwba;
			h = pdisp->QueryInterface(IID_PPV_ARGS(&pwba));
			if (FAILED(h)) { continue; }
			HWND hwba;
			if (SUCCEEDED(pwba->get_HWND((LONG_PTR*)&hwba))) {
					CComPtr<IServiceProvider> psp;
					h = pwba->QueryInterface(IID_PPV_ARGS(&psp));
					if (FAILED(h)) { continue; }
					CComPtr<IShellBrowser> psb;
					h = psp->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&psb));
					if (FAILED(h)) { continue; }
					CComPtr<IShellView> psv;
					h = psb->QueryActiveShellView(&psv);
					if (FAILED(h)) { continue; }
					psv->Refresh();
					CComPtr<IFolderView2> pfv;
					h = psv->QueryInterface(IID_PPV_ARGS(&pfv));
					if (FAILED(h)) { continue; }
					CComPtr<IShellItemArray> items;
					if (SUCCEEDED(pfv->Items(SVGIO_ALLVIEW, IID_PPV_ARGS(&items)))) {
						DWORD count = 0;
						items->GetCount(&count);
						for (DWORD i = 0; i < count; i++) {
							IShellItem* _item;
							if (SUCCEEDED(items->GetItemAt(i, &_item))) {
								LPWSTR filename = NULL;
								if (SUCCEEDED(_item->GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING, &filename))) {
									ATL::CPath fpath(filename);
									fpath.Canonicalize();
									if (fpath.m_strPath.CompareNoCase(folder_to_select) == 0) {
										try {
											if (SUCCEEDED(pfv->SelectItem(i, SVSI_EDIT | SVSI_SELECT | SVSI_DESELECTOTHERS | SVSI_ENSUREVISIBLE))) {
											}
										}
										catch (const std::exception&) {

										}
										break;
									}
								}
							}
						}
					}
				
			}
		}
	}
	CoUninitialize();

}
