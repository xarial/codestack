#include "stdafx.h"
#import "sldworks.tlb"
#include <iostream>

int main()
{
    ::CoInitialize(NULL);
    CComPtr<SldWorks::ISldWorks> pSwApp;

    if (SUCCEEDED(pSwApp.CoCreateInstance(
        __uuidof(SldWorks::SldWorks), NULL, CLSCTX_LOCAL_SERVER)))
    {
        pSwApp->Visible = TRUE;
        _bstr_t revNmb = pSwApp->RevisionNumber();

        std::cout << revNmb;
    }

    pSwApp = NULL;
    ::CoUninitialize();

    //wait for input (do not close console to see results)
    std::cin.get();

    return 0;
}
