#include "stdafx.h"
#import "SwDocumentMgr.dll" raw_interfaces_only
#include <iostream>

#define SW_DM_KEY L"[CompanyName]:swdocmgr_general-00000-{31 times},swdocmgr_previews-00000-{31 times},swdocmgr_dimxpert-00000-{31 times},swdocmgr_geometry-00000-{31 times},swdocmgr_xml-00000-{31 times},swdocmgr_tessellation-00000-{31 times}"

int main()
{
    CoInitialize(NULL);

    CComPtr pClassFactory;

    if (SUCCEEDED(pClassFactory.CoCreateInstance(
        __uuidof(SwDocumentMgr::SwDMClassFactory), NULL, CLSCTX_INPROC_SERVER)))
    {
        CComPtr pDmApp;

        if (SUCCEEDED(pClassFactory->GetApplication(SW_DM_KEY, &pDmApp)))
        {
            long latestVers;

            HRESULT r = pDmApp->GetLatestSupportedFileVersion(&latestVers);

            if (SUCCEEDED(r))
            {
                std::cout << latestVers;
            }
            else
            {
                std::cout << "Failed to get version";
            }
        }

        pDmApp = NULL;
        pClassFactory = NULL;
        ::CoUninitialize();
    }
    else
    {
        std::cout << "Document Manager SDK is not installed";
    }
    
    std::cin.get();

    return 0;
}