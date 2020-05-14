#include "stdafx.h"
#import "sldworks.tlb"
#include <iostream>
#include <windows.h>
#include <string>
#include <chrono>
#include <thread>

HRESULT StartSwProcess(LPCWSTR appPath, int& prcId)
{
    prcId = -1;

    STARTUPINFO si;
    PROCESS_INFORMATION pi;

    ZeroMemory(&si, sizeof(si));

    HRESULT res = E_FAIL;
    
    if(CreateProcess(L"C:\\Program Files\\SOLIDWORKS Corp\\SOLIDWORKS\\SLDWORKS.exe",
        L"", NULL, NULL, FALSE, 0,
        NULL, NULL, &si, &pi))
    {
        CloseHandle(pi.hProcess);
        CloseHandle(pi.hThread);

        prcId = pi.dwProcessId;

        res = S_OK;
    }
    
    return res;
}

HRESULT GetSwAppFromProcess(int prcId, SldWorks::ISldWorks** pSwApp)
{
    HRESULT res = E_FAIL;

    CComPtr<IRunningObjectTable> pRot;
    CComPtr<IBindCtx> pBindingContext;

    if (SUCCEEDED(CreateBindCtx(0, &pBindingContext)))
    {
        if (GetRunningObjectTable(0, &pRot) == S_OK)
        {
            CComPtr<IEnumMoniker> pEnumMoniker;
            if (SUCCEEDED(pRot->EnumRunning(&pEnumMoniker)))
            {
                WCHAR szMonikerName[30];
                swprintf_s(szMonikerName, 30, L"SolidWorks_PID_%d", prcId);

                ULONG fetched;
                CComPtr<IMoniker> pMon;

                while (pEnumMoniker->Next(1, &pMon, &fetched) == S_OK)
                {
                    LPOLESTR pName;
                    pMon->GetDisplayName(pBindingContext, NULL, &pName);

                    if (wcscmp(pName, szMonikerName) == 0)
                    {
                        CComPtr<IUnknown> pUnk;

                        if (SUCCEEDED(pRot->GetObjectW(pMon, &pUnk)))
                        {
                            if (SUCCEEDED(pUnk->QueryInterface(_uuidof(SldWorks::ISldWorks), (void**)pSwApp)))
                            {    
                                res = S_OK;
                                break;
                            }
                        }
                    }

                    pMon = NULL;
                }
            }
        }
    }

    pRot = NULL;
    pBindingContext = NULL;

    return res;
}

HRESULT ConnectToSwApp(LPCWSTR appPath, SldWorks::ISldWorks** pSwApp, int timeoutSec) 
{
    HRESULT res = E_FAIL;

    int prcId;

    if (SUCCEEDED(StartSwProcess(appPath, prcId)))
    {
        auto start = std::chrono::high_resolution_clock::now();
        
        while (FAILED(GetSwAppFromProcess(prcId, pSwApp)))
        {
            std::this_thread::sleep_for(std::chrono::milliseconds(200));
            auto end = std::chrono::high_resolution_clock::now();
            std::chrono::duration<double, std::milli> elapsed = end - start;

            if (elapsed.count() > timeoutSec * 1000)
            {
                throw std::runtime_error("Timeout");
            }
        }

        res = S_OK;
    }

    return res;
}

int main()
{
    ::CoInitialize(NULL);
    
    CComPtr<SldWorks::ISldWorks> pSwApp;

    try 
    {
        if (SUCCEEDED(ConnectToSwApp(L"C:\\Program Files\\SOLIDWORKS Corp\\SOLIDWORKS (2)\\SLDWORKS.exe", 
            &pSwApp, 10))) 
        {
            _bstr_t revNmb = pSwApp->RevisionNumber();
            std::cout << revNmb;
        }
    }
    catch (std::runtime_error& e) 
    {
        std::cout << e.what() << std::endl;
    }

    pSwApp = NULL;
    
    ::CoUninitialize();

    //wait for input (do not close console to see results)
    std::cin.get();
    
    return 0;
}

