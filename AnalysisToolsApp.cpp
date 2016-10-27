/////////////////////////////////////////////////////////////////////////////
// AnalysisToolsApp.cpp

#include "stdafx.h"
#include "AnalysisToolsApp.h"

BEGIN_MESSAGE_MAP(CAnalysisToolsApp, CWinApp)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CAnalysisToolsApp construction

CAnalysisToolsApp::CAnalysisToolsApp()
{
}

// The one and only CAnalysisToolsApp object
CAnalysisToolsApp theApp;

const GUID CDECL BASED_CODE _tlid = { 0x8E069405, 0x7078, 0x499A, { 0xA8, 0x44, 0x5D, 0xB9, 0xB0, 0xC, 0x92, 0x52 } };
const WORD _wVerMajor = 1;
const WORD _wVerMinor = 0;

/////////////////////////////////////////////////////////////////////////////
// CAnalysisToolsApp initialization

BOOL CAnalysisToolsApp::InitInstance()
{
  CWinApp::InitInstance();

  // Register all OLE server (factories) as running.  This enables the
  //  OLE libraries to create objects from other applications.
  COleObjectFactory::RegisterAll();

  return TRUE;
}

int CAnalysisToolsApp::ExitInstance()
{
  return CWinApp::ExitInstance();
}

// DllGetClassObject - Returns class factory
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
  AFX_MANAGE_STATE(AfxGetStaticModuleState());
  return AfxDllGetClassObject(rclsid, riid, ppv);
}

// DllCanUnloadNow - Allows COM to unload DLL
STDAPI DllCanUnloadNow(void)
{
  AFX_MANAGE_STATE(AfxGetStaticModuleState());
  return AfxDllCanUnloadNow();
}

// DllRegisterServer - Adds entries to the system registry
STDAPI DllRegisterServer(void)
{
  AFX_MANAGE_STATE(AfxGetStaticModuleState());

  if (!AfxOleRegisterTypeLib(AfxGetInstanceHandle(), _tlid))
    return SELFREG_E_TYPELIB;

  if (!COleObjectFactory::UpdateRegistryAll())
    return SELFREG_E_CLASS;

  return S_OK;
}

// DllUnregisterServer - Removes entries from the system registry
STDAPI DllUnregisterServer(void)
{
  AFX_MANAGE_STATE(AfxGetStaticModuleState());

  if (!AfxOleUnregisterTypeLib(_tlid, _wVerMajor, _wVerMinor))
    return SELFREG_E_TYPELIB;

  if (!COleObjectFactory::UpdateRegistryAll(FALSE))
    return SELFREG_E_CLASS;

  return S_OK;
}
