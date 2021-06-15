#include <ShlObj.h>
#include <Intshcut.h>

typedef HRESULT(__stdcall* DllGetClassObj)(IN REFCLSID rclsid,IN REFIID riid, OUT LPVOID FAR* ppv);

HRESULT __stdcall CreateInstanceFromDll(
  LPCTSTR szDllName,
  IN REFCLSID rclsid,
  IUnknown* pUnkOuter,
  IN REFIID riid,
  OUT LPVOID FAR* ppv)
{
  HRESULT hr = REGDB_E_READREGDB;
  HMODULE hMod = ::LoadLibrary(szDllName);
  if (hMod)
  {
      DllGetClassObj getClassObjFunc = (DllGetClassObj)::GetProcAddress(hMod, "DllGetClassObject");
      if (getClassObjFunc)
      {
          IClassFactory* pIFactory;
          hr = getClassObjFunc(rclsid, IID_IClassFactory, (LPVOID*)&pIFactory);
          if (SUCCEEDED(hr))
          {
              hr = pIFactory->CreateInstance(pUnkOuter, riid, ppv);
              pIFactory->Release();
          }
      }
      ::FreeLibrary(hMod);
  }
  return hr;
}

int main()
{
	IUniformResourceLocator* pHook = nullptr;
	HRESULT hr = CreateInstanceFromDll(L"ieframe.dll", CLSID_InternetShortcut, NULL, IID_IUniformResourceLocator, (LPVOID*)&pHook);
}
