#include <windows.h>
#include <stdio.h>
#include "main.h"
#include "msvbvm60.tlh"

//for Scheduled Tasks 1.0 API (win 95, NT4, 2k, XP)
//https://msdn.microsoft.com/en-us/library/windows/desktop/aa446831(v=vs.85).aspx
#include <initguid.h>
#include <ole2.h>
#include <mstask.h>
#include <msterr.h>
#include <wchar.h>

#define TASKS_TO_RETRIEVE          5


BOOL APIENTRY DllMain( HANDLE hModule, DWORD  ul_reason_for_call,  LPVOID lpReserved){ 
	//if(ul_reason_for_call==1){
	return TRUE;
}


void GetTaskDetails(FILE* f, ITaskScheduler *pITS, LPCWSTR lpcwszTaskName)
{
  
  HRESULT hr = S_OK;
  ITask *pITask = 0;

  hr = pITS->Activate(lpcwszTaskName, IID_ITask, (IUnknown**) &pITask);
  
  if (FAILED(hr))
  {
     fwprintf(f, L"Failed calling ITaskScheduler::Activate; error = 0x%x\n",hr);
     return;
  }

  LPWSTR lpwszApplicationName;
  hr = pITask->GetApplicationName(&lpwszApplicationName);

  if (FAILED(hr))
  {
     fwprintf(f, L"Failed calling ITask::GetApplicationName error = 0x%x\n", hr);
     lpwszApplicationName = 0;
  }

  LPWSTR lpwszParameters;
  hr = pITask->GetParameters(&lpwszParameters);

  if (FAILED(hr))
  {
     fwprintf(f, L"Failed calling ITask::GetApplicationName error = 0x%x\n", hr);
	 lpwszParameters = 0;
  }

  pITask->Release();

  if(lpwszApplicationName){
	  fwprintf(f, L"\t-Exe: %s\n", lpwszApplicationName);
	  CoTaskMemFree(lpwszApplicationName);
  }

  if(lpwszParameters){
	  fwprintf(f, L"\t-Params: %s\n", lpwszParameters);
	  CoTaskMemFree(lpwszParameters);
  }

}

int __stdcall EnumTasks(char* outPath){
  
  int cnt=0;
  HRESULT hr = S_OK;
  ITaskScheduler *pITS;
 
  //dont call this from a vb utilized dll no need..
  //hr = CoInitialize(NULL); 
  //if (FAILED(hr)) return hr;

  hr = CoCreateInstance(CLSID_CTaskScheduler,
						  NULL,
						  CLSCTX_INPROC_SERVER,
						  IID_ITaskScheduler,
						  (void **) &pITS);

  if (FAILED(hr)) return 0;
  
  IEnumWorkItems *pIEnum;
  hr = pITS->Enum(&pIEnum);
  if (FAILED(hr)) return 0;
  
  FILE *f = fopen(outPath,"w");
  if(f == NULL) return 0;

  LPWSTR *lpwszNames;
  DWORD dwFetchedTasks = 0;
  while (SUCCEEDED(pIEnum->Next(TASKS_TO_RETRIEVE, &lpwszNames, &dwFetchedTasks)) && (dwFetchedTasks != 0))
  {

	while (dwFetchedTasks)
    {
       fwprintf(f, L"%s\n", lpwszNames[--dwFetchedTasks]);
	   GetTaskDetails(f, pITS, lpwszNames[dwFetchedTasks]);
       CoTaskMemFree(lpwszNames[dwFetchedTasks]);
	   fputc(5,f); //parsing marker..
	   cnt++;
    }
    CoTaskMemFree(lpwszNames);
  }
  
  fclose(f);
  pITS->Release();
  pIEnum->Release();
  return cnt;
}


int __stdcall EnumMutex(char* outPath){
	
	int cnt=0;

    EnablePrivilege(SE_DEBUG_NAME);
    HMODULE hNtDll = LoadLibrary(TEXT("ntdll.dll"));
    if (!hNtDll) return -1;

    PZWQUERYSYSTEMINFORMATION ZwQuerySystemInformation =  (PZWQUERYSYSTEMINFORMATION)GetProcAddress(hNtDll, "ZwQuerySystemInformation");
    PZWDUPLICATEOBJECT ZwDuplicateObject = (PZWDUPLICATEOBJECT)GetProcAddress(hNtDll, "ZwDuplicateObject");
    PZWQUERYOBJECT ZwQueryObject = (PZWQUERYOBJECT)GetProcAddress(hNtDll, "ZwQueryObject");

	if( (int)ZwQuerySystemInformation == 0 || (int)ZwDuplicateObject  == 0 || (int)ZwQueryObject == 0) return -2;

	ULONG n = 0x1000;
    PULONG p = new ULONG[n];

	FILE *f = fopen(outPath,"w");
	if(f == NULL) return -3;

	while (ZwQuerySystemInformation(SystemHandleInformation, p, n * sizeof *p, 0) == STATUS_INFO_LENGTH_MISMATCH){
		delete [] p;
		p = new ULONG[n *= 2];
	}

    PSYSTEM_HANDLE_INFORMATION h = PSYSTEM_HANDLE_INFORMATION(p + 1);

	for (ULONG i = 0; i < *p; i++){

            HANDLE hObject;
			OBJECT_BASIC_INFORMATION obi;
            HANDLE hProcess = OpenProcess(PROCESS_DUP_HANDLE, FALSE, h[i].ProcessId);

			if (ZwDuplicateObject(hProcess, HANDLE(h[i].Handle), NtCurrentProcess(), &hObject, 0, 0, DUPLICATE_SAME_ATTRIBUTES)!= STATUS_SUCCESS){ 
                continue;
			}

            ZwQueryObject(hObject, ObjectBasicInformation, &obi, sizeof obi, &n);

            n = obi.TypeInformationLength + 2;
            POBJECT_TYPE_INFORMATION oti = POBJECT_TYPE_INFORMATION(new CHAR[n]);
            ZwQueryObject(hObject, ObjectTypeInformation, oti, n, &n);
            
			if(oti[0].Name.Length > 0 && wcscmp(oti[0].Name.Buffer,L"Mutant")==0){
				n = obi.NameInformationLength == 0 ? MAX_PATH * sizeof (WCHAR) : obi.NameInformationLength;
				POBJECT_NAME_INFORMATION oni = POBJECT_NAME_INFORMATION(new CHAR[n]);
				NTSTATUS rv = ZwQueryObject(hObject, ObjectNameInformation, oni, n, &n);
				if (NT_SUCCESS(rv)){
					if(oni[0].Name.Length > 0){
						fprintf(f,"%ld ", h[i].ProcessId);
						fprintf(f,"%.*ws\r\n", oni[0].Name.Length / 2, oni[0].Name.Buffer);
						cnt++;
					}
				}
			}

            CloseHandle(hObject);
            CloseHandle(hProcess);              
    }

    delete [] p;
	fclose(f);
	return cnt;

}


void addStr(_CollectionPtr p , char* str){
	_variant_t vv;
	vv.SetString(str);
	p->Add(&vv.GetVARIANT());
}

int __stdcall EnumMutex2(_CollectionPtr *pColl){
	
	int cnt=0;
	char buf[600];

	if(pColl==0 || *pColl == 0) return -4;

    EnablePrivilege(SE_DEBUG_NAME);
    HMODULE hNtDll = LoadLibrary(TEXT("ntdll.dll"));
    if (!hNtDll) return -1;

    PZWQUERYSYSTEMINFORMATION ZwQuerySystemInformation =  (PZWQUERYSYSTEMINFORMATION)GetProcAddress(hNtDll, "ZwQuerySystemInformation");
    PZWDUPLICATEOBJECT ZwDuplicateObject = (PZWDUPLICATEOBJECT)GetProcAddress(hNtDll, "ZwDuplicateObject");
    PZWQUERYOBJECT ZwQueryObject = (PZWQUERYOBJECT)GetProcAddress(hNtDll, "ZwQueryObject");

	if( (int)ZwQuerySystemInformation == 0 || (int)ZwDuplicateObject  == 0 || (int)ZwQueryObject == 0) return -2;

	ULONG n = 0x1000;
    PULONG p = new ULONG[n];

	while (ZwQuerySystemInformation(SystemHandleInformation, p, n * sizeof *p, 0) == STATUS_INFO_LENGTH_MISMATCH){
		delete [] p;
		p = new ULONG[n *= 2];
	}

    PSYSTEM_HANDLE_INFORMATION h = PSYSTEM_HANDLE_INFORMATION(p + 1);

	for (ULONG i = 0; i < *p; i++){

            HANDLE hObject;
			OBJECT_BASIC_INFORMATION obi;
            HANDLE hProcess = OpenProcess(PROCESS_DUP_HANDLE, FALSE, h[i].ProcessId);

			if (ZwDuplicateObject(hProcess, HANDLE(h[i].Handle), NtCurrentProcess(), &hObject, 0, 0, DUPLICATE_SAME_ATTRIBUTES)!= STATUS_SUCCESS){ 
                continue;
			}

            ZwQueryObject(hObject, ObjectBasicInformation, &obi, sizeof obi, &n);

            n = obi.TypeInformationLength + 2;
            POBJECT_TYPE_INFORMATION oti = POBJECT_TYPE_INFORMATION(new CHAR[n]);
            ZwQueryObject(hObject, ObjectTypeInformation, oti, n, &n);
            
			if(oti[0].Name.Length > 0 && wcscmp(oti[0].Name.Buffer,L"Mutant")==0){
				n = obi.NameInformationLength == 0 ? MAX_PATH * sizeof (WCHAR) : obi.NameInformationLength;
				POBJECT_NAME_INFORMATION oni = POBJECT_NAME_INFORMATION(new CHAR[n]);
				NTSTATUS rv = ZwQueryObject(hObject, ObjectNameInformation, oni, n, &n);
				if (NT_SUCCESS(rv)){
					if(oni[0].Name.Length > 0){
						_snprintf(buf, sizeof(buf)-1, "%ld %.*ws", h[i].ProcessId, oni[0].Name.Length / 2, oni[0].Name.Buffer);
						addStr(*pColl,buf);
						cnt++;
					}
				}
			}

            CloseHandle(hObject);
            CloseHandle(hProcess);              
    }

    delete [] p;
	return cnt;

}


