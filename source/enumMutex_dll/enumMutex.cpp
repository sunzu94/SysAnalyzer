#include <windows.h>
#include <stdio.h>
#include "main.h"
								  
/*cheesy unicode to ascii conversion 
void Convert(char* buf, char* wBuf, int wLen, int bLen=255){
	
	for(int i=0,j=0; i<wLen;i++){
		if(i>bLen) return;
		if(wBuf[i] != 0) buf[j++] =  wBuf[i];
	}
	
}*/
		
BOOL APIENTRY DllMain( HANDLE hModule, DWORD  ul_reason_for_call,  LPVOID lpReserved){ 
	//if(ul_reason_for_call==1){
	return TRUE;
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


