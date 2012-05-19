/*
'License:   GPL
'Copyright: 2005 iDefense a Verisign Company
'Site:      http://labs.idefense.com
'
'Author:    David Zimmer <david@idefense.com, dzzie@yahoo.com>
'
'         This program is free software; you can redistribute it and/or modify it
'         under the terms of the GNU General Public License as published by the Free
'         Software Foundation; either version 2 of the License, or (at your option)
'         any later version.
'
'         This program is distributed in the hope that it will be useful, but WITHOUT
'         ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
'         FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for
'         more details.
'
'         You should have received a copy of the GNU General Public License along with
'         this program; if not, write to the Free Software Foundation, Inc., 59 Temple
'         Place, Suite 330, Boston, MA 02111-1307 USA
*/

typedef struct{
    int dwFlag;
    int cbSize;
    int lpData;
} cpyData;

//basically used to give us a function pointer with right prototype
//and 24 byte empty buffer inline which we assemble commands into in the
//hook proceedure. 
//#define BLOCK _asm int 3
#define ALLOC_THUNK(prototype) __declspec(naked) prototype { _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop}	   

ALLOC_THUNK( HMODULE  __stdcall Real_LoadLibraryA(LPCSTR a0) );
ALLOC_THUNK( BOOL     __stdcall Real_WriteFile(HANDLE a0,LPCVOID a1,DWORD a2,LPDWORD a3,LPOVERLAPPED a4) ); 
ALLOC_THUNK( HANDLE   __stdcall Real_CreateFileA(LPCSTR a0,DWORD a1,DWORD a2,LPSECURITY_ATTRIBUTES a3,DWORD a4,DWORD a5,HANDLE a6) );
ALLOC_THUNK( HMODULE  __stdcall Real_LoadLibraryExA(LPCSTR a0,HANDLE a1,DWORD a2) );
ALLOC_THUNK( HMODULE  __stdcall Real_LoadLibraryExW(LPCWSTR a0,HANDLE a1,DWORD a2) );
ALLOC_THUNK( HMODULE  __stdcall Real_LoadLibraryW(LPCWSTR a0) );
ALLOC_THUNK( BOOL	  __stdcall Real_WriteFileEx(HANDLE a0,LPCVOID a1,DWORD a2,LPOVERLAPPED a3,LPOVERLAPPED_COMPLETION_ROUTINE a4)) ;
ALLOC_THUNK( HFILE    __stdcall Real__lclose(HFILE a0));
ALLOC_THUNK( HFILE	  __stdcall Real__lcreat(LPCSTR a0,int a1));
ALLOC_THUNK( HFILE	  __stdcall Real__lopen(LPCSTR a0,int a1));
ALLOC_THUNK( UINT	  __stdcall Real__lread(HFILE a0,LPVOID a1,UINT a2));
ALLOC_THUNK( UINT	  __stdcall Real__lwrite(HFILE a0,LPCSTR a1,UINT a2));
ALLOC_THUNK( BOOL	  __stdcall Real_CreateProcessA(LPCSTR a0,LPSTR a1,LPSECURITY_ATTRIBUTES a2,LPSECURITY_ATTRIBUTES a3,BOOL a4,DWORD a5,LPVOID a6,LPCSTR a7,struct _STARTUPINFOA* a8,LPPROCESS_INFORMATION a9));
ALLOC_THUNK( UINT	  __stdcall Real_WinExec(LPCSTR a0,UINT a1));
ALLOC_THUNK( BOOL	  __stdcall Real_DeleteFileA(LPCSTR a0));
ALLOC_THUNK( void	  __stdcall Real_ExitProcess(UINT a0));
ALLOC_THUNK( void	  __stdcall Real_ExitThread(DWORD a0));
ALLOC_THUNK( FARPROC  __stdcall Real_GetProcAddress(HMODULE a0,LPCSTR a1));
ALLOC_THUNK( DWORD	  __stdcall Real_WaitForSingleObject(HANDLE a0,DWORD a1));
ALLOC_THUNK( HANDLE	  __stdcall Real_CreateRemoteThread(HANDLE a0,LPSECURITY_ATTRIBUTES a1,DWORD a2,LPTHREAD_START_ROUTINE a3,LPVOID a4,DWORD a5,LPDWORD a6));
ALLOC_THUNK( HANDLE	  __stdcall Real_OpenProcess(DWORD a0,BOOL a1,DWORD a2));
ALLOC_THUNK( BOOL	  __stdcall Real_WriteProcessMemory(HANDLE a0,LPVOID a1,LPVOID a2,DWORD a3,LPDWORD a4));
ALLOC_THUNK( HMODULE  __stdcall Real_GetModuleHandleA(LPCSTR a0));
ALLOC_THUNK( SOCKET	  __stdcall Real_accept(SOCKET a0,sockaddr* a1,int* a2));
ALLOC_THUNK( int	  __stdcall Real_bind(SOCKET a0,SOCKADDR_IN* a1,int a2));
ALLOC_THUNK( int	  __stdcall Real_closesocket(SOCKET a0));
ALLOC_THUNK( int	  __stdcall Real_connect(SOCKET a0,SOCKADDR_IN* a1,int a2));
ALLOC_THUNK( hostent* __stdcall Real_gethostbyaddr(char* a0,int a1,int a2));
ALLOC_THUNK( hostent* __stdcall Real_gethostbyname(char* a0));
ALLOC_THUNK( int	  __stdcall Real_gethostname(char* a0,int a1));
ALLOC_THUNK( int	  __stdcall Real_listen(SOCKET a0,int a1));
ALLOC_THUNK( int	  __stdcall Real_recv(SOCKET a0,char* a1,int a2,int a3));
ALLOC_THUNK( int	  __stdcall Real_send(SOCKET a0,char* a1,int a2,int a3));
ALLOC_THUNK( int	  __stdcall Real_shutdown(SOCKET a0,int a1));
ALLOC_THUNK( SOCKET   __stdcall Real_socket(int a0,int a1,int a2));
ALLOC_THUNK( SOCKET   __stdcall Real_WSASocketA(int a0,int a1,int a2,struct _WSAPROTOCOL_INFOA* a3,GROUP a4,DWORD a5));
ALLOC_THUNK( int				Real_system(const char* cmd));
ALLOC_THUNK( FILE*				Real_fopen(const char* cmd, const char* mode));
ALLOC_THUNK( size_t				Real_fwrite(const void* a0, size_t a1, size_t a2, FILE* a3));
ALLOC_THUNK( int	  __stdcall Real_URLDownloadToFileA(int a0,char* a1, char* a2, DWORD a3, int a4));
ALLOC_THUNK( int	  __stdcall Real_URLDownloadToCacheFile(int a0,char* a1, char* a2, DWORD a3, DWORD a4, int a5));
ALLOC_THUNK( LPSTR    __stdcall Real_GetCommandLineA( VOID ) );
ALLOC_THUNK( BOOL     __stdcall Real_IsDebuggerPresent(VOID) );
ALLOC_THUNK( void               Real___setargv(void) );
ALLOC_THUNK( BOOL     __stdcall Real_GetVersionExA( LPOSVERSIONINFOA a0 ) );
ALLOC_THUNK( HGLOBAL  __stdcall Real_GlobalAlloc( UINT a0, DWORD a1 ) );
ALLOC_THUNK( DWORD    __stdcall Real_GetCurrentProcessId( VOID ) );
ALLOC_THUNK( BOOL     __stdcall Real_DebugActiveProcess( DWORD a0 ) );
ALLOC_THUNK( BOOL     __stdcall Real_ReadFile( HANDLE a0, LPVOID a1, DWORD a2, LPDWORD a3, LPOVERLAPPED a4 ) );
ALLOC_THUNK( VOID     __stdcall Real_GetSystemTime( LPSYSTEMTIME a0 ) );
ALLOC_THUNK( HANDLE   __stdcall Real_CreateMutex(int a0, int a1, int a2) );
ALLOC_THUNK( BOOL     __stdcall Real_ReadProcessMemory( HANDLE a0, PVOID64 a1, PVOID64 a2, DWORD a3, LPDWORD a4 ) );
ALLOC_THUNK( DWORD    __stdcall Real_GetVersion(void) );
ALLOC_THUNK( BOOL     __stdcall Real_CopyFile(char* lpExistingFile, char* lpNewFile, BOOL bFailIfExists) );
ALLOC_THUNK( BOOL	  __stdcall Real_InternetGetConnectedState( LPDWORD a0, DWORD a1 ) );

ALLOC_THUNK( int __stdcall Real_RegCreateKeyA ( HKEY a0, LPCSTR a1, PHKEY a2 ) );
ALLOC_THUNK( int __stdcall Real_RegDeleteKeyA ( HKEY a0, LPCSTR a1 ) );
ALLOC_THUNK( int __stdcall Real_RegDeleteValueA ( HKEY a0, LPCSTR a1 ) );
ALLOC_THUNK( int __stdcall Real_RegEnumKeyA ( HKEY a0, DWORD a1, LPSTR a2, DWORD a3 ) );
ALLOC_THUNK( int __stdcall Real_RegEnumValueA ( HKEY a0, DWORD a1, LPSTR a2, LPDWORD a3, LPDWORD a4, LPDWORD a5, LPBYTE a6, LPDWORD a7 ) );
ALLOC_THUNK( int __stdcall Real_RegQueryValueA ( HKEY a0, LPCSTR a1, LPSTR a2, PLONG   a3 ) );
ALLOC_THUNK( int __stdcall Real_RegSetValueA ( HKEY a0, LPCSTR a1, DWORD a2, LPCSTR a3, DWORD a4 ) );
ALLOC_THUNK( int __stdcall Real_RegCreateKeyExA ( HKEY a0, LPCSTR a1, DWORD a2, LPSTR a3, DWORD a4, REGSAM a5, LPSECURITY_ATTRIBUTES a6, PHKEY a7, LPDWORD a8 ) );
ALLOC_THUNK( int __stdcall Real_RegOpenKeyA ( HKEY a0, LPCSTR a1, PHKEY a2 ) );
ALLOC_THUNK( int __stdcall Real_RegOpenKeyExA ( HKEY a0, LPCSTR a1, DWORD a2, REGSAM a3, PHKEY a4 ) );
ALLOC_THUNK( int __stdcall Real_RegQueryValueExA ( HKEY a0, LPCSTR a1, LPDWORD a2, LPDWORD a3, LPBYTE a4, LPDWORD a5 ) );
ALLOC_THUNK( int __stdcall Real_RegSetValueExA ( HKEY a0, LPCSTR a1, DWORD a2, DWORD a3, CONST BYTE* a4, DWORD a5 ) );

ALLOC_THUNK( VOID __stdcall Real_Sleep( DWORD a0 ) );
ALLOC_THUNK( DWORD __stdcall Real_GetTickCount( VOID ) );

void msg(char);
void LogAPI(const char*, ...);

DWORD (__stdcall *GetModuleFileNameExA)(HANDLE hProcess, HMODULE hModule, LPSTR lpFilename, DWORD nSize);



bool Warned=false;
HWND hServer=0;
int DumpAt=0;

void GetDllPath(char* buf){ //returns full path of dll
	
	HINSTANCE hLib = LoadLibrary("psapi.dll");
	Real_GetProcAddress(hLib, "GetModuleFileNameExA");

	_asm mov GetModuleFileNameExA, eax

	if( (int)GetModuleFileNameExA == 0){
		 strcpy(buf, "api_log.dll");
		 return;
	}

	HANDLE hProc = Real_OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, 0 , GetCurrentProcessId() );
	GetModuleFileNameExA(hProc, Real_GetModuleHandleA("api_log.dll") , buf, MAX_PATH);
	CloseHandle(hProc);

}


int bInstr(char *buf, char *match, int bufLen, int matchLen){

	int i, j;

	for(i=0; i < bufLen ; i++){
		
		if(buf[i] == match[0]){
			for(j=1; j < matchLen; j++){
				if(buf[i+j] != match[j]) break;
			}
			if(j==matchLen) return i;
		}

	}

	return -1;
}


void Seek_n_Destroy_AntiVmWare(char *calledFrom, int searchSz){

	 
	char *vmdetect="\xB8\x68\x58\x4D\x56\x8B\x5D";
	void *disable ="\xB8\x68\x61\x68\x61";
	DWORD r;
   
	int ret = bInstr( calledFrom, vmdetect, searchSz, 7);
	

	if(ret > 0){
	        				
			 int *offset = (int *)( (int)calledFrom + ret);
			 
		     HANDLE h = OpenProcess(PROCESS_ALL_ACCESS, -1 , GetCurrentProcessId());
			 WriteProcessMemory(h, offset, disable, 5, &r);
			 CloseHandle(h);

			 char buf[400];
			 if(h > 0){
				sprintf(buf, "*****   Anti-Vmware Code Disabled at offset %x", (int)calledFrom + ret);
			 }else{
				sprintf(buf, "*****   Anti-Vmware Code disable FAILED at offset %x", (int)calledFrom + ret);
			 }

			 LogAPI(buf);

	}

}

char* ipfromlng(SOCKADDR_IN* sck){
	
	char *ip = (char*)malloc(16);
	unsigned char *x=0;

    _asm{
		 mov eax, [sck]
		 add eax,4
		 mov x,eax
	}

	sprintf(ip,"%d.%d.%d.%d\x00", x[0], x[1], x[2], x[3]);
	
	return ip;

}

void GetHive(HKEY hive, char* buf){

	switch((int)hive){
		case 0x80000000:
				 strcpy(buf, "HKCR\\");
				 break;
		
		case 0x80000001:
				 strcpy(buf, "HKCU\\");
				 break;

		case 0x80000002:
					 strcpy(buf, "HKLM\\");
					 break;

		case 0x80000003:
					 strcpy(buf, " HKU\\");
					 break;

		case 0x80000004 :
					 strcpy(buf, "HKPD\\");
					 break;

		case 0x80000005 :
					 strcpy(buf, "HKPD\\");
					 break;

		case 0x80000006 :
					 strcpy(buf, "HKCC\\");
					 break;
	
		default:
					 //sprintf(buf, "%x", (int)hive);
					 buf[0] = 0;

	};
}




void FindVBWindow(){
	char *vbIDEClassName = "ThunderFormDC" ;
	char *vbEXEClassName = "ThunderRT6FormDC" ;
	char *vbWindowCaption = "ApiLogger" ;

	hServer = FindWindowA( vbIDEClassName, vbWindowCaption );
	if(hServer==0) hServer = FindWindowA( vbEXEClassName, vbWindowCaption );

	if(hServer==0){
		if(!Warned){
			MessageBox(0,"Could not find msg window","",0);
			Warned=true;
		}
	}
	else{
		if(!Warned){
			//first time we are being called we could do stuff here...
			Warned=true;

		}
	}	

} 

int msg(char *Buffer){
  
  if(hServer==0) FindVBWindow();
  
  cpyData cpStructData;
  
  cpStructData.cbSize = strlen(Buffer) ;
  cpStructData.lpData = (int)Buffer;
  cpStructData.dwFlag = 3;
  
  return SendMessage(hServer, WM_COPYDATA, 0,(LPARAM)&cpStructData);

} 

void LogAPI(const char *format, ...)
{
	DWORD dwErr = GetLastError();
		
	if(format){
		char buf[1024]; 
		va_list args; 
		va_start(args,format); 
		try{
 			 _vsnprintf(buf,1024,format,args);
			 msg(buf);
		}
		catch(...){}
	}

	SetLastError(dwErr);
}


__declspec(naked) int CalledFrom(){ 
	
	_asm{
			 mov eax, [ebp+4]  //return address of parent function (were nekkid)
			 ret
	}
	
}

 

