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

#define _WIN32_WINNT 0x0401  //for IsDebuggerPresent 
#include <windows.h>
#include <stdio.h>
#include <string.h>
#include <stdlib.h>
#include <stdarg.h>
#include <wininet.h>
//#include <Winsock2.h>
#include <tlhelp32.h>

#pragma warning(disable:4996)
#pragma comment(lib, "Wininet.lib")
void InstallHooks(void);

 

#include "NtHookEngine.h"
#include "main.h"   //contains a bunch of library functions in it too..

//todo:  
//      block everyway you can find to delete files
//      protect analysis apps from OpenProcess                5.17.12
//      include process name in writeprocessmemory dumps
//      hook toolhelp snapshots and hide analysis apps.
//      getmodulehandle - hide api_log.dll                    5.17.12
//      hook SetWindowsHook/Ex
//      NtCreateThreadEx?  http://chmag.in/article/mar2011/remote-thread-execution-system-process-using-ntcreatethreadex-vista-windows-7
//      QueueUserAPC? 
//      SetThreadContext
//      main app: hardcode scan, look for new (untrusted) dlls in new processes
//                hardcore scan, look for RWE memory sections in a process that arent in a module.
//
//      config options: ignore/allow Sleep, normal/advance GetTickCount
//                      allow/block OpenProcess?

bool Installed =false;

void Closing(void){ msg("***** Injected Process Terminated *****"); exit(0);}
	
extern "C" __declspec (dllexport) int NullSub(void){ return 1;}

//Config options..these must all default to 0 because default windProc response = 0 if unhandled by client..
int noSleep = 0;
int noRegistry = 0;
int blockOpenProcess = 0;
int noGetProc = 0;
int queryGetTick = 0;
int blockDebugControl = 0;
int ignoreExitProcess = 0;
int myPID = 0;

BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  ul_reason_for_call, 
                       LPVOID lpReserved
					 )
{

    if(!Installed){
		 Installed=true;
		 InstallHooks();
		 atexit(Closing);
	}

	return TRUE;
}

char *strlower(char *s)		
{
  char *cp;
  if ( !(cp=s) )
    return NULL;

  while ( *s != 0 ) {
    *s = tolower( *s );
    s++;
  }
  return cp;
}

char* findProcessByPid(int pid){
	
	PROCESSENTRY32 pe;
    HANDLE hSnap;
	int cnt=0;
    char buf[200];

    pe.dwSize = sizeof(pe);
    hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    
    Process32First( hSnap, &pe);
    if( pe.th32ProcessID == pid ) return strlower(strdup(pe.szExeFile));

    while( Process32Next(hSnap, &pe) ){
		if( pe.th32ProcessID == pid ) return strlower(strdup(pe.szExeFile));
	}

	sprintf(buf, "-- pid %x not in ToolHelp Api! --", pid);
	
	return strdup(buf);

}


//___________________________________________________hook implementations _________

BOOL __stdcall My_CloseHandle(HANDLE a0)
{
    
	LogAPI("%x     CloseHandle(h=%x)", CalledFrom(), a0);

    BOOL ret = 0;
    try {
        ret = Real_CloseHandle(a0);
    } 
	catch(...){	} 
    return ret;
}


int My_ZwQuerySystemInformation(int SystemInformationClass, int SystemInformation, int SystemInformationLength, int ReturnLength){

	//todo if SystemProcessInformation rename tool processes to bs..
	LogAPI("%x     ZwQuerySystemInformation(class=%x)", CalledFrom(), SystemInformationClass);

	return Real_ZwQuerySystemInformation(SystemInformationClass, SystemInformation, SystemInformationLength, ReturnLength);

}

int My_ZwSystemDebugControl( int Command, int InputBuffer, int InputBufferLength,int OutputBuffer, int OutputBufferLength, int ReturnLength){

	char* blk = blockDebugControl ? "BLOCKED" : "";

	int ret = 0;
	
	if(blockDebugControl == 0){ //causes endless loop in sample to ignore this..
		ret = Real_ZwSystemDebugControl( Command,  InputBuffer, InputBufferLength,OutputBuffer, OutputBufferLength, ReturnLength);
	}
	
	LogAPI("%x     ZwSystemDebugControl(cmd=%x, dest=%x, size=%x, src=%x, sz=%x) = %x - %s", CalledFrom(), Command, OutputBuffer, OutputBufferLength, InputBuffer, InputBufferLength, ret, blk);

	return ret;

}

VOID __stdcall My_Sleep( DWORD a0 )
{
	
	DWORD minToLog = 3000; //if its under 3 seconds who cares just do it dont spam me with it...
	
	if( a0 > minToLog){
		if(noSleep){
			LogAPI("%x     Sleep(%x) - IGNORED", CalledFrom(), a0);
			return;
		}

		LogAPI("%x     Sleep(%x)", CalledFrom(), a0);
	}

	Real_Sleep(a0);
	return;

}

DWORD __stdcall My_GetTickCount( VOID )
{

	DWORD  ret = 0;
	int verbose = 0; //this can be called a metric shit ton of times...

	if(queryGetTick){
		ret = msg("***config:getTickValue");
		if(ret!=0){
			if(verbose) LogAPI("%x     GetTickCount() OVERRIDDEN = %x", CalledFrom(), ret);
			return ret;
		}
	}
	
	try{
		ret = Real_GetTickCount();
	}
	catch(...){}

	if(verbose) LogAPI("%x     GetTickCount() = %x", CalledFrom(), ret);

	return ret;
}


HANDLE __stdcall My_CreateFileA(LPCSTR a0,DWORD a1,DWORD a2,LPSECURITY_ATTRIBUTES a3,DWORD a4,DWORD a5,HANDLE a6)
{
	
	char *calledFrom=0;

	LogAPI("%x     CreateFileA(%s)", CalledFrom(), a0);

    HANDLE ret = 0;
    try{
        ret = Real_CreateFileA(a0, a1, a2, a3, a4, a5, a6);
    }
	catch(...){} 
	
	return ret;

}

BOOL __stdcall My_WriteFile(HANDLE a0,LPCVOID a1,DWORD a2,LPDWORD a3,LPOVERLAPPED a4)
{
    
	LogAPI("%x     WriteFile(h=%x)", CalledFrom(), a0);

    BOOL ret = 0;
    try {
        ret = Real_WriteFile(a0, a1, a2, a3, a4);
    } 
	catch(...){	} 
    return ret;
}
 
HFILE __stdcall My__lcreat(LPCSTR a0,int a1)
{

	LogAPI("%x     _lcreat(%s,%x)", CalledFrom(), a0, a1);

    HFILE ret = 0;
    try {
        ret = Real__lcreat(a0, a1);
    } 
	catch(...){	} 
    return ret;
}

HFILE __stdcall My__lopen(LPCSTR a0, int a1)
{
   
	
	LogAPI("%x     _lopen(%s,%x)", CalledFrom(), a0, a1);

    HFILE ret = 0;
    try {
        ret = Real__lopen(a0, a1);
    }
	catch(...){	} 

    return ret;
}

UINT __stdcall My__lread(HFILE a0,LPVOID a1,UINT a2)
{
   
	LogAPI("%x     _lread(%x,%x,%x)", CalledFrom(), a0, a1, a2);

    UINT ret = 0;
    try {
        ret = Real__lread(a0, a1, a2);
    }
	catch(...){	} 

    return ret;
}

UINT __stdcall My__lwrite(HFILE a0,LPCSTR a1,UINT a2)
{
    
	LogAPI("%x     _lwrite(h=%x)", CalledFrom(), a0);

    UINT ret = 0;
    try {
        ret = Real__lwrite(a0, a1, a2);
    }
	catch(...){	} 

    return ret;
}




BOOL __stdcall My_WriteFileEx(HANDLE a0,LPCVOID a1,DWORD a2,LPOVERLAPPED a3,LPOVERLAPPED_COMPLETION_ROUTINE a4)
{
  
    LogAPI("%x     WriteFileEx(h=%x)", CalledFrom(), a0);

    BOOL ret = 0;
    try {
        ret = Real_WriteFileEx(a0, a1, a2, a3, a4);
    }
	catch(...){	} 

    return ret;
}

DWORD __stdcall My_WaitForSingleObject(HANDLE a0,DWORD a1)
{
   
	LogAPI("%x     WaitForSingleObject(%x,%x)", CalledFrom(), a0, a1);

    DWORD ret = 0;
    try {
        ret = Real_WaitForSingleObject(a0, a1);
    }
	catch(...){	} 

    return ret;
}


//_________ws2_32__________________________________________________________

SOCKET __stdcall My_accept(SOCKET a0,sockaddr* a1,int* a2)
{

    SOCKET ret = 0;
    try {
        ret = Real_accept(a0, a1, a2);
    }
	catch(...){	} 

	LogAPI("%x     accept(%x,%x,%x) = %x", CalledFrom(), a0, a1, a2, ret);

    return ret;
}

int __stdcall My_bind(SOCKET a0,SOCKADDR_IN* a1, int a2)
{
    
	LogAPI("%x     bind(%x, port=%ld)", CalledFrom(), a0, htons(a1->sin_port) );

    int ret = 0;
    try {
        ret = Real_bind(a0, a1, a2);
    }
	catch(...){	} 

    return ret;
}

int __stdcall My_closesocket(SOCKET a0)
{
    
	LogAPI("%x     closesocket(%x)", CalledFrom(), a0);

    int ret = 0;
    try {
        ret = Real_closesocket(a0);
    }
	catch(...){	} 

    return ret;
}

int __stdcall My_connect(SOCKET a0,SOCKADDR_IN* a1,int a2)
{
    
	char* ip=0;	
	ip=ipfromlng(a1);
	
	LogAPI("%x     connect(s=%x, host=%s:%d )", CalledFrom(), a0, ip, htons(a1->sin_port) );
	
	free(ip);

    int ret = 0;
    try {
        ret = Real_connect(a0, a1, a2);
    }
	catch(...){	} 

    return ret;
}

hostent* __stdcall My_gethostbyaddr(char* a0,int a1,int a2)
{
    
	LogAPI("%x     gethostbyaddr(%s)", CalledFrom(), a0);

    hostent* ret = 0;
    try {
        ret = Real_gethostbyaddr(a0, a1, a2);
    }
	catch(...){	} 

    return ret;
}

hostent* __stdcall My_gethostbyname(char* a0)
{
	LogAPI("%x     gethostbyname(%s)", CalledFrom(), a0);

    hostent* ret = 0;
    try {
        ret = Real_gethostbyname(a0);
    }
	catch(...){	} 

    return ret;
}

int __stdcall My_listen(SOCKET a0,int a1)
{
    
	LogAPI("%x     listen(h=%x )", CalledFrom(), a0);

    int ret = 0;
    try {
        ret = Real_listen(a0, a1);
    }
	catch(...){	} 

    return ret;
}

int __stdcall My_recv(SOCKET a0,char* a1,int a2,int a3)
{

    int ret = 0;
    try {
        ret = Real_recv(a0, a1, a2, a3);
    } 
	catch(...){	} 

	LogAPI("%x     recv(h=%x, buf=%x) = %x bytes", CalledFrom(), a0, a1, ret);

    return ret;
}

int __stdcall My_send(SOCKET a0,char* a1,int a2,int a3)
{
    
	LogAPI("%x     send(h=%x, buf=%x, sz=%x)", CalledFrom(), a0, a1, a2);
    int ret = 0;

    try {
        ret = Real_send(a0, a1, a2, a3);
	}
	catch(...){	} 

    return ret;
}

int __stdcall My_shutdown(SOCKET a0,int a1)
{
    
	LogAPI("%x     shutdown()",  CalledFrom());

    int ret = 0;
    try {
        ret = Real_shutdown(a0, a1);
    }
	catch(...){	} 

    return ret;
}

SOCKET __stdcall My_socket(int a0,int a1,int a2)
{

    SOCKET ret = 0;
    try {
        ret = Real_socket(a0, a1, a2);
    }
	catch(...){	} 

	LogAPI("%x     socket(family=%x,type=%x,proto=%x) = %x", CalledFrom(), a0, a1, a2, ret);

    return ret;
}

//untested
int My_URLDownloadToFileA(int a0,char* a1, char* a2, DWORD a3, int a4)
{
	
    SOCKET ret = 0;
    try {
        ret = Real_URLDownloadToFileA(a0, a1, a2, a3, a4);
    }
	catch(...){	} 

	char* sret = (ret == S_OK) ? "OK" : "FAILED";

	LogAPI("%x     URLDownloadToFile(%s, %s) = %s", CalledFrom(), a1, a2, sret);


    return ret;
}

//untested
int My_URLDownloadToCacheFile(int a0,char* a1, char* a2, DWORD a3, DWORD a4, int a5)
{
	
    SOCKET ret = 0;
    try {
        ret = Real_URLDownloadToCacheFile(a0, a1, a2, a3, a4, a5);
    }
	catch(...){	} 

	char* sret = (ret == S_OK) ? "OK" : "FAILED";

	LogAPI("%x     URLDownloadToCacheFile(%s, %s)", CalledFrom(), a1, a2, sret);

    return ret;
}

void __stdcall My_ExitProcess(UINT a0)
{
    
	char* s = ignoreExitProcess ? " - IGNORED" : "";

	LogAPI("%x     ExitProcess() %s", CalledFrom(),s);

    if(ignoreExitProcess==0) Real_ExitProcess(a0);

}

void __stdcall My_ExitThread(DWORD a0)
{
    
	LogAPI("%x     ExitThread()", CalledFrom());

    try {
        Real_ExitThread(a0);
    }
	catch(...){	} 

}

HANDLE __stdcall My_OpenProcess(DWORD a0,BOOL a1,DWORD a2)
{

	HANDLE ret = 0;
	int i=0;
	
	char *target = findProcessByPid(a2);

	if(blockOpenProcess){
		LogAPI("%x     OpenProcess(pid:%x) %s -  BLOCKED", CalledFrom(), a2, target );
		free(target);
		return 0;
	}

	char* tools[] = {"api_logger.exe","sysanalyzer.exe","ollydbg.exe","windump.exe","sniff_hit.exe",0};

	while(tools[i]){
		if(strcmp(tools[i],target) == 0){
			LogAPI("%x     OpenProcess(%s) -  PROTECTED", CalledFrom(), target );
			free(target);
			return 0;
		}
		i++;
	}

    try {
        ret = Real_OpenProcess(a0, a1, a2);
    }
	catch(...){	} 

	LogAPI("%x     OpenProcess(pid=%ld) = 0x%x  - %s", CalledFrom(), a2, ret, target );
	free(target);

    return ret;
}

HMODULE __stdcall My_GetModuleHandleA(char* a0)
{
	/* nice idea but vmware hook.dll freaks out or i suck one of the two...
	char *my = strlower(strdup(a0)); //may not be a writable string in which case strlower would crash us..

	if( strcmp(my, "api_log.dll") == 0 || strcmp(my, "api_log") == 0){
		LogAPI("%x     GetModuleHandleA(%s) - HIDDEN", CalledFrom(), a0);
		free(my);
		return 0;
	}
	
	free(my);*/

	LogAPI("%x     GetModuleHandleA(%s)", CalledFrom(), a0);

    HMODULE ret = 0;
    try {
        ret = Real_GetModuleHandleA(a0);
    }
	catch(...){	} 

    return ret;
}



UINT __stdcall My_WinExec(LPCSTR a0,UINT a1)
{

	LogAPI("%x     WinExec(%s,%x)", CalledFrom(), a0, a1);

    UINT ret = 0;
    try {
        ret = Real_WinExec(a0, a1);
    }
	catch(...){	} 

    return ret;


}

BOOL __stdcall My_DeleteFileA(LPCSTR a0)
{
	
 	LogAPI("%x     Skipping DeleteFileA(%s)", CalledFrom(), a0); //deleting is never cool nonet or not
	return 0;
	 

}

BOOL __stdcall My_CreateProcessA(LPCSTR a0,LPSTR a1,LPSECURITY_ATTRIBUTES a2,LPSECURITY_ATTRIBUTES a3,BOOL a4,DWORD a5,LPVOID a6,LPCSTR a7,struct _STARTUPINFOA* si,LPPROCESS_INFORMATION pi)
{

	//type def unsigned long (__stdcall  *lpfnLoadLib)(void *);

	char* flags = a5 == CREATE_SUSPENDED ? "CREATE_SUSPENDED" : "";

	int buflen, ret, lpfnLoadLib ; 
	unsigned long writeLen, hThread;
	HANDLE hProcess, lpdllPath;
   
	char dllPath[MAX_PATH]; // = "api_log.dll\x00";

    BOOL retv = 0;
    try {

		if(a0 && strstr(a0,"git.exe") > 0) return 0;
		if(a0 && !a1)	LogAPI("%x     CreateProcessA(%s)", CalledFrom(), a0);
		if(a1 && !a0)	LogAPI("%x     CreateProcessA("", %s)", CalledFrom(), a1);
        if(a1 && a0)	LogAPI("%x     CreateProcessA(%s, %s)", CalledFrom(), a0, a1);

		retv = Real_CreateProcessA(a0, a1, NULL, NULL, FALSE, CREATE_SUSPENDED, NULL, NULL, si, pi);
		
		GetDllPath( (char*)dllPath );
		buflen = strlen(dllPath);

		LogAPI("*****   Injecting %s into new process", dllPath);

		hProcess = Real_OpenProcess(PROCESS_ALL_ACCESS, 0, pi->dwProcessId);
		LogAPI("*****   OpenProcess Handle=%x",hProcess);
              
		lpdllPath = VirtualAllocEx(hProcess, 0, buflen, MEM_COMMIT, PAGE_READWRITE);
		LogAPI("*****   Remote Allocation base: %x", lpdllPath);
        
		ret = Real_WriteProcessMemory(hProcess, lpdllPath, dllPath, buflen, &writeLen);
		LogAPI("*****   WriteProcessMemory=%x BufLen=%x  BytesWritten:%x", ret, buflen, writeLen);
            
		lpfnLoadLib = (int)Real_GetProcAddress(Real_GetModuleHandleA("kernel32.dll"), "LoadLibraryA");
		
		//_asm mov lpfnLoadLib, eax

		LogAPI("*****   LoadLibraryA=%x",lpfnLoadLib);
    
		ret = (int)Real_CreateRemoteThread(hProcess, 0, 0, (LPTHREAD_START_ROUTINE)lpfnLoadLib, lpdllPath, 0, &hThread);
		LogAPI("*****   CreateRemoteThread=%x" , ret);
            
	    if(a5 != CREATE_SUSPENDED) ResumeThread(pi->hThread);

		/*if(strlen(flags) > 1){
			LogAPI("%x     CreateProcessA(%s,%s,%x,%s, %s) hProc=%x hThread=%x", CalledFrom(), a0, a1, a6, a7, flags, pi->hProcess, pi->hThread);
		}else{
			LogAPI("%x     CreateProcessA(%s,%s,%x,%s,flags=0x%x) hProc=%x hThread=%x", CalledFrom(), a0, a1, a6, a7, a5, pi->hProcess, pi->hThread);
		}*/

		Real_CloseHandle(hProcess);

	}
	catch(...){	} 

    return retv;



}

HANDLE __stdcall My_CreateRemoteThread(HANDLE a0,LPSECURITY_ATTRIBUTES a1,DWORD a2,LPTHREAD_START_ROUTINE a3,LPVOID a4,DWORD a5,LPDWORD a6)
{
	

	LogAPI("%x     CreateRemoteThread(h=%x, start=%x)", CalledFrom(), a0,a3);

    HANDLE ret = 0;
    try {
        ret = Real_CreateRemoteThread(a0, a1, a2, a3, a4, a5, a6);
    }
	catch(...){	} 

    return ret;

}

BOOL __stdcall My_WriteProcessMemory(HANDLE a0,LPVOID a1,LPVOID a2,DWORD a3,LPDWORD a4)
{

	/*  
	    BOOL WINAPI WriteProcessMemory(
		  __in   HANDLE hProcess,
		  __in   LPVOID lpBaseAddress,
		  __in   LPCVOID lpBuffer,
		  __in   SIZE_T nSize,
		  __out  SIZE_T *lpNumberOfBytesWritten
		);
	*/

	char buf[255];
	DWORD written=0;

	//todo lookup handle and relate back to which process name it was handed out for...
	sprintf(buf, "c:\\wpm_h_%x_mem_%x.bin", a0, a1);
	HANDLE h = Real_CreateFileA(buf, GENERIC_READ|GENERIC_WRITE ,0,0,OPEN_ALWAYS,FILE_ATTRIBUTE_NORMAL,0); 
	/*Real_*/WriteFile(h,a2,a3,&written,0);
	Real_CloseHandle(h);

	LogAPI("%x     WriteProcessMemory(h=%x,base=%x,buf=%x,len=%x) Saved as %s", CalledFrom(), a0,a1,a2,a3,buf);


    BOOL ret = 0;
    try {
		
		//hexdump( (unsigned char*) a2, a3 );
        ret = Real_WriteProcessMemory(a0, a1, a2, a3, a4);

    }
	catch(...){	} 

    return ret;
}

 
// ________________________________________________  monitored ________________

HMODULE __stdcall My_LoadLibraryA(char* a0)
{

	
	HMODULE ret =0;
	try {
		ret = Real_LoadLibraryA(a0);
	}
	catch(...){	} 

	LogAPI("%x     LoadLibraryA(%s)=%x", CalledFrom(),  a0, ret);

	return ret;

}


 
FARPROC __stdcall My_GetProcAddress(HMODULE a0,LPCSTR a1)
{
	
	//Real_GetProcAddress is 
	//directly used in other code in here, if you want to
	//disable logging just comment out LogAPI line and not
	//the actual addhook call

    FARPROC ret = 0;
    try {
        ret = Real_GetProcAddress(a0, a1);
    }
	catch(...){	} 
	
	if(noGetProc==0) LogAPI("%x     GetProcAddress(%s)", CalledFrom(), a1);

    return ret;
}


//-----------------------------------------------------------------

LPSTR __stdcall My_GetCommandLineA( VOID )
{
	
	LogAPI("%x     GetCommandLineA()", CalledFrom() );

	LPSTR  ret = 0;
	try{
		ret = Real_GetCommandLineA();
	}
	catch(...){}
	
	return ret;
}

BOOL __stdcall My_IsDebuggerPresent(void)
{
	LogAPI("%x     IsDebuggerPresent() = 0", CalledFrom() );

	return 0;

	/*BOOL  ret = 0;
	try{
		ret = Real_IsDebuggerPresent();
	}
	catch(...){}
	
	return ret;*/
}

BOOL __stdcall My_GetVersionExA( LPOSVERSIONINFOA a0 )
{

	LogAPI("%x     GetVersionExA()", CalledFrom() );

	BOOL  ret = 0;
	try{
		ret = Real_GetVersionExA(a0);
	}
	catch(...){}

	return ret;
}

HGLOBAL __stdcall My_GlobalAlloc( UINT a0, DWORD a1 )
{
	
	LogAPI("%x     GlobalAlloc()", CalledFrom() );

	HGLOBAL  ret = 0;
	try{
		ret = Real_GlobalAlloc(a0,a1);
	}
	catch(...){}

	return ret;
}
DWORD __stdcall My_GetCurrentProcessId( VOID )
{
	
	
	DWORD  ret = 0;
	try{
		ret = Real_GetCurrentProcessId();
	}
	catch(...){}
	
	LogAPI("%x     GetCurrentProcessId()=%d", CalledFrom(), ret);

	return ret;
}
BOOL __stdcall My_DebugActiveProcess( DWORD a0 )
{
	
	LogAPI("%x     DebugActiveProcess()", CalledFrom() );

	BOOL  ret = 0;
	try{
		ret = Real_DebugActiveProcess(a0);
	}
	catch(...){}

	return ret;
}
BOOL __stdcall My_ReadFile( HANDLE a0, LPVOID a1, DWORD a2, LPDWORD a3, LPOVERLAPPED a4 )
{
	
	LogAPI("%x     ReadFile()", CalledFrom() );

	BOOL  ret = 0;
	try{
		ret = Real_ReadFile(a0,a1,a2,a3,a4);
	}
	catch(...){}

	return ret;
}

VOID __stdcall My_GetSystemTime( LPSYSTEMTIME a0 )
{
	
	LogAPI("%x     GetSystemTime()", CalledFrom() );


	try{
		Real_GetSystemTime(a0);
	}
	catch(...){}


}

HANDLE __stdcall My_CreateMutex(int a0, int a1, int a2){


	HANDLE ret = 0;
	try{
		ret = Real_CreateMutex(a0,a1,a2);
	}
	catch(...){}
	
	LogAPI("%x     CreateMutex(%s) = 0x%x", CalledFrom(), a2, ret );

	return ret;

}

BOOL __stdcall My_ReadProcessMemory( HANDLE a0, PVOID64 a1, PVOID64 a2, DWORD a3, LPDWORD a4 )
{

	LogAPI("%x     ReadProcessMemory(h=%x)", CalledFrom(), a0);

	BOOL  ret = 0;
	try{
		ret = Real_ReadProcessMemory(a0,a1,a2,a3,a4);
	}
	catch(...){}

	return ret;
}


DWORD __stdcall My_GetVersion(void)
{
	LogAPI("%x     GetVersion()", CalledFrom());

	DWORD  ret = 0;
	try{
		ret = Real_GetVersion();
	}
	catch(...){}

	return ret;
}


BOOL My_CopyFile(char* a0, char* a1, BOOL a2){

	LogAPI("%x     Copy(%s->%s)", CalledFrom(), a0, a1 );

	BOOL  ret = 0;
	try{
		ret = Real_CopyFile(a0,a1,a2);
	}
	catch(...){}

	return ret;
}


BOOL __stdcall My_InternetGetConnectedState( LPDWORD a0, DWORD a1 ){

	LogAPI("%x     InternetGetConnectedState()", CalledFrom() );

	BOOL  ret = 0;
	try{
		ret = Real_InternetGetConnectedState(a0,a1);
	}
	catch(...){}

	return 1; //vm machines mess this up sometimes

}

//------------------------------------------------------------------
int __stdcall My_RegCreateKeyA ( HKEY a0, LPCSTR a1, PHKEY a2 )
{
	char h[6];
	GetHive(a0,h);
	LogAPI("%x     RegCreateKeyA (%s%s)", CalledFrom() ,h, a1 );

	
	int  ret = 0;
	try{
		ret = Real_RegCreateKeyA (a0,a1,a2);
	}
	catch(...){}

	return ret;
}
int __stdcall My_RegDeleteKeyA ( HKEY a0, LPCSTR a1 )
{
	char h[6];
	GetHive(a0,h);

	LogAPI("%x     RegDeleteKeyA (%s%s)", CalledFrom(),h, a1 );

	int  ret = 0;
	try{
		ret = Real_RegDeleteKeyA (a0,a1);
	}
	catch(...){}

	return ret;
}
int __stdcall My_RegDeleteValueA ( HKEY a0, LPCSTR a1 )
{
	char h[6];
	GetHive(a0,h);

	LogAPI("%x     RegDeleteValueA (%s%s)", CalledFrom(), h, a1 );

	int  ret = 0;
	try{
		ret = Real_RegDeleteValueA (a0,a1);
	}
	catch(...){}

	return ret;
}
int __stdcall My_RegEnumKeyA ( HKEY a0, DWORD a1, LPSTR a2, DWORD a3 )
{
	char h[6];
	GetHive(a0,h);

	int  ret = 0;
	try{
		ret = Real_RegEnumKeyA (a0,a1,a2,a3);
	}
	catch(...){}
	
	LogAPI("%x     RegEnumKeyA(%s%s)", CalledFrom(), h, a2 );

	return ret;
}
int __stdcall My_RegEnumValueA ( HKEY a0, DWORD a1, LPSTR a2, LPDWORD a3, LPDWORD a4, LPDWORD a5, LPBYTE a6, LPDWORD a7 )
{
	char h[6];
	GetHive(a0,h);

	int  ret = 0;
	try{
		ret = Real_RegEnumValueA (a0,a1,a2,a3,a4,a5,a6,a7);
	}
	catch(...){}
	
	LogAPI("%x     RegEnumValueA %s%s)", CalledFrom(), h, a2 );

	return ret;
}
int __stdcall My_RegQueryValueA ( HKEY a0, LPCSTR a1, LPSTR a2, PLONG   a3 )
{
	char h[6];
	GetHive(a0, h);

	int  ret = 0;
	try{
		ret = Real_RegQueryValueA (a0,a1,a2,a3);
	}
	catch(...){}
	
	LogAPI("%x     RegQueryValueA (%s%s,%s)", CalledFrom(), h, a1, a2 );

	return ret;
}
int __stdcall My_RegSetValueA ( HKEY a0, LPCSTR a1, DWORD a2, LPCSTR a3, DWORD a4 )
{
	char h[6];
	GetHive(a0,h);

	LogAPI("%x     RegSetValueA (%s%s,%s)", CalledFrom(), h, a1,a3 );

	int  ret = 0;
	try{
		ret = Real_RegSetValueA (a0,a1,a2,a3,a4);
	}
	catch(...){}

	return ret;
}


int __stdcall My_RegCreateKeyExA ( HKEY a0, LPCSTR a1, DWORD a2, LPSTR a3, DWORD a4, REGSAM a5, LPSECURITY_ATTRIBUTES a6, PHKEY a7, LPDWORD a8 )
{
	char h[6];
	GetHive(a0,h);

	LogAPI("%x     RegCreateKeyExA (%s%s,%s)", CalledFrom(), h, a1 , a3 );

	int  ret = 0;
	try{
		ret = Real_RegCreateKeyExA (a0,a1,a2,a3,a4,a5,a6,a7,a8);
	}
	catch(...){}

	return ret;
}
int __stdcall My_RegOpenKeyA ( HKEY a0, LPCSTR a1, PHKEY a2 )
{
	char h[6];
	GetHive(a0,h);

	LogAPI("%x     RegOpenKeyA (%s%s)", CalledFrom(), h, a1 );

	int  ret = 0;
	try{
		ret = Real_RegOpenKeyA (a0,a1,a2);
	}
	catch(...){}

	return ret;
}
int __stdcall My_RegOpenKeyExA ( HKEY a0, LPCSTR a1, DWORD a2, REGSAM a3, PHKEY a4 )
{
	char h[6];
	GetHive(a0,h);

	LogAPI("%x     RegOpenKeyExA (%s%s)", CalledFrom(), h, a1 );

	int  ret = 0;
	try{
		ret = Real_RegOpenKeyExA (a0,a1,a2,a3,a4);
	}
	catch(...){}

	return ret;
}
int __stdcall My_RegQueryValueExA ( HKEY a0, LPCSTR a1, LPDWORD a2, LPDWORD a3, LPBYTE a4, LPDWORD a5 )
{
	char h[6];
	GetHive(a0,h);

	int  ret = 0;
	try{
		ret = Real_RegQueryValueExA (a0,a1,a2,a3,a4,a5);
	}
	catch(...){}
	
	LogAPI("%x     RegQueryValueExA (%s%s)", CalledFrom(), h, a1 );

	return ret;
}
int __stdcall My_RegSetValueExA ( HKEY a0, LPCSTR a1, DWORD a2, DWORD a3, CONST BYTE* a4, DWORD a5 )
{
	char h[6];
	GetHive(a0,h);

	LogAPI("%x     RegSetValueExA (%s%s)", CalledFrom(), h, a1 );

	int  ret = 0;
	try{
		ret = Real_RegSetValueExA (a0,a1,a2,a3,a4,a5);
	}
	catch(...){}

	return ret;
}

//--------------------------------------------------------------


//_______________________________________________ install hooks fx 

bool InstallHook( void* real, void* hook, int* thunk, char* name, enum hookType ht){
	if( HookFunction((ULONG_PTR) real, (ULONG_PTR)hook, name, ht) ){ 
		*thunk = (int)GetOriginalFunction((ULONG_PTR) hook);
		return true;
	}
	return false;
}

HMODULE hKernelBase = 0;

void DoHook(void* real, void* hook, int* thunk, char* name){

	void *lpReal = 0;
	
	if(hKernelBase != 0){//its Vista+, see if the export exists there if its in both, 
		if(Real_GetProcAddress == NULL){
			lpReal = (void*)GetProcAddress(hKernelBase, name); //k32 is just a forwarder which we cant hook...
		}else{
			lpReal = (void*)Real_GetProcAddress(hKernelBase, name); 
		}
	}
	
	if(lpReal == 0) lpReal = real;

	if(!InstallHook( lpReal, hook, thunk, name, ht_auto ) ){
		LogAPI("Install %s hook failed...\r\nError: %s\r\n", name, GetHookError());
	}

	 
}



//Macro wrapper to build DoHook() call
#define ADDHOOK(name) DoHook( name, My_##name, (int*)&Real_##name, #name );
	
int ConfigHandlerThreadProc(int x){
	
	noSleep = msg("***config:noSleep");
	noGetProc = msg("***config:noGetProc");
	noRegistry = msg("***config:noRegistry");
	queryGetTick = msg("***config:queryGetTick");
	blockOpenProcess = msg("***config:blockOpenProcess");
	blockDebugControl = msg("***config:blockDebugControl");
    ignoreExitProcess = msg("***config:ignoreExitProcess");
    logLevel = msg("***config:hooklibLogLevel");

	if(noSleep) msg("OPTION_SET = noSleep");
	if(noRegistry) msg("OPTION_SET = noRegistry");
	if(noGetProc) msg("OPTION_SET = noGetProc");
	if(queryGetTick) msg("OPTION_SET = queryGetTick");
	if(blockOpenProcess) msg("OPTION_SET = blockOpenProcess");
	if(blockDebugControl) msg("OPTION_SET = blockDebugControl");
	if(ignoreExitProcess) msg("OPTION_SET = ignoreExitProcess");
    if(logLevel > 0) LogAPI("OPTION_SET = hooklibLogLevel = %d", logLevel);
	return 1;

}

void HookEngineDebugMessage(char* msg){
	LogAPI("Debug> %s", msg);
}

void InstallHooks(void)
{

	logLevel = 0;
	debugMsgHandler = HookEngineDebugMessage;

	myPID = GetCurrentProcessId();
	msg("***** Installing Hooks *****");	
	LogAPI("***config:handler:%x", ConfigHandlerThreadProc);
	ConfigHandlerThreadProc(0); //first one we do automatically to stay in sync...

	//DO NOT HOOK GetProcAddress or GetModuleHandle we use them below (not in hook engine)..
	hKernelBase = GetModuleHandle("kernelbase.dll");

	ADDHOOK(CreateFileA);
	ADDHOOK(_lcreat);
	ADDHOOK(_lopen);
	ADDHOOK(CreateProcessA);
	ADDHOOK(WinExec);
	ADDHOOK(ExitProcess);
	ADDHOOK(ExitThread);
	ADDHOOK(CreateRemoteThread);
	ADDHOOK(OpenProcess);
	ADDHOOK(WriteProcessMemory);
	ADDHOOK(accept);
	ADDHOOK(bind);
	ADDHOOK(closesocket);
	ADDHOOK(connect);
	ADDHOOK(gethostbyaddr);
	ADDHOOK(gethostbyname);
	
	ADDHOOK(listen);
	ADDHOOK(recv);
	ADDHOOK(send);
	ADDHOOK(shutdown);
	ADDHOOK(socket);
	ADDHOOK(URLDownloadToFileA);    //todo: in sclog this had to go manual lookup test me..
	ADDHOOK(URLDownloadToCacheFile); // ""
	//ADDHOOK(GetCommandLineA);   //useful for finding end of packer
	ADDHOOK(IsDebuggerPresent);
	
	//ADDHOOK(GetVersionExA);
	//ADDHOOK(GlobalAlloc)
	ADDHOOK(DebugActiveProcess)
	ADDHOOK(GetSystemTime)
	ADDHOOK(CreateMutex)
	ADDHOOK(ReadProcessMemory)
	ADDHOOK(CopyFile)
	ADDHOOK(InternetGetConnectedState)

	//these can add allot of noise 
	if(noRegistry==0){
		ADDHOOK(RegCreateKeyA) 
		ADDHOOK(RegDeleteKeyA) 
		ADDHOOK(RegDeleteValueA) 
		ADDHOOK(RegEnumKeyA) 
		ADDHOOK(RegEnumValueA)
		//ADDHOOK(RegQueryValueA) spamy
		ADDHOOK(RegSetValueA)
		ADDHOOK(RegCreateKeyExA)
		ADDHOOK(RegOpenKeyA)
		ADDHOOK(RegOpenKeyExA)
		//ADDHOOK(RegQueryValueExA) spamy
		ADDHOOK(RegSetValueExA)
	}

	ADDHOOK(Sleep)
	ADDHOOK(GetTickCount)
	ADDHOOK(CloseHandle)


	void* real = GetProcAddress( GetModuleHandleA("ntdll.dll"), "ZwQuerySystemInformation");
	/*if ( !InstallHook( real, My_ZwQuerySystemInformation, Real_ZwQuerySystemInformation) ){ 
		msg("Install hook ZwQuerySystemInformation failed...Error: \r\n");
		ExitProcess(0);
	}

	real = Real_GetProcAddress( Real_GetModuleHandleA("ntdll.dll"), "ZwSystemDebugControl");
	if ( !InstallHook( real, My_ZwSystemDebugControl, Real_ZwSystemDebugControl) ){ 
		msg("Install hook ZwSystemDebugControl failed...Error: \r\n");
		ExitProcess(0);
	}*/

	real = GetProcAddress( GetModuleHandleA("ntdll.dll"), "NtSystemDebugControl");
	if ( !InstallHook( real, My_ZwSystemDebugControl, (int*)&Real_ZwSystemDebugControl,"ZwSystemDebugControl", ht_jmp ) ){ 
		msg("Install hook NtSystemDebugControl failed...Error: \r\n");
	}

	
}
