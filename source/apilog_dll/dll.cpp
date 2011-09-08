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
#include <Winsock2.h>
#include <stdio.h>
#include <string.h>
#include <stdlib.h>
#include <stdarg.h>
#include <wininet.h>

void InstallHooks(void);

extern "C" void __setargv(void);

#include "hooker.h"
#include "main.h"   //contains a bunch of library functions in it too..



bool Installed =false;

void Closing(void){ msg("***** Injected Process Terminated *****"); }
	

BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  ul_reason_for_call, 
                       LPVOID lpReserved
					 )
{

    if(!Installed){
		 Installed=true;
		 InstallHooks();
		 atexit(Closing);

		 WSADATA WsaDat;			 
  		 WSAStartup(MAKEWORD(1,1), &WsaDat);
	}

	return TRUE;
}




//___________________________________________________hook implementations _________


HANDLE __stdcall My_CreateFileA(LPCSTR a0,DWORD a1,DWORD a2,LPSECURITY_ATTRIBUTES a3,DWORD a4,DWORD a5,HANDLE a6)
{
	
	char *calledFrom=0;

	LogAPI("%x     CreateFileA(%s)", CalledFrom(), a0);

    HANDLE ret = 0;
    try{
        ret = Real_CreateFileA(a0, a1, a2, a3, a4, a5, a6);
    }
	catch(...){} 
	
	if(a0 && strstr(a0,"NTICE") > 0 ){ //to many gaobots = this		
		
		_asm {
			mov eax, [ebp+4]    ;//return addr on stack
			sub eax, 900h       ;//vmware code in 900h buffer before ret addr
   			mov calledFrom, eax
		} 

		Seek_n_Destroy_AntiVmWare(calledFrom, 0x900);
		return (HANDLE)-1;

	}
	else{
		return ret;
	}

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
	LogAPI("%x     accept(%x,%x,%x)", CalledFrom(), a0, a1, a2);

    SOCKET ret = 0;
    try {
        ret = Real_accept(a0, a1, a2);
    }
	catch(...){	} 

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
	
	LogAPI("%x     connect( %s:%d )", CalledFrom(), ip, htons(a1->sin_port) );
	
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
    
	LogAPI("%x     gethostbyaddr(%x)", CalledFrom(), a0);

    hostent* ret = 0;
    try {
        ret = Real_gethostbyaddr(a0, a1, a2);
    }
	catch(...){	} 

    return ret;
}

hostent* __stdcall My_gethostbyname(char* a0)
{
	LogAPI("%x     gethostbyname(%x)", CalledFrom(), a0);

    hostent* ret = 0;
    try {
        ret = Real_gethostbyname(a0);
    }
	catch(...){	} 

    return ret;
}

int __stdcall My_gethostname(char* a0,int a1)
{
	LogAPI("%x     gethostname(%x)", CalledFrom(), a0);

    int ret = 0;
    try {
        ret = Real_gethostname(a0, a1);
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
    LogAPI("%x     recv(h=%x)", CalledFrom(), a0);

    int ret = 0;
    try {
        ret = Real_recv(a0, a1, a2, a3);

		if(ret>0){
			//hexdump((unsigned char*)a1,ret);
		}

    } 
	catch(...){	} 

    return ret;
}

int __stdcall My_send(SOCKET a0,char* a1,int a2,int a3)
{
    
	LogAPI("%x     send(h=%x)", CalledFrom(), a0);
    int ret = 0;

    try {

		//if(a2>0 && *a1 !=0)	//hexdump((unsigned char*)a1,a2);
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
	
	LogAPI("%x     socket(family=%x,type=%x,proto=%x)", CalledFrom(), a0, a1, a2);

    SOCKET ret = 0;
    try {
        ret = Real_socket(a0, a1, a2);
    }
	catch(...){	} 

    return ret;
}

SOCKET __stdcall My_WSASocketA(int a0,int a1,int a2,struct _WSAPROTOCOL_INFOA* a3,GROUP a4,DWORD a5)
{
    
	LogAPI("%x     WSASocketA(fam=%x,typ=%x,proto=%x)", CalledFrom(), a0, a1, a2);

    SOCKET ret = 0;
    try {
        ret = Real_WSASocketA(a0, a1, a2, a3, a4, a5);
    }
	catch(...){	} 

    return ret;
}



//untested
int My_URLDownloadToFileA(int a0,char* a1, char* a2, DWORD a3, int a4)
{
	
	LogAPI("%x     URLDownloadToFile(%s)", CalledFrom(), a1);

    SOCKET ret = 0;
    try {
        ret = Real_URLDownloadToFileA(a0, a1, a2, a3, a4);
    }
	catch(...){	} 

    return ret;
}

//untested
int My_URLDownloadToCacheFile(int a0,char* a1, char* a2, DWORD a3, DWORD a4, int a5)
{
	
	LogAPI("%x     URLDownloadToCacheFile(%s)", CalledFrom(), a1);

    SOCKET ret = 0;
    try {
        ret = Real_URLDownloadToCacheFile(a0, a1, a2, a3, a4, a5);
    }
	catch(...){	} 

    return ret;
}

void __stdcall My_ExitProcess(UINT a0)
{
    
	LogAPI("%x     ExitProcess()", CalledFrom());

    try {
        Real_ExitProcess(a0);
    }
	catch(...){	} 

}

void __stdcall My_ExitThread(DWORD a0)
{
    
	LogAPI("%x     ExitThread()", CalledFrom());

    try {
        Real_ExitThread(a0);
    }
	catch(...){	} 

}

FILE* __stdcall My_fopen(const char* a0, const char* a1)
{

	LogAPI("%x     fopen(%s)", CalledFrom(), a0);

	FILE* rt=0;
    try {
        rt = Real_fopen(a0,a1);
    }
	catch(...){	} 

	return rt;
}

size_t __stdcall My_fwrite(const void* a0, size_t a1, size_t a2, FILE* a3)
{

	LogAPI("%x     fwrite(h=%x)", CalledFrom(), a3);

	size_t rt=0;
    try {
        rt = Real_fwrite(a0,a1,a2,a3);
    }
	catch(...){	} 

	return rt;
}

HANDLE __stdcall My_OpenProcess(DWORD a0,BOOL a1,DWORD a2)
{
	LogAPI("%x     OpenProcess(pid=%ld)", CalledFrom(), a2);

    HANDLE ret = 0;
    try {
        ret = Real_OpenProcess(a0, a1, a2);
    }
	catch(...){	} 

    return ret;
}

HMODULE __stdcall My_GetModuleHandleA(LPCSTR a0)
{
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

	unsigned long (__stdcall  *lpfnLoadLib)(void *);

	int buflen, ret ; 
	unsigned long writeLen, hThread;
	HANDLE hProcess, lpdllPath;
   
	char dllPath[MAX_PATH]; // = "api_log.dll\x00";

	LogAPI("%x     CreateProcessA(%s,%s,%x,%s)", CalledFrom(), a0, a1, a6, a7);

    BOOL retv = 0;
    try {
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
            
		Real_GetProcAddress(Real_GetModuleHandleA("kernel32.dll"), "LoadLibraryA");
		
		_asm mov lpfnLoadLib, eax

		LogAPI("*****   LoadLibraryA=%x",lpfnLoadLib);
    
		ret = (int)Real_CreateRemoteThread(hProcess, 0, 0, lpfnLoadLib, lpdllPath, 0, &hThread);
		LogAPI("*****   CreateRemoteThread=%x" , ret);
            
	    ResumeThread(pi->hThread);

    }
	catch(...){	} 

    return retv;



}

int My_system(const char* cmd)
{
    
	
	LogAPI("%x     system(%s)",  CalledFrom(), cmd);

	int ret=0;
	try {
        ret = Real_system(cmd);
    }
	catch(...){	} 

    return ret;

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

	LogAPI("%x     WriteProcessMemory(h=%x,len=%x)", CalledFrom(), a0, a3);

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
	
	LogAPI("%x     GetProcAddress(%s)", CalledFrom(), a1);

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
	LogAPI("%x     IsDebuggerPresent()", CalledFrom() );

	BOOL  ret = 0;
	try{
		ret = Real_IsDebuggerPresent();
	}
	catch(...){}
	
	return ret;
}

void My___setargv(void){

	LogAPI("%x     __setargv()", CalledFrom() );

	
	try{
		Real___setargv();
	}
	catch(...){}
	
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

	LogAPI("%x     CreateMutex(%s)", CalledFrom(), a2 );

	HANDLE ret = 0;
	try{
		ret = Real_CreateMutex(a0,a1,a2);
	}
	catch(...){}
	
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

void DoHook(void* real, void* hook, void* thunk, char* name){

	char err[400];

	if ( !InstallHook( real, hook, thunk) ){ //try to install the real hook here
		sprintf(err,"***** Install %s hook failed...Error: %s", name, &lastError);
		msg(err);
	} 

}


//Macro wrapper to build DoHook() call
#define ADDHOOK(name) DoHook( name, My_##name, Real_##name, #name );	


void InstallHooks(void)
{

	msg("***** Installing Hooks *****");	
 
	ADDHOOK(LoadLibraryA); 
	ADDHOOK(WriteFile);
	ADDHOOK(CreateFileA);
	ADDHOOK(WriteFileEx);
	ADDHOOK(_lcreat);
	ADDHOOK(_lopen);
	ADDHOOK(_lread);
	ADDHOOK(_lwrite);
	ADDHOOK(CreateProcessA);
	ADDHOOK(WinExec);
	ADDHOOK(ExitProcess);
	ADDHOOK(ExitThread);
	ADDHOOK(GetProcAddress);     //logging disabled hook proc (spam)
	ADDHOOK(WaitForSingleObject);
	ADDHOOK(CreateRemoteThread);
	ADDHOOK(OpenProcess);
	ADDHOOK(WriteProcessMemory);
	ADDHOOK(GetModuleHandleA);
	ADDHOOK(accept);
	ADDHOOK(bind);
	ADDHOOK(closesocket);
	ADDHOOK(connect);
	ADDHOOK(gethostbyaddr);
	ADDHOOK(gethostbyname);
	ADDHOOK(gethostname);
	ADDHOOK(listen);
	ADDHOOK(recv);
	ADDHOOK(send);
	ADDHOOK(shutdown);
	ADDHOOK(socket);
	ADDHOOK(WSASocketA);
	ADDHOOK(system);
	ADDHOOK(fopen);
	ADDHOOK(fwrite);
	ADDHOOK(URLDownloadToFileA);
	ADDHOOK(URLDownloadToCacheFile);
	ADDHOOK(GetCommandLineA);   //useful for finding end of packer
	ADDHOOK(IsDebuggerPresent);
	ADDHOOK(__setargv);
	ADDHOOK(GetVersionExA);
	ADDHOOK(GlobalAlloc)
	ADDHOOK(GetCurrentProcessId)
	ADDHOOK(DebugActiveProcess)
	ADDHOOK(ReadFile)
	ADDHOOK(GetSystemTime)
	ADDHOOK(CreateMutex)
	ADDHOOK(ReadProcessMemory)
	//ADDHOOK(GetVersion)
	ADDHOOK(CopyFile)
	ADDHOOK(InternetGetConnectedState)

	//these can add allot of noise 
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




