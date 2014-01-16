#include "windows.h"
#include "stdio.h"
#include <tlhelp32.h> 
#include <conio.h>

void init(STARTUPINFO *si, PROCESS_INFORMATION *pi)
{
	memset(si,0, sizeof(STARTUPINFO));
	memset(pi,0, sizeof(PROCESS_INFORMATION));
	si->cb = sizeof(STARTUPINFO);
}

void main(int argc, char** argv)
{
	STARTUPINFO si;
	PROCESS_INFORMATION pi;

	//printf("Loading api_log.dll\n", LoadLibrary("D:\\_Installs\\iDef\\github\\SysAnalyzer\\api_log.dll"));

	char* exe = "c:\\windows\\notepad.exe";

	init(&si, &pi);
	BOOL ret = CreateProcess(exe,NULL,NULL,NULL,0,0,0,0,&si,&pi);
	printf("Starting notepad by name: %d  pid = %x\n", ret, pi.dwProcessId);

	init(&si, &pi);
	ret = CreateProcess(NULL,exe,NULL,NULL,0,0,0,0,&si,&pi);
	printf("Starting notepad by cmdline: %d  pid = %x\n", ret, pi.dwProcessId);

	int h = (int)ShellExecute(0, "open", exe, NULL, NULL, 1);
	printf("Starting notepad by ShellExecute: %d \n", h);

	getch();
        
}
 
