;InnoSetupVersion=4.2.6

[Setup]
AppName=SysAnalyzer
AppVerName=SysAnalyzer 1.0
DefaultDirName=c:\iDEFENSE\SysAnalyzer\
DefaultGroupName=SysAnalyzer
OutputBaseFilename=SysAnalyzer_Setup
OutputDir=./

[Files]
Source: ./dependancy\spSubclass.dll; DestDir: {win}; Flags: uninsneveruninstall regserver
Source: ./dependancy\TABCTL32.OCX; DestDir: {win}; Flags: uninsneveruninstall regserver
Source: ./dependancy\vbDevKit.dll; DestDir: {win}; Flags: uninsneveruninstall regserver
Source: ./dependancy\adoKit.dll; DestDir: {win}; Flags: uninsneveruninstall regserver
Source: ./dependancy\mscomctl.ocx; DestDir: {win}; Flags: uninsneveruninstall regserver
Source: ./dependancy\MSWINSCK.OCX; DestDir: {win}; Flags: uninsneveruninstall regserver
Source: ./\source\apilog_dll\injector\Project1.vbw; DestDir: {app}\source\apilog_dll\injector
Source: ./\source\apilog_dll\injector\hook.ico; DestDir: {app}\source\apilog_dll\injector
Source: ./\source\apilog_dll\injector\Form2.frx; DestDir: {app}\source\apilog_dll\injector
Source: ./\source\apilog_dll\injector\Form2.frm; DestDir: {app}\source\apilog_dll\injector
Source: ./\source\apilog_dll\injector\Project1.vbp; DestDir: {app}\source\apilog_dll\injector
Source: ./\source\apilog_dll\dll.dsw; DestDir: {app}\source\apilog_dll
Source: ./\source\apilog_dll\dll.ncb; DestDir: {app}\source\apilog_dll
Source: ./\source\apilog_dll\hooker.h; DestDir: {app}\source\apilog_dll
Source: ./\source\apilog_dll\hooker.lib; DestDir: {app}\source\apilog_dll
Source: ./\source\apilog_dll\main.cpp; DestDir: {app}\source\apilog_dll
Source: ./\source\apilog_dll\ReadMe.txt; DestDir: {app}\source\apilog_dll
Source: ./\source\apilog_dll\main.h; DestDir: {app}\source\apilog_dll
Source: ./\source\apilog_dll\dll.cpp; DestDir: {app}\source\apilog_dll
Source: ./\source\apilog_dll\dll.opt; DestDir: {app}\source\apilog_dll
Source: ./\source\apilog_dll\dll.dsp; DestDir: {app}\source\apilog_dll
Source: ./\source\dirwatch_dll\dir_watch_dll.cpp; DestDir: {app}\source\dirwatch_dll
Source: ./\source\dirwatch_dll\dir_watch_dll.def; DestDir: {app}\source\dirwatch_dll
Source: ./\source\dirwatch_dll\dir_watch_dll.dsp; DestDir: {app}\source\dirwatch_dll
Source: ./\source\dirwatch_dll\dir_watch_dll.dsw; DestDir: {app}\source\dirwatch_dll
Source: ./\source\dirwatch_dll\README.txt; DestDir: {app}\source\dirwatch_dll
Source: ./\source\proc_analyzer\Form1.frm; DestDir: {app}\source\proc_analyzer
Source: ./\source\proc_analyzer\pa.ico; DestDir: {app}\source\proc_analyzer
Source: ./\source\proc_analyzer\Project1.vbp; DestDir: {app}\source\proc_analyzer
Source: ./\source\proc_analyzer\Project1.vbw; DestDir: {app}\source\proc_analyzer
Source: ./\source\proc_analyzer\CDumpFix.cls; DestDir: {app}\source\proc_analyzer
Source: ./\source\proc_analyzer\Module1.bas; DestDir: {app}\source\proc_analyzer
Source: ./\source\proc_analyzer\CStrings.cls; DestDir: {app}\source\proc_analyzer
Source: ./\source\proc_analyzer\CExploitScanner.cls; DestDir: {app}\source\proc_analyzer
Source: ./\source\proc_analyzer\Form1.frx; DestDir: {app}\source\proc_analyzer
Source: ./\source\sysanalyzer\wormy.ico; DestDir: {app}\source\sysanalyzer
Source: ./\source\sysanalyzer\CModule.cls; DestDir: {app}\source\sysanalyzer
Source: ./\source\sysanalyzer\CProcess.cls; DestDir: {app}\source\sysanalyzer
Source: ./\source\sysanalyzer\CProcessInfo.cls; DestDir: {app}\source\sysanalyzer
Source: ./\source\sysanalyzer\CProcessPort.cls; DestDir: {app}\source\sysanalyzer
Source: ./\source\sysanalyzer\image.bmp; DestDir: {app}\source\sysanalyzer
Source: ./\source\sysanalyzer\CRegDiff.cls; DestDir: {app}\source\sysanalyzer
Source: ./\source\sysanalyzer\frmWizard.frm; DestDir: {app}\source\sysanalyzer
Source: ./\source\sysanalyzer\CSysDiff.cls; DestDir: {app}\source\sysanalyzer
Source: ./\source\sysanalyzer\Form1.frm; DestDir: {app}\source\sysanalyzer
Source: ./\source\sysanalyzer\Form1.frx; DestDir: {app}\source\sysanalyzer
Source: ./\source\sysanalyzer\Project1.vbp; DestDir: {app}\source\sysanalyzer
Source: ./\source\sysanalyzer\frmReport.frm; DestDir: {app}\source\sysanalyzer
Source: ./\source\sysanalyzer\FileVer.bas; DestDir: {app}\source\sysanalyzer
Source: ./\source\sysanalyzer\Project1.vbw; DestDir: {app}\source\sysanalyzer
Source: ./\source\sysanalyzer\modDllInject.bas; DestDir: {app}\source\sysanalyzer
Source: ./\source\sysanalyzer\CProcessPorts.cls; DestDir: {app}\source\sysanalyzer
Source: ./\source\sysanalyzer\frmApiLogger.frm; DestDir: {app}\source\sysanalyzer
Source: ./\source\sysanalyzer\frmWizard.frx; DestDir: {app}\source\sysanalyzer
Source: ./\source\sysanalyzer\frmDirWatch.frm; DestDir: {app}\source\sysanalyzer
Source: ./\source\sysanalyzer\CKnownFile.cls; DestDir: {app}\source\sysanalyzer
Source: ./\source\sysanalyzer\frmMarkKnown.frm; DestDir: {app}\source\sysanalyzer
Source: ./\source\sysanalyzer\frmKnownFiles.frm; DestDir: {app}\source\sysanalyzer
Source: ./\source\sysanalyzer\frmKnownFiles.frx; DestDir: {app}\source\sysanalyzer
Source: ./\source\apilog_dll\parse_h\Project1.vbp; DestDir: {app}\source\apilog_dll\parse_h
Source: ./\source\apilog_dll\parse_h\Form1.frm; DestDir: {app}\source\apilog_dll\parse_h
Source: ./\source\apilog_dll\parse_h\parse_h.exe; DestDir: {app}\source\apilog_dll\parse_h
Source: ./\source\apilog_dll\parse_h\Project1.vbw; DestDir: {app}\source\apilog_dll\parse_h
Source: ./\api_log.dll; DestDir: {app}; Flags: ignoreversion
Source: ./\api_logger.exe; DestDir: {app}; Flags: ignoreversion
Source: ./\dir_watch.dll; DestDir: {app}; Flags: ignoreversion
Source: ./\exploit_sigs.txt; DestDir: {app}
Source: ./\proc_analyzer.exe; DestDir: {app}; Flags: ignoreversion
Source: ./\safe_test1.exe; DestDir: {app}
Source: ./\sniff_hit.exe; DestDir: {app}; Flags: ignoreversion
Source: ./\sysAnalyzer.exe; DestDir: {app}; Flags: ignoreversion
Source: ./\SysAnalyzer_help.chm; DestDir: {app}
Source: ./\SysAnalyzer.pdb; DestDir: {app}
Source: ./\known_files.mdb; DestDir: {app}; Flags: uninsneveruninstall
Source: dirwatch_ui.exe; DestDir: {app}; Flags: ignoreversion
Source: source\dirwatch_ui\clsCmnDlg.cls; DestDir: {app}\source\dirwatch_ui
Source: source\dirwatch_ui\dir_watch.dll; DestDir: {app}\source\dirwatch_ui
Source: source\dirwatch_ui\FileVer.bas; DestDir: {app}\source\dirwatch_ui
Source: source\dirwatch_ui\Form1.frm; DestDir: {app}\source\dirwatch_ui
Source: source\dirwatch_ui\Form1.frx; DestDir: {app}\source\dirwatch_ui
Source: source\dirwatch_ui\frmDirWatch.frm; DestDir: {app}\source\dirwatch_ui
Source: source\dirwatch_ui\Project1.vbp; DestDir: {app}\source\dirwatch_ui
Source: source\dirwatch_ui\Project1.vbw; DestDir: {app}\source\dirwatch_ui
Source: source\dirwatch_ui\simple-fso..bas; DestDir: {app}\source\dirwatch_ui
Source: windump.exe; DestDir: {app}
Source: WinPcap_4_1_2.exe; DestDir: {app}
Source: loadlib.exe; DestDir: {app}
Source: source\sysanalyzer\frmMemoryMap.frm; DestDir: {app}\source\sysanalyzer
Source: source\sysanalyzer\frmInjectionScan.frm; DestDir: {app}\source\sysanalyzer
Source: source\sysanalyzer\CMemory.cls; DestDir: {app}\source\sysanalyzer
Source: source\apilog_dll\injector\frmListProcess.frm; DestDir: {app}\source\apilog_dll\injector
Source: ShellExt.exe; DestDir: {app}

[Dirs]
Name: {app}\source
Name: {app}\source\apilog_dll
Name: {app}\source\apilog_dll\parse_h
Name: {app}\source\apilog_dll\injector
Name: {app}\source\dirwatch_dll
Name: {app}\source\proc_analyzer
Name: {app}\source\sysanalyzer
Name: {app}\source\dirwatch_ui

[Icons]
Name: {group}\SysAnalyzer; Filename: {app}\sysAnalyzer.exe
Name: {group}\ProcAnalyzer; Filename: {app}\proc_analyzer.exe
Name: {group}\Source\SysAnalyzer.vbp; Filename: {app}\source\sysanalyzer\Project1.vbp
Name: {group}\Source\Api_Log.dsw; Filename: {app}\source\apilog_dll\dll.dsw
Name: {group}\Source\DirWatch.dsw; Filename: {app}\source\dirwatch_dll\dir_watch_dll.dsw
Name: {group}\Source\ProcAnalyzer.vbp; Filename: {app}\source\proc_analyzer\Project1.vbp
Name: {group}\Source\ApiLogger.vbp; Filename: {app}\source\apilog_dll\injector\Project1.vbp
Name: {group}\Source\DirWatchUI.vbp; Filename: {app}\source\dirwatch_ui\Project1.vbp
Name: {group}\ApiLogger; Filename: {app}\api_logger.exe
Name: {group}\Sniff_Hit; Filename: {app}\sniff_hit.exe
Name: {group}\Uninstall; Filename: {app}\unins000.exe; WorkingDir: {app}
Name: {userdesktop}\SysAnalyzer; Filename: {app}\sysAnalyzer.exe; WorkingDir: {app}; IconIndex: 0
Name: {group}\Test Binary; Filename: {app}\sysAnalyzer.exe; Parameters: safe_test1.exe; WorkingDir: {app}; IconFilename: {app}\safe_test1.exe
Name: {group}\Help File; Filename: {app}\SysAnalyzer_help.chm; WorkingDir: {app}
Name: {userdesktop}\DirWatch; Filename: {app}\dirwatch_ui.exe; IconIndex: 0
Name: {userdesktop}\Sniffhit; Filename: {app}\sniff_hit.exe; IconIndex: 0
Name: {userdesktop}\ApiLogger; Filename: {app}\api_logger.exe; IconIndex: 0

[CustomMessages]
NameAndVersion=%1 version %2
AdditionalIcons=Additional icons:
CreateDesktopIcon=Create a &desktop icon
CreateQuickLaunchIcon=Create a &Quick Launch icon
ProgramOnTheWeb=%1 on the Web
UninstallProgram=Uninstall %1
LaunchProgram=Launch %1
AssocFileExtension=&Associate %1 with the %2 file extension
AssocingFileExtension=Associating %1 with the %2 file extension...
[Run]
Filename: {app}\WinPcap_4_1_2.exe; StatusMsg: Installing WinPcap Packet Sniffer Driver; Flags: postinstall
