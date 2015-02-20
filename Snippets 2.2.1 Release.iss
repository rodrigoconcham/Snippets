; InnoScript Version 11.5  Build 2
; Randem Systems, Inc.
; Copyright (c) 2002 - 2014, Randem Systems, Inc.
; Website:  http://www.randemsystems.com
; Support:  http://www.randemsystems.com/support/
; OS: Windows NT 6.1 build 7601 (Service Pack 1)

; Derived from VB VBP Project File

; Designed for Inno Setup Version: 5.5.5 (a)
; Installed Inno Setup Version: 5.5.5 (a)

; Date: enero 12, 2015

; Local Machine Settings. Use these settings as a template for your installation folders

; {app}           : C:\Program Files\Randem Systems\InnoScript
; {appdata}       : C:\Users\rodrigocm\AppData\Roaming\Randem Systems\InnoScript\
; {localappdata}  : C:\Users\rodrigocm\AppData\Local\Randem Systems\InnoScript\
; {cf}            : C:\Program Files\Common Files\Randem Systems
; {tmp}           : C:\Users\rodrigocm\AppData\Local\Temp\
; {commonappdata} : C:\ProgramData\Randem Systems\InnoScript\Release\
; {pf}            : C:\Program Files\

;              VB Runtime Files Folder:   C:\Program Files\Randem Systems\InnoScript 11\VB 6 Redist Files\
;     Visual Basic Project File (.vbp):   C:\Snippets\Codebank.vbp
; Inno Setup Script Output File (.iss):   C:\Snippets\Scripts\Snippets 2.2.1 Release.iss
;:   C:\temp3\Templates\temp3.tpl
;:   C:\Users\rodrigocm\AppData\Local\Randem Systems\InnoScript\Release\Templates\JetWXP.tpl
;:   C:\Users\rodrigocm\AppData\Local\Randem Systems\InnoScript\Release\Templates\JetWVista.tpl
;:   C:\Users\rodrigocm\AppData\Local\Randem Systems\InnoScript\Release\Templates\JetWin7.tpl
;:   C:\Users\rodrigocm\AppData\Local\Randem Systems\InnoScript\Release\Templates\JetWin8.tpl

; ------------------------
;        References
; ------------------------

; OLE Automation - (StdOle2.tlb)
; Microsoft DAO 3.0 Object Library - (DAO350.DLL)


; --------------------------
;        Components
; --------------------------

; Microsoft Common Dialog Control 6.0 (SP6) - (COMDLG32.OCX)
; Microsoft Windows Common Controls 6.0 (SP6) - (mscomctl.ocx)
; ArsBdy - (arsbody.ocx)
; ArsTabB - (arstabb.ocx)


[Setup]
SetupLogging=Yes
AppId=Snippets

;------------------------------------------------------------------------------------------------------------------------
; Taken from VBP/VBG Project File Parameters AppName, AppName AppVersion and Company
;------------------------------------------------------------------------------------------------------------------------

AppName=Snippets
AppVerName=Codebank
AppPublisher=Snippets

;------------------------------------------------------------------------------------------------------------------------

AppVersion=0.0.5
VersionInfoVersion=0.0.5
AllowNoIcons=yes
DefaultGroupName=Snippets\Snippets
DefaultDirName={pf}\Snippets
AppCopyright=Public domain
PrivilegesRequired=Admin
MinVersion=0,5.01
OnlyBelowVersion=0,6.4
Compression=lzma
OutputBaseFilename=SnippetsRelease005

[Tasks]
Name: Desktopicon; Description: {cm:CreateDesktopIcon}; GroupDescription: {cm:AdditionalIcons}
Name: DatabaseSupport; Description: Install Database Support; GroupDescription: Install Database Support:

[Files]
Source: C:\Program Files\Randem Systems\InnoScript 11\VB 6 Redist Files\StdOle2.tlb; DestDir: {sys}; Flags:  regtypelib restartreplace sharedfile uninsneveruninstall; OnlyBelowVersion: 0,6.0;
Source: C:\Program Files\Randem Systems\InnoScript 11\VB 6 Redist Files\MSVBVM60.DLL; DestDir: {sys}; Flags:  regserver restartreplace sharedfile uninsneveruninstall; OnlyBelowVersion: 0,6.0;
Source: C:\Program Files\Randem Systems\InnoScript 11\VB 6 Redist Files\OLEAUT32.DLL; DestDir: {sys}; Flags:  restartreplace sharedfile uninsneveruninstall; OnlyBelowVersion: 0,6.0;
Source: C:\Program Files\Randem Systems\InnoScript 11\VB 6 Redist Files\OLEPRO32.DLL; DestDir: {sys}; Flags:  regserver restartreplace sharedfile uninsneveruninstall; OnlyBelowVersion: 0,6.0;
Source: C:\Program Files\Randem Systems\InnoScript 11\VB 6 Redist Files\ASYCFILT.DLL; DestDir: {sys}; Flags:  restartreplace sharedfile uninsneveruninstall; OnlyBelowVersion: 0,6.0;
Source: C:\Program Files\Randem Systems\InnoScript 11\VB 6 Redist Files\COMCAT.DLL; DestDir: {sys}; Flags:  restartreplace sharedfile uninsneveruninstall; OnlyBelowVersion: 0,6.0;
Source: C:\Program Files\Randem Systems\InnoScript 11\VB 6 Redist Files\VB5DB.DLL; DestDir: {sys}; Flags:  restartreplace sharedfile uninsneveruninstall; 
Source: C:\Program Files\TC UP\PLUGINS\Tools\ColSel\COMDLG32.OCX; DestDir: {sys}; Flags:  restartreplace sharedfile uninsneveruninstall; OnlyBelowVersion: 0,6.0;
Source: C:\Program Files\Common Files\microsoft shared\DAO\DAO350.DLL; DestDir: DestDir: {sys}; Flags:  restartreplace sharedfile uninsneveruninstall; OnlyBelowVersion: 0,6.0;
Source: Z:\Shared code\Codigo Fuente06-08-04\Codigo Fuente\cobranza\instalador\mscomctl.ocx; DestDir: {sys}; Flags:  restartreplace sharedfile uninsneveruninstall; OnlyBelowVersion: 0,6.0; 
;Source: Z:\Shared code\Codigo Fuente06-08-04\Codigo Fuente\cobranza\instalador\ArSBody.ocx; DestDir: {sys}; Flags:  restartreplace sharedfile; 
;Source: Z:\Shared code\Codigo Fuente06-08-04\Codigo Fuente\cobranza\instalador\runtime\ArsTabB.ocx; DestDir: {sys}; Flags:  restartreplace sharedfile; 

;Source:  Z:\Shared code\codigo fuente06-08-04\codigo fuente\cobranza\instalador\msjet35.dll; DestDir: {sys}; MinVersion: 4.0,4.0sp5; Flags:  regserver restartreplace sharedfile
;Source:  Z:\Shared code\codigo fuente06-08-04\codigo fuente\cobranza\instalador\msjter35.dll; DestDir: {sys}; MinVersion: 4.0,4.0sp5; Flags:  restartreplace sharedfile
;Source:  Z:\Shared code\codigo fuente06-08-04\codigo fuente\cobranza\instalador\msjint35.dll; DestDir: {sys}; MinVersion: 4.0,4.0sp5; Flags:  sharedfile


Source: C:\Snippets\Snippets.exe; DestDir: {app}; Flags:  ignoreversion restartreplace; 
Source: C:\Snippets\Codebank.mdb; DestDir: {app}
;
; Download MDAC and JET files from the Download menu of InnoScript or http://www.randemsystems.com/osupdatersupport.html
; * * * Downloaded files MUST be placed in the \Scripts\Output\Support\ folder to be found * * *
;
Source: ..\Scripts\Output\Support\mdac_typ_20.exe; DestDir: {tmp}; Flags: ignoreversion nocompression deleteafterinstall; MinVersion: 4.0,4.0; OnlyBelowVersion: 0,6.0; Tasks: DatabaseSupport
;Source: ..\Scripts\Output\Support\Jet40SP8_WXP.exe; DestDir: {tmp}; Flags: ignoreversion nocompression deleteafterinstall; MinVersion: 0,5.01; OnlyBelowVersion: 0,6.0; Tasks: DatabaseSupport
;
; Download MDAC and JET files from the Download menu of InnoScript or http://www.randemsystems.com/osupdatersupport.html
; * * * Downloaded files MUST be placed in the \Scripts\Output\Support\ folder to be found * * *
;
;
; Download MDAC and JET files from the Download menu of InnoScript or http://www.randemsystems.com/osupdatersupport.html
; * * * Downloaded files MUST be placed in the \Scripts\Output\Support\ folder to be found * * *
;
;
; Download MDAC and JET files from the Download menu of InnoScript or http://www.randemsystems.com/osupdatersupport.html
; * * * Downloaded files MUST be placed in the \Scripts\Output\Support\ folder to be found * * *
;

[INI]
Filename: {app}\Snippets.url; Section: InternetShortcut; Key: URL; String: 

[Icons]
Name: {group}\Codebank ; Filename : {app}\Snippets.exe; WorkingDir: {app};
;Name: {group}\{cm:ProgramOnTheWeb, Codebank }; Filename: {app}\Snippets.url;
Name: {group}\{cm:UninstallProgram, Codebank }; Filename: {uninstallexe};
Name: {commondesktop}\Codebank ; Filename: {app}\Snippets.exe; Tasks: Desktopicon ; WorkingDir: {app};

[Run]
Filename: {tmp}\mdac_typ_20.exe; Parameters: "/Q /C:""setup /QNT"""; WorkingDir: {tmp}; Flags: skipifdoesntexist; MinVersion: 4.0,4.0; OnlyBelowVersion: 0,5.1; Tasks: DatabaseSupport
;Filename: {tmp}\Jet40SP8_WXP.exe; Parameters: /Q; WorkingDir: {tmp}; Flags: skipifdoesntexist; MinVersion: 0,5.01; OnlyBelowVersion: 0,6.0; Tasks: DatabaseSupport
Filename: {app}\Snippets.exe; Description: {cm:LaunchProgram, Codebank }; Flags: nowait postinstall skipifsilent runascurrentuser; WorkingDir: {app}

[UninstallDelete]
Type: files; Name: {app}\Snippets.url
Type: dirifempty; Name: {app}

[Registry]

