[Setup]
AppName=YaraWorkBench
AppVerName=YaraWorkBench
DefaultDirName=c:\YaraWorkBench
DefaultGroupName=YaraWorkBench
UninstallDisplayIcon={app}\unins000.exe
OutputDir=./
OutputBaseFilename=YaraWorkBench_Setup

[Files]
Source: yara_workbench\dependancies\vbDevKit.dll; DestDir: {win};                   Flags: regserver ignoreversion
Source: D:\_code\vbYara\yara_workbench\dependancies\vbCapstone.dll; DestDir: {app}; Flags: ignoreversion
Source: D:\_code\vbYara\yara_workbench\dependancies\UTypes.dll; DestDir: {app};     Flags: ignoreversion
Source: D:\_code\vbYara\yara_workbench\dependancies\vbUtypes.dll; DestDir: {app};   Flags: regserver    ignoreversion
Source: D:\_code\vbYara\yara_workbench\dependancies\capstoneAX.dll; DestDir: {app}; Flags: regserver    ignoreversion
Source: D:\_code\vbYara\yara_workbench\dependancies\sppe3.dll; DestDir: {app};      Flags: regserver    ignoreversion
Source: D:\_code\vbYara\yara_workbench\dependancies\scivb2.ocx; DestDir: {app};     Flags: regserver    ignoreversion
Source: D:\_code\vbYara\yara_workbench\dependancies\hexed.ocx; DestDir: {app};      Flags: regserver    ignoreversion
Source: D:\_code\vbYara\yara_workbench\dependancies\SciLexer.dll; DestDir: {app};   Flags: ignoreversion
Source: D:\_code\vbYara\yara_workbench\dependancies\MSCOMCTL.OCX; DestDir: {sys};   Flags: regserver onlyifdoesntexist sharedfile uninsneveruninstall
Source: D:\_code\vbYara\yara_workbench\dependancies\capstone.dll; DestDir: {app};   Flags: ignoreversion
Source: D:\_code\vbYara\yara_workbench\ywb.exe; DestDir: {app};                     Flags: ignoreversion
Source: D:\_code\vbYara\libyara.dll; DestDir: {app};                                Flags: ignoreversion
Source: D:\_code\vbYara\yhelp.dll; DestDir: {app};                                  Flags: ignoreversion
Source: yara_workbench\Credits.txt; DestDir: {app};
Source: D:\_code\vbYara\yara_workbench\test.yar; DestDir: {app}
Source: D:\_code\vbYara\yara_workbench\dependancies\java.hilighter; DestDir: {app}
Source: yara_workbench\intellisense\cuckoo.txt; DestDir: {app}\intellisense\
Source: yara_workbench\intellisense\dotnet.txt; DestDir: {app}\intellisense\
Source: yara_workbench\intellisense\elf.txt; DestDir: {app}\intellisense\
Source: yara_workbench\intellisense\hash.txt; DestDir: {app}\intellisense\
Source: yara_workbench\intellisense\magic.txt; DestDir: {app}\intellisense\
Source: yara_workbench\intellisense\math.txt; DestDir: {app}\intellisense\
Source: yara_workbench\intellisense\pe.txt; DestDir: {app}\intellisense\
Source: yara_workbench\intellisense\time.txt; DestDir: {app}\intellisense\
Source: yara_workbench\library\isPe.txt; DestDir: {app}\library\
Source: yara_workbench\library\sect_entropy.txt; DestDir: {app}\library\
Source: yara_workbench\alt_vers\yara3.5.exe; DestDir: {app}\alt_vers\

[Icons]
Name: {group}\Yara WorkBench; Filename: {app}\ywb.exe
Name: {group}\Readme.txt; Filename: {app}\Readme.txt
Name: {group}\Uninstall; Filename: {app}\unins000.exe
Name: {userdesktop}\Yara WorkBench; Filename: {app}\ywb.exe; IconIndex: 0

[Dirs]
Name: {app}\intellisense
Name: {app}\library
Name: {app}\alt_vers
