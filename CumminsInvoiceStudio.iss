[Setup]
AppId={{B9323FC5-9D5D-43E8-88CF-7A33F2515EF4}
AppName=Cummins Invoice Studio
AppVerName=Cummins Invoice Studio 1.0.8
AppVersion=1.0.8
AppPublisher=Eduard Osipov
AppPublisherURL=mailto:edosipov@gmail.com
AppSupportURL=mailto:edosipov@gmail.com
AppContact=edosipov@gmail.com
VersionInfoVersion=1.0.8.0
VersionInfoProductVersion=1.0.8.0
VersionInfoCompany=Eduard Osipov
VersionInfoDescription=Cummins PDF invoice parser and Excel export tool
VersionInfoProductName=Cummins Invoice Studio
VersionInfoCopyright=Eduard Osipov
DefaultDirName={autopf}\Cummins Invoice Studio
DefaultGroupName=Cummins Invoice Studio
OutputDir=dist
OutputBaseFilename=CumminsInvoiceStudio-Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
SetupIconFile=assets\app_icon.ico
UninstallDisplayIcon={app}\CumminsInvoiceStudio.exe

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "dist\CumminsInvoiceStudio\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Cummins Invoice Studio"; Filename: "{app}\CumminsInvoiceStudio.exe"; IconFilename: "{app}\app_icon.ico"
Name: "{autodesktop}\Cummins Invoice Studio"; Filename: "{app}\CumminsInvoiceStudio.exe"; Tasks: desktopicon; IconFilename: "{app}\app_icon.ico"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional icons:"

[Run]
Filename: "{app}\CumminsInvoiceStudio.exe"; Description: "Launch Cummins Invoice Studio"; Flags: nowait postinstall skipifsilent
