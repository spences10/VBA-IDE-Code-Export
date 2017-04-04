AppName={#product}
VersionInfoProductName={#product}
AppVerName={#product} {#version}
AppPublisher={#company}
VersionInfoCompany={#company}
AppCopyright={#yearspan} {#company}
VersionInfoCopyright={#yearspan} {#company}
VersionInfoDescription=Excel addin.
VersionInfoVersion={#longversion}
VersionInfoProductVersion={#longversion}
VersionInfoTextVersion={#version}

; Make this setup program work with 32-bit and 64-bit Windows
ArchitecturesAllowed=x86 x64
ArchitecturesInstallIn64BitMode=x64

; Always write a log file
SetupLogging=true

; Addins do not need a program group and no user-configurable
; installation folder.
DisableProgramGroupPage=true
DisableDirPage=true
CreateAppDir=true
AppendDefaultDirName=false
DisableReadyPage=true

; Allow normal users to install the addin into their profile.
; This directive also ensures that the uninstall information is
; stored in the user profile rather than a system folder (which
; would require administrative rights).
PrivilegesRequired=lowest

; The destination folder is also determined by code since
; different language versions of Excel expect addins in
; localized folders.
DefaultDirName={userappdata}\Microsoft\AddIns\

; The uninstall information is to be put into a subfolder of
; the installation in the user's profile. If this were not set,
; the uninstall information would be put into the Windows system
; folder by InnoSetup.
UninstallFilesDir={code:GetDestDir}\{#product}\uninstall

InternalCompressLevel=max
SolidCompression=true

