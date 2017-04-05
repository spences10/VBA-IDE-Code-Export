#define VERSION "0.3.0"              ; The version number
#define LONGVERSION "0.3.0"          ; The version in four-number format
#define YEARSPAN "2016-2017"         ; The year(s) of publication
                                     ; (e.g., 2014-2015)
#define PRODUCT "VBA IDE Code Export"
#define COMPANY "Contributors of VBA IDE Code Export"

#define SOURCEDIR "src"              ; The folder with the addin files
                                     ; (relative path)

#define LOGFILE "INST-LOG.TXT"       ; The name of the log file.

AppPublisherURL=https://github.com/spences10/VBA-IDE-Code-Export
AppSupportURL=https://github.com/spences10/VBA-IDE-Code-Export/issues
AppUpdatesURL=https://github.com/spences10/VBA-IDE-Code-Export/releases
OutputBaseFilename=CodeExport_setup_{#version}
OutputDir=deploy

; If you want to display a license file, uncomment the following
; line and put in the correct file name.
LicenseFile=LICENSE

; Icons and banners
; The install icon and banner do not need to be included
; in the setup package; they are only required during compilation
; of the InnoSetup script.
; SetupIconFile=\images\icon.ico
; WizardImageFile=\images\installbanner.bmp
; WizardSmallImageFile=..\images\icon.bmp
; WizardImageStretch=false
; WizardImageBackColor=clWhite

; The uninstall icon must be included in the setup package
; and placed in the output folder.
; UninstallDisplayIcon={app}\icon.png

; Specific AppID
; Use InnoSetup's Generate UID command from the Tools menu
; to generate a unique ID. Make sure to have this ID start
; with TWO curly brackets.
; Never change this UID after the addin has been deployed.
AppId={{4ED72357-D953-4156-9F26-1CA15F055161}}

; vim: set ts=2 sts=2 sw=2 noet tw=60 fo+=lj cms=;%s
