; ExcelAddinInstaller
; InnoSetup script to install and activate native Excel addins.
; Originally developed for Daniel's XL Toolbox (xltoolbox.sf.net).
; Requires the InnoSetup Preprocessor (ISPP).
; Copyright (C) 2013  Daniel Kraus <http://github.com/bovender>
; 
; This program is free software: you can redistribute it and/or modify
; it under the terms of the GNU General Public License as published by
; the Free Software Foundation, either version 3 of the License, or
; (at your option) any later version.
; 
; This program is distributed in the hope that it will be useful,
; but WITHOUT ANY WARRANTY; without even the implied warranty of
; MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.	See the
; GNU General Public License for more details.
; 
; You should have received a copy of the GNU General Public License
; along with this program.	If not, see <http://www.gnu.org/licenses/>.

[Setup]
#ifexist "config.iss"        
  #include "config.iss"
#endif
#include "inc/setup.iss"

[Files]
; The include file makes adds all .XLA and .XLAM files contained in the
; SOURCEDIR to the project.
#include "inc/files-addins.iss"

; Define any additional files in the custom files.iss file.
#ifexist "files.iss"
  #include "files.iss"
#endif

[Tasks]
; We make it optional for users to have the addin activated for use in
; Excel. In most cases, this will be left enabled by users (everything
; else does not make sense).
Name: ActivateAddin; Description: {cm:taskActivate}; 

; Define any additional tasks in the custom tasks.iss file.
#ifexist "tasks.iss"
  #include "tasks.iss"
#endif

[Code]
#include "inc/code.iss"

[Languages]
Name: English; MessagesFile: compiler:Default.isl; 
Name: Deutsch; MessagesFile: compiler:Languages\German.isl; 
Name: Dansk; MessagesFile: compiler:Languages\Danish.isl; 
#ifexist "languages.iss"
  #include "languages.iss"
#endif

[CustomMessages]
#include "inc/messages.iss"

; Define any additional messages in the custom messages.iss file.
#ifexist "messages.iss"
  #include "messages.iss"
#endif

; vim: set ts=2 sts=2 sw=2 noet tw=60 fo+=lj cms=;%s 
