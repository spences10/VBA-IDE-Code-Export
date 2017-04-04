{
		GLOBAL CONSTANTS
}
const 
	{
		To be able to automatically remove addins from Excel''s list of
		activated addins during uninstall, we keep a central repository
		of addins in the registry. If Your Company deploys only one addin
		(determined by the AppID), this repository will of course contain
		only this one addin.
	}
	AddinsRegistry = 'Software\{#SetupSetting("AppPublisher")}\{#SetupSetting("AppID")}';

	{
		MinExcelVersion and MaxExcelVersion provide boundaries for
		the minimum and maximum versions of Excel that an add-in
		can be installed for. The setup code will loop through
		the Office registry entries for these versions.
		The installer checks if any of these Excel versions
		exist; if not, the addin may still be extracted from the
		setup package, but no activation will take place. The
		user will be informed about this.
	}
	MinExcelVersion = 9;   { Excel 2000 }
	MaxExcelVersion = 16;  { Excel 2016 }

	{
		Maximum number of addins to check. This is a safety
		variable that serves to prevent infinite loops.
	}
	MaxAddins = 255;

	{
		Windows API constants
	}
	WM_CLOSE = $10;
	MAX_PATH = 250;

{
	GLOBAL VARIABLES
}
var
	ExcelNotInstalled: boolean;
	ExcelExePath: string;
	OkToCopyLog: boolean;

	{
		The destination folder is determined at run-time
		depending on the Excel locale. It is cached in this
		variable by the GetDestDir() function.
	}
	DestDir: string;
						
	{
		If a localized Addins path could not be found in the
		registry, the addin needs to be registered with Excel
		using the full path. This global variable is set by the
		GetDestDir() function.	
	}
	RegisterWithFullPath: boolean;


{
	Looks up the localized Addins folder in the registry.
	This function is used in the [Files] section of the
	script, where function calls always expect a parametrized
	function. This function does not require a parameter.
	Therefore, a dummy parameter is defined.
}
function GetDestDir(dummyparameter: string): string;
var
	Addins: string;
  s: string;
	CallName: string;
	i: integer;
begin
	if DestDir = '' then
	begin
		CallName := 'GetDestDir(' + dummyparameter + '): ';
		log(CallName+'Trying to find addins folder name');

		{
			Note the trailing backslash
		}
		DestDir := ExpandConstant('{userappdata}\Microsoft\');

		{
			Loop through possible version numbers of Excel and find out if
			any installed version uses an addin folder other than "addins".
			This can be the case with international versions (other than English).
			If an addin folder name that is different from "Addins" is found,
			the addin will be installed into a dedicated folder.
		}
    for i := 8 to 32 do
    begin
      s := '';
      if RegQueryStringValue(HKEY_CURRENT_USER, 'Software\Microsoft\Office\'
        +IntToStr(i)+'.0\Common\General', 'AddIns', s) then
			begin
        if Length(Addins) > 0 then
        begin
					{
						If the Addins variable has been set already and we encounter
          	a different name, reset everything in order to use a dedicated
          	name further on.
					}
          if s <> Addins then
          begin
						{
							Set the Addins variable to a zero-length string to force
							using a dedicated dir name later.
						}
  				  log(CallName+'Found alternative Addins key for version '+IntToStr(i)+': "'+s+'"');
            Addins := '';
						{
							Once a single dir name that is different from "Addins" was
							found, we can exit the loop.
						}
    				break;
          end 
        end
        else
        begin
					{
						Addins variable has zero length: Set it to the current value of s
					}
				  log(CallName+'Found first Addins key: version '+IntToStr(i)+', "'+s+'"');
          Addins := s;
        end
			end
    end;
    
		{
			Check if the Addins variable contains something now; if not, use
			a default value ('XL Toolbox')
		}
    if Addins = '' then
		begin
      DestDir := ExpandConstant('{userappdata}\Microsoft\Addins\');
			RegisterWithFullPath := true;
			log(CallName+'Using dedicated folder: "'+DestDir+'"');  
		end
		else
		begin
      DestDir := ExpandConstant('{userappdata}\Microsoft\' + Addins);
			RegisterWithFullPath := false;
			log(CallName+'Installing to default Addins folder: ' + DestDir);  
		end;
  end;
	result := DestDir; 
end;

{
	Helper function to convert boolean values into a string
	(used for logging).
}
function BoolToStr(b: boolean): string;
begin
	if b then result := 'TRUE' else result := 'FALSE'
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
	if CurStep = ssDone then OkToCopyLog := True;
end;


{
	This function is called by InnoSetup.
	It determines whether the "Tasks" page of the Wizard
	shall be displayed or not. If Excel is not installed
	on the system, this page will be skipped (returns TRUE).
}
function ShouldSkipPage(PageID: Integer): Boolean;
begin
	Result := (PageID = wpSelectTasks) and ExcelNotInstalled;
end;


{
	This function checks if any version of Excel from "MinVer"
	to "MaxVer" is installed by looking up an Excel-specific
	registry key. It is used in conjunction with a "Check"
	parameter in the [Files] section, and also by the
	InitializeWizard() function below. This can be used to
	install certain files only for specific versions of Excel.
	See Wikipedia for a list of Excel versions.
}
function IsExcelVersionInstalled(MinVer: integer; MaxVer: integer): boolean;
var
	xl: integer;
begin
	for xl := MinVer to MaxVer do
		if RegKeyExists(HKEY_CURRENT_USER,
		'Software\Microsoft\Office\'+IntToStr(xl)+'.0\Excel\Options') then
		begin
			Result := true;
		end;
	Log('IsExcelVersionInstalled('+IntToStr(MinVer)+', '+
		IntToStr(MaxVer)+') returns '+BoolToStr(result));
end;

{
===========================================================
== Activating/deactivating addins
===========================================================
}

{
	Returns a numbered value in the form expected for Excel
	addins in the Windows registry.	For example, the list of
	addins is stored as "OPEN", "OPEN1", "OPEN2" etc. 
}
function GetNumberedValue(KeyName: string; Num: integer): string;
begin
	if Num = 0 then
	begin
		Result := KeyName;
	end
	else
	begin
		Result := KeyName+IntToStr(Num);
	end
end;

{
	Activates an addin for use in Excel by adding an "OPEN"
	value with the addin's name to the Excel registry key.
	The MinVer and MaxVer parameters indicate the
	Excel versions for which an addin shall be installed.
}
procedure ActivateAddin(MinVer: integer; MaxVer: integer);
var
	xl, i: integer;
	Key, Value: string;
	AddinName: string;
	CallName: String;
begin
	{
		Only proceed if the task 'Activate addin' is selected.
	}
	if not IsTaskSelected('ActivateAddin') then Exit;

	{
		Activated Excel addins are usually stored in the registry
		with their file name only (without path). If the standard
		addins folder could not be determined by the GetDestDir()
		function, the entire path is used instead.
		The GestDestDir() function will have set the global
		variable RegisterWithFullPath accordingly.
	}
	if RegisterWithFullPath then
	begin
		AddinName := '"' + CurrentFileName + '"';
	end
	else
	begin
		AddinName := '"' + ExtractFileName(CurrentFileName) + '"';
	end;

	{
		Build the string for the log file.
	}
	CallName := 'RegisterAddin('+IntToStr(MinVer)+', '+
		IntToStr(MaxVer)+'): ';
	Log(CallName+'Processing ' + AddInName);

	{
		Create a registry key for each Excel version that this addin
		is to be activated for.
	}
	xl := MinVer;
	repeat
		{
			This registry key scheme is used by Excel 2000 and
			newer. Excel 97 uses a different scheme; but since
			Excel 97 is really ancient, we ignore it here.
		}
		Key := 'Software\Microsoft\Office\'+IntToStr(xl)+'.0\Excel\Options';

		{
			If the registry key exists, assume that the Excel version
			is installed.
		}
		if RegKeyExists(HKEY_CURRENT_USER, Key) then
		begin
			{
				Find the next available OPEN value
			}
			i := 0;
			repeat
				{
					Check if a registry value OPEN[i] exists (OPEN, OPEN1,
					OPEN2, ...).  If it does not exist, we can write the addin
					name into "OPEN[i+1]" and finish.	If it does exist, check
					if it contains our addin. If so, finish. If not, increase i
					and try again.
				}
				if RegQueryStringValue(HKEY_CURRENT_USER, Key, 
					GetNumberedValue('OPEN', i), Value) then
				begin
					{
						Check if the value happens to match the current
						addin's name.
					}
					if uppercase(Value) = uppercase(AddinName) then
					begin
						{
							The addin is activated already, so we can quit.
						}
							break;
					end
					else
					begin
						{
							Increase the iterator variable to test the next
							OPEN[i] value.
						}
						i := i+1
					end
				end
				else
				begin
					{
						The OPEN[i] value does not exist yet:
						Write to the log file that we are modifying the
						registry now.
					}
					Log(CallName+'Writing registry: '+Key+' ==> '+IntToStr(i));
					RegWriteStringValue(HKEY_CURRENT_USER, Key, 
						GetNumberedValue('OPEN', i), AddinName);
					break;
				end;
			until i > MaxAddins
		end;

		{
			Move on to the next Excel version.
		}
		xl := xl+1;
	until xl > MaxVer;

	{
		Now, save the addin name in a repository in order to be able
		to remove it from Excel's list of active addins during
		uninstall. This loop is very similar to the one above.
	}
	i := 0;
	repeat
		if RegQueryStringValue(HKEY_CURRENT_USER, AddinsRegistry, 
			GetNumberedValue('Addin', i), Value) then
		begin
			{
				A numbered value for the current value of i
				exists already.
			}
			if uppercase(Value) = uppercase(AddinName) then
			begin
				{
					The addin is already contained in this repository;
					break out of the loop.
				}
				break;
			end
			else
			begin
				{
					Check the next numbered value
				}
				i := i+1;
			end
		end
		else
		begin
			{
				Numbered value Addin[i] does not exist yet:
				Store the addin in the repository.
				Also store the Excel versions for which it was activated.
			}
			Log(CallName+'Writing registry: '+Key+' ==> '+IntToStr(i));
			RegWriteStringValue(HKEY_CURRENT_USER, AddinsRegistry,
				GetNumberedValue('Addin', i), AddinName);
			RegWriteStringValue(HKEY_CURRENT_USER, AddinsRegistry,
				GetNumberedValue('Addin', i)+'_MinVer', IntToStr(MinVer));
			RegWriteStringValue(HKEY_CURRENT_USER, AddinsRegistry,
				GetNumberedValue('Addin', i)+'_MaxVer', IntToStr(MaxVer));
			break;
		end;
	until i > MaxAddins;
end;

{
	This function is called by the DeinitializeUninstall()
	procedure. It loops through the Excel registry keys for the
	Excel versions indicated by MinVer and MaxVer and removes
	occurrences of "AddinName".	If an occurrence is found, the
	corresponding "OPENx" value is deleted, and all values
	"OPENy", "OPENz" etc.	are renamed as "OPENx", "OPENy" and
	so on to preserve consecutive numbering of "OPEN" values.
}
procedure DeactivateAddin(AddinName: string; 
	MinVer: integer; MaxVer: integer);
var
	xl, i, j: integer;
	Key, Value: string;
begin
	{
		Make sure the addin name is enclosed in double quotes.
	}
	if AddinName[1] <> '"' then AddinName := '"'+AddinName+'"';

	{
		Loop through Excel versions.
	}
	xl := MinVer;
	repeat
		Key := 'Software\Microsoft\Office\'+IntToStr(xl)+'.0\Excel\Options';

		{
			Check if the Excel version denoted by xl exists.
		}
		if RegKeyExists(HKEY_CURRENT_USER, Key) then
		begin
			i := 0;
			repeat
				{
					Check if an OPEN[i] key exists that holds our addin's name.
				}
				if RegQueryStringValue(HKEY_CURRENT_USER, Key, 
					GetNumberedValue('OPEN', i), Value) then
				begin
					if uppercase(Value) = uppercase(AddinName) then
					begin
						{
							Addin was found: remove the registry key.
						}
						RegDeleteValue(HKEY_CURRENT_USER, Key, 
						GetNumberedValue('OPEN', i));

						{
							Adjust the numbering of any remaining OPEN keys.
						}
						j := i+1;
						repeat
						if RegQueryStringValue(HKEY_CURRENT_USER, Key,
							GetNumberedValue('OPEN', j), Value) then
						begin
							RegWriteStringValue(HKEY_CURRENT_USER, Key, 
							GetNumberedValue('OPEN', j-1), Value)
						end
						else
						begin
							{
								If the jth OPEN value does not exist, we
								need to delete the previous one, but only if
								j is bigger than i+1
							}
							if j > i+1 then
							begin
								RegDeleteValue(HKEY_CURRENT_USER, Key, 
									GetNumberedValue('OPEN', j-1));
							end;
							break;
						end;

						{
							Move to the next numbered OPEN[i] key whose
							number needs to be adjusted.
						}
						j := j+1;
						until j > MaxAddins;
					end;

					{
						Move to the next numbered OPEN[i] key.
					}
					i := i+1;
				end
				else
				begin
					{
						No more values, so let's finish.
					}
					break;
				end;
			until i > MaxAddins;
		end;

		{
			Move on to the next Excel version.
		}
		xl := xl+1;
	until xl > MaxVer;
end;


{
	This procedure is called by the InnoSetup uninstaller once
	the uninstall process is completed. Each addin that is
	found in the addin repository that was created during
	installation ("AddinsRegistry" constant) is deactivated
	and unregistered by removing it from Excel's registry key.
}
procedure DeinitializeUninstall();
var
	i: integer;
	Value: string;
	VerInfo: string;
	MinVer, MaxVer: integer;
begin
	i := 0;
	repeat
		{
			Is there an "AddIn[i]" value stored in our personal registry key?
		}
		if RegQueryStringValue(HKEY_CURRENT_USER, AddinsRegistry, 
			GetNumberedValue('Addin', i), Value) then
		begin
			{
				Obtain the minimum and maximum Excel version numbers
				for which this addin was registered.
			}
			if RegQueryStringValue(HKEY_CURRENT_USER, AddinsRegistry,
				GetNumberedValue('Addin', i)+'_MinVer', VerInfo) then
			begin
				MinVer := StrToInt(VerInfo);
			end
			else
			begin
				{
					Fallback if the registry value is not found.
				}
				MinVer := 9;
			end;
			if RegQueryStringValue(HKEY_CURRENT_USER, AddinsRegistry,
				GetNumberedValue('Addin', i)+'_MaxVer', VerInfo) then
			begin
				MaxVer := StrToInt(VerInfo);
			end
			else
			begin
				{
					Fallback if the registry value is not found.
				}
				MaxVer := 20; 
			end;

			DeactivateAddin(Value, MinVer, MaxVer);

			{
				Remove the addin from the repository.
			}
			RegDeleteValue(HKEY_CURRENT_USER, AddinsRegistry, GetNumberedValue('Addin', i));
			RegDeleteValue(HKEY_CURRENT_USER, AddinsRegistry, GetNumberedValue('Addin', i)+'_MinVer');
			RegDeleteValue(HKEY_CURRENT_USER, AddinsRegistry, GetNumberedValue('Addin', i)+'_MaxVer')
		end
		else
		begin
			{
				If there is no Addin[i] value in the repository for
				the current iteration, do not expect there to be any
				other values with higher numbers, and break out of
				the loop.
			}
			break;
		end;

		{
			Move to the next Addin[i] value in the repository.
		}
		i := i+1;
	until i > MaxAddins;

	{
		If the repository is now empty, remove the registry key
		in order to leave no trace of the installation on the
		system.
	}
	RegDeleteKeyIfEmpty(HKEY_CURRENT_USER, AddinsRegistry);
end;



{
	This function is used in the [Files] section to determine
	which of the add-ins (.XLA for 2000-2003 or .XLAM for
	2007-2013) shall be copied to the target system. If no
	Excel is installed so far, return TRUE to copy both;
	otherwise, if Excel is installed, copy the one that is
	appropriate for the installed Excel version.
}
function ShouldInstallFile(MinVer: integer; MaxVer: integer): boolean;
begin
	if ExcelNotInstalled then
	begin
		result := true;
	end
	else
	begin
		result := IsExcelVersionInstalled(MinVer, MaxVer);
	end
end;


{
	Determines if a 64-bit version of Excel is installed.
	Note that there may be a 32-bit version of Excel running
	on a 64-bit version of Windows.
}
function Excel64(): boolean;
var
	bitness: string;
	callname: string;
begin
	callname := 'Excel64: ';
	if RegQueryStringValue(HKEY_LOCAL_MACHINE,
		'Software\Wow6432Node\Microsoft\Office\14.0\Outlook',
		'Bitness', bitness) then
	begin
		Log(callname+'Wow6432 registry key: Bitness = ' + bitness);
		if bitness = 'x64' then Result := true;
	end
	else
	begin
		Log(callname+'Wow6432 registry key not present');
	end;
end;


{ =========================================================== }
{ == Win32 API
{ =========================================================== }

#include "win32.iss"


{ =========================================================== }
{ == Pre install
{ =========================================================== }

function GetUninstallRegKeyStr: string;
var
	AppID: string;
begin
{
	The following code requires the InnoSetup Preprocessor (ISPP).
	NB: The double curly brackets at the beginning of the AppID string
	are not properly processed by ISPP: They remain as-is.
	However, in the registry, the AppID has just a single leading
	curly bracket, so we need to remove the first one.
}
	AppID := Copy('{#SetupSetting("AppID")}', 2, 255);
	result :=
		'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\'+AppId+'_is1';
end;

{
	This function checks the command-line parameters for the
	/UPDATE switch which indicates that setup was started by
	an update procedure of the Excel addin. If the switch is
	present, Excel is shut down silently, rather than
	displaying a message to the user. When the installation is
	complete, Excel is started again. It is up to the addin to
	use or not use this switch and implement appropriate
	behavior before and after the installation process.
}
function InitializeSetup(): boolean;
var
	i: LongInt;
	hWnd: LongInt;
	IsUpdate: boolean;
	bCancel: boolean;
	CallName: string;
begin
	CallName:='InitializeSetup(): ';
	Log(CallName+'Examining command line...');

	for i := 1 to ParamCount do
	begin
		if uppercase(ParamStr(i)) = '/UPDATE' then
		begin
			Log(CallName+'UPDATE switch found');
			hwnd := FindWindowByClassName('XLMAIN');
			if hWnd<>0 then
			begin
				{
					Before we attempt to shut down Excel, we fetch its
					executable path, so that we can restart it later.
				}
				ExcelExePath := GetProcessExePath(Hwnd);
				Log(CallName+'Sending WM_CLOSE to XLMAIN...');
				SendMessage(hWnd, WM_CLOSE, 0, 0); { WM_CLOSE: $10 }

				{
					After sending the WM_CLOSE message, we need to give
					Excel a second to completely shut down and remove the
					Mutex; otherwise, the Setup program will abort if
					started with /SP- /SILENT 
					A delay of 500 ms is sufficient
				}
				Sleep(500);
				IsUpdate := true;
			end
		end
	end;

{
	If this setup program is *not* run as an update, we need
	to check if Excel is running and if so, tell the user that
	we are going to attempt to close it.
}

	if IsUpdate then
	begin
		Log(CallName+'Running as UPDATE --> exiting with Result := TRUE');
		{
			Returning True indicates that InnoSetup may continue.
		}
		Result := true
	end
	else
	begin
		hWnd := FindWindowByClassName('XLMAIN');

		{
			If Excel is running, hWnd is different from 0.
		}
		while (hWnd <> 0) and not bCancel do
		begin
			{
				Inform the user that Excel needs to be closed.
			}
			if SuppressibleMsgBox(CustomMessage('ExcelIsRunning'), 
				mbConfirmation, MB_OKCANCEL, IDOK) <> IDOK then
			begin
				Log(CallName+'Excel running - user aborted setup.');
				bCancel := true;
			end
			else
			begin
				Log(CallName+'Excel running - attempting to close...');

				{
					Use the Win32 api to shut down Excel.
				}
				ExcelExePath := GetProcessExePath(Hwnd);
				SendMessage(hWnd, WM_CLOSE, 0, 0); { WM_CLOSE: $10 }
				Sleep(200);

				{
					Check if Excel has been shut down successfully.
				}
				hWnd := FindWindowByClassName('XLMAIN');
			end;
		end;

		{
			If Excel is no longer running, hWnd will be 0 at this
			point, so we can return True to indicate to InnoSetup
			that it may proceed.
		}
		Result := (hWnd = 0);
	end;
end;


{
	This function is called by InnoSetup during the
	initialization step of the setup process. It checks if
	Excel is installed on the system at all.
}
procedure InitializeWizard();
var
	CallName: string;
begin
	CallName := 'InitializeWizard(): ';
	Log(CallName);

	if not IsExcelVersionInstalled(MinExcelVersion, MaxExcelVersion) then
	begin
		Log(CallName+'Excel not installed!');
		CreateOutputMsgPage(wpWelcome, CustomMessage('ExcelNotInstalled'),
		CustomMessage('ExcelNotInstalledCaption'),
		CustomMessage('ExcelNotInstalledExplanation'));
		ExcelNotInstalled := true;
	end;
end;


{
	===========================================================
	== Post install
	===========================================================
}

{
	If this setup program is operating in "Update" mode (with
	the /UPDATE switch), the InitializeSetup() function placed
	the path of the executable Excel image file in the global
	variable ExcelExePath. We can use it to automatically
	restart Excel.
}
procedure DeinitializeSetup();
var
	ExitCode: Integer;
	CallName: String;
begin
	CallName := 'DeinitializeSetup(): ';

	{
		Restart Excel if it has been running before.
	}
	if Length(ExcelExePath) > 0 then
	begin
		Log(CallName+'Restarting Excel...');
		Log(CallName+'--> '+ExcelExePath);
		Exec(ExcelExePath, '', '', SW_SHOW, ewNoWait, ExitCode);
	end
	else
	begin
		Log(CallName+'Excel was not running when setup was started.');
	end;

	{
		Copy the log file to the installation
	}
	if OkToCopyLog then
		FileCopy(
			ExpandConstant('{log}'),
			AddBackslash(GetDestDir(''))+'{#product}\{#logfile}',
			false);
	{
		The following line requires administrator privileges
		during setup. It is therefore commented out until a way
		is found to delete the log file after savely without
		administrator privileges. When uncommenting, make sure
		to replace the '>>' with closing curly braces.
	}
	{ RestartReplace(ExpandConstant('{log>>'), ''); }
end;

{ vim: set ft=pascal ts=2 sts=2 sw=2 noet tw=60 fo+=lj  :}
