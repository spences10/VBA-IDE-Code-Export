{
	===========================================================
	== Win32 API Calls																				 
	===========================================================
}

function GetProcessID(hProcess: LongInt): LongInt;
external 'GetProcessId@kernel32.dll stdcall delayload setuponly';

function GetWindowThreadProcessId(hWnd: LongInt; var lpdwProcessId: LongInt): LongInt;
external 'GetWindowThreadProcessId@user32.dll stdcall delayload setuponly';

function OpenProcess(dwDesiredAccess: LongInt; bInheritHandle: LongInt; dwProcessId: LongInt): LongInt;
external 'OpenProcess@kernel32.dll stdcall delayload setuponly';

function CloseHandle(hObject: LongInt): LongInt;
external 'CloseHandle@kernel32.dll stdcall delayload setuponly';

function GetProcessImageFileName(hProcess: longint; lpImageFileName: string; nSize: LongInt): LongInt;
external 'GetProcessImageFileNameA@psapi.dll stdcall delayload setuponly';

function GetLogicalDriveStrings(nBufferLength: LongInt; lpBuffer: string): LongInt;
external 'GetLogicalDriveStringsA@kernel32.dll stdcall delayload setuponly';

function QueryDosDevice(lpDeviceName: string; lpTargetPath: string; ucchMax: LongInt): LongInt;
external 'QueryDosDeviceA@kernel32.dll stdcall delayload setuponly';



{ =========================================================== }
{ == Initializing
{ =========================================================== }

{
	Identifies the process that owns hWnd and returns the
	executable path that belongs to it. We need this to be
	able to close Excel and automatically restart it after
	setup has completed.
}
function GetProcessExePath(hWnd: LongInt): string;
var
	ProcID: LongInt;
	hProc: LongInt;
	FileName: string;
	FileNameLen: LongInt;
	StrLen: longInt;
	Drives: string;
	DriveName, DeviceName: string;
	iNull: integer;
	CallName: string;
begin
	CallName := 'GetProcessExePath('+IntToStr(hWnd)+'): ';

	{ Identify the process that owns the hWnd Window }
	GetWindowThreadProcessId(hWnd, ProcID);

	if ProcID <> 0 then
	begin
		Log(CallName+'Found process ID');

		{ Get a handle for the process }
		hProc := OpenProcess($400, 0, ProcID);

		if hProc <> 0 then
		begin
			Log(CallName+'Obtained process handle');
			FileNameLen := MAX_PATH;
			FileName := StringOfChar(#0, FileNameLen);
			StrLen := GetProcessImageFileName(hProc, FileName, FileNameLen)

			if StrLen <> 0 then
			begin
				FileName := Copy(FileName, 1, Pos(#0, FileName)-1);

				{ The FileName that we got has an MS-DOS device name in it,
				which we need to resolve now. }
				{ First we obtain a list of all available drives; then
				we iterate through the list, obtain the MS-DOS device name
				for each of the drives and check whether the MS-DOS device
				name occurs in our FileName. If yes, we replace it with the
				drive letter. }

				Drives := StringOfChar(#0, MAX_PATH);
				StrLen := GetLogicalDriveStrings(MAX_PATH, Drives);
				Drives := Copy(Drives, 1, StrLen);

				while Length(Drives)>0 do
				begin
					{ Extract a NULL-terminated substring }
					iNull := Pos(#0, Drives);
					if iNull = 0 then iNull := MAX_PATH;
					DriveName := copy(Drives, 1, iNull-2);

					{ Convert the "C:\" style into a device path }
					DeviceName := StringOfChar(#0, MAX_PATH);
					StrLen := QueryDosDevice(DriveName, DeviceName, MAX_PATH);
					DeviceName := Copy(DeviceName, 1, Pos(#0, DeviceName)-1);

					{ Check if we have found "our" device path;
						if so, replace it with a "C:\" style path }
					if Pos(DeviceName, FileName) = 1 then
					begin
						StringChangeEx(FileName, DeviceName, DriveName, false);
						Log(CallName+'Path: '+FileName);
						Result := FileName;
						Exit { the while loop }
					end;
					Delete(Drives, 1, iNull);
				end;
			end;
			Log(CallName+'Closing handle');
			CloseHandle(hProc);
		end
		else
		begin
			Log(CallName+'*** No handle ***');
		end
	end
	else
	begin
		Log(CallName+'*** Process ID not found ***');
	end;
end;

{ vim: set ft=pascal ts=2 sts=2 sw=2 noet tw=60 fo+=lj :}
