'THIS PROGRAM CREATES THE NECESSARY FILES/FOLDERS FOR OUR POST SNAPSHOT
'SAVES FILES IN DIR ENVIRONMENT VARIABLES POINT TO
'RUNS WINDIFF WITH DIRS BEING COMPARED

'-------------------------THIS CODE FINDS THE LAST PRE RUN AND SAVES TO VARIABLES-----------------------------------------
Dim AppName 'Name for Application
Dim AppVer  'Version Of Application
Dim objFSO
Dim strComputer
Dim objReg
dim strOld
set objFSO = CreateObject("Scripting.FileSystemObject")
set WshShell = CreateObject("WScript.Shell") 'Create an instance of the wshShell object

Const HKCU = &H80000001 'HKEY_CURRENT_USER
Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
strComputer = "."

Set objReg = GetObject("winmgmts:{impersonationlevel=impersonate,(Security)}\\" & strComputer & "\root\default:StdRegProv")

objReg.GetStringValue HKLM, "Software\CSAT", "Last_Tested", strOld

'THIS CODE SETS PATH WHERE PRE RUSULTS ARE LOCATED
DirStrPre = strOld & "\Pre"
'THIS CODE SETS THE PATH WHERE TO SAVE POST SNAPSHOTS
DirStrPost = strOld & "\Post"


'-----------------------------THIS CODE RUNS AUTORUN AND SAVES TO PRE FOLDER----------------------------------------------
Dim Autorun
Autorun = "E:\SysinternalsSuite\autorun.bat"
if objFSO.FileExists("e:\SysinternalsSuite\autorun.txt") = True then
	objFSO.deletefile("e:\SysinternalsSuite\autorun.txt")
end if
Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run ("cmd.exe /c " & Autorun)
WScript.Sleep(10000)

'-----------------------------THIS CODE RUNS THE DIR AND SAVES TO POST FOLDER----------------------------------------------
Dim Dircmd
DirCmd = "Dir C:\ /s > " & DirStrPOST  & "\dir.txt" 'variable of dos command to rum dir
Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run ("cmd.exe /c " & DirCmd) 'run command prompt flagged auto close with dir command

'-----------------------------THIS CODE SAVES THE DOT NET CONFIGURATION TO A FILE-----------------------------------------
Dim DotNet
DotNet = ("C:\Windows\Microsoft.NET\Framework\v2.0.50727\caspol.exe -all -lp -lg -lf > " & DirStrPost & "\DotNet.txt")
Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run ("cmd.exe /c " & DotNet)

'-----------------------------THIS CODE RUNS HIJACK THIS-----------------------------------------
Dim HijackThis
Dim KillHijack
Dim KillNotepad
HijackThis = ("E:\HijackThis\HijackThis.exe /autolog")
KillHijack = ("E:\SysinternalsSuite\pskill HijackThis")
KillNotepad = ("E:\SysinternalsSuite\pskill notepad")
Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run ("cmd.exe /c " & HijackThis)
WScript.Sleep(60000)
Set HijackWshShell = WScript.CreateObject("WScript.Shell")
HijackWshShell.Run ("cmd.exe /c " & KillHijack)
Set NotepadWshShell = WScript.CreateObject("WScript.Shell")
NotepadWshShell.Run ("cmd.exe /c " & killNotepad)
objFSO.MoveFile "e:\HijackThis\hijackthis.log", dirstrpost & "\hijackthis.log"

'------------------THIS CODE RUNS GOGOLD.BAT WHICH RUNS GOLD DISK AND SAVES TO PROGRAM FOLDER-------------------------
Set WshShell = WScript.CreateObject("WScript.Shell")
wshShell.Run ("cmd.exe /c cscript E:\Scripts\W7_Script.vbs > " & DirStrPost & "\results.txt"),,TRUE
wshShell.Run ("cmd.exe /c cscript E:\Scripts\Internet_Explorer_8_CSAT_Script.vbs > " & DirStrPost & "\ie_results.txt"),,TRUE

'--------------------------------THIS IS MY TIMER IT WILL BE USED IN A LOOP--------------------------------------------P
dim timer 
timer = (cstr(minute(time)))
Dim EndTime
If timer < 57 then
	Endtime = timer + 3
elseif timer = 57 then
	endtime = 00
elseif timer = 58 then
	endtime = 01
else
	endtime = 02
end if

'--------------------THIS CODE WAITS FOR THE DONE OUTPUT FILE TO EXIST AND THEN RESULT FILES TO PRE------------------
dim done
done = "no"
Set WshShell = WScript.CreateObject("Wscript.Shell")
do until done = "yes"
	'wait 1 minute before checking for the file (avoids serious performance issues)
	wscript.sleep (10000)
	if (objfso.FileExists("e:\SysinternalsSuite\done.txt")) = True then
		objFSO.MoveFile "e:\SysinternalsSuite\autorun.txt", dirstrpost & "\autorun.txt"
		objFSO.MoveFile "e:\SysinternalsSuite\listdlls.txt", dirstrpost & "\listdlls.txt"
		objFSO.MoveFile "e:\SysinternalsSuite\loggedonsessions.txt", dirstrpost & "\loggedonsessions.txt"
		objFSO.MoveFile "e:\SysinternalsSuite\pipelist.txt", dirstrpost & "\pipelist.txt"
		objFSO.MoveFile "e:\SysinternalsSuite\psfile.txt", dirstrpost & "\psfile.txt"
		objFSO.MoveFile "e:\SysinternalsSuite\pslist.txt", dirstrpost & "\pslist.txt"
		objFSO.MoveFile "e:\SysinternalsSuite\psloggedon.txt", dirstrpost & "\psloggedon.txt"
		objFSO.MoveFile "e:\SysinternalsSuite\psloglist.txt", dirstrpost & "\psloglist.txt"
		objFSO.MoveFile "e:\SysinternalsSuite\psservice.txt", dirstrpost & "\psservice.txt"
		objFSO.MoveFile "e:\SysinternalsSuite\netstat.txt", dirstrpost & "\netstat.txt"
		done = "yes"
	end if
loop

'--------------------THIS CODE WAITS FOR THE DONE OUTPUT FILE TO EXIST AND THEN RESULT FILES TO POST------------------
dim done1
done = "no"
do until done1 = "yes"
	'wait 1 minute before checking for the file (avoids serious performance issues)
	wscript.sleep (60000)
	if (objfso.FileExists(DirStrPost & "\results.txt")) = True then
		done1 = "yes"
	end if
loop
'--------------------THIS CODE RUNS WINDIFF COMPARING THE TWO DIRECTORIES------------------
Dim DiffRun
DiffRun = "e:\WINDIFF\windiff " & dirstrpre & " " & dirstrpost 
wshShell.run (diffrun)