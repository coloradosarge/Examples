'THIS PROGRAM CREATES THE NECESSARY FILES/FOLDERS FOR OUR PRE SNAPSHOT
'ALSO SAVES APPNAME AND APPVER TO ENVIRONMENT VARIABLES FOR FUTURE USE
'msgbox "This script has been updated for use with Gold Disk v2.  It can take up to 20 minutes to run."

'-------------------------THIS CODE PROMPTS USER FOR INPUT AND SAVES TO VARIABLES-----------------------------------------
Dim AppName 'Name for Application
Dim AppVer  'Version Of Application
Dim objReg  'Object for manipulating registry
Dim boolResult 'True if finds correct key
Dim arrKeyList 'List of Registry Keys
Dim strKey     'Individual Registry Keys
Dim DirStr  'Holds variable of directory for results
Dim Result  'Result of setting registry key
Dim strComputer
AppName = InputBox("What is the Application Name?  Do Not Use Spaces!!!") 'Prompt user for App Name
AppVer = InputBox("What is the Applications Version Number?  Do Not use spaces") 'Prompt user for Ver Number
AppName = Trim(appname)
appver = Trim(appver)

Const HKCU = &H80000001 'HKEY_CURRENT_USER
Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
strComputer = "."

Set objReg = GetObject("winmgmts:{impersonationlevel=impersonate,(Security)}\\" & strComputer & "\root\default:StdRegProv")

'------------------THIS CODE TAKES VARIABLES AND MAKES THEM WINDOWS ENVIRONMENT VARIABLES---------------------------------
boolResult = False
DirStr = "e:\results\" & AppName & "_" & AppVer & "-" & cstr(Month(Date)) & cstr(Year(Date)) 'string for main folder
DirStrPre = "e:\results\" & AppName & "_" & AppVer & "-" & cstr(Month(Date)) & cstr(Year(Date)) & "\Pre" 'string for Pre folder
DirStrPost = "e:\results\" & AppName & "_" & AppVer & "-" & cstr(Month(Date)) & cstr(Year(Date)) & "\Post" 'string for Post folder
objReg.EnumValues HKLM, "Software\CSAT", arrKeyList
If isArray(arrKeyList) Then
   For Each strKey in arrKeyList
      If strKey = "Last_Tested" Then
         boolResult = True
      End If
   Next
End If

If boolResult = False Then
   objReg.CreateKey HKLM, "Software\CSAT"
End If

Result = objReg.SetStringValue(HKLM, "Software\CSAT", "Last_Tested", DirStr)
      
set WshShell = CreateObject("WScript.Shell") 'Create an instance of the wshShell object
'set oEnv=WshShell.Environment("System") 'choose environment variable
'oEnv("StigApp") = AppName 'Set envvironment variable StigApp to the name of user inputted appname
'oEnv("StigVer") = AppVer 'Set envvironment variable StigVer to the name of user inputted Appver

'---------------------------------THIS CODE MAKES FOLDER TO STORE STIG RESULTS-------------------------------------------
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.CreateFolder(Dirstr) 'create app folder
Set objFolder = objFSO.CreateFolder(DirstrPre) 'create pre
Set objFolder = objFSO.CreateFolder(DirstrPost) 'create post

'-----------------------------THIS CODE RUNS AUTORUN AND SAVES TO PRE FOLDER----------------------------------------------
Dim Autorun
Autorun = "E:\SysinternalsSuite\autorun.bat"
if objFSO.FileExists("e:\SysinternalsSuite\autorun.txt") = True then
	objFSO.deletefile("e:\SysinternalsSuite\autorun.txt")
end if
Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run ("cmd.exe /c " & Autorun)
WScript.Sleep(10000)

'-----------------------------THIS CODE RUNS THE DIR AND SAVES TO PRE FOLDER----------------------------------------------
Dim Dircmd
DirCmd = "Dir C:\ /s > " & DirStrPre  & "\dir.txt" 'variable of dos command to rum dir
Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run ("cmd.exe /c " & DirCmd) 'run command prompt flagged auto close with dir command


'-----------------------------THIS CODE SAVES THE DOT NET CONFIGURATION TO A FILE-----------------------------------------
Dim DotNet
DotNet = ("C:\Windows\Microsoft.NET\Framework\v2.0.50727\caspol.exe -all -lp -lg -lf > " & DirStrPre & "\DotNet.txt")
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
objFSO.MoveFile "e:\HijackThis\hijackthis.log", dirstrpre & "\hijackthis.log"

'------------------THIS CODE RUNS GOGOLD.BAT WHICH RUNS GOLD DISK AND SAVES TO PROGRAM FOLDER-------------------------
Set WshShell = WScript.CreateObject("WScript.Shell")
wshShell.Run ("cmd.exe /c cscript E:\Scripts\W7_Script.vbs > " & DirStrPre & "\results.txt"),,TRUE
wshShell.Run ("cmd.exe /c cscript E:\Scripts\Internet_Explorer_8_CSAT_Script.vbs > " & DirStrPre & "\ie_results.txt"),,TRUE

'--------------------------------THIS IS MY TIMER IT WILL BE USED IN A LOOP--------------------------------------------
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
		objFSO.MoveFile "e:\SysinternalsSuite\autorun.txt", dirstrpre & "\autorun.txt"
		objFSO.MoveFile "e:\SysinternalsSuite\listdlls.txt", dirstrpre & "\listdlls.txt"
		objFSO.MoveFile "e:\SysinternalsSuite\loggedonsessions.txt", dirstrpre & "\loggedonsessions.txt"
		objFSO.MoveFile "e:\SysinternalsSuite\pipelist.txt", dirstrpre & "\pipelist.txt"
		objFSO.MoveFile "e:\SysinternalsSuite\psfile.txt", dirstrpre & "\psfile.txt"
		objFSO.MoveFile "e:\SysinternalsSuite\pslist.txt", dirstrpre & "\pslist.txt"
		objFSO.MoveFile "e:\SysinternalsSuite\psloggedon.txt", dirstrpre & "\psloggedon.txt"
		objFSO.MoveFile "e:\SysinternalsSuite\psloglist.txt", dirstrpre & "\psloglist.txt"
		objFSO.MoveFile "e:\SysinternalsSuite\psservice.txt", dirstrpre & "\psservice.txt"
		objFSO.MoveFile "e:\SysinternalsSuite\netstat.txt", dirstrpre & "\netstat.txt"
		done = "yes"
	end if
loop

'--------------------THIS CODE WAITS FOR THE GOLD DISK OUTPUT FILE TO EXIST AND THEN MOVES IT TO PRE------------------
dim done1
done1 = "no"
do until done1 = "yes"
	'wait 1 minute before checking for the file (avoids serious performance issues)
	wscript.sleep (10000)
	if (objfso.FileExists(DirStrPre & "\results.txt")) = True then
		done1 = "yes"
		msgbox "You may now install the software!"
	end if
loop
