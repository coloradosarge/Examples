'THIS PROGRAM IS USED TO PERFORM A PRE OF SOFTWARE IMMEDIATLY AFTER A POST HAS BEEN DONE

Dim AppName 'Name for Application
Dim AppVer  'Version Of Application
Dim strComputer
Dim objReg
Dim strOld

Const HKCU = &H80000001 'HKEY_CURRENT_USER
Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
strComputer = "."

Set objReg = GetObject("winmgmts:{impersonationlevel=impersonate,(Security)}\\" & strComputer & "\root\default:StdRegProv")

objReg.GetStringValue HKLM, "Software\CSAT", "Last_Tested", strOld

'THIS CODE WILL CAPTURE ENV VARIABLES AND SAVE THEM LOCALLY
set WshShell = CreateObject("WScript.Shell") 'Create an instance of the wshShell object

'THIS CODE PROMPTS USER FOR INPUT AND SAVES TO VARIABLES
Dim NewAppName 'Name for Application
Dim NewAppVer  'Version Of Application
NewAppName = InputBox("What is the Application Name?  NO SPACES!") 'Prompt user for App Name
NewAppVer = InputBox("What is the Applications Version Number?  NO SPACES!") 'Prompt user for Ver Number
newappver = trim(newappver)
newappname = trim(newappname)

'THIS CODE TAKES VARIABLES AND MAKES THEM WINDOWS ENVIRONMENT VARIABLES
boolResult = False
DirStr = "e:\results\" & NewAppName & "_" & NewAppVer & "-" & cstr(Month(Date)) & cstr(Year(Date)) 'string for main folder
DirStrPre = "e:\results\" & NewAppName & "_" & NewAppVer & "-" & cstr(Month(Date)) & cstr(Year(Date)) & "\Pre" 'string for Pre folder
DirStrPost = "e:\results\" & NewAppName & "_" & NewAppVer & "-" & cstr(Month(Date)) & cstr(Year(Date)) & "\Post" 'string for Post folder
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

'THIS CODE MAKES FOLDER TO STORE STIG RESULTS
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.CreateFolder(Dirstr) 'create app folder
Set objFolder = objFSO.CreateFolder(DirstrPre) 'create pre
Set objFolder = objFSO.CreateFolder(DirstrPost) 'create post

strOld = strOld & "\Post"

'THIS CODE TAKES OLD POST AND COPIES TO NEW PRE
Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run ("cmd.exe /c copy " & strOld & "\*.* " & dirstrpre) 

Msgbox "You may now install the software"
