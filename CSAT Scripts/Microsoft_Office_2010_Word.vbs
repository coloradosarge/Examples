Option Explicit
On Error Resume Next
Dim strComputer
Dim objReg
Dim arrHive(216)
Dim arrKeyPath(216)
Dim arrKeyPath2(216)
Dim arrSubKey(216)
Dim arrSubKey2(216)
Dim arrValue(216)
Dim arrFinding(216)
Dim arrExpectedValue(216)
Dim arrExpectedValue2(216)
Dim arrType(216)
Dim arrCategory(216)
Dim arrTitle(216)
Dim arrKeyList
DIm strSubkey
Dim intCompare
Dim intCount
Dim boolResult
Dim finding
Dim x
Dim RegEx

Const HKCU = &H80000001 'HKEY_CURRENT_USER
Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
set RegEx = New RegExp

x = 0
arrFinding(x) = "DTOO104"
arrKeyPath(x) = "Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_HTTP_USERNAME_PASSWORD_DISABLE"
arrKeyPath2(x) = "x"
arrSubKey(x) = "winword.exe"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Disable user name and password"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO111"
arrKeyPath(x) = "Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_SAFE_BINDTOOBJECT"
arrKeyPath2(x) = "x"
arrSubKey(x) = "winword.exe"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Enable IE Bind to Object"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO117"
arrKeyPath(x) = "Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_UNC_SAVEDFILECHECK"
arrKeyPath2(x) = "x"
arrSubKey(x) = "winword.exe"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Saved from URL"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO123"
arrKeyPath(x) = "Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_VALIDATE_NAVIGATE_URL"
arrKeyPath2(x) = "x"
arrSubKey(x) = "winword.exe"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Block Navigation to URL from Office"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO129"
arrKeyPath(x) = "Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_WEBOC_POPUPMANAGEMENT"
arrKeyPath2(x) = "x"
arrSubKey(x) = "winword.exe"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Block Pop-Ups"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO131"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\word\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "NoTBPromptUnsignedAddin"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Trust Bar Notifications"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO210"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\word\security\fileblock"
arrKeyPath2(x) = "x"
arrSubKey(x) = "Word12BetaFilesFromConverters"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Block opening of pre-release versions "
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO133"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\word\security\trusted locations"
arrKeyPath2(x) = "x"
arrSubKey(x) = "AllLocationsDisabled"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Disable all trusted locations"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO142"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\word\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "WordBypassEncryptedMacroScan"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Force Scan Encr. Macros in open XML"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO134"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\word\security\trusted locations"
arrKeyPath2(x) = "x"
arrSubKey(x) = "AllowNetworkLocations"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Disable Allow trusted locations"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO139"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\word\options"
arrKeyPath2(x) = "x"
arrSubKey(x) = "DefaultFormat"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = ""
arrExpectedValue2(x) = "x"
arrType(x) = "StringBlank"
arrTitle(x) = "Save files default format"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO146"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\word\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "AccessVBOM"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Disable Trust access to VB Project Macros"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO304"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\word\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "VBAWarnings"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "2"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "VBA Macro Warning settings"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO302"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\word\options"
arrKeyPath2(x) = "x"
arrSubKey(x) = "DontUpdateLinks"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Auto update of links at open"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO303"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\word\options\vpref"
arrKeyPath2(x) = "x"
arrSubKey(x) = "fWarnRevisions_1125_1"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Warn before printing"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO126"
arrKeyPath(x) = "Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_ADDON_MANAGEMENT"
arrKeyPath2(x) = "x"
arrSubKey(x) = "winword.exe"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Add-on Management"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO209"
arrKeyPath(x) = "Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_ZONE_ELEVATION"
arrKeyPath2(x) = "x"
arrSubKey(x) = "winword.exe"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Zone Elevation Protection"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO211"
arrKeyPath(x) = "Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_RESTRICT_ACTIVEXINSTALL"
arrKeyPath2(x) = "x"
arrSubKey(x) = "winword.exe"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Restrict ActiveX Install"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO132"
arrKeyPath(x) = "Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_RESTRICT_FILEDOWNLOAD"
arrKeyPath2(x) = "x"
arrSubKey(x) = "winword.exe"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Restrict File Download"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO124"
arrKeyPath(x) = "Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_WINDOW_RESTRICTIONS"
arrKeyPath2(x) = "x"
arrSubKey(x) = "winword.exe"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Scripted Window Security"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO127"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\word\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "RequireAddinSig"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Add-ins are signed by Trusted Word"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO128"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\Word\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "EnableDEP"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Data Execution Prevention"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO119"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\word\security\filevalidation"
arrKeyPath2(x) = "x"
arrSubKey(x) = "EnableOnLoad"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Turn off file validation"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO110"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\word\security\fileblock"
arrKeyPath2(x) = "x"
arrSubKey(x) = "OpenInProtectedView"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Set default file block behavior"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO121"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\word\security\protectedview"
arrKeyPath2(x) = "x"
arrSubKey(x) = "DisableInternetFilesInPV"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Files from the Internet zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO288"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\word\security\protectedview"
arrKeyPath2(x) = "x"
arrSubKey(x) = "DisableUnsafeLocationsInPV"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Files in unsafe locations"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO292"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\word\security\filevalidation"
arrKeyPath2(x) = "Software\Policies\Microsoft\Office\14.0\word\security\filevalidation"
arrSubKey(x) = "OpenInProtectedView"
arrSubKey2(x) = "DisableEditFromPV"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "1"
arrType(x) = "DualValueEx"
arrTitle(x) = "Set document behavior"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO293"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\word\security\protectedview"
arrKeyPath2(x) = "x"
arrSubKey(x) = "DisableAttachmentsInPV"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Turn off Protected View for attachments"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO305"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\toolbars\word"
arrKeyPath2(x) = "x"
arrSubKey(x) = "NoExtensibilityCustomizationFromDocument"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "UI extending from documents and templates"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO328"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\research\translation"
arrKeyPath2(x) = "x"
arrSubKey(x) = "UseOnline"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Use online translation dictionaries"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO333"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\word\security\fileblock"
arrKeyPath2(x) = "x"
arrSubKey(x) = "Word2Files"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "2"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Word 2 and earlier binary documents"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO334"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\word\security\fileblock"
arrKeyPath2(x) = "x"
arrSubKey(x) = "Word2000Files"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "4"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Word 2000 binary documents and templates"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO336"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\word\security\fileblock"
arrKeyPath2(x) = "x"
arrSubKey(x) = "Word60Files"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "2"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Word 6.0 binary documents and templates"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO337"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\word\security\fileblock"
arrKeyPath2(x) = "x"
arrSubKey(x) = "Word95Files"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "2"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Word 95 binary documents and templates"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO338"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\word\security\fileblock"
arrKeyPath2(x) = "x"
arrSubKey(x) = "Word97Files"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "4"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Word 97 binary documents and templates"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO339"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\word\security\fileblock"
arrKeyPath2(x) = "x"
arrSubKey(x) = "WordXPFiles"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "4"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Word XP binary documents and templates"
arrCategory(x) = "II"

WScript.Echo "Performing " & x & " security checks on Microsoft Office 2010 Word" & vbCrLf

strComputer = "."

Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
For intCount = 0 to x
   If (StrComp(arrType(intCount), "Value", vbTextCompare)) = 0 Then
      objReg.GetDWORDValue arrHive(intCount), arrKeyPath(intCount), arrSubkey(intCount), arrValue(intCount)
      intCompare = StrComp(arrValue(intCount), arrExpectedValue(intCount), vbTextCompare)
      If intCompare = 0 Then
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)  
      Else
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If


   Elseif (StrComp(arrType(intCount), "Exists", vbTextCompare)) = 0 Then
      objReg.EnumValues arrHive(intCount), arrKeyPath(intCount), arrKeyList
      boolResult = false 
      For Each strSubkey In arrKeyList
         If (StrComp(arrSubkey(intCount),strSubkey,vbTextCompare)) = 0 Then
            boolResult = true
            Exit For
         End If
      Next
      If Not boolResult Then
        WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      Else 
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If


   ElseIf (StrComp(arrType(intCount), "ValueNotEx", vbTextCompare)) = 0 Then
      boolResult = false
      objReg.EnumValues arrHive(intCount), arrKeyPath(intCount), arrKeyList
      If(isArray(arrKeyList)) Then
         For Each strSubkey In arrKeyList
            If (StrComp(arrSubkey(intCount),strSubkey,vbTextCompare)) = 0 Then
               boolResult = true
               Exit For
            End If
         Next
      End If
      objReg.GetDWORDValue arrHive(intCount), arrKeyPath(intCount), arrSubkey(intCount), arrValue(intCount)
      If boolResult = false Then
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)  
      ElseIf CDbl(arrValue(intCount)) = CDbl(arrExpectedValue(intCount)) Then
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)  
      Else
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If


   ElseIf (StrComp(arrType(intCount), "DualValueEx", vbTextCompare)) = 0 Then
      finding = "closed"
      boolResult = false
      objReg.EnumValues arrHive(intCount), arrKeyPath(intCount), arrKeyList
      If(isArray(arrKeyList)) Then
         For Each strSubkey In arrKeyList
            If (StrComp(arrSubkey(intCount),strSubkey,vbTextCompare)) = 0 Then
               boolResult = true
               Exit For
            End If
         Next
      End If
      If boolResult = true Then
         boolResult = false
         If(isArray(arrKeyList)) Then
            For Each strSubkey In arrKeyList
               If (StrComp(arrSubkey(intCount),strSubkey,vbTextCompare)) = 0 Then
                  boolResult = true
                  Exit For
               End If
            Next
         End If
      End If
      
      objReg.GetDWORDValue arrHive(intCount), arrKeyPath(intCount), arrSubkey(intCount), arrValue(intCount)
      If CLng(arrvalue(intCount)) = CLng(arrExpectedValue(intCount)) Then
         finding = "closed"
      Else 
         finding = "open"
      End If
      objReg.GetDWORDValue arrHive(intCount), arrKeyPath2(intCount), arrSubkey2(intCount), arrValue(intCount)
      If finding = "closed" Then
         If CLng(arrvalue(intCount)) = CLng(arrExpectedValue2(intCount)) Then
            finding = "closed"
         Else 
            finding = "open"
         End If
      End If
      If boolResult = false Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      ElseIf finding = "closed" Then
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      Else
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If


   ElseIf (StrComp(arrType(intCount), "DualValueOr", vbTextCompare)) = 0 Then
      finding = "closed"
      boolResult = false
      objReg.EnumValues arrHive(intCount), arrKeyPath(intCount), arrKeyList
      If(isArray(arrKeyList)) Then
         For Each strSubkey In arrKeyList
            If (StrComp(arrSubkey(intCount),strSubkey,vbTextCompare)) = 0 Then
               boolResult = true
               Exit For
            End If
         Next
      End If
      If boolResult = true Then
         boolResult = false
         If(isArray(arrKeyList)) Then
            For Each strSubkey In arrKeyList
               If (StrComp(arrSubkey(intCount),strSubkey,vbTextCompare)) = 0 Then
                  boolResult = true
                  Exit For
               End If
            Next
         End If
      End If
      
      objReg.GetStringValue arrHive(intCount), arrKeyPath(intCount), arrSubkey(intCount), arrValue(intCount)
      If StrComp(arrValue(intCount), arrExpectedValue(intCount), vbTextCompare) = 0 Then
         finding = "closed"
      Else 
         finding = "open"
      End If
      objReg.GetDWORDValue arrHive(intCount), arrKeyPath2(intCount), arrSubkey2(intCount), arrValue(intCount)
      If finding = "open" Then
         If CLng(arrvalue(intCount)) = CLng(arrExpectedValue2(intCount)) Then
            finding = "closed"
         Else 
            finding = "open"
         End If
      End If
      If boolResult = false Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      ElseIf finding = "closed" Then
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      Else
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If


   Elseif (StrComp(arrType(intCount), "String", vbTextCompare)) = 0 Then
      finding = "open"
      objReg.GetStringValue arrHive(intCount), arrKeyPath(intCount), arrSubkey(intCount), arrValue(intCount)
      intCompare = StrComp(arrValue(intCount), arrExpectedValue(intCount), vbTextCompare)
      If intCompare = 0 Then finding = "closed" End If
      If finding = "closed" Then
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)  
      Else
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If

   Elseif (StrComp(arrType(intCount), "StringBlank", vbTextCompare)) = 0 Then
      finding = "open"
      boolResult = False
      objReg.EnumValues arrHive(intCount), arrKeyPath(intCount), arrKeyList
      If(isArray(arrKeyList)) Then
         For Each strSubkey In arrKeyList
            If (StrComp(arrSubkey(intCount),strSubkey,vbTextCompare)) = 0 Then
               boolResult = true
               Exit For
            End If
         Next
      End If
      If boolResult = True Then
      RegEx.Pattern = "^\s+$"
         RegEx.IgnoreCase = True
         objReg.GetStringValue arrHive(intCount), arrKeyPath(intCount), arrSubkey(intCount), arrValue(intCount)
         If RegEx.Test(arrValue(intCount)) Then
            finding = "closed"
         End If
      End If
      If finding = "closed" Then
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)  
      Else
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If


   ElseIf (StrComp(arrType(intCount), "StringEx", vbTextCompare)) = 0 Then
      finding = "closed"
      boolResult = false
      objReg.EnumValues arrHive(intCount), arrKeyPath(intCount), arrKeyList
      If(isArray(arrKeyList)) Then
         For Each strSubkey In arrKeyList
            If (StrComp(arrSubkey(intCount),strSubkey,vbTextCompare)) = 0 Then
               boolResult = true
               Exit For
            End If
         Next
      End If
      objReg.GetStringValue arrHive(intCount), arrKeyPath(intCount), arrSubkey(intCount), arrValue(intCount)
      If boolResult = false then
         finding = "open"
      ElseIf StrComp(arrValue(intCount), arrExpectedValue(intCount), vbTextCompare) = 0 Then
         finding = "closed" 
      Else
         finding = "open"
      End If
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      End If

    ElseIf (StrComp(arrType(intCount), "ValueEx", vbTextCompare)) = 0 Then
      boolResult = false
      objReg.EnumValues arrHive(intCount), arrKeyPath(intCount), arrKeyList
      If(isArray(arrKeyList)) Then
         For Each strSubkey In arrKeyList
            If (StrComp(arrSubkey(intCount),strSubkey,vbTextCompare)) = 0 Then
               boolResult = true
               Exit For
            End If
         Next
      End If
      objReg.GetDWORDValue arrHive(intCount), arrKeyPath(intCount), arrSubkey(intCount), arrValue(intCount)
      If boolResult = false Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)   
      ElseIf CDbl(arrValue(intCount)) = CDbl(arrExpectedValue(intCount)) Then
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)  
      Else
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If
   End If
Next
