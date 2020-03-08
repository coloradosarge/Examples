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

Const HKCU = &H80000001 'HKEY_CURRENT_USER
Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE

x = 0
arrFinding(x) = "DTOO131"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\infoPath\security"
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
arrFinding(x) = "DTOO133"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\infopath\security\trusted locations"
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
arrFinding(x) = "DTOO157"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\infoPath\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "GradualUpgradeRedirection"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "2"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "SharePoint Services Gradual Upgrade"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO167"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\infoPath\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "EMailFormsRunCodeAndScript"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Forms Opening behavior - EMail /w code"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO176"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\infopath\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "EmailFormsBeaconingUI"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Email forms beaconing UI "
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO169"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\infopath\deployment"
arrKeyPath2(x) = "x"
arrSubKey(x) = "CacheMailXSN"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Disable dynamic caching / form template"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO173"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\infopath\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "EnableFullTrustEmailForms"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "E-Mail forms from Full Trust Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO172"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\infopath\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "EnableInternetEMailForms"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Disable eMail forms - Internet Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO171"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\infopath\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "EnableRestrictedEMailForms"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Disable eMail forms in Restricted"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO159"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\infopath\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "RunFullTrustSolutions"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Fully trusted solutions access"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO158"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\infopath\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "AllowInternetSolutions"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Disable Open Solutions / Internet Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO168"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\infopath\deployment"
arrKeyPath2(x) = "x"
arrSubKey(x) = "MailXSNwithXML"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Disable sending template w/ eMail form"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO170"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\infopath"
arrKeyPath2(x) = "x"
arrSubKey(x) = "DisableInfoPath2003EmailForms"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "2003 forms as email"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO164"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\infopath\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "InfoPathBeaconingUI"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Beaconing UI / forms opening"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO165"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\infopath\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "EditorActiveXBeaconingUI"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Beaconing UI /forms opened Activex"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO156"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\infopath\editor\offline"
arrKeyPath2(x) = "x"
arrSubKey(x) = "CachedModeStatus"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "2"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Offline Mode - InfoPath"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO160"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\infopath\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "DisallowAttachmentCustomization"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Prevent Unsafe File Att. - InfoPath"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO127"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\infopath\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "RequireAddinSig"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Add-ins are signed by Trusted Publisher"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO128"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\infopath\security"
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
arrFinding(x) = "DTOO294"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\infopath\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "EnableIntranetEMailForms"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "E-mail forms from the Intranet"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO295"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\options\mail"
arrKeyPath2(x) = "x"
arrSubKey(x) = "DisableInfopathForms"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "InfoPath e-mail forms in Outlook"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO296"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\infopath\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "RunManagedCodeFromInternet"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Managed code from the Internet"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO297"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\infopath\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "SignatureWarning"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "A form is digitally signed"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO305"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\toolbars\infopath"
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
arrFinding(x) = "DTOO309"
arrKeyPath(x) = "Software\Policies\Microsoft\office\14.0\infopath\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "APTCA_AllowList"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "APTCA Assembly Allow List Enforcement"
arrCategory(x) = "II"


WScript.Echo "Performing " & x & " security checks on Microsoft Office 2010 InfoPath" & vbCrLf

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
