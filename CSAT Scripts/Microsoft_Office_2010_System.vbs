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
arrFinding(x) = "DTOO191"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\Common\Security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "UFIControls"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "Exists"
arrTitle(x) = "ActiveX Control Initialization For Office"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO196"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\security\trusted locations"
arrKeyPath2(x) = "x"
arrSubKey(x) = "Allow User Locations"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Mix of Policy and User Locations"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO181"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\internet"
arrKeyPath2(x) = "x"
arrSubKey(x) = "AllowPNG"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "PNG as Output Format"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO212"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\Common\Blog"
arrKeyPath2(x) = "x"
arrSubKey(x) = "DisableBlog"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Control Blogging"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO200"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\drm"
arrKeyPath2(x) = "x"
arrSubKey(x) = "IncludeHTML"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Allow users to read with browsers"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO177"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\internet"
arrKeyPath2(x) = "x"
arrSubKey(x) = "DisableDownloadCenterAccess"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Disable Updates from Office Online Site"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO186"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\trustcenter"
arrKeyPath2(x) = "x"
arrSubKey(x) = "TrustBar"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Trust Bar Notifications"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO207"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\documentinformationpanel"
arrKeyPath2(x) = "x"
arrSubKey(x) = "Beaconing"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Document Info Beaconing UI"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO184"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common"
arrKeyPath2(x) = "x"
arrSubKey(x) = "QMEnable"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Cust. Experience Improvement Program"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO190"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "DefaultEncryption12"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "Microsoft Enhanced RSA and AES Cryptographic Provider,AES 256,256"
arrExpectedValue2(x) = "x"
arrType(x) = "StringEx"
arrTitle(x) = "Encr. type for Password Protected files"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "VDTOO189"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "OpenXMLEncryption"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "Microsoft Enhanced RSA and AES Cryptographic Provider,AES 256,256"
arrExpectedValue2(x) = "x"
arrType(x) = "StringEx"
arrTitle(x) = "Encr. Type / Pwd. Prot. Open XML files "
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO182"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\ptwatson"
arrKeyPath2(x) = "x"
arrSubKey(x) = "PTWOptIn"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Improve Proofing Tools"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO194"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "DisableHyperLinkWarning"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Disable hyperlink warnings - Office"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO206"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\fixedformat"
arrKeyPath2(x) = "x"
arrSubKey(x) = "DisableFixedFormatDocProperties"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Incl. Doc. properties for PDF and XPS"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO198"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\services\fax"
arrKeyPath2(x) = "x"
arrSubKey(x) = "NoFax"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Disable Internet Fax Feature - Office"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO202"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\drm"
arrKeyPath2(x) = "x"
arrSubKey(x) = "DisablePassportCertification"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Disable Passport Service - Office"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO183"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\general"
arrKeyPath2(x) = "x"
arrSubKey(x) = "ShownFirstRunOptin"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Enable Opt-In Wizard on first run"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO195"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "DisablePasswordUI"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Disable Password to Open UI"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO197"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\Common\Smart Tag"
arrKeyPath2(x) = "x"
arrSubKey(x) = "NeverLoadManifests"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Disable Manifests - Office"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO208"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\portal"
arrKeyPath2(x) = "x"
arrSubKey(x) = "LinkPublishingDisabled"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Office client polling from Office Server"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO201"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\drm"
arrKeyPath2(x) = "x"
arrSubKey(x) = "RequireConnection"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Always req. connection to verify perms"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO185"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common"
arrKeyPath2(x) = "x"
arrSubKey(x) = "UpdateReliabilityData"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Do not receive Automatic small updates"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO193"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\Common\Security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "AutomationSecurity"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "2"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Automation Security"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO203"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\signatures"
arrKeyPath2(x) = "x"
arrSubKey(x) = "XPCompatibleSignatureFormat"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Run ActiveX Controls and Plugins - Restricted Sites Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO192"
arrKeyPath(x) = "Software\Policies\Microsoft\VBA\Security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "LoadControlsInForms"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "Exists"
arrTitle(x) = "Load controls for forms3"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO179"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\internet"
arrKeyPath2(x) = "x"
arrSubKey(x) = "OpenDocumentsReadWriteWhileBrowsing"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Open as Read/Write when browsing"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO199"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\drm"
arrKeyPath2(x) = "x"
arrSubKey(x) = "DisableCreation"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Permissions on managed content"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO178"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\internet"
arrKeyPath2(x) = "x"
arrSubKey(x) = "DisableCustomerSubmittedUpload"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Prevent Upload to Office Online - Office"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO188"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "OpenXMLEncryptProperty"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Allow Meta Refresh - Restricted Sites Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO187"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "DRMEncryptProperty"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Protect metadata / rights managed docs"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO180"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\internet"
arrKeyPath2(x) = "x"
arrSubKey(x) = "RelyOnVML"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Don't Rely on VML / IE graphics -Office"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO204"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\signatures"
arrKeyPath2(x) = "x"
arrSubKey(x) = "SuppressExtSigningSvcs"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "External Signature Services menu"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO306"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\internet"
arrKeyPath2(x) = "x"
arrSubKey(x) = "DisableTemplatesOnTheWeb"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Disable hyperlinks to web templates"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO307"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\officeliveworkspace"
arrKeyPath2(x) = "x"
arrSubKey(x) = "TurnOffOfficeLiveWorkspaceIntegration"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Office Live Workspace Integration"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO311"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\general"
arrKeyPath2(x) = "x"
arrSubKey(x) = "FilterDigitalSignatureCert"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Key Usage Filtering"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO345"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\internet"
arrKeyPath2(x) = "x"
arrSubKey(x) = "UseOnlineContent"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Online content options"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO312"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\internet"
arrKeyPath2(x) = "x"
arrSubKey(x) = "DisableCustomerSubmittedDownload"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Customer-submitted templates downloads "
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO321"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "EncryptDocProps"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Encrypt document properties"
arrCategory(x) = "II"

WScript.Echo "Performing " & x & " security checks on Microsoft Office 2010 System" & vbCrLf

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
