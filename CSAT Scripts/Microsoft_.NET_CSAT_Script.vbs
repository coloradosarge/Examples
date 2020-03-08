Dim objEnterpriseXmlDoc
Dim objMachineXmlDoc
Dim objUserXmlDoc
Dim colElemList
Dim objPermissionSetElement
Dim strPermissionSetName
Dim objShell
Dim strUserProfile
Dim arrEnterpriseDefaultCodeGroups(1)
Dim arrUserDefaultCodeGroups(1)
Dim arrMachineDefaultCodeGroups(12)
Dim arrDefaultPermissionSets(7)
Dim arrGoodMembership(3)
Dim arrIsolatedStoragePerms(3)
Dim arrFinding(100)
Dim arrPermissionSet(100)
Dim arrCategory(100)
Dim arrTitle(100)
Dim arrType(100)
Dim arrAttribute(100)
Dim arrArraySize(100)
Dim arrArraySize2(100)
Dim arrPermArray()
Dim arrChildArray()
Dim arrEntCheckPermSet()
Dim arrMacCheckPermSet()
Dim arrUsrCheckPermSet()
Dim i
Dim j
Dim k
Dim objElement
Dim objIPermission
Dim finding
Dim strMessage
Dim x
Dim intCount
Dim blnAggFinding
Dim RegEx
Dim objReg
Dim arrKeyList
Dim strComputer
Dim strSubkey
Dim objFSO
Dim objFolder
Dim objFolder2
Dim NetFolders()
Dim netVersion()
Dim entBool()
Dim entFile()
Dim macBool()
Dim macFile()
Dim usrBool()
Dim usrFile()
Dim objSubFolder
Dim objSubFolder2
Dim versionCount
Dim currentDirectory
Dim wshShell

Set objShell = CreateObject("Wscript.Shell")
strUserProfile = objShell.ExpandEnvironmentStrings("%USERPROFILE%")

strComputer = "."

Set RegEx = New RegExp

Set objReg = GetObject("winmgmts:{impersonationlevel=impersonate,(Security)}\\" & strComputer & "\root\default:StdRegProv")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder("C:\Windows\Microsoft.NET\Framework")
If objFSO.FolderExists(strUserProfile & "\Application Data\Microsoft\CLR Security Config") Then
   Set objFolder2 = objFSO.GetFolder(strUserProfile & "\Application Data\Microsoft\CLR Security Config")
Else
   Set objFolder2 = objFSO.GetFolder(left(WScript.ScriptFullName,(Len(WScript.ScriptFullName))-(len(WScript.ScriptName))))
End If

Const HKCU = &H80000001 'HKEY_CURRENT_USER
Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
Const HKU = &H80000003  'HKEY_USERS

arrEnterpriseDefaultCodeGroups(0) = "All_Code"
arrUserDefaultCodeGroups(0) = "All_Code"
arrMachineDefaultCodeGroups(0) = "All_Code"
arrMachineDefaultCodeGroups(1) = "My_Computer_Zone"
arrMachineDefaultCodeGroups(2) = "Microsoft_Strong_Name"
arrMachineDefaultCodeGroups(3) = "ECMA_Strong_Name"
arrMachineDefaultCodeGroups(4) = "LocalIntranet_Zone"
arrMachineDefaultCodeGroups(5) = "Intranet_Same_Site_Access"
arrMachineDefaultCodeGroups(6) = "Intranet_Same_Directory_Access"
arrMachineDefaultCodeGroups(7) = "Internet_Zone"
arrMachineDefaultCodeGroups(8) = "Internet_Same_Site_Access"
arrMachineDefaultCodeGroups(9) = "Restricted_Zone"
arrMachineDefaultCodeGroups(10) = "Trusted_Zone"
arrMachineDefaultCodeGroups(11) = "Trusted_Same_Site_Access"

arrDefaultPermissionSets(0) = "FullTrust"
arrDefaultPermissionSets(1) = "SkipVerification"
arrDefaultPermissionSets(2) = "Execution"
arrDefaultPermissionSets(3) = "Nothing"
arrDefaultPermissionSets(4) = "LocalIntranet"'
arrDefaultPermissionSets(5) = "Internet"
arrDefaultPermissionSets(6) = "Everything"

arrGoodMembership(0) = "StrongNameMembershipCondition"
arrGoodMembership(1) = "PublisherMembershipCondition"
arrGoodMembership(2) = "HashMembershipCondition"

Redim Preserve arrPermsArray(100, 4)
Redim Preserve arrChildArray(100,4)


Sub setFiles(intCheck)
   Set objEnterpriseXmlDoc = CreateObject("Msxml2.DOMDocument.6.0")
   Set objMachineXmlDoc = CreateObject("Msxml2.DOMDocument.6.0")
   Set objUserXmlDoc = CreateObject("MsXml2.DOMDocument.6.0")
   if entBool(intCheck) = true Then
      objEnterpriseXmlDoc.load(entFile(intCheck))
      if objEnterpriseXmlDoc.parseError.errorCode <> 0 Then
         WScript.Echo "Error Loading Enterprise Configuration"
         WScript.Quit(1)
      End If
   End If
   If macBool(intCheck) = true Then
      objMachineXmlDoc.load(macFile(intCheck))
      if objMachineXmlDoc.parseError.errorCode <> 0 Then
         WScript.Echo "Error Loading Machine Configuration"
         WScript.Quit(1)
      End If
   End If
   If usrBool(intCheck) = true Then
      objUserXmlDoc.load(usrFile(intCheck))
      if objUserXmlDoc.parseError.errorCode <> 0 Then
         WScript.Echo "Error Loading User Configuration"
         WScript.Echo strUserProfile
         WScript.Quit(1)
      End If
   End If
End Sub

Sub ComparePermissionSets(xmlDoc, arrTargetArray, arrDefaultCodeGroups, strPermissionType, strLevel)
    Dim strPermSet
    Dim strMembership
    Dim objMembershipElement
    Dim blnDefaultCodeGroup
    Dim strCodeGroupName
    finding = false
    Set ColElemList = xmlDoc.getElementsByTagName("CodeGroup")
    For Each objElement in colElemList
       blnDefaultCodeGroup = false
       strCodeGroupName = objElement.getAttribute("Name")
       strPermSet = objElement.getAttribute("PermissionSetName")
       For Each objMembershipElement in objElement.ChildNodes
          If objMembershipElement.nodeName = "IMembershipCondition" Then
             strMembership = objMembershipElement.getAttribute("class")
             Exit For
          End If
       Next
       
       For i = 0 to UBound(arrTargetArray)
          blnDefaultCodeGroup = false
          If strPermSet = arrTargetArray(i) Then
             For j = 0 to UBound(arrDefaultCodeGroups)
                If arrDefaultCodeGroups(j) = strCodeGroupName Then
                   blnDefaultCodeGroup = true
                End If
             Next
             If blnDefaultCodeGroup = false Then
                Finding = true
                For j = 0 to UBound(arrGoodMembership)
                   If strMembership = arrGoodMembership(j) Then
                      finding = false
                      Exit For
                   End If  
                Next
                If finding = true Then
                   strMessage = strMessage & vbCrLf & "      " & strCodeGroupName & " Code Group in " &strLevel & " policy has unauthorized " & strPermissionType
                End If
	     End If
          End If
       Next
    Next
End Sub

Sub CodeGroupAttribute(xmlDoc, strAttribute, strValue)
    Dim strAttributeValue
    finding = false
    Set ColElemList = xmlDoc.getElementsByTagName("CodeGroup")
    For Each objElement in colElemList
       strAttributeValue = objElement.getAttribute(strAttribute)
       if strAttributeValue = strValue Then
          finding=true
          strMessage = strMessage & vbCrLf & "      " & objElement.getAttribute("Name") & " has the attribute " & strAttribute & " which is set to " & strValue 
       End If
    Next
End Sub

Sub TypeFilterAttribute(xmlDoc, strAttribute, strValue)
    Dim strAttributeValue
    finding = false
    Set ColElemList = xmlDoc.getElementsByTagName("*")
    For Each objElement in colElemList
       strAttributeValue = objElement.getAttribute(strAttribute)
       if strAttributeValue = strValue Then
          finding=true
          strMessage = strMessage & vbCrLf & "      " & objElement.getAttribute("Name") & " has the attribute " & strAttribute & " which is set to " & strValue 
       End If
    Next
End Sub

Sub NonDefaultMembership(xmlDoc, strAttribute, strValue, arrTargetArray)
    Dim strMembership
    Dim objMembershipElement
    Dim blnDefaultCodeGroup
    Dim strCodeGroupName
    finding = false
    Set ColElemList = xmlDoc.getElementsByTagName("CodeGroup")
    For Each objElement in colElemList
       blnDefaultCodeGroup = false
       strCodeGroupName = objElement.getAttribute("Name")
       blnDefaultCodeGroup = false
       For i = 0 to UBound(arrTargetArray)
          If objElement.getAttribute("Name") = arrTargetArray(i) Then
              blnDefaultCodeGroup = true
              Exit For
          End If
       Next
       If blnDefaultCodeGroup = false Then
          For Each objMembershipElement in objElement.ChildNodes
             If objMembershipElement.nodeName = "IMembershipCondition" Then
                strMembership = objMembershipElement.getAttribute(strAttribute)
                If strMembership = strValue Then
                   finding = true
                   strMessage = strMessage & vbCrLf & "     " & strCodeGroupName & " has unauthorized membership of type " & strValue
                   Exit For
                End If
             End If
          Next
       End If
    Next
End Sub

Sub RegExCodeGroupAttribute(xmlDoc, strAttribute, strValue)
    Dim strAttributeValue
    finding = false
    RegEx.Pattern = strValue
    RegEx.IgnoreCase =  true
    Set ColElemList = xmlDoc.getElementsByTagName("CodeGroup")
    For Each objElement in colElemList
       strAttributeValue = objElement.getAttribute(strAttribute)
       If Not isNull(strAttributeValue) Then
          If RegEx.Test(strAttributeValue) Then
             finding=true
             strMessage = strMessage & vbCrLf & "      " & objElement.getAttribute("Name") & " has the attribute " & strAttribute & " which is set to " & strAttributeValue
          End If
       End If
    Next
End Sub

Sub DualCodeGroupAttribute(xmlDoc, strAttribute, strValue, strAttribute1, strValue1, arrTargetArray)
    Dim strAttributeValue
    finding = false
    Dim blnDefaultCodeGroup
    Set ColElemList = xmlDoc.getElementsByTagName("CodeGroup")
    For Each objElement in colElemList
       blnDefaultCodeGroup = false
       For i = 0 to UBound(arrTargetArray)
          If objElement.getAttribute("Name") = arrTargetArray(i) Then
              blnDefaultCodeGroup = true
              Exit For
          End If
       Next
       If blnDefaultCodeGroup = false Then
          strAttributeValue = objElement.getAttribute(strAttribute)
          if strAttributeValue = strValue Then
             finding=true
             strMessage = strMessage & vbCrLf & "      " & objElement.getAttribute("Name") & " Code Group has the attribute " & strAttribute & " which is set to " & strValue 
          End If
       End If
    Next
    For Each objElement in colElemList
       blnDefaultCodeGroup = false
       For i = 0 to UBound(arrTargetArray)
          If objElement.getAttribute("Name") = arrTargetArray(i) Then
              blnDefaultCodeGroup = true
              Exit For
          End If
       Next
       If blnDefaultCodeGroup = false Then
          if strAttributeValue = strValue1 Then
             finding=true
             strMessage = strMessage & vbCrLf & "      " & objElement.getAttribute("Name") & " Code Group has the attribute " & strAttribute1 & " which is set to " & strValue1 
          End If
       End If
    Next
End Sub

Sub getApplicablePermissions(xmlDoc, arrTargetArray, strPermissionType)
   Dim strName
   i = 0
   Redim arrTargetArray(0)
   Set colElemList = xmlDoc.getElementsByTagName("PermissionSet")
   For Each objElement in colElemList
      strName = objElement.getAttribute("Name")
      For Each objIPermission in objElement.ChildNodes
         strTest = objIPermission.getAttribute("class")
         If strTest = strPermissionType Then
            Redim Preserve arrTargetArray(i)
            arrTargetArray(i) = strName
            i = i + 1
         End If
      Next
      If strName = "FullTrust" Then
         Redim Preserve arrTargetArray(i)
         arrTargetArray(i) = strName
         i = i + 1
      End If 
   Next
End Sub

Sub getApplicablePermissionsSpecific(xmlDoc, arrTargetArray, strPermissionType, strAttribute)
   Dim strName
   Dim strElementAttribute
   i = 0
   Redim arrTargetArray(0)
   Set colElemList = xmlDoc.getElementsByTagName("PermissionSet")
   For Each objElement in colElemList
      strName = objElement.getAttribute("Name")
      For Each objIPermission in objElement.ChildNodes
         strTest = objIPermission.getAttribute("class")
         If strTest = strPermissionType Then
            if objIPermission.getAttribute("Unrestricted") = "true" Then
               Redim Preserve arrTargetArray(i)
               arrTargetArray(i) = strName
               i = i + 1
            Else
               for j = 0 to arrArraySize(intCount)
                  strElementAttribute = objIPermission.getAttribute(strAttribute)
                  If Not isNull(strElementAttribute) Then
                     RegEx.Pattern = arrPermsArray(intCount,j)
                     RegEx.IgnoreCase = True
                     If RegEx.Test(strElementAttribute) Then
                        Redim Preserve arrTargetArray(i)
                        arrTargetArray(i) = strName
                        i = i + 1
                        Exit For
                     End If
                  End If
                Next
            End If
         End If
      Next
      If strName = "FullTrust" Then
         Redim Preserve arrTargetArray(i)
         arrTargetArray(i) = strName
         i = i + 1
      End If 
   Next
End Sub

Sub getApplicablePermissionsChild(xmlDoc, arrTargetArray, strPermissionType, strAttribute)
   Dim strName
   Dim strElementAttribute
   Dim objChildElement
   i = 0
   Redim arrTargetArray(0)
   Set colElemList = xmlDoc.getElementsByTagName("PermissionSet")
   For Each objElement in colElemList
      strName = objElement.getAttribute("Name")
      For Each objIPermission in objElement.ChildNodes
         strTest = objIPermission.getAttribute("class")
         If strTest = strPermissionType Then
            if objIPermission.getAttribute("Unrestricted") = "true" Then
               Redim Preserve arrTargetArray(i)
               arrTargetArray(i) = strName
               i = i + 1
            Else
               For Each objChildElement in objIPermission.ChildNodes
                  for j = 0 to arrArraySize(x)
                     strElementAttribute = objChildElement.getAttribute(strAttribute)
                     If Not isNull(strElementAttribute) Then
                        For k = 0 to arrArraySize(intCount)
                           If objChildElement.nodeName = arrChildArray(intCount,k) Then
                              RegEx.Pattern = arrPermsArray(intCount,j)
                              RegEx.IgnoreCase = True
                              If RegEx.Test(strElementAttribute) Then
                                 Redim Preserve arrTargetArray(i)
                                 arrTargetArray(i) = strName
                                 i = i + 1
                                 Exit For
                              End If
                           End If
                        Next
                     End If
                   Next
               Next
            End If
         End If
      Next
      If strName = "FullTrust" Then
         Redim Preserve arrTargetArray(i)
         arrTargetArray(i) = strName
         i = i + 1
      End If 
   Next
End Sub

Sub DuplicateNames(XmlDoc1, XmlDoc2, XmlDoc3, intCheck)
    Dim arrNames()
    Dim arrCompare
    Dim blnDefaultCodeGroup
    i = 0
    Redim arrNames(0)
    If entBool(intCheck) = true Then
        Set colElemList = XmlDoc1.getElementsByTagName("CodeGroup")
        For Each objElement in colElemList
            blnDefaultCodeGroup = false
            If i > UBound(arrNames) Then
                Redim Preserve arrNames(i)
            End If
            For j = 0 to UBound(arrEnterpriseDefaultCodeGroups)
               If objElement.getAttribute("Name") = arrEnterpriseDefaultCodeGroups(j) Then
                   blnDefaultCodeGroup = true
               End If
            Next
            For j = 0 to UBound(arrMachineDefaultCodeGroups)
               If objElement.getAttribute("Name") = arrMachineDefaultCodeGroups(j) Then
                   blnDefaultCodeGroup = true
                   End If
            Next
            For j = 0 to UBound(arrUserDefaultCodeGroups)
               If objElement.getAttribute("Name") = arrUserDefaultCodeGroups(j) Then
                   blnDefaultCodeGroup = true
               End If
            Next
            If blnDefaultCodeGroup = false Then
                arrNames(i) = objElement.getAttribute("Name")
                i = i + 1
            End If
         Next
    End If
    If macBool(intCheck) = true Then
        Set colElemList = XmlDoc2.getElementsByTagName("CodeGroup")
        For Each objElement in colElemList
            blnDefaultCodeGroup = false
            If i > UBound(arrNames) Then
                Redim Preserve arrNames(i)
            End If
            For j = 0 to UBound(arrEnterpriseDefaultCodeGroups)
                If objElement.getAttribute("Name") = arrEnterpriseDefaultCodeGroups(j) Then
                    blnDefaultCodeGroup = true
                End If
            Next
            For j = 0 to UBound(arrMachineDefaultCodeGroups)
                If objElement.getAttribute("Name") = arrMachineDefaultCodeGroups(j) Then
                    blnDefaultCodeGroup = true
                End If
            Next
            For j = 0 to UBound(arrUserDefaultCodeGroups)
                If objElement.getAttribute("Name") = arrUserDefaultCodeGroups(j) Then
                    blnDefaultCodeGroup = true
                End If
            Next
            If blnDefaultCodeGroup = false Then
                arrNames(i) = objElement.getAttribute("Name")
                i = i + 1
            End If
        Next
   End If
   If usrBool(intCheck) = true Then
       Set colElemList = XmlDoc3.getElementsByTagName("CodeGroup")
       For Each objElement in colElemList
           blnDefaultCodeGroup = false
           If i > UBound(arrNames) Then
               Redim Preserve arrNames(i)
           End If
           For j = 0 to UBound(arrEnterpriseDefaultCodeGroups)
              If objElement.getAttribute("Name") = arrEnterpriseDefaultCodeGroups(j) Then
                  blnDefaultCodeGroup = true
              End If
           Next
           For j = 0 to UBound(arrMachineDefaultCodeGroups)
               If objElement.getAttribute("Name") = arrMachineDefaultCodeGroups(j) Then
                   blnDefaultCodeGroup = true
               End If
           Next
           For j = 0 to UBound(arrUserDefaultCodeGroups)
              If objElement.getAttribute("Name") = arrUserDefaultCodeGroups(j) Then
                  blnDefaultCodeGroup = true
              End If
           Next
           If blnDefaultCodeGroup = false Then
               arrNames(i) = objElement.getAttribute("Name")
               i = i + 1
           End If
       Next
    End If
    arrCompare = arrNames
    For i = 0 to UBound(arrNames)
        k = 0
        For j = 0 to UBound(arrCompare)
            If arrNames(i) = arrCompare(j) Then
                k = k + 1
             End If
        Next
        If k > 1 Then
            RegEx.Pattern = ".*" & arrNames(i) & ".*"
            RegEx.IgnoreCase = true
            finding = true
            If Not RegEx.Test(strMessage) Then
                strMessage = strMessage & vbCrLf & "      " & arrNames(i) & " is duplicate and has " & k & " instances"
            End If
        End If
    Next
End Sub
   

x = 0
arrFinding(x) = "APPNET0001"
arrPermissionSet(x) = "FileIOPermission"
arrType(x) = "PermissionSetGeneric"
arrTitle(x) = "File IO Permissions are granted to unauthorized applications"
arrAttribute(x) = ""
arrCategory(x) = "II"
arrArraySize(x) = "0"

x = x + 1
arrFinding(x) = "APPNET0003"
arrPermissionSet(x) = "IsolatedStorageFilePermission"
arrType(x) = "PermissionSetSpecific"
arrTitle(x) = "The Isolated Storage permissions are granted to unauthorized applications"
arrAttribute(x) = "Allowed"
arrCategory(x) = "II" 
arrPermsArray(x,0) = ".*AdministerIsolatedStorageByUser.*"
arrPermsArray(x,1) = ".*AssemblyIsolationByUser.*"
arrPermsArray(x,2) = ".*AssemblyIsolationByRoamingUser.*"
arrArraySize(x) = 2

x = x + 1
arrFinding(x) = "APPNET0004"
arrPermissionSet(x) = "UIPermission"
arrType(x) = "PermissionSetSpecific"
arrTitle(x) = "User Interface Windowing permissions are granted to unauthorized applications"
arrAttribute(x) = "Window"
arrCategory(x) = "II"
arrPermsArray(x,0) = ".*AllWindows.*"
arrPermsArray(x,1) = ".*SafeTopLevelWindows.*"
arrArraySize(x) = 1

x = x + 1
arrFinding(x) = "APPNET0005"
arrPermissionSet(x) = "UIPermission"
arrType(x) = "PermissionSetSpecific"
arrTitle(x) = "User Interface Clipboard permissions are granted to unauthorized applications"
arrAttribute(x) = "Clipboard"
arrCategory(x) = "II"
arrPermsArray(x,0) = ".*AllClipboard.*"
arrPermsArray(x,1) = ".*OwnClipboard.*"
arrArraySize(x) = 1

x = x + 1
arrFinding(x) = "APPNET0006"
arrPermissionSet(x) = "ReflectionPermission"
arrType(x) = "PermissionSetSpecific"
arrTitle(x) = "Reflection permissions are granted to unauthorized applications"
arrAttribute(x) = "Flags"
arrCategory(x) = "II"
arrPermsArray(x,0) = ".*TypeInformation.*"
arrPermsArray(x,1) = ".*MemberAccess.*"
arrArraySize(x) = 1

x = x + 1
arrFinding(x) = "APPNET0007"
arrPermissionSet(x) = "PrintingPermission"
arrType(x) = "PermissionSetSpecific"
arrTitle(x) = "Printing permissions are granted to unauthorized applications"
arrAttribute(x) = "Level"
arrCategory(x) = "II"
arrPermsArray(x,0) = ".*AllPrinting.*"
arrArraySize(x) = 0

x = x + 1
arrFinding(x) = "APPNET0008"
arrPermissionSet(x) = "DnsPermission"
arrType(x) = "PermissionSetSpecific"
arrTitle(x) = "DNS permissions are granted to unauthorized applications"
arrAttribute(x) = "Level"
arrCategory(x) = "II"
arrPermsArray(x,0) = ".*AllPrinting.*"
arrArraySize(x) = -1

x = x + 1
arrFinding(x) = "APPNET0009"
arrPermissionSet(x) = "SocketPermission"
arrType(x) = "PermissionSetGeneric"
arrTitle(x) = "Socket Permissions are granted to unauthorized applications"
arrAttribute(x) = ""
arrCategory(x) = "II"
arrArraySize(x) = "0"

x = x + 1
arrFinding(x) = "APPNET0010"
arrPermissionSet(x) = "WebPermission"
arrType(x) = "PermissionSetGeneric"
arrTitle(x) = "Web Permissions are granted to unauthorized applications"
arrAttribute(x) = ""
arrCategory(x) = "II"
arrArraySize(x) = "0"

x = x + 1
arrFinding(x) = "APPNET0011"
arrPermissionSet(x) = "MessageQueuePermission"
arrType(x) = "PermissionSetChild"
arrTitle(x) = "Message Queue Permissions are granted to unauthorized applications"
arrAttribute(x) = "access"
arrCategory(x) = "II"
arrPermsArray(x,0) = ".*Administer.*"
arrPermsArray(x,1) = ".*Browse.*"
arrArraySize(x) = 1
arrChildArray(x, 0) = "Path"
arrChildArray(x, 1) = "Criteria"
arrArraySize2(x) = 1

x = x + 1
arrFinding(x) = "APPNET0012"
arrPermissionSet(x) = "ServiceControllerPermission"
arrType(x) = "PermissionSetGeneric"
arrTitle(x) = "Service Controller Permissions are granted to unauthorized applications"
arrAttribute(x) = ""
arrCategory(x) = "II"
arrArraySize(x) = "0"

x = x + 1
arrFinding(x) = "APPNET0013"
arrPermissionSet(x) = "SqlClientPermission"
arrType(x) = "PermissionSetGeneric"
arrTitle(x) = "Sql Client Permissions are granted to unauthorized applications"
arrAttribute(x) = ""
arrCategory(x) = "II"
arrArraySize(x) = "0"

x = x + 1
arrFinding(x) = "APPNET0014"
arrPermissionSet(x) = "SecurityPermission"
arrType(x) = "PermissionSetSpecific"
arrTitle(x) = "Security Permissions (Extend Infrastructure) are granted to unauthorized applications"
arrAttribute(x) = "Flags"
arrCategory(x) = "II"
arrPermsArray(x,0) = ".*Infrastructure.*"
arrArraySize(x) = 0

x = x + 1
arrFinding(x) = "APPNET0015"
arrPermissionSet(x) = "SecurityPermission"
arrType(x) = "PermissionSetSpecific"
arrTitle(x) = "Security Permissions (Enable Remoting Configuration) are granted to unauthorized applications"
arrAttribute(x) = "Flags"
arrCategory(x) = "II"
arrPermsArray(x,0) = ".*RemotingConfiguration.*"
arrArraySize(x) = 0

x = x + 1
arrFinding(x) = "APPNET0016"
arrPermissionSet(x) = "SecurityPermission"
arrType(x) = "PermissionSetSpecific"
arrTitle(x) = "Security Permissions (Enable Serialization Formatter) are granted to unauthorized applications"
arrAttribute(x) = "Flags"
arrCategory(x) = "II"
arrPermsArray(x,0) = ".*SerializationFormatter.*"
arrArraySize(x) = 0

x = x + 1
arrFinding(x) = "APPNET0017"
arrPermissionSet(x) = "SecurityPermission"
arrType(x) = "PermissionSetSpecific"
arrTitle(x) = "Security Permissions (Enable Thread Control) are granted to unauthorized applications"
arrAttribute(x) = "Flags"
arrCategory(x) = "II"
arrPermsArray(x,0) = ".*ControlThread.*"
arrArraySize(x) = 0

x = x + 1
arrFinding(x) = "APPNET0018"
arrPermissionSet(x) = "SecurityPermission"
arrType(x) = "PermissionSetSpecific"
arrTitle(x) = "Security Permissions (Allow Principal Control) are granted to unauthorized applications"
arrAttribute(x) = "Flags"
arrCategory(x) = "II"
arrPermsArray(x,0) = ".*ControlPrincipal.*"
arrArraySize(x) = 0

x = x + 1
arrFinding(x) = "APPNET0019"
arrPermissionSet(x) = "SecurityPermission"
arrType(x) = "PermissionSetSpecific"
arrTitle(x) = "Security Permissions (Enable Assembly Execution) are granted to unauthorized applications"
arrAttribute(x) = "Flags"
arrCategory(x) = "II"
arrPermsArray(x,0) = ".*Execution.*"
arrArraySize(x) = 0

x = x + 1
arrFinding(x) = "APPNET0020"
arrPermissionSet(x) = "SecurityPermission"
arrType(x) = "PermissionSetSpecific"
arrTitle(x) = "Security Permissions (Skip Verification) are granted to unauthorized applications"
arrAttribute(x) = "Flags"
arrCategory(x) = "II"
arrPermsArray(x,0) = ".*SkipVerification.*"
arrArraySize(x) = 0

x = x + 1
arrFinding(x) = "APPNET0021"
arrPermissionSet(x) = "SecurityPermission"
arrType(x) = "PermissionSetSpecific"
arrTitle(x) = "Security Permissions (Allow Calls to Unmanaged Assemblies) are granted to unauthorized applications"
arrAttribute(x) = "Flags"
arrCategory(x) = "II"
arrPermsArray(x,0) = ".*UnmanagedCode.*"
arrArraySize(x) = 0

x = x + 1
arrFinding(x) = "APPNET0022"
arrPermissionSet(x) = "SecurityPermission"
arrType(x) = "PermissionSetSpecific"
arrTitle(x) = "Security Permissions (Allow Policy Control) are granted to unauthorized applications"
arrAttribute(x) = "Flags"
arrCategory(x) = "II"
arrPermsArray(x,0) = ".*ControlPolicy.*"
arrArraySize(x) = 0

x = x + 1
arrFinding(x) = "APPNET0023"
arrPermissionSet(x) = "SecurityPermission"
arrType(x) = "PermissionSetSpecific"
arrTitle(x) = "Security Permissions (Allow Domain Policy Control) are granted to unauthorized applications"
arrAttribute(x) = "Flags"
arrCategory(x) = "II"
arrPermsArray(x,0) = ".*ControlDomainPolicy.*"
arrArraySize(x) = 0

x = x + 1
arrFinding(x) = "APPNET0024"
arrPermissionSet(x) = "SecurityPermission"
arrType(x) = "PermissionSetSpecific"
arrTitle(x) = "Security Permissions (Allow Evidence Control) are granted to unauthorized applications"
arrAttribute(x) = "Flags"
arrCategory(x) = "II"
arrPermsArray(x,0) = ".*ControlEvidence.*"
arrArraySize(x) = 0

x = x + 1
arrFinding(x) = "APPNET0025"
arrPermissionSet(x) = "SecurityPermission"
arrType(x) = "PermissionSetSpecific"
arrTitle(x) = "Security Permissions (Assert Any Permission that Has Been Granted) are granted to unauthorized applications"
arrAttribute(x) = "Flags"
arrCategory(x) = "II"
arrPermsArray(x,0) = ".*Assertion.*"
arrArraySize(x) = 0

x = x + 1
arrFinding(x) = "APPNET0026"
arrPermissionSet(x) = "PerformanceCounterPermission"
arrType(x) = "PermissionSetGeneric"
arrTitle(x) = "Performance Counter Permissions are granted to unauthorized applications"
arrAttribute(x) = ""
arrCategory(x) = "II"
arrArraySize(x) = "0"

x = x + 1
arrFinding(x) = "APPNET0027"
arrPermissionSet(x) = "EnvironmentPermission"
arrType(x) = "PermissionSetGeneric"
arrTitle(x) = "Environment Variable Permissions are granted to unauthorized applications"
arrAttribute(x) = ""
arrCategory(x) = "II"
arrArraySize(x) = "0"

x = x + 1
arrFinding(x) = "APPNET0028"
arrPermissionSet(x) = "EventLogPermission"
arrType(x) = "PermissionSetGeneric"
arrTitle(x) = "Event Log Permissions are granted to unauthorized applications"
arrAttribute(x) = ""
arrCategory(x) = "II"
arrArraySize(x) = "0"

x = x + 1
arrFinding(x) = "APPNET0029"
arrPermissionSet(x) = "RegistryPermission"
arrType(x) = "PermissionSetGeneric"
arrTitle(x) = "Registry Permissions are granted to unauthorized applications"
arrAttribute(x) = ""
arrCategory(x) = "II"
arrArraySize(x) = "0"

x = x + 1
arrFinding(x) = "APPNET0030"
arrPermissionSet(x) = "DirectoryServicesPermission"
arrType(x) = "PermissionSetChild"
arrTitle(x) = "Directory Services Permissions are granted to unauthorized applications"
arrAttribute(x) = "access"
arrCategory(x) = "II"
arrPermsArray(x,0) = ".*Write.*"
arrPermsArray(x,1) = ".*Browse.*"
arrArraySize(x) = 1
Redim Preserve arrChildArray(100,2)
arrChildArray(x, 0) = "Path"
arrArraySize2(x) = 0

x = x + 1
arrFinding(x) = "APPNET0031"
arrPermissionSet(x) = "Software\Microsoft\StrongName\Verification"
arrType(x) = "AssemblyList"
arrTitle(x) = "No Strong Name Verification for Assemblys"
arrAttribute(x) = HKLM
arrCategory(x) = "II"
arrArraySize(x) = "0"

x = x + 1
arrFinding(x) = "APPNET0032"
arrPermissionSet(x) = "FirstMatchCodeGroup"
arrType(x) = "CodeGroupAttribute"
arrTitle(x) = "First Match Code Groups Exist"
arrAttribute(x) = "class"
arrCategory(x) = "II"
arrArraySize(x) = "0"

x = x + 1
arrFinding(x) = "APPNET0033"
arrPermissionSet(x) = "NetCodeGroup"
arrType(x) = "DualCodeGroupAttribute"
arrTitle(x) = "Non-Default Net or File Code Groups Exist"
arrAttribute(x) = "class"
arrCategory(x) = "II"
arrArraySize(x) = "0"
arrPermsArray(x, 0) = "FileCodeGroup"
arrArraySize(x) = 0

x = x + 1
arrFinding(x) = "APPNET0035"
arrPermissionSet(x) = ".*LevelFinal.*"
arrType(x) = "RegExCodeGroupAttribute"
arrTitle(x) = "Level Final Code Group Attributes Exist"
arrAttribute(x) = "Attributes"
arrCategory(x) = "II"
arrArraySize(x) = "0"

x = x + 1
arrFinding(x) = "APPNET0041"
arrPermissionSet(x) = "ZoneMembershipCondition"
arrType(x) = "NonDefaultMembership"
arrTitle(x) = "Code Group Zone Membership Condition"
arrAttribute(x) = "class"
arrCategory(x) = "II"
arrArraySize(x) = "0"

x = x + 1
arrFinding(x) = "APPNET0045"
arrPermissionSet(x) = "Software\Microsoft\.NETFramework\Security\Policy"
arrType(x) = "CASPolSecurity"
arrTitle(x) = "CASPol Security is Turned Off"
arrAttribute(x) = "GlobalSettings"
arrCategory(x) = "II"
arrArraySize(x) = "0"

x = x + 1
arrFinding(x) = "APPNET0046"
arrPermissionSet(x) = 160
arrType(x) = "CertBitMask"
arrTitle(x) = "Administering the Windows Environment for Test Root Certificates"
arrAttribute(x) = 0
arrCategory(x) = "II"
arrArraySize(x) = "0"

x = x + 1
arrFinding(x) = "APPNET0047"
arrPermissionSet(x) = 256
arrType(x) = "CertBitMask"
arrTitle(x) = "Administering the Windows Environment for Expired Certificates"
arrAttribute(x) = 0
arrCategory(x) = "II"
arrArraySize(x) = "0"

x = x + 1
arrFinding(x) = "APPNET0048"
arrPermissionSet(x) = "PublisherMembershipCondition"
arrType(x) = "NonDefaultMembership"
arrTitle(x) = "Code Group Zone Membership Condition"
arrAttribute(x) = "class"
arrCategory(x) = "II"
arrArraySize(x) = "0"

x = x + 1
arrFinding(x) = "APPNET0049"
arrPermissionSet(x) = 512
arrType(x) = "CertBitMask"
arrTitle(x) = "Administering the Windows Environment for Revoked Certificates"
arrAttribute(x) = 0
arrCategory(x) = "II"
arrArraySize(x) = "0"

x = x + 1
arrFinding(x) = "APPNET0050"
arrPermissionSet(x) = 15360
arrType(x) = "CertBitMask"
arrTitle(x) = "Administering the Windows Environment for Unknown Certificate Status"
arrAttribute(x) = 0
arrCategory(x) = "II"
arrArraySize(x) = "0"

x = x + 1
arrFinding(x) = "APPNET0051"
arrPermissionSet(x) = 196608
arrType(x) = "CertBitMask"
arrTitle(x) = "Administering the Windows Environment for Time Stamped Certificate Revocation"
arrAttribute(x) = 65536
arrCategory(x) = "II"
arrArraySize(x) = "0"

x = x + 1
arrFinding(x) = "APPNET0052"
arrPermissionSet(x) = "StrongNameMembershipCondition"
arrType(x) = "NonDefaultMembership"
arrTitle(x) = "Code Group Strong Name Membership Condition"
arrAttribute(x) = "class"
arrCategory(x) = "II"
arrArraySize(x) = "0"

x = x + 1
arrFinding(x) = "APPNET0054"
arrPermissionSet(x) = ""
arrType(x) = "DuplicateNames"
arrTitle(x) = "Administering CAS Policy for Group Names"
arrAttribute(x) = ""
arrCategory(x) = "II"
arrArraySize(x) = "0"

x = x + 1
arrFinding(x) = "APPNET0055"
arrPermissionSet(x) = ""
arrType(x) = "Closed"
arrTitle(x) = "Administering CAS Policy and Policy Configuration File Backups"
arrAttribute(x) = ""
arrCategory(x) = "II"
arrArraySize(x) = "0"

x = x + 1
arrFinding(x) = "APPNET0060"
arrPermissionSet(x) = "typefilterlevel"
arrType(x) = "TypeFilterAttribute"
arrTitle(x) = "Remote Services Authentication and Encryption"
arrAttribute(x) = "Full"
arrCategory(x) = "II"
arrArraySize(x) = "0"

x = x + 1
arrFinding(x) = "APPNET0061"
arrPermissionSet(x) = ""
arrType(x) = "Closed"
arrTitle(x) = "Unsupported .Net Framework Versions"
arrAttribute(x) = ""
arrCategory(x) = "II"
arrArraySize(x) = "0"


Sub Checks(intCheck)
    For intCount = 0 to x
        If arrType(intCount) = "PermissionSetGeneric" Then
        blnAggFinding = false
        strMessage = ""
        If entBool(intCheck) = true Then
            getApplicablePermissions objEnterpriseXmlDoc, arrEntCheckPermSet, arrPermissionSet(intCount)
            ComparePermissionSets objEnterpriseXmlDoc, arrEntCheckPermSet, arrEnterpriseDefaultCodeGroups, arrPermissionSet(intCount), "Enterprise"
            If finding = true Then
                blnAggFinding = true
            End If
        End If
        If macBool(intCheck) = true Then
            getApplicablePermissions objMachineXmlDoc, arrMacCheckPermSet, arrPermissionSet(intCount)
            ComparePermissionSets objMachineXmlDoc, arrMacCheckPermSet, arrMachineDefaultCodeGroups, arrPermissionSet(intCount), "Machine"
            If finding = true Then
                blnAggFinding = true
            End If
        End If
        If usrBool(intCheck) = true Then
            getApplicablePermissions objUserXmlDoc, arrUsrCheckPermSet, arrPermissionSet(intCount)
            ComparePermissionSets objUserXmlDoc, arrUsrCheckPermSet, arrUserDefaultCodeGroups, arrPermissionSet(intCount), "User"
            If finding = true Then
                blnAggFinding = true
            End If
        End If
        If blnAggFinding = true Then
           WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)  & strMessage
        Else
           WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
        End If

        ElseIf arrType(intCount) = "PermissionSetSpecific" Then
            blnAggFinding = false
            strMessage = ""
            If entBool(intCheck) = true Then
                getApplicablePermissionsSpecific objEnterpriseXmlDoc, arrEntCheckPermSet, arrPermissionSet(intCount), arrAttribute(intCount)
                ComparePermissionSets objEnterpriseXmlDoc, arrEntCheckPermSet, arrEnterpriseDefaultCodeGroups, arrPermissionSet(intCount), "Enterprise"
                If finding = true Then
                    blnAggFinding = true
                End If
            End If
            If macBool(intCheck) = true Then
                getApplicablePermissionsSpecific objMachineXmlDoc, arrMacCheckPermSet, arrPermissionSet(intCount), arrAttribute(intCount)
                ComparePermissionSets objMachineXmlDoc, arrMacCheckPermSet, arrMachineDefaultCodeGroups, arrPermissionSet(intCount), "Machine"
                If finding = true Then
                    blnAggFinding = true
                End If
            End If
            If usrBool(intCheck) = true Then
                getApplicablePermissionsSpecific objUserXmlDoc, arrUsrCheckPermSet, arrPermissionSet(intCount), arrAttribute(intCount)
                ComparePermissionSets objUserXmlDoc, arrUsrCheckPermSet, arrUserDefaultCodeGroups, arrPermissionSet(intCount), "User"
                If finding = true Then
                    blnAggFinding = true
                End If
            End If
            If blnAggFinding = true Then
                WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)  & strMessage
            Else
                WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
            End If

        ElseIf arrType(intCount) = "PermissionSetChild" Then
            blnAggFinding = false
            strMessage = ""
            If entBool(intCheck) = true Then
                getApplicablePermissionsChild objEnterpriseXmlDoc, arrEntCheckPermSet, arrPermissionSet(intCount), arrAttribute(intCount)
                ComparePermissionSets objEnterpriseXmlDoc, arrEntCheckPermSet, arrEnterpriseDefaultCodeGroups, arrPermissionSet(intCount), "Enterprise"
                If finding = true Then
                    blnAggFinding = true
                End If
            End If
            If macBool(intCheck) = true Then
                getApplicablePermissionsChild objMachineXmlDoc, arrMacCheckPermSet, arrPermissionSet(intCount), arrAttribute(intCount)
                ComparePermissionSets objMachineXmlDoc, arrMacCheckPermSet, arrMachineDefaultCodeGroups, arrPermissionSet(intCount), "Machine"
                If finding = true Then
                    blnAggFinding = true
                End If
            End If
            If usrBool(intCheck) = true Then
                 getApplicablePermissionsChild objUserXmlDoc, arrUsrCheckPermSet, arrPermissionSet(intCount), arrAttribute(intCount)
                 ComparePermissionSets objUserXmlDoc, arrUsrCheckPermSet, arrUserDefaultCodeGroups, arrPermissionSet(intCount), "User"
                 If finding = true Then
                     blnAggFinding = true
                 End If
            End If
            If blnAggFinding = true Then
               WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)  & strMessage
            Else
               WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
            End If

        ElseIf arrType(intCount) = "AssemblyList" Then
            finding = false
            blnAggFinding = false
            strMessage = ""
            objReg.EnumKey arrAttribute(intCount), arrPermissionSet(intCount), arrKeyList
            If(isArray(arrKeyList)) Then
                For Each strSubkey In arrKeyList
                    finding = true
                    strMessage = strMessage & vbCRLF & "      " & strSubkey & " assembly exists and does not require strong name verification"
                    Exit For
                Next
            End If
            If finding = true Then
                blnAggFinding = true
            End If
            If blnAggFinding = true Then
                WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)  & strMessage
            Else
                WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
            End If

        ElseIf arrType(intCount) = "CodeGroupAttribute" Then
            finding = false
            blnAggFinding = false
            strMessage = ""
            If entBool(intCheck) = true Then
                CodeGroupAttribute objEnterpriseXmlDoc, arrAttribute(intCount), arrPermissionSet(intCount)
                If finding = true Then
                    blnAggFinding = true
                End If
            End If
            If macBool(intCheck) = true Then
                CodeGroupAttribute objMachineXmlDoc, arrAttribute(intCount), arrPermissionSet(intCount)
                If finding = true Then
                    blnAggFinding = true
                End If
            End If
            If usrBool(intCheck) = true Then
                CodeGroupAttribute objUserXmlDoc, arrAttribute(intCount), arrPermissionSet(intCount)
                If finding = true Then
                    blnAggFinding = true
                End If
            End If
            If blnAggFinding = true Then
               WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)  & strMessage
            Else
               WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
            End If

        ElseIf arrType(intCount) = "DualCodeGroupAttribute" Then
            finding = false
            blnAggFinding = false
            strMessage = ""
            If entBool(intCheck) = true Then
                DualCodeGroupAttribute objEnterpriseXmlDoc, arrAttribute(intCount), arrPermissionSet(intCount), arrAttribute(intCount), arrPermsArray(intCount, 0), arrEnterpriseDefaultCodeGroups 
                If finding = true Then
                    blnAggFinding = true
                End If
            End If
            If macBool(intCheck) = true Then
                DualCodeGroupAttribute objMachineXmlDoc, arrAttribute(intCount), arrPermissionSet(intCount), arrAttribute(intCount), arrPermsArray(intCount, 0), arrMachineDefaultCodeGroups
                If finding = true Then
                    blnAggFinding = true
                End If
            End If
            If usrBool(intCheck) = true Then
                DualCodeGroupAttribute objUserXmlDoc, arrAttribute(intCount), arrPermissionSet(intCount), arrAttribute(intCount), arrPermsArray(intCount, 0), arrUserDefaultCodeGroups
                If finding = true Then
                    blnAggFinding = true
                End If
            End If
            If blnAggFinding = true Then
               WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)  & strMessage
            Else
               WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
            End If

         ElseIf arrType(intCount) = "RegExCodeGroupAttribute" Then
            finding = false
            blnAggFinding = false
            strMessage = ""
            If entBool(intCheck) = true Then
                RegExCodeGroupAttribute objEnterpriseXmlDoc, arrAttribute(intCount), arrPermissionSet(intCount)
                If finding = true Then
                    blnAggFinding = true
                End If
            End If
            If macBool(intCheck) = true Then
                RegExCodeGroupAttribute objMachineXmlDoc, arrAttribute(intCount), arrPermissionSet(intCount)
                If finding = true Then
                    blnAggFinding = true
                End If
            End If
            If usrBool(intCheck) = true Then
                RegExCodeGroupAttribute objUserXmlDoc, arrAttribute(intCount), arrPermissionSet(intCount)
                If finding = true Then
                    blnAggFinding = true
                End If
            End If
            If blnAggFinding = true Then
               WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)  & strMessage
            Else
               WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
            End If

        ElseIf arrType(intCount) = "NonDefaultMembership" Then
            finding = false
            blnAggFinding = false
            strMessage = ""
            If entBool(intCheck) = true Then
                NonDefaultMembership objEnterpriseXmlDoc, arrAttribute(intCount), arrPermissionSet(intCount), arrEnterpriseDefaultCodeGroups
                If finding = true Then
                    blnAggFinding = true
                End If
            End If
            If macBool(intCheck) = true Then
                NonDefaultMembership objMachineXmlDoc, arrAttribute(intCount), arrPermissionSet(intCount), arrMachineDefaultCodeGroups
                If finding = true Then
                    blnAggFinding = true
                End If
            End If
            If usrBool(intCheck) = true Then
                NonDefaultMembership objUserXmlDoc, arrAttribute(intCount), arrPermissionSet(intCount), arrUserDefaultCodeGroups
                If finding = true Then
                    blnAggFinding = true
                End If
            End If
            If blnAggFinding = true Then
               WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)  & strMessage
            Else
               WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
            End If

        ElseIf arrType(intCount) = "CASPolSecurity" Then
            finding = false
            blnAggFinding = false
            strMessage = ""
            objReg.EnumValues HKLM, arrPermissionSet(intCount), arrKeyList
            If(isArray(arrKeyList)) Then
                For Each strSubkey In arrKeyList
                    If (StrComp(arrAttribute(intCount),strSubkey,vbTextCompare)) = 0 Then
                       finding = true
                       strMessage = strMessage & vbCrLf & "      CASPol Security is turned off."
                       Exit For
                    End If
                Next
            End If
   
            If finding = true Then
               blnAggFinding = true
            End If
            If blnAggFinding = true Then
               WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)  & strMessage
            Else
               WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
            End If

        ElseIf arrType(intCount) = "CertBitMask" Then
            finding = false
            blnAggFinding = false
            Dim strValue
            strMessage = ""
            objReg.GetDWORDValue HKCU, "Software\Microsoft\Windows\CurrentVersion\WinTrust\Trust Providers\Software Publishing", "State", strValue
            If (strValue And arrPermissionSet(intCount)) = arrAttribute(intCount) Then
               finding = false
            Else
               finding = true
               blnAggFinding = true
               strMessage = vbCrLf & "      " & arrTitle(intCount) & " is set incorrectly"
            End If
            If blnAggFinding = false Then
               objReg.GetDWORDValue HKU, "S-1-5-18\Software\Microsoft\Windows\CurrentVersion\WinTrust\Trust Providers\Software Publishing", "State", strValue
               If (strValue And arrPermissionSet(intCount)) = arrAttribute(intCount) Then
                  finding = false
               Else
                  finding = true
                  blnAggFinding = true
                  strMessage = vbCrLf & "      " & arrTitle(intCount) & " is set incorrectly"
               End If
            End If
            If blnAggFinding = false Then
               objReg.GetDWORDValue HKU, "S-1-5-19\Software\Microsoft\Windows\CurrentVersion\WinTrust\Trust Providers\Software Publishing", "State", strValue     
               If (strValue And arrPermissionSet(intCount)) = arrAttribute(intCount) Then
                   finding = false
               Else
                   finding = true
                   blnAggFinding = true
                  strMessage = vbCrLf & "      " & arrTitle(intCount) & " is set incorrectly"
               End If
            End If
            If blnAggFinding = false Then
                objReg.GetDWORDValue HKU, "S-1-5-20\Software\Microsoft\Windows\CurrentVersion\WinTrust\Trust Providers\Software Publishing", "State", strValue
                If (strValue And arrPermissionSet(intCount)) = arrAttribute(intCount) Then
                    finding = false
                Else
                    finding = true
                    blnAggFinding = true
                    strMessage = vbCrLf & "      " & arrTitle(intCount) & " is set incorrectly"
                End If
            End If
   
            If finding = true Then
                blnAggFinding = true
            End If
            If blnAggFinding = true Then
                WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)  & strMessage
            Else
                WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
            End If

        ElseIf arrType(intCount) = "DuplicateNames" Then
            finding = false
            blnAggFinding = false
            strMessage = ""
            DuplicateNames objEnterpriseXmlDoc, objMachineXmlDoc, objUserXmlDoc, intCheck
            If finding = true Then
                blnAggFinding = true
            End If
            If blnAggFinding = true Then
                WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)  & strMessage
            Else
                WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
            End If

        ElseIf arrType(intCount) = "Closed" Then
            blnAggFinding = false
            If blnAggFinding = true Then
                WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)  & strMessage
            Else
                WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
            End If

        ElseIf arrType(intCount) = "TypeFilterAttribute" Then
            finding = false
            blnAggFinding = false     
            strMessage = ""
            TypeFilterAttribute objEnterpriseXmlDoc, arrAttribute(intCount), arrPermissionSet(intCount)
            If finding = true Then
                blnAggFinding = true
            End If
            TypeFilterAttribute objMachineXmlDoc, arrAttribute(intCount), arrPermissionSet(intCount)
            If finding = true Then
                blnAggFinding = true
            End If
            TypeFilterAttribute objUserXmlDoc, arrAttribute(intCount), arrPermissionSet(intCount)
            If finding = true Then
                blnAggFinding = true
            End If
            If blnAggFinding = true Then
                WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)  & strMessage
            Else
                WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
            End If
        End If
    Next
End Sub

versionCount = 0
Redim netVersion(0)
Redim entBool(0)
Redim entFile(0)
Redim macBool(0)
Redim macFile(0)
Redim usrBool(0)
Redim usrFile(0)
Dim dotNetCaspol
Set wshShell = WScript.CreateObject("WScript.shell")
For Each objSubFolder in objFolder.Subfolders
   RegEx.Pattern = "v\d.*"
   RegEx.IgnoreCase = true
   If RegEx.Test(objSubFolder.Name) Then
      If objFSO.FileExists("C:\Windows\Microsoft.NET\Framework\" & objSubFolder.Name & "\caspol.exe") Then
         dotNetCaspol= "C:\Windows\Microsoft.NET\Framework\" & objSubFolder.Name & "\caspol.exe"
         Redim Preserve netVersion(versionCount)
         Redim Preserve entBool(versionCount)
         Redim Preserve entFile(versionCount)
         Redim Preserve macBool(versionCount)
         Redim Preserve macFile(versionCount)
         Redim Preserve usrBool(versionCount)
         Redim Preserve usrFile(versionCount)
         netVersion(versionCount) = objSubFolder.Name
         If objFSO.FileExists("C:\Windows\Microsoft.NET\Framework\" & objSubFolder.Name & "\CONFIG\security.config") Then
            macBool(versionCount) = true
            macFile(versionCount) = "C:\Windows\Microsoft.NET\Framework\" & objSubFolder.Name & "\CONFIG\security.config"
         Else
            wshShell.Run "cmd /c " & dotNetCaspol & " -q -m -reset",1,True
            macBool(versionCount) = true
            macFile(versionCount) = "C:\Windows\Microsoft.NET\Framework\" & objSubFolder.Name & "\CONFIG\security.config"
         End If   
         If objFSO.FileExists("C:\Windows\Microsoft.NET\Framework\" & objSubFolder.Name & "\CONFIG\enterprisesec.config") Then
            entBool(versionCount) = true
            entFile(versionCount) = "C:\Windows\Microsoft.NET\Framework\" & objSubFolder.Name & "\CONFIG\enterprisesec.config" 
         Else
            wshShell.Run "cmd /c " & dotNetCaspol & " -q -en -reset",1,True
            entBool(versionCount) = true
            entFile(versionCount) = "C:\Windows\Microsoft.NET\Framework\" & objSubFolder.Name & "\CONFIG\enterprisesec.config " 
         End If 
         If Not objFSO.FolderExists(strUserProfile & "\Application Data\Microsoft\CLR Security Config") Then
            wshShell.Run "cmd /c " & dotNetCaspol & " -q -u -reset",1,True
            Set objFolder2 = objFSO.GetFolder(strUserProfile & "\Application Data\Microsoft\CLR Security Config")
         End If
         RegEx.Pattern = objSubFolder.Name & ".*"
         RegEx.IgnoreCase = true
         For Each objSubFolder2 in objFolder2.Subfolders
            If RegEx.Test(objSubFolder2.Name) Then
               If objFSO.FileExists(strUserProfile & "\Application Data\Microsoft\CLR Security Config\" & objSubFolder2.Name & "\security.config") Then
                  usrBool(versionCount) = true
                  usrFile(versionCount) = strUserProfile & "\Application Data\Microsoft\CLR Security Config\" & objSubFolder2.Name & "\security.config"
               Else
                  wshShell.Run "cmd /c " & dotNetCaspol & " -q -u -reset",1,True
                  usrBool(versionCount) = true
                  usrFile(versionCount) = strUserProfile & "\Application Data\Microsoft\CLR Security Config\" & objSubFolder2.Name & "\security.config"
               End If 
            End If
         Next 
         versionCount = versionCount + 1   
      End If
   End If
Next

If objFSO.FolderExists("C:\Windows\Microsoft.NET\Framework64") Then
   Set objFolder = objFSO.GetFolder("C:\Windows\Microsoft.NET\Framework64")
   For Each objSubFolder in objFolder.Subfolders
       RegEx.Pattern = "v\d.*"
       RegEx.IgnoreCase = true
       If RegEx.Test(objSubFolder.Name) Then
           If objFSO.FileExists("C:\Windows\Microsoft.NET\Framework64\" & objSubFolder.Name & "\caspol.exe") Then
               dotNetCaspol= "C:\Windows\Microsoft.NET\Framework64\" & objSubFolder.Name & "\caspol.exe"
               Redim Preserve netVersion(versionCount)
               Redim Preserve entBool(versionCount)
               Redim Preserve entFile(versionCount)
               Redim Preserve macBool(versionCount)
               Redim Preserve macFile(versionCount)
               Redim Preserve usrBool(versionCount)
               Redim Preserve usrFile(versionCount)
               netVersion(versionCount) = objSubFolder.Name & " 64bit"
               If objFSO.FileExists("C:\Windows\Microsoft.NET\Framework64\" & objSubFolder.Name & "\CONFIG\security.config") Then
                  macBool(versionCount) = true
                  macFile(versionCount) = "C:\Windows\Microsoft.NET\Framework64\" & objSubFolder.Name & "\CONFIG\security.config"
               Else
                  wshShell.Run "cmd /c " & dotNetCaspol & " -q -m -reset",1,True
                  macBool(versionCount) = true
                  macFile(versionCount) = "C:\Windows\Microsoft.NET\Framework64\" & objSubFolder.Name & "\CONFIG\security.config"
               End If   
               If objFSO.FileExists("C:\Windows\Microsoft.NET\Framework64\" & objSubFolder.Name & "\CONFIG\enterprisesec.config") Then
                  entBool(versionCount) = true
                  entFile(versionCount) = "C:\Windows\Microsoft.NET\Framework64\" & objSubFolder.Name & "\CONFIG\enterprisesec.config" 
               Else
                  wshShell.Run "cmd /c " & dotNetCaspol & " -q -en -reset",1,True
                  entBool(versionCount) = true
                  entFile(versionCount) = "C:\Windows\Microsoft.NET\Framework64\" & objSubFolder.Name & "\CONFIG\enterprisesec.config "
               End If 
               If Not objFSO.FolderExists(strUserProfile & "\Application Data\Microsoft\CLR Security Config") Then
	           wshShell.Run "cmd /c " & dotNetCaspol & " -q -u -reset",1,True
                   Set objFolder2 = objFSO.GetFolder(strUserProfile & "\Application Data\Microsoft\CLR Security Config")
               End If
               RegEx.Pattern = objSubFolder.Name & ".*"
               RegEx.IgnoreCase = true
               For Each objSubFolder2 in objFolder2.Subfolders
                   If RegEx.Test(objSubFolder2.Name) Then
                       If objFSO.FileExists(strUserProfile & "\Application Data\Microsoft\CLR Security Config\" & objSubFolder2.Name & "\64bit\security.config") Then
                           usrBool(versionCount) = true
                           usrFile(versionCount) = strUserProfile & "\Application Data\Microsoft\CLR Security Config\" & objSubFolder2.Name & "\64bit\security.config"
                       Else
                           wshShell.Run "cmd /c " & dotNetCaspol & " -q -u -reset",1,True
                           usrBool(versionCount) = true
                           usrFile(versionCount) = strUserProfile & "\Application Data\Microsoft\CLR Security Config\" & objSubFolder2.Name & "\security.config"
                       End If 
                   End If
               Next 
               versionCount = versionCount + 1   
           End If
       End If
    Next
End If

WScript.Echo

For versionCount = 0 to UBound(netVersion)
   setFiles(versionCount)
   WScript.Echo "Performing Checks on Microsoft.NET " & netVersion(versionCount)
   checks(versionCount)
   WScript.Echo
   WScript.Echo
Next
      