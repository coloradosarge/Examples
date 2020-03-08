Option Explicit
On Error Resume Next
Dim strComputer
Dim objReg
Dim objReg2
Dim arrHive(600)
Dim arrKeyPath(600)
Dim arrKey(600)
Dim arrSubKey(600)
Dim arrSubKey2(600)
Dim arrSubKey3(600)
Dim arrValue(600)
Dim arrValue2(600)
Dim arrValue3(600)
Dim arrFinding(600)
Dim arrExpectedValue(600)
Dim arrExpectedValue2(600)
Dim arrExpectedValue3(600)
Dim arrType(600)
Dim arrFile1(600)
Dim arrFile2(600)
Dim arrFile3(600)
Dim arrTitle(600)
Dim arrCategory(600)
Dim arrKeyList
Dim strSubkey
Dim intCompare
Dim intCompare2
Dim intCount
Dim secFile1
Dim secFile2
Dim SecFile3
Dim objSD1
Dim objSD2
Dim objSD3
Dim boolResult
Dim finding
Dim objACE
Dim colDrives
Dim objDrive
Dim objWMIService
Dim result
Dim strDrive
Dim RegEx
Dim wshShell
Dim strNextLine
Dim objTextFile
Dim objFSO
Dim i
Dim secLines()
Dim colItems
Dim WshNetwork
Dim objItem
Dim objUser
Dim dtmLastLogin
Dim RegEx2
Dim RegEx3
Dim arrMultiReg
Dim objGroup
Dim arrServices(110,2)
Dim objService
Dim colServiceList
Dim strMessage
Dim arrRemoteReg(11)
Dim j
Dim arrBinary()
Dim pfxLine
Dim objOutFile
Dim pfxLines()
Dim auditLines()
Dim globalFind
Dim x
Dim colShares
Dim objShare
Dim objShareSecurity
Dim fileVer
Dim ServicePack
Dim colOSItems
Dim colKB
Dim objKB
Dim colRoles
Dim objRole
Dim intYear
Dim intMonth
Dim intDay
Dim strDate
Dim datAVDate
Dim colPrinters
Dim objPrinter
Dim colProducts
Dim objProduct
Dim certLines()
Dim compFilter
Dim colComputerSystems
Dim objComputerSystem
Dim objPolicy
Dim colRules
Dim objRule
Dim testValue1
Dim testValue2
Dim testValue3
Dim actualValue
Dim h
Dim l
Dim objFolder
Dim colSubFolders

globalFind = "no"

arrServices(0,0) = "AeLookupService"
arrServices(0,1) = "Manual"
arrServices(1,0) = "AppIDSvc"
arrServices(1,1) = "Manual"
arrServices(2,0) = "AppInfo"
arrServices(2,1) = "Manual"
arrServices(3,0) = "ALG"
arrServices(3,1) = "Manual"
arrServices(4,0) = "AppMgmt"
arrServices(4,1) = "Manual"
arrServices(5,0) = "BITS"
arrServices(5,1) = "Manual"
arrServices(6,0) = "CertPropSvc"
arrServices(6,1) = "Manual"
arrServices(7,0) = "KeyIso"
arrServices(7,1) = "Manual"
arrServices(8,0) = "EventSystem"
arrServices(8,1) = "Auto"
arrServices(9,0) = "COMSysApp"
arrServices(9,1) = "Manual"
arrServices(10,0) = "Browser"
arrServices(10,1) = "Disabled"
arrServices(11,0) = "VaultSvc"
arrServices(11,1) = "Manual"
arrServices(12,0) = "CryptSvc"
arrServices(12,1) = "Auto"
arrServices(13,0) = "DcomLaunch"
arrServices(13,1) = "Auto"
arrServices(14,0) = "UxSms"
arrServices(14,1) = "Auto"
arrServices(15,0) = "Dhcp"
arrServices(15,1) = "Auto"
arrServices(16,0) = "DPS"
arrServices(16,1) = "Auto"
arrServices(17,0) = "WdiServiceHost"
arrServices(17,1) = "Manual"
arrServices(18,0) = "WdiSystemHost"
arrServices(18,1) = "Manual"
arrServices(19,0) = "defragsvc"
arrServices(19,1) = "Manual"
arrServices(20,0) = "TrkWks"
arrServices(20,1) = "Auto"
arrServices(21,0) = "MSDTC"
arrServices(21,1) = "Auto"
arrServices(22,0) = "Dnscache"
arrServices(22,1) = "Auto"
arrServices(23,0) = "EFS"
arrServices(23,1) = "Manual"
arrServices(24,0) = "EapHost"
arrServices(24,1) = "Manual"
arrServices(25,0) = "fdPHost"
arrServices(25,1) = "Manual"
arrServices(26,0) = "FDResPub"
arrServices(26,1) = "Manual"
arrServices(27,0) = "gpsvc"
arrServices(27,1) = "Auto"
arrServices(28,0) = "hkmsvc"
arrServices(28,1) = "Manual"
arrServices(29,0) = "hidserv"
arrServices(29,1) = "Manual"
arrServices(30,0) = "IKEEXT"
arrServices(30,1) = "Manual"
arrServices(31,0) = "UIODetect"
arrServices(31,1) = "Manual"
arrServices(32,0) = "SharedAccess"
arrServices(32,1) = "Disabled"
arrServices(33,0) = "iphlpsvc"
arrServices(33,1) = "Auto"
arrServices(34,0) = "PolicyAgent"
arrServices(34,1) = "Manual"
arrServices(35,0) = "KtmRm"
arrServices(35,1) = "Manual"
arrServices(36,0) = "lltdsvc"
arrServices(36,1) = "Manual"
arrServices(37,0) = "clr_optimization_v2.0.50727_64"
arrServices(37,1) = "Manual"
arrServices(38,0) = "clr_optimization_v2.0.50727_32"
arrServices(38,1) = "Manual"
arrServices(39,0) = "FCRegSvc"
arrServices(39,1) = "Manual"
arrServices(40,0) = "MSiSCSI"
arrServices(40,1) = "Auto"
arrServices(41,0) = "swprv"
arrServices(41,1) = "Manual"
arrServices(42,0) = "MMCSS"
arrServices(42,1) = "Manual"
arrServices(43,0) = "Netlogon"
arrServices(43,1) = "Manual"
arrServices(44,0) = "napagent"
arrServices(44,1) = "Manual"
arrServices(45,0) = "Netman"
arrServices(45,1) = "Manual"
arrServices(46,0) = "netprofm"
arrServices(46,1) = "Manual"
arrServices(47,0) = "NlaSvc"
arrServices(47,1) = "Auto"
arrServices(48,0) = "nsi"
arrServices(48,1) = "Auto"
arrServices(49,0) = "PerfHost"
arrServices(49,1) = "Manual"
arrServices(50,0) = "pla"
arrServices(50,1) = "Manual"
arrServices(51,0) = "PlugPlay"
arrServices(51,1) = "Auto"
arrServices(52,0) = "IPBusEnum"
arrServices(52,1) = "Disabled"
arrServices(53,0) = "WPDBusEnum"
arrServices(53,1) = "Manual"
arrServices(54,0) = "Power"
arrServices(54,1) = "Auto"
arrServices(55,0) = "Spooler"
arrServices(55,1) = "Auto"
arrServices(56,0) = "wercplsupport"
arrServices(56,1) = "Manual"
arrServices(57,0) = "ProtectedStorage"
arrServices(57,1) = "Manual"
arrServices(58,0) = "RasAuto"
arrServices(58,1) = "Manual"
arrServices(59,0) = "RasMan"
arrServices(59,1) = "Manual"
arrServices(60,0) = "SessionEnv"
arrServices(60,1) = "Manual"
arrServices(61,0) = "TermService"
arrServices(61,1) = "Manual"
arrServices(62,0) = "UmRdpService"
arrServices(62,1) = "Manual"
arrServices(63,0) = "RpcSs"
arrServices(63,1) = "Auto"
arrServices(64,0) = "RpcLocator"
arrServices(64,1) = "Manual"
arrServices(65,0) = "RemoteRegistry"
arrServices(65,1) = "Auto"
arrServices(66,0) = "RSoPProv"
arrServices(66,1) = "Manual"
arrServices(67,0) = "RemoteAccess"
arrServices(67,1) = "Disabled"
arrServices(68,0) = "RpcEptMapper"
arrServices(68,1) = "Auto"
arrServices(69,0) = "seclogon"
arrServices(69,1) = "Manual"
arrServices(70,0) = "SstpSvc"
arrServices(70,1) = "Manual"
arrServices(71,0) = "SamSs"
arrServices(71,1) = "Auto"
arrServices(72,0) = "LanmanServer"
arrServices(72,1) = "Auto"
arrServices(73,0) = "ShellHWDetection"
arrServices(73,1) = "Auto"
arrServices(74,0) = "SCardSvr"
arrServices(74,1) = "Manual"
arrServices(75,0) = "SCPolicySvc"
arrServices(75,1) = "Auto"
arrServices(76,0) = "SNMPTRAP"
arrServices(76,1) = "Manual"
arrServices(77,0) = "sppsvc"
arrServices(77,1) = "Auto"
arrServices(78,0) = "sacsvr"
arrServices(78,1) = "Manual"
arrServices(79,0) = "appuinotify"
arrServices(79,1) = "Manual"
arrServices(80,0) = "SSDPSRV"
arrServices(80,1) = "Disabled"
arrServices(81,0) = "SENS"
arrServices(81,1) = "Auto"
arrServices(82,0) = "Schedule"
arrServices(82,1) = "Auto"
arrServices(83,0) = "lmhosts"
arrServices(83,1) = "Auto"
arrServices(84,0) = "TapiSvc"
arrServices(84,1) = "Manual"
arrServices(85,0) = "THREADORDER"
arrServices(85,1) = "Manual"
arrServices(86,0) = "TBS"
arrServices(86,1) = "Manual"
arrServices(87,0) = "upnphost"
arrServices(87,1) = "Disabled"
arrServices(88,0) = "ProfSvc"
arrServices(88,1) = "Auto"
arrServices(89,0) = "vds"
arrServices(89,1) = "Manual"
arrServices(90,0) = "VSS"
arrServices(90,1) = "Manual"
arrServices(91,0) = "AudioSrv"
arrServices(91,1) = "Manual"
arrServices(92,0) = "AudioEndpointBuilder"
arrServices(92,1) = "Manual"
arrServices(93,0) = "WcsPlugInService"
arrServices(93,1) = "Manual"
arrServices(94,0) = "wudfsvc"
arrServices(94,1) = "Manual"
arrServices(95,0) = "WerSvc"
arrServices(95,1) = "Manual"
arrServices(96,0) = "Wecsvc"
arrServices(96,1) = "Manual"
arrServices(97,0) = "eventlog"
arrServices(97,1) = "Auto"
arrServices(98,0) = "MpsSvc"
arrServices(98,1) = "Auto"
arrServices(99,0) = "FontCache"
arrServices(99,1) = "Manual"
arrServices(100,0) = "msiserver"
arrServices(100,1) = "Manual"
arrServices(101,0) = "Winmgmt"
arrServices(101,1) = "Auto"
arrServices(102,0) = "TrustedInstaller"
arrServices(102,1) = "Manual"
arrServices(103,0) = "WinRM"
arrServices(103,1) = "Auto"
arrServices(104,0) = "W32Time"
arrServices(104,1) = "Auto"
arrServices(105,0) = "wuauserv"
arrServices(105,1) = "Auto"
arrServices(106,0) = "WinHttpAutoProxySvc"
arrServices(106,1) = "Manual"
arrServices(107,0) = "dot3svc"
arrServices(107,1) = "Manual"
arrServices(108,0) = "wmiApSrv"
arrServices(108,1) = "Manual"
arrServices(109,0) = "LanmanWorkstation"
arrServices(109,1) = "Auto"

arrRemoteReg(0) = "Software\Microsoft\OLAP Server"
arrRemoteReg(1) = "Software\Microsoft\Windows NT\CurrentVersion\Perflib"
arrRemoteReg(2) = "Software\Microsoft\Windows NT\CurrentVersion\Print"
arrRemoteReg(3) = "Software\Microsoft\Windows NT\CurrentVersion\Windows"
arrRemoteReg(4) = "System\CurrentControlSet\Control\ContentIndex"
arrRemoteReg(5) = "System\CurrentControlSet\Control\Print\Printers"
arrRemoteReg(6) = "System\CurrentControlSet\Control\Terminal Server"
arrRemoteReg(7) = "System\CurrentControlSet\Control\Terminal Server\UserConfig"
arrRemoteReg(8) = "System\CurrentControlSet\Control\Terminal Server\DefaultUserConfiguration"
arrRemoteReg(9) = "System\CurrentControlSet\Services\Eventlog"
arrRemoteReg(10) = "System\CurrentControlSet\Services\Sysmonlog"

Const HKCU = &H80000001 'HKEY_CURRENT_USER
Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
Const PWDNEVEREXPIRES = &H10000
Const PWDREQUIRED = &H20

WScript.Echo "Performing prescript activites....."

x = 0

x = x + 1
arrFinding(x) = "V-21883 - IAVAI 2009-A-0106"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Cisco VPN Client"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-21926 - IAVA 2009-A-0112"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Blackberry Desktop Manager Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-22092 - IAVA 2009-A-0133"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Roxio Creator Image Parsing Integer Overflow Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-22181 - IAVA 2010-B-0002"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Remote Vulnerabilities in Sun Java System Directory Server"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-22633 - IAVA 2010-A-0017"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "CiscoWorks Remote Code Execution Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-22671 - IAVA 2010-B-0009"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Adobe Products XML Processing Information Disclosure Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-22672 - IAVA 2010-B-0010"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "HP Network Node Manager Arbitrary Command Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-22679 - IAVA 2010-A-0025"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Quartz.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7600.16490"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft DirectShow Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-22681 - IAVA 2010-A-0027"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Office Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-22682 - IAVA 2010-A-0028"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Remote Vulnerabilities in Microsoft Office PowerPoint"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-22694 - IAVA 2010-A-0036"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Symantec Products"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-22695 - IAVA 2010-A-0035"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Oracle WebLogic Server Remote Command Execution Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-23997 - IAVA 2010-A-0066"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in VMware Products"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-24010 - IAVA 2010-B-0033"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "HP Operations Manager Buffer Overflow Vulnerability"
arrCategory(x) = "I"

arrFinding(x) = "V-24163 - IAVA 2010-B-0037"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in HP OpenView Network Node Manager"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-24168 - IAVA 2010-B-0039"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Program Files\Windows Mail\Msoe.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7600.20659"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Outlook Express and Windows Mail Remote Code Exection Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-24366 - IAVA 2010-B-0045"
arrKeyPath(x) = "Default"
arrSubKey(x) = "Authsspi.dll"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\winsxs\"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "7.5.7600.20659"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "WPFFileVer"
arrTitle(x) = "Microsoft Internet Information Services Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-24367 - IAVA 2010-B-0046"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\Microsoft.NET\Framework\v2.0.50727\System.security.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "2.0.50727.4434"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft .NET Framework Data Tampering Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-24371 - IAVA 2010-A-0078"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Asycfilt.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7600.16544"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Multiple Vulnerabilities in Microsoft Windows"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-24377 - IAVA 2010-A-0079"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Micrsoft Office SharePoint"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-24847 - IAVA 2010-B-0053"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\cdd.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7600.16595"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Canonical Display Driver Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-24850 - IAVA 2010-A-0094"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulerabilities in Microsoft Office Access"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-24852 - IAVA 2010-A-0093"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Office Outlook Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-25026 - IAVA 2010-B-0057"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Citrix Online Plug-in and ICA Client Remote Code Execution Vulnerabilities"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-25067 - IAVA 2010-A-0103"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\SysWOW64\iccvid.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1.10.0.13"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Cinepak Codec Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-25074 - IAVA 2010-B-0064"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Rtutils.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7600.16617"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Multiple Vulnerabilities in Micrososoft Windows Tracing Feature for Services"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-25184 - IAVA 2010-B-0075"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in VxWorks"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-25352 - IAVA 2010-A-0132"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "BlackBerry Desktop Software Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-25353 - IAVA 2010-A-0120"
arrKeyPath(x) = "Default"
arrSubKey(x) = "x"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Inetsrv\asp.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "7.5.7600.16620"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Multiple Vulnerabilities in Microsoft Internet Information Services (IIS)"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-25362 - IAVA 2010-A-0124"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Spoolsv.exe"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7600.16661"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows Print Spooler Remote Code Execution Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-25510 - IAVA 2010-A-0145"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Microsoft Office Word"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-25516 - IAVA 2010-A-0140"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Wmp.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "12.0.7600.16667"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows Media Player Remote Code Execution Vulneraiblity"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-25520 - IAVA 2010-A-0141"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Wmpmde.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "12.0.7600.16661"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows Media Player Network Sharing Service Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-25528 - IAVA 2010-A-0135"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ServicePack"
arrTitle(x) = "Microsoft Windows Embedded OpenType Font Engine vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-25530 - IAVA 2010-A-0134"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Structuredquery.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "7.0.7600.16587"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows COM Validation Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-25532 - IAVA 2010-B-0091"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\SysWOW64\Mfc40.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "4.1.0.6151"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Foundation Classes Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-25534 - IAVA 2010-B-0090"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Comctl32.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "5.82.7600.16661"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows Common Control Library Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-25619 - IAVA 2010-B-0098"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Intel Xeon Baseboard Management Component (BMC) Privilege Escalation Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-25765 - IAVA 2010-A-0164"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "CiscoWorks Common Services Remote Code Execution Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-25835 - IAVA 2010-A-0168"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in VMware Products"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-25844 - IAVA 2010-A-0171"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Office Publisher Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-25845 - IAVA 2010-A-0173"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Program Files\Windows Mail\wab.exe"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7600.16684"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows Address Book Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-25851 - IAVA 2010-B-0117"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Consent.exe"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7600.16688"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows Consent User Interface Elevation of Privilege Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-25855 - IAVA 2010-B-0170"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Microsoft Office"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-25862 - IAVA 2010-B-0110"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Taskschd.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7600.16699"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Task Scheduler Elevation of Privilege Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-25866 - IAVA 2010-B-0118"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "HP StorageWorks Unauthorized Access Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-25868 - IAVA 2011-B-0001"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "HP Multiple LaserJet Printers Information Disclosure Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-25869 - IAVA 2011-B-0002"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "IBM WebSphere Service Registry and Repository Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-25885 - IAVA 2011-A-0003"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "BlackBerry Attachment Service PDF Distiller Remote Buffer Overflow Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-25887 - IAVA 2011-A-0004"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Program Files (x86)\Common Files\System\msadc\Msadco.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7600.16688"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Data Access Components Remote Code Execution Vulnerabilities"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26049 - IAVA 2011-A-0011"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Symantec Products"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-26050 - IAVA 2011-B-0013"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in IBM DB2"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-26056 - IAVA 2011-B-0017"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\JScript.dll"
arrFile2(x) = "C:\Windows\System32\Vbscript.dll"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "5.8.7601.17535"
arrExpectedValue2(x) = "5.8.7601.17535"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft JScript and VBScript Scripting Information Disclosure Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26065 - IAVA 2011-A-0022"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Ntdll.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7600.16695"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Multiple Vulnerabilities in Microsoft Windows Kernel"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-26067 - IAVA 2011-A-0021"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Kerberos.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17527"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Multiple Vulnerabilities in Microsoft Windows Kerberos"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26075 - IAVA 2011-B-0020"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Computer Associates ARCserver Password Security Bypass Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-26077 - IAVA 2011-B-0021"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulerablilities in IBM Tivoli Access Manager"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-26082 - IAVA 2011-B-0025"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Malware Protection Engine Local Privilege Escalation Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26088 - IAVA 2011-A-0031"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Encdec.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17528"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Multiple Vulnerabilities in Microsoft Windows Media"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26091 - IAVA 2011-A-0033"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Mstsc.exe"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7600.16722"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Remote Desktop Connection Client Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26092 - IAVA 2011-B-0034"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Groove Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26509 - IAVA 2011-B-0045"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Fxscover.exe"
arrFile2(x) = "C:\Windows\System32\Mfc42.dll"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17559"
arrExpectedValue2(x) = "6.6.8064.0"
arrExpectedValue3(x) = "0"
arrType(x) = "DualFileVer"
arrTitle(x) = "Microsoft Windows Fax Cover Page Editor Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26512 - IAVA 2011-B-0046"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Remote Code Execution Vulnerability in Microsoft Foundation Class (MFC) Library"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26514 - IAVA 2011-A-0039"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\dnsapi.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17570"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft DNS Resolution Remote Code Execution Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-26521 - IAVA 2011-A-0050"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Drivers\Srv.sys"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17565"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft SMB Server Remote Code Execution Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-26525 - IAVA 2011-A-0047"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Microsoft Office PowerPoint"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26526 - IAVA 2011-A-0048"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\JScript.dll"
arrFile2(x) = "C:\Windows\System32\Vbscript.dll"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "5.8.7601.17562"
arrExpectedValue2(x) = "5.8.7601.17562"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Windows Scripting Memory Realloation Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26527 - IAVA 2011-A-0045"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Microsoft Office"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-27158 - IAVA 2011-A-0066"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in VMWare Products"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-27639 - IAVA 2011-B-0060"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Apache Portable Runtime Denial of Service Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-28308 - IAVA 2011-A-0072"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "IBM Tivoli Management Framework Remote Code Execution Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-28311 - IAVA 2011-A-0075"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in VMWare Products"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-28583 - IAVA 2011-A-0086"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Excel Remote Code Execution Vulnerabilities"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-28584 - IAVA 2011-A-0085"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Forefront Threat Management Gateway Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-28592 - IAVA 2011-A-0079"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Drivers\Mrxsmb.sys"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17605"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft SMB Client Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-28593 - IAVA 2011-A-0087"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Drivers\Dfsc.sys"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7600.16804"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Distributed File System Remote Code Execution Vulnerabilities"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-28594 - IAVA 2011-A-0082"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\Microsoft.NET\Framework\v2.0.50727\Mscorlib.dll"
arrFile2(x) = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\clr.dll"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "2.0.50727.4961"
arrExpectedValue2(x) = "4.0.30319.235"
arrExpectedValue3(x) = "0"
arrType(x) = "DualFileVer"
arrTitle(x) = "Microsoft.NET Framework Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-28597 - IAVA 2011-A-0081"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Oleaut32.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17567"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows OLE Automatic Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-28598 - IAVA 2011-A-0078"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Drivers\Srvnet.sys"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17608"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Server Message Block (SMB) Denial of Service Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-28599 - IAVA 2011-B-0069"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "HP OpenView Storage Data Protector Remote Code Execution Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-28601 - IAVA 2011-B-0064"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft XML Editor Information Disclosure Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-28611 - IAVA 2011-B-0067"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Drivers\Afd.sys"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17603"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows Ancillary Function Driver Privilege Escalation Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-28617 - IAVA 2011-B-0065"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Inetcomm.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17609"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft MHTML Information Disclosure Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-28950 - IAVA 2011-B-0071"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Adobe LiveCycle and BlazeDS"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-29343 - IAVA 2011-B-0075"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in HP OpenView Storage Data Protector"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-29384 - IAVA 2011-A-0100"
arrKeyPath(x) = "Default"
arrSubKey(x) = "Bthport.sys"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\winsxs\"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17607"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "WPFFileVer"
arrTitle(x) = "Microsoft Windows Bluetooth Stack Remote Code Execution Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-29387 - IAVA 2011-A-0098"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Visio Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-29527 - IAVA 2011-B-0084"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Citrix EdgeSight Remote Code Execution Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-29565 - IAVA 2011-B-0087"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Sybase Products"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-29567 - IAVA 2011-B-0091"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "HP Operations Manager Arbitrary File Deletion Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-29735 - IAVA 2011-A-0109"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Adobe Flash Media Server Memory Corruption Remote Denial of Service Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-29745 - IAVA 2011-B-0104"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Ntoskrnl.exe"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17640"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows Kernel Remote Denial of Service Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-29760 - IAVA 2011-B-0115"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Conhost.exe"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17641"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows Client/Server Run-time Subsystem Elevation of Privilege Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-29770 - IAVA 2011-B-0095"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Adobe RoboHelp Cross-Site Scripting Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-29776 - IAVA 2011-B-0097"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Microsoft Office Visio"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-29780 - IAVA 2011-B-0099"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Report Viewer Information Disclosure Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-29781 - IAVA 2011-B-0100"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft ASP.NET Chart Control Informaiton Disclosure Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-29783 - IAVA 2011-B-0101"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Program Files (x86)\Common Files\System\Ole DB\Msdaosp.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17632"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows Data Access Components Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-30070 - IAVA 2011-B-0111"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Citrix Access Gateway Cross-Site Scripting Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-30239 - IAVA 2011-B-0115"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Microsoft Office SharePoint"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-30245 - IAVA 2011-A-0124"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Microsoft Office Excel"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-30246 - IAVA 2011-A-0125"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Microsoft Office"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-30402 - IAVA 2011-B-0138"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Oleacc.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "7.0.0.0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Active Accessibility Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-30403 - IAVA 2011-B-0124"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Psisdecd.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.6.7601.17669"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows Media Center Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-30544 - IAVA 2011-A-0148"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "CiscoWorks Common Services Remote Code Execution Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-30588 - IAVA 2011-A-0149"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in VMware vCenter Server 4.1 and vCenter Update Manager 4.1"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-30608 - IAVA 2011-B-0135"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Program Files\Common Files\System\Wab32.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17699"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Windows Mail and Windows Meeting Space Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-30610 - IAVA 2011-B-0138"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerablities in HP OpenView Network Node Manager"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-30769 - IAVA 2011-A-0160"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in VMware vCenter Server 4.0 and vCenter Update Manager 4.0"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-30773 - IAVA 2011-B-0144"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Adobe Flex Cross Site Scripting Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-30822 - IAVA 2011-B-0146"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Office Input Method Editor (IME) Privilege Escalation Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-30826 - IAVA 2011-A-0171"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Encdec.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17708"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows Media Memory Corruption Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-30831 - IAVA 2011-A-0166"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Remote Code Execution Vulnerabilities in Microsoft Office PowerPoint"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-30834 - IAVA 2011-A-0163"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Office Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-30835 - IAVA 2011-A-0162"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Csrsrv.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17713"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows Client/Server Run-time Subsystem Elevation of Privilege Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-30836 - IAVA 2011-A-0161"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Active Directory Remote Code Execution Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-30840 - IAVA 2011-A-0175"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "RSA SecurID Software Token Insecure Library Loading Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-30927 - IAVA 2012-A-0001"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\ServiceModelReg.exe"
arrFile2(x) = "C:\Program Files\Reference Assemblies\Microsoft\Framework\v3.5\System.Web.Extensions.dll"
arrFile3(x) = "C:\Windows\Microsoft.NET\Framework\v2.0.50727\aspnet_wp.exe"
arrHive(x) = HKLM
arrExpectedValue(x) = "4.0.30319.272"
arrExpectedValue2(x) = "3.5.30729.4927"
arrExpectedValue3(x) = "2.0.50727.4971"
arrType(x) = "TriFileVer"
arrTitle(x) = "Multiple Vulnerabilities in Microsoft .NET Framework"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-30932 - IAVA 2012-A-0002"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\IME\IMEJP10\Imjpapi.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "10.1.7601.17658"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows Components Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-30998 - IAVA 2012-A-0003"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Ntdll.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17725"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows Kernel Security Bypass Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-31000 - IAVA 2012-A-0005"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Qdvd.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.6.7601.17713"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Multiple Remote Code Execution Vulnerabilities in Microsoft Windows Media Player"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-31005 - IAVA 2012-B-0005"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "HP Printer and Digital Senders Remote Firmware Update (RFU) Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-31007 - IAVA 2012-B-0003"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Information Disclosure Vulnerability in Microsoft AntiXSS Library"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-31010 - IAVA 2012-A-0007"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Packager.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17727"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows Remote Code Execution Vulnerability."
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-31054 - IAVA 2012-B-0006"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Schannel.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17725"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft SSL/TLS Information Disclosure Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-31249 - IAVA 2012-A-0019"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Symantec pcAnywhere"
arrCategory(x) = "I"


x = x + 1
arrFinding(x) = "V-31348 - IAVA 2012-A-0026"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Msvcrt.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17744"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows C Run-Time Library Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-31356 - IAVA 2012-B-0022"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Adobe RoboHelp XSS Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-31831 - IAVA 2012-B-0027"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "RSA SecureID Software Token Converter Buffer Overflow Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-31884 - IAVA 2012-A-0038"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Expression Design Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-31885 - IAVA 2012-A-0039"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Drivers\Rdpwd.sys"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17779"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Multiple Vulnerabilities in Microsoft Remote Desktop Protocol"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-31891 - IAVA 2012-A-0042"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Visual Studio Elevation of Privilege Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-31892 - IAVA 2012-B-0030"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Security Vulnerabilities in IBM DB2"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-31902 - IAVA 2012-A-0049"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in VMware View"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-31904 - IAVA 2012-B-0034"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "VMware vCenter Orchestrator Password Disclosure Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-31972 - IAVA 2012-B-0038"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in HP Onboard Administrator"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-31982 - IAVA 2012-A-0059"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Windows Common Controls Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-31983 - IAVA 2012-A-0060"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Wintrust.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17787"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-31984 - IAVA 2012-B-0041"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Office Works File Convertor Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-32042 - IAVA 2012-B-0046"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Security Vulnerabilities in Sourcefire Defense Center"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-32178 - IAVA 2012-B-0048"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in HP Systems Insight Manager"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-32304 - IAVA 2012-A-0079"
arrKeyPath(x) = "C:\Windows\winsxs\"
arrSubKey(x) = "gdiplus.dll"
arrSubKey2(x) = "6.1.7601.17825"
arrSubkey3(x) = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\WPF\PresentationCore.dll"
arrFile1(x) = "C:\Windows\System32\Dwrite.dll"
arrFile2(x) = "C:\Program Files\Windows Journal\JNTFiltr.dll"
arrFile3(x) = "C:\Windows\System32\Win32k.sys"
arrHive(x) = "4.0.30319.275"
arrExpectedValue(x) = "6.1.7601.17789"
arrExpectedValue2(x) = "6.1.7601.17803"
arrExpectedValue3(x) = "6.1.7601.17803"
arrType(x) = "OctFileVer"
arrTitle(x) = "Combined Security Update for Microsoft Office, Windows, .NET and Silverlight"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-32305 - IAVA 2012-A-0080"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\Microsoft.NET\Framework\v2.0.50727\Mscorlib.dll"
arrFile2(x) = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\clr.dll"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "2.0.50727.4971"
arrExpectedValue2(x) = "4.0.30319.269"
arrExpectedValue3(x) = "0"
arrType(x) = "DualFileVer"
arrTitle(x) = "Multiple Remote Code Execution Vulnerabilities in Microsoft .NET Framework"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-32308 - IAVA 2012-B-0049"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Adobe Illustrator"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-32311 - IAVA 2012-B-0052"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Drivers\Partmgr.sys"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17796"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows Partition Manager Privilege Escalation Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-32316 - IAVA 2012-A-0083"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Program Files (x86)\Microsoft\Office\Office14\Winword.exe"
arrFile2(x) = "C:\Program Files\Microsoft\Office\Office14\Winword.exe"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "12.0.6661.5000"
arrExpectedValue2(x) = "12.0.6661.5000"
arrExpectedValue3(x) = "0"
arrType(x) = "DualFileVer"
arrTitle(x) = "Microsoft Office Word Remote Code Execution Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-32826 - IAVA 2012-A-0092"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Drivers\Rdpwd.sys"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17830"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Remote Desktop Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-33046 - IAVA 2012-A-0104"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Cisco AnyConnect Secure Mobility Client"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-33048 - IAVA 2012-B-0063"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Security Vulnerabilities in IBM DB2"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-33309 - IAVA 2012-A-0110"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Shell32.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17859"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows Shell Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-33310 - IAVA 2012-A-0108"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Schannel.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17856"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft TLS Protocol Information Disclosure Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-33311 - IAVA 2012-A-0109"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Visual Basic for Applications Remote Code Execution Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-33313 - IAVA 2012-A-0107"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Program Files\Common Files\System\ado\msado15.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17857"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Data Access Components (MDAC) Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-33315 - IAVA 2012-B-0067"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Symantec LiveUpdate Privilege Escalation Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-33396 - IAVA 2012-A-0125"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Symantec System Recovery Arbitrary Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-33555 - IAVA 2012-B-0074"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Cross-Site Scritping Vulnerabilities in HP Newtork Node Manager"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-33555 - IAVA 2012-B-0074"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Cross-Site Scripting Vulnerabilities in HP Network Node Manager i(NNMi)"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-33652 - IAVA 2012-B-0075"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Office Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-33654 - IAVA 2012-A-0130"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\jscript.dll"
arrFile2(x) = "C:\Windows\System32\vbscript.dll"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "5.8.7601.17866"
arrExpectedValue2(x) = "5.8.7601.17866"
arrExpectedValue3(x) = "0"
arrType(x) = "DualFileVer"
arrTitle(x) = "Microsoft JScript and VBScript Engines Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-33657 - IAVA 2012-A-0137"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Localspl.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17841"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Multiple Vulnerabilities in Microsoft Windows Networking Components"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-33659 - IAVA 2012-A-0132"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Windows Common Controls Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-33661 - IAVA 2012-B-0077"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Citrix Access Gateway Plug-in for Windows"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-33670 - IAVA 2012-B-0080"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Cisco IP Communicator Bypass Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-33691 - IAVA 2012-A-0140"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "McAfee Smartfilter Administration Remote Code Execution Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-33692 - IAVA 2012-B-0083"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple SQL Injection Vulnerabilities in Ipswitch WhatsUp Gold"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-33777 - IAVA 2012-B-0085"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Adobe Photoshop"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-33786 - IAVA 2012-B-0089"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft System Center Configuration Manager Cross Site Scripting Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-33787 - IAVA 2012-B-0090"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Visual Studio Team Foundation Server Cross Site Scripting Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-33792 - IAVA 2012-B-0146"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in VMware vCenter Update Manager 4.1"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-33793 - IAVA 2012-B-0147"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in VMware vCenter Server 4.1"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-33808 - IAVA 2012-B-0094"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Citrix Receiver and Online Plug-in for Windows Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-33809 - IAVA 2012-B-0092"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "ISC DHCP Denial of Service Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-33969 - IAVA 2012-B-0095"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "IBM Tivoli Federated Identity Manager Validation Bypass Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-33972 - IAVA 2012-B-0097"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "IBM Rational Business Developer Information Disclosure Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-34012 - IAVA 2012-B-0100"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Adobe Products Code Signing Certificate Revocation"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-34013 - IAVA 2012-B-0099"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Adobe Enterprise Products Code Signing Certificate Revocation"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-34177 - IAVA 2012-A-0160"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft SQL Server Cross-Site Scripting Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-34180 - IAVA 2012-B-0103"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Kerberos.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17095"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Kerberos Denial Of Service Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-34182 - IAVA 2012-B-0102"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Works 9 Remote Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-34185 - IAVA 2012-B-0101"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "HP Network Node Manager i (NNMi) Information Disclosure Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-34340 - IAVA 2012-A-0177"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in EMC NetWorker Module"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-34341 - IAVA 2012-B-0105"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Cisco WebEx WRF Player"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-34512 - IAVA 2012-B-0106"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Computer Associates ARCserve Backup"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-34953 - IAVA 2012-B-0111"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\inetsrv\iisreg.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "7.5.7600.17855"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Multiple Vulnerabilities in Microsoft Internet Information Services (IIS)"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-34955 - IAVA 2012-A-0184"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "2000"
arrFile1(x) = "C:\Windows\Microsoft.NET\Framework\v2.0.50727\System.dll"
arrFile2(x) = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\clr.dll"
arrFile3(x) = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.dll"
arrHive(x) = HKLM
arrExpectedValue(x) = "2.0.50727.5466"
arrExpectedValue2(x) = "4.0.30319.296"
arrExpectedValue3(x) = "4.0.30319.18014"
arrType(x) = "NetTriFileVer"
arrTitle(x) = "Multiple Remote Code Execution Vulnerabilities in Microsoft .NET Framework"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-34956 - IAVA 2012-A-0185"
arrKeyPath(x) = "Default"
arrSubKey(x) = "Synceng.dll"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\system32\synceng.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17959"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Multiple Vulnerabilities in Microsoft Windows Shell"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-34959 - IAVA 2012-A-0188"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in VMWare Player"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-34960 - IAVA 2012-A-0187"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in VMware Workstation"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-35030 - IAVA 2012-A-0192"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ""
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Symantec Scan Engine Memory Corruption Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-35051 - IAVA 2012-B-0118"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in IBM DB2"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-35052 - IAVA 2012-B-0117"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Autonomy Keyview affecting Symantec Products"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-35487 - IAVA 2012-B-0124"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Dpnaddr.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17514"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows DirectPlay Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-35488 - IAVA 2012-A-0196"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Kernel32.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17135"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows File Handling Component Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-35492 - IAVA 2012-A-0194"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Office Word Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-35495 - IAVA 2012-A-0197"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "IBM Informix Dynamic Server Buffer Overflow Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-35496 - IAVA 2012-B-0125"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "HP Network Node Manager i Remote Unauthorized Access Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-36444 - IAVA 2013-A-0004"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Msxml3.dll"
arrFile2(x) = "C:\Windows\System32\Msxml6.dll"
arrFile3(x) = "0"
arrHive(x) = HKLMarrExpectedValue(x) = "8.110.7601.17988"
arrExpectedValue2(x) = "6.30.7601.17988"
arrExpectedValue3(x) = "0"
arrType(x) = "DualFileVer"
arrTitle(x) = "Multiple Vulnerabilities in Microsoft XML Core Services"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-36447 - IAVA 2013-B-0001"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\v3.5\System.data.services.dll"
arrFile2(x) = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.Data.Services.dll"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "3.5.30729.5451"
arrExpectedValue2(x) = "4.0.30319.297"
arrExpectedValue3(x) = "0"
arrType(x) = "DualFileVer"
arrTitle(x) = "Microsoft Open Data Protocol Denial of Service"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-36448 - IAVA 2013-B-0002"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft System Center Operations Manager Privilege Escalation Vulnerabilities"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-36449 - IAVA 2013-B-0004"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Symantec Enterprise Security Manager (ESM) Privilege Escalatin Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-36450 - IAVA 2012-B-0003"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\NCrypt.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.18007"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows Security Bypass Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-36453 - IAVA 2012-A-0006"
arrKeyPath(x) = "Default"
arrSubKey(x) = "x"
arrSubKey2(x) = "x"
arrSubkey3(x) = "x"
arrFile1(x) = "C:\Windows\Microsoft.NET\Framework\v2.0.50727\system.dll"
arrFile2(x) = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.Security.dll"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "2.0.50727.4927"
arrExpectedValue2(x) = "4.0.30319.1001"
arrExpectedValue3(x) = "0"
arrType(x) = "DualFileVer"
arrTitle(x) = "Multiple Vulnerabilities in Microsoft .NET Framework"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-36637 - IAVA 2013-A-0025"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Cross Site Scripting Vulnerabilities in Red Hat JBoss Enterprise Portal Platform"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-36644 - IAVA 2013-B-0008"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in IBM WebSphere Application Server"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-36785 - IAVA 2013-B-0011"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in IBM Tivoli Storage Manager Client"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-36789 - IAVA 2013-B-0012"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "VMware vCenter 4.1 Server and vSphere 4.1 Client Memory Corruption Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-36818 - IAVA 2013-A-0033"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "McAfee VirusScan Enterprise and Host Intrusion Prevention Privilege Escalation Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-36821 - IAVA 2013-A-0041"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Kernel32.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17965"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows Client/Server Run-time Subsystem (CSRSS) Privilege Escalation Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-36822 - IAVA 2013-A-0040"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "2000"
arrFile1(x) = "C:\Windows\Microsoft.NET\Framework\v2.0.50727\System.Windows.Forms.dll"
arrFile2(x) = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.Windows.Forms.dll"
arrFile3(x) = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.Windows.Forms.dll"
arrHive(x) = HKLM
arrExpectedValue(x) = "2.0.50727.5468"
arrExpectedValue2(x) = "4.0.30319.1002"
arrExpectedValue3(x) = "4.0.30319.18036"
arrType(x) = "NetTriFileVer"
arrTitle(x) = "Microsoft .NET Framework Privilege Escalation Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-37068 - IAVA 2013-A-0052"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "CKEditor Cross-Site Scripting Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-37404 - IAVA 2013-A-0063"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Drivers\Usb8023.sys"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.18076"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows Kernel-Mode Drivers Privilege Escalation Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-37405 - IAVA 2013-A-0064"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Program Files (x86)\Microsoft Silverlight\sllauncher.exe"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "5.1.20125.0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Silverlight Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-37411 - IAVA 2013-B-0027"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft OneNote Information Disclosure Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-37412 - IAVA 2013-B-0028"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Visio Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-37413 - IAVA 2013-B-0024"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple HP LaserJet Pro Printers Information Disclosure Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-37414 - IAVA 2013-B-0023"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Stunnel Remote Buffer Overflow Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-37451 - IAVA 2013-B-0030"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in NVIDIA Windows Display Driver"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-37602 - IAVA 2013-A-0072"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "Cert\s+Hash\(sha1\)\:\s+f9\s+2b\s+e5\s+26\s+6c\s+c0\s+5d\s+b2\s+dc\s+0d\s+c3\s+f2\s+dc\s+74\s+e0\s+2d\s+ef\s+d9\s+49\s+cb"
arrExpectedValue2(x) = "Cert\s+Hash\(sha1\)\:\s+c6\s+9f\s+28\s+c8\s+25\s+13\s+9e\s+65\s+a6\s+46\s+c4\s+34\s+ac\s+a5\s+a1\s+d2\s+00\s+29\s+5d\s+b1"
arrExpectedValue3(x) = "Cert\s+Hash\(sha1\)\:\s+4d\s+85\s+47\s+b7\s+f8\s+64\s+13\s+2a\s+7f\s+62\s+d9\s+b7\s+5b\s+06\s+85\s+21\s+f1\s+0b\s+68\s+e3"
arrType(x) = "TriUntrustedCertificates"
arrTitle(x) = "Microsoft WIndows Digital Security Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-37605 - IAVA 2013-A-0077"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulerabilities in OpenSSL"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-37608 - IAVA 2013-A-0081"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Active Directory Denial of Service Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-37609 - IAVA 2013-A-0080"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Ntoskrnl.exe"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.18113"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows Kernel Privilege Escalation Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-37613 - IAVA 2013-A-0083"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Office HTMl Sanitization Privilege Escalation Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-37618 - IAVA 2013-A-0082"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Mstscax.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.18079"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows Remote Desktop Client Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-37619 - IAVA 2013-B-0035"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in PostgreSQL"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-37763 - IAVA 2013-A-0098"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in McAfee ePolicy Orchestrator"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-37764 - IAVA 2013-B-0043"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "HP Multiple LaserJet Printers Information Disclosure Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-37765 - IAVA 2013-A-0096"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "HP Managed Printing Administration Cross Site Scripting Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-37766 - IAVA 2013-B-0041"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in VMware vCenter Server 5.1"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-37767 - IAVA 2013-A-0097"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Red Hat JBoss Enterprise Portal Platform"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-37937 - IAVA 2013-A-0107"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Microsoft Publisher Remote Code Execution Vulnerabilities"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-37938 - IAVA 2013-B-0051"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Lync Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-37941 - IAVA 2013-B-0052"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Visio Information Disclosure Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-37947 - IAVA 2013-B-0047"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Apache Tomcat"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-37949 - IAVA 2013-B-0054"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Wind River VxWorks"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-38289 - IAVA 2013-A-0110"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Malware Protection Engine Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-38435 - IAVA 2013-B-0059"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in IDA Pro"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-38436 - IAVA 2013-B-0058"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Apple QuickTime"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-38875 - IAVA 2013-B-0063"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Trimble SketchUp"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-39066 - IAVA 2013-A-0117"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "McAfee Agent ePO Extension SQL Injection Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-39072 - IAVA 2013-A-0120"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Win32spl.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.18142"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows Print Spooler Privilege Escalation Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-39073 - IAVA 2013-A-0121"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Office Remote Code Execution Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-39109 - IAVA 2013-A-0123"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Sybase EAServer"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-39134 - IAVA 2013-A-0127"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Symantec Endpoint Protection Remote Buffer Overflow Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-39199 - IAVA 2013-A-0135"
arrKeyPath(x) = "Default"
arrSubKey(x) = "Gdiplus.dll"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\winsxs\"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.18120"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "WPFFileVer"
arrTitle(x) = "Microsoft GDI+ Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-39200 - IAVA 2013-A-0134"
arrKeyPath(x) = "Default"
arrSubKey(x) = "Qedit.dll"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\winsxs\"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.18175"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "WPFFileVer"
arrTitle(x) = "Microsoft DirectShow Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-39210 - IAVA 2013-A-0137"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Windows Defender Privilege Escalation Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-39211 - IAVA 2013-B-0071"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "x"
arrFile1(x) = "C:\Windows\Microsoft.NET\Framework\v2.0.50727\System.configuration.dll"
arrFile2(x) = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.configuration.dll"
arrFile3(x) = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.configuration.dll"
arrHive(x) = HKLM
arrExpectedValue(x) = "2.0.50727.5476"
arrExpectedValue2(x) = "4.0.30319.1015"
arrExpectedValue3(x) = "4.0.30319.18060"
arrType(x) = "NetTriFileVer"
arrTitle(x) = "Multiple Remote Code Execution Vulnerabilities in Microsoft .NET Framework"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-39212 - IAVA 2013-B-0072"
arrKeyPath(x) = "Default"
arrSubKey(x) = "Wmvdecod.dll"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\system32\wmvdecod.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.18221"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows Media Format Runtime Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-39335 - IAVA 2013-B-0073"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "HP Network Node Manage i (NNMi) Unauthorized Access Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-39573 - IAVA 2013-A-0146"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Security Vulnerabilities in Apache HTTP Server"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-39815 - IAVA 2013-A-0150"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Remote Code Execution Vulnerabilities in HP SiteScope"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-39824 - IAVA 2013-B-0076"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Apache OpenOffice"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-39912 - IAVA 2013-B-0084"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Symantec Encryption Desktop Privilege Escalation Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-39913 - IAVA 2013-A-0156"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Symantec Backup Exec Arbitrary Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-39914 - IAVA 2013-B-0080"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "HP LaserJet Pro Printers Information Disclosure Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-40034 - IAVA 2013-A-0163"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Rpcrt4.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.18205"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows Remote Procedure Call (RPC) Elevation of Privilege Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-40035 - IAVA 2013-A-0161"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\drivers\Tcpip.sys"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.18203"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft ICMPv6 Denial of Service Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-40037 - IAVA 2013-A-0164"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Windows Unicode Scripts Processor Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-40045 - IAVA 2013-B-0088"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\kernel32.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.18015"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Multiple Privilege Escalation Vulnerabilities in Microsoft WindowS Kernel"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-40108 - IAVA 2013-B-0093"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in PHP"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-40109 - IAVA 2013-B-0090"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in DotNetNuke"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-40163 - IAVA 2013-A-0166"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Security Vulnerabilities in RealNetworks RealPlayer"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-40287 - IAVA 2013-B-0099"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Remote Code Execution Vulnerabilities in Microsoft Access"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-40288 - IAVA 2013-A-0177"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerablities in Red Hat JBoss Enterprise Application Platform"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-40289 - IAVA 2013-A-0178"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Microsoft Office"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-40295 - IAVA 2013-A-0171"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Remote Code Execution Vulnerabilities in Microsoft Excel"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-40296 - IAVA 2013-A-0169"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Adobe Reader and Acrobat"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-40300 - IAVA 2013-B-0103"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Csrsrv.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.18229"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows Service Control Manager Privilege Escalation Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-40301 - IAVA 2013-B-0102"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Office Input Method Editor (IME) Privilege Escalation Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-40302 - IAVA 2013-B-0101"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft FrontPage Information Disclosure Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-40303 - IAVA 2013-B-0100"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Active Directory Denial of Service Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-40374 - IAVA 2013-B-0106"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in WordPress"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-40488 - IAVA 2013-B-0108"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Multiple Hewlett Packard (HP) Network Devices"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-40489 - IAVA 2013-A-0183"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Apache Struts"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-40490 - IAVA 2013-B-0109"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "HP System Management Homepage (SMH) Command Injection Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-40710 - IAVA 2013-B-0110"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Splunk Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-40753 - IAVA 2013-A-0187"
arrKeyPath(x) = "Default"
arrSubKey(x) = "x"
arrSubKey2(x) = "x"
arrSubkey3(x) = "x"
arrFile1(x) = "C:\Windows\Microsoft.NET\Framework\v2.0.50727\System.security.dll"
arrFile2(x) = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.Security.dll"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "2.0.50727.5475"
arrExpectedValue2(x) = "4.0.30319.1016"
arrExpectedValue3(x) = "0"
arrType(x) = "DualFileVer"
arrTitle(x) = "Multiple Vulnerabilities in Microsoft .NET Framework"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-40754 - IAVA 2013-A-0186"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Adobe Reader and Acrobat Javascript Security Control Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-40756 - IAVA 2013-B-0115"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Office Word Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-40757 - IAVA 2013-B-0114"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Microsoft Office Excel"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-40758 - IAVA 2013-B-0111"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Adobe RoboHelp 10 Memory Corruption Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-40760 - IAVA 2013-A-0189"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Comctl32.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "5.82.7601.18021"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows Common Control Library Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-40761 - IAVA 2013-B-0113"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "McAfee Agent Denial of Service Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-40763 - IAVA 2013-A-0190"
arrKeyPath(x) = "Default"
arrSubKey(x) = "Usbser.sys"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\winsxs"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.18247"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Multiple Vulnerabilities in Microsoft Windows Kernel-Mode Drivers"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-40764 - IAVA 2013-B-0117"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Program Files (x86)\Microsoft Silverlight\sllauncher.exe"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "5.1.20913.0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Silverlight Information Disclosure Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-40781 - IAVA 2013-A-0195"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Oracle Sun Systems Product Suite"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-40782 - IAVA 2013-A-0201"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Oracle MySQL Products"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-40783 - IAVA 2013-A-0200"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Oracle Java"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-40784 - IAVA 2013-A-0197"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Oracle E-Business Suite Remote Code Execution Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-40785 - IAVA 2013-A-0198"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Oracle Enterprise Manager"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-40786 - IAVA 2013-A-0199"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Oracle Fusion Middleware"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-40787 - IAVA 2013-A-0196"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Oracle Database"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-40789 - IAVA 2013-B-0118"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "BlackBerry Enterprise Service (BES) Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-41065 - IAVA 2013-A-0202"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in VMware vCenter Server 5.0"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-41364 - IAVA 2013-B-0121"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "EMC Networker Information Disclosure Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-41645 - IAVA 2013-B-0123"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Wireshark Denial of Service Vulnerabilities"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-41646 - IAVA 2013-A-0207"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in HP LoadRunner"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-42291 - IAVA 2013-A-0208"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Adobe ColdFusion"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-42293 - IAVA 2013-A-0213"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "KB2900986"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "Hotfix"
arrTitle(x) = "Cumulative Security Update of Microsoft ActiveX Kill Bits"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-42294 - IAVA 2013-A-0214"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Gdi32.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.18275"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft GDI Memory Corruption Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-42295 - IAVA 2013-A-0216"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Office Outlook Information Disclosure Vulnerability"
arrCategory(x) = "II"


x = x + 1
arrFinding(x) = "V-42297 - IAVA 2013-A-0212"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "ISC BIND Security Bypass Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-42298 - IAVA 2013-A-0217"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Sybase Adaptive Server Enterprise"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-42302 - IAVA 2013-B-0127"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\drivers\Afd.sys"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.18272"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows Ancillary Function Driver Information Disclosure Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-42304 - IAVA 2013-B-0128"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Crypt32.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.18277"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows Digital Signature Denial of Service Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-42306 - IAVA 2013-B-0126"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Microsoft Office"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-42307 - IAVA 2013-B-0125"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Joomla! Multiple Cross Site Scripting Vulnerabilities"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-42381 - IAVA 2013-B-0132"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Google Chrome Memory Corruption Vulnerability"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-42429 - IAVA 2013-B-0133"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Drupal Core"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-42580 - IAVA 2013-A-0228"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\cscript.exe"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "5.8.7601.18283"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-42581 - IAVA 2013-A-0227"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Wmi.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.17787"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Microsoft Windows (WinVerifyTrust) Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-42582 - IAVA 2013-A-0232"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Win32k.sys"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1.7601.18300"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "Multiple Vulnerabilities in Microsoft Windows Kernel-Mode Drivers"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-42584 - IAVA 2013-A-0223"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Windows\System32\Urlmon.dll"
arrFile2(x) = "11.0.9600.16476"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "8.0.7601.18305"
arrExpectedValue2(x) = "9.0.8112.16526"
arrExpectedValue3(x) = "10.0.9200.16750"
arrType(x) = "IE11FileVer"
arrTitle(x) = "Cumulative Security Update for Microsoft Internet Explorer"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-42585 - IAVA 2013-B-0135"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Office Security Bypass Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-42589 - IAVA 2013-B-0134"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft Office Information Disclosure Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-42590 - IAVA 2013-A-0224"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft ASP.NET SignalR Privilege Escalation Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-42592 - IAVA 2013-A-0231"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Microsoft Exchange Server"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-42593 - IAVA 2013-A-0225"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Microsoft GDI Remote Code Execution Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-42594 - IAVA 2013-A-0230"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Adobe Shockwave Player"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-42595 - IAVA 2013-A-0229"
arrKeyPath(x) = "Software\Macromedia\FlashPlayer"
arrSubKey(x) = "CurrentVersion"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "11.7.700.260"
arrExpectedValue2(x) = "12.0.0.38"
arrExpectedValue3(x) = "0"
arrType(x) = "FlashDualStringLess"
arrTitle(x) = "Multiple Vulnerabilties in Adobe Flash Player"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-42596 - IAVA 2013-A-0233"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Mozilla Products"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-42597 - IAVA 2013-B-137"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Multiple Vulnerabilities in Google Chrome"
arrCategory(x) = "I"

WScript.Echo vbCrLf & "Performing " & x & " security checks on Windows 2008 R2 64bit" & vbCrLf

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set WshNetwork = CreateObject("Wscript.Network")

strComputer = wshNetwork.ComputerName

Set objReg = GetObject("winmgmts:{impersonationlevel=impersonate,(Security)}\\" & strComputer & "\root\default:StdRegProv")
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate,(Security)}\\" & strComputer & "\root\cimv2")
Set colDrives = objWMIService.ExecQuery("Select * from Win32_LogicalDisk")
Set colItems = objWMIService.ExecQuery("Select * from Win32_UserAccount Where LocalAccount=True")
Set colServiceList = objWMIService.ExecQuery("Select * from Win32_Service")
Set colRoles = objWMIService.ExecQuery ("Select * from Win32_ServerFeature")
Set colProducts = objWMIService.ExecQuery("Select * from Win32_Product")
Set colKB = objWMIService.ExecQuery("Select * from Win32_QuickFixEngineering")
Set colShares = objWMIService.ExecQuery("Select * from Win32_Share")
Set colPrinters = objWMIService.ExecQuery("Select * from Win32_Printer Where Shared = 1")
Set colOSItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
Set colComputerSystems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
Set objPolicy = CreateObject("HNetCfg.FwPolicy2")
Set colRules = objPolicy.Rules
For Each objItem in colOSItems
   ServicePack = objItem.ServicePackMajorVersion
Next
Set RegEx = New RegExp
Set RegEx2 = New RegExp
Set RegEx3 = New RegExp
Set wshShell = WScript.CreateObject("WScript.shell")
wshShell.run "secedit /analyze /db stig.db /cfg 2008R2.inf /log stig.txt",,True 
wshShell.run "cmd /c more stig.txt > secedit_results.txt",,True
wshShell.run "cmd /c auditpol /get /category:* > audit_results.txt",,True
wshShell.run "cmd /c certutil -store disallowed > untrusted_certificates_results.txt",,True
wshShell.run "cmd /c certutil -store > dod_certificates_results.txt",,True
wshShell.run "cmd /c certutil -store -enterprise >> dod_certificates_results.txt",,True
wshShell.run "cmd /c certutil -store CA >> dod_certificates_results.txt",,True
wshShell.run "cmd /c certutil -store Root >> dod_certificates_results.txt",,True

If objFSO.FileExists("dir_results.txt") Then
   objFSO.DeleteFile("dir_results.txt")
End If

if globalFind = "yes" Then
   For Each objDrive in colDrives
      If objDrive.DriveType = 3 Then
        strDrive = objDrive.DeviceID & "\"
        wshShell.run "cmd /c dir " & strDrive & " /s >> dir_results.txt",,True
      End If
   Next
End If

i = 0

Set objTextFile = objFSO.OpenTextFile("secedit_results.txt", 1)
Do Until objTextFile.AtEndOfStream
    Redim Preserve secLines(i)
    secLines(i) = objTextFile.ReadLine
    i = i + 1
Loop

objTextFile.Close

if globalFind = "yes" Then

   Set objTextFile = objFSO.OpenTextFile("dir_results.txt", 1)
   Set objOutFile = objFSO.CreateTextFile("pfx_results.txt", True)

   RegEx.Pattern = ".*\.pfx.*"
   RegEx2.Pattern = ".*\.p12.*" 
   RegEx.IgnoreCase = True
   RegEx2.IgnoreCase = True

   Do Until objTextFile.AtEndOfStream
       pfxLine = objTextFile.ReadLine
       if RegEx.Test(pfxLine) Or RegEx2.Test(pfxLine) Then
          objOutFile.WriteLine(pfxLine)
       End If
   Loop

   objTextFile.Close
   objOutFIle.Close

   Set objTextFile = objFSO.OpenTextFile("pfx_results.txt", 1)
   i = 0

   Do Until objTextFile.AtEndOfStream
      Redim Preserve pfxLines(i)
      pfxLines(i) = objTextFile.ReadLine
      i = i + 1
  Loop

   objTextFile.Close

End If

Set objTextFile = objFSO.OpenTextFile("audit_results.txt", 1)
i = 0

Do Until objTextFile.AtEndOfStream
    Redim Preserve auditLines(i)
    auditLines(i) = objTextFile.ReadLine
    i = i + 1
Loop

objTextFile.Close

Set objTextFile = objFSO.OpenTextFile("untrusted_certificates_results.txt", 1)
i = 0

Do Until objTextFile.AtEndOfStream
    Redim Preserve certLines(i)
    certLines(i) = objTextFile.ReadLine
    i = i + 1
Loop

objTextFile.Close

Set objTextFile = objFSO.OpenTextFile("dod_certificates_results.txt", 1)
i = 0

Do Until objTextFile.AtEndOfStream
    Redim Preserve dodLines(i)
    dodLines(i) = objTextFile.ReadLine
    i = i + 1
Loop

objTextFile.Close

objFSO.DeleteFile("stig.txt")

For intCount = 0 to x
   If (StrComp(arrType(intCount), "Value", vbTextCompare)) = 0 Then
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
      ElseIf CLng(arrValue(intCount)) = CLng(arrExpectedValue(intCount)) Then
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      Else
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If


   ElseIf (StrComp(arrType(intCount), "ValueGreater", vbTextCompare)) = 0 Then
      objReg.GetDWORDValue arrHive(intCount), arrKeyPath(intCount), arrSubkey(intCount), arrValue(intCount)
      If CLng(arrvalue(intCount)) >= CLng(arrExpectedValue(intCount)) Then
        finding = "closed" 
      Else
        finding = "open"
      End If
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      End If


     ElseIf (StrComp(arrType(intCount), "ValueLess", vbTextCompare)) = 0 Then
       objReg.GetDWORDValue arrHive(intCount), arrKeyPath(intCount), arrSubkey(intCount), arrValue(intCount)
       If CLng(arrvalue(intCount)) <= CLng(arrExpectedValue(intCount)) Then
         finding = "closed" 
      Else
         finding = "open"
      End If
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If


   ElseIf (StrComp(arrType(intCount), "ValueLessEx", vbTextCompare)) = 0 Then
      finding = "closed"
      boolResult = false
      objReg.EnumValues arrHive(intCount), arrKeyPath(intCount), arrKeyList
      For Each strSubkey In arrKeyList
         If (StrComp(arrSubkey(intCount),strSubkey,vbTextCompare)) = 0 Then
            boolResult = true
            Exit For
         End If
      Next
      objReg.GetDWORDValue arrHive(intCount), arrKeyPath(intCount), arrSubkey(intCount), arrValue(intCount)
      If boolResult = false then
         finding = "open"
      ElseIf CLng(arrvalue(intCount)) <= CLng(arrExpectedValue(intCount)) Then
         finding = "closed" 
      Else
         finding = "open"
      End If
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If

   ElseIf (StrComp(arrType(intCount), "ValueLessExZero", vbTextCompare)) = 0 Then
      finding = "open"
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
      If boolResult = false then
         finding = "open"
      ElseIf CLng(arrvalue(intCount)) = 0 Then
         finding = "open"   
      ElseIf CLng(arrvalue(intCount)) <= CLng(arrExpectedValue(intCount)) Then
         finding = "closed" 
      Else
         finding = "open"
      End If
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If


   ElseIf (StrComp(arrType(intCount), "ValueGreaterEx", vbTextCompare)) = 0 Then
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
      objReg.GetDWORDValue arrHive(intCount), arrKeyPath(intCount), arrSubkey(intCount), arrValue(intCount)
      If boolResult = false then
         finding = "open"
      ElseIf CLng(arrvalue(intCount)) >= CLng(arrExpectedValue(intCount)) Then
         finding = "closed" 
      Else
         finding = "open"
      End If
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If

    Elseif (StrComp(arrType(intCount), "Exists", vbTextCompare)) = 0 Then
      objReg.EnumValues arrHive(intCount), arrKeyPath(intCount), arrKeyList
      boolResult = false 
      If(isArray(arrKeyList)) Then
         For Each strSubkey In arrKeyList
            If (StrComp(arrSubkey(intCount),strSubkey,vbTextCompare)) = 0 Then
               boolResult = true
               Exit For
            End If
         Next
      End If
      If Not boolResult Then
        WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      Else 
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If

   Elseif (StrComp(arrType(intCount), "String", vbTextCompare)) = 0 Then
      objReg.GetStringValue arrHive(intCount), arrKeyPath(intCount), arrSubkey(intCount), arrValue(intCount)
      intCompare = StrComp(arrValue(intCount), arrExpectedValue(intCount), vbTextCompare)
      If intCompare = 0 Then
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      Else
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If

   Elseif (StrComp(arrType(intCount), "StringGreater", vbTextCompare)) = 0 Then
      strMessage = ""
      objReg.GetStringValue arrHive(intCount), arrKeyPath(intCount), arrSubkey(intCount), arrValue(intCount)
      If CLng(arrValue(intCount)) >= CLng(arrExpectedValue(intCount)) Then
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)  
      Else
	     strMessage = strMessage & vbCrLf & "    Registry key is version: " & arrValue(intCount) & " Expected: " & arrExpectedValue(intCount)
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) & strMessage
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


   Elseif (StrComp(arrType(intCount),"OS_Version", vbTextCompare)) = 0 Then
      objReg.GetStringValue arrHive(intCount), arrKeyPath(intCount), arrSubkey(intCount), arrValue(intCount)
      objReg.GetStringValue arrHive(intCount), arrKeyPath(intCount), arrSubkey2(intCount), arrValue2(intCount)
      intCompare = StrComp(arrValue(intCount), arrExpectedValue(intCount), vbTextCompare)
      intCompare2 = StrComp(arrValue2(intCount), arrExpectedValue2(intCount), vbTextCompare)
      If intCompare > -1 And intCompare2 > -1 Then
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      Else
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If
   

   Elseif(StrComp(arrType(intCount), "LogPerms", vbTextCompare)) = 0 Then
      finding = "closed"
      Set secFile1 = objWMIService.Get("Win32_LogicalFileSecuritySetting='" & arrFile1(intCount) & "'")
      Set secFile2 = objWMIService.Get("Win32_LogicalFileSecuritySetting='" & arrFile2(intCount) & "'") 
      Set secFile3 = objWMIService.Get("Win32_LogicalFileSecuritySetting='" & arrFile3(intCount) & "'")
      result = secFile1.GetSecurityDescriptor(objSD1)
      result = secFile2.GetSecurityDescriptor(objSD2)
      result = secFile3.GetSecurityDescriptor(objSD3)
      If finding = "closed" Then
         For Each objACE in objSD1.DACL
            If objACE.Trustee.Name = "Administrators" And (objACE.AccessMask XOr 1179817) = 0  And objACE.AceType = 0 Then
            ElseIf objACE.Trustee.Name = "eventlog" And (objACE.AccessMask XOr 2032127) = 0  And objACE.AceType = 0 Then
            ElseIf objACE.Trustee.Name = "SYSTEM" And (objACE.AccessMask XOr 2032127) = 0  And objACE.AceType = 0 Then
            ElseIf objACE.Trustee.Name = "Auditors" And (objACE.AccessMask XOr 2032127) = 0 And objACE.AceType = 0 Then
            Else
               finding = "open"
            End If 
         Next
      End If
      If finding = "closed" Then
         For Each objACE in objSD2.DACL
            If objACE.Trustee.Name = "Administrators" And (objACE.AccessMask XOr 1179817) = 0  And objACE.AceType = 0 Then
            ElseIf objACE.Trustee.Name = "eventlog" And (objACE.AccessMask XOr 2032127) = 0  And objACE.AceType = 0 Then
            ElseIf objACE.Trustee.Name = "SYSTEM" And (objACE.AccessMask XOr 2032127) = 0  And objACE.AceType = 0 Then
            ElseIf objACE.Trustee.Name = "Auditors" And (objACE.AccessMask XOr 2032127) = 0  And objACE.AceType = 0 Then
            Else
               finding = "open"
            End If 
         Next
      End If
      If finding = "closed" Then
         For Each objACE in objSD3.DACL
            If objACE.Trustee.Name = "Administrators" And (objACE.AccessMask XOr 1179817) = 0  And objACE.AceType = 0 Then
            ElseIf objACE.Trustee.Name = "eventlog" And (objACE.AccessMask XOr 2032127) = 0  And objACE.AceType = 0 Then
            ElseIf objACE.Trustee.Name = "SYSTEM" And (objACE.AccessMask XOr 2032127) = 0  And objACE.AceType = 0 Then
            ElseIf objACE.Trustee.Name = "Auditors" And (objACE.AccessMask XOr 2032127) = 0 And objACE.AceType = 0 Then
            Else
               finding = "open"
            End If 
         Next
      End If
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If
   

   ElseIf(StrComp(arrType(intCount), "VolAudit", vbTextCompare)) = 0 Then
      For Each objDrive in colDrives
         If objDrive.DriveType = 3 Then
            strDrive = objDrive.DeviceID & "\\"
            Set secFile1 = objWMIService.Get("Win32_LogicalFileSecuritySetting='" & strDrive & "'")
            result = secFile1.GetSecurityDescriptor(objSD1)
            finding = "open"
            If isArray(objSD1.SACL) Then
               For Each objACE in objSD1.SACL
                  If objACE.Trustee.Name = "Everyone" And (objACE.AccessMask XOr 983551) = 0 And objACE.AceType = 2 Then
                     finding = "closed"
                  End If
                  If finding = "closed" Then
                     Exit For
                  End If  
               Next
            Else
               finding = "open"
            End If
         End If
      Next
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If


   ElseIf(StrComp(arrType(intCount), "FileSystem", vbTextCompare)) = 0 Then
      For Each objDrive in colDrives
         If objDrive.DriveType = 3 Then
            finding = "closed"
            If objDrive.FileSystem <> "NTFS" Or isNull(objDrive.FileSystem) Then
               finding = "open"
            End If
         End If
         If finding = "open" Then
            Exit For
         End If
      Next
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If


   ElseIf(StrComp(arrType(intCount), "RegAudit", vbTextCompare)) = 0 Then
      finding = "open"
      objReg.GetSecurityDescriptor HKLM, arrSubkey(intCount), objSD1
      objReg.GetSecurityDescriptor HKLM, arrSubkey2(intCount), objSD2
      If Not isArray(objSD1.SACL) or Not isArray(objSD2.SACL) Then
      Else
         For Each objACE In objSD1.SACL
            If objACE.Trustee.Name = "Everyone" And (objACE.AccessMask XOr 983103) = 0 And objACE.AceType = 2 Then
               finding = "closed"
               Exit For
            End If
         Next
         If finding = "closed" Then
            finding = "open"
            For Each objACE In objSD2.SACL
              If objACE.Trustee.Name = "Everyone" And (objACE.AccessMask XOr 983103) = 0 And objACE.AceType = 2 Then
                 finding = "closed"
                 Exit For
              End If
            Next
         End If  
      End If
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If  


   ElseIf(StrComp(arrType(intCount), "RegStringRegExp", vbTextCompare)) = 0 Then
      objReg.GetStringValue arrHive(intCount), arrKeyPath(intCount), arrSubkey(intCount), arrValue(intCount)
      RegEx.Pattern = ".*USG.*"
      If RegEx.Test(arrValue(intCount)) Then
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      Else
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If


   ElseIf(StrComp(arrType(intCount), "SeceditRegExp", vbTextCompare)) = 0 Then
      finding = "closed"
      RegEx.Pattern = arrExpectedValue(intCount)  
      For i = 0 to UBound(secLines)
         if RegEx.Test(secLines(i)) Then
            finding = "open"
         End If
      Next
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If 


   ElseIf(StrComp(arrType(intCount), "StaleUsers", vbTextCompare)) = 0 Then
      finding = "closed"
      strMessage = ""
      For Each objItem in colItems
        dtmLastLogin = ""
        On Error Resume Next
        Set objUser = GetObject("WinNT://" & strComputer & "/" & objItem.Name & ",user")
        dtmLastLogin = objUser.lastLogin
        RegEx.Pattern = ".*S-1-5-21.*50[1|0]"
        RegEx2.Pattern = ".*IUSR.*"
        if RegEx.Test(objItem.SID) Then
        Elseif RegEx2.Test(objItem.Name) Then
        Elseif Abs(DateDiff("d", Now, dtmLastLogin)) > 35 Then
           finding="open"
           If isNull(Abs(DateDiff("d", Now, dtmLastLogin))) Then
              strMessage = strMessage & vbCRLF & "     User " & objItem.Name & " has never logged in" 
           Else
              strMessage = strMessage & vbCRLF & "     User " & objItem.Name & " has not logged in for " & Abs(DateDiff("d", Now, dtmLastLogin)) & " days" & vbCRLF
           End If
        End If
      Next
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) & strMessage
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If 


   ElseIf(StrComp(arrType(intCount), "AccountExpires", vbTextCompare)) = 0 Then
      finding = "closed"
      strMessage = ""
      For Each objItem in colItems
        On Error Resume Next
        Set objUser = GetObject("WinNT://" & strComputer & "/" & objItem.Name & ",user")
        dtmLastLogin = objUser.lastLogin
        RegEx.Pattern = ".*S-1-5-21.*500"
        RegEx2.Pattern = ".*IUSR.*"
        if RegEx.Test(objItem.SID) Then
        Elseif RegEx2.Test(objItem.Name) Then
        ElseIf cLng(objUser.userFlags) And PWDNEVEREXPIRES Then
           finding="open"
           strMessage = strMessage & vbCRLF & "    " & objUser.Name & "'s account is not set to expire"
        End If
      Next
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) & strMessage
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If 


   ElseIf(StrComp(arrType(intCount), "PasswordRequired", vbTextCompare)) = 0 Then
      finding = "closed"
      strMessage = ""
      For Each objItem in colItems
        On Error Resume Next
        Set objUser = GetObject("WinNT://" & strComputer & "/" & objItem.Name & ",user")
        dtmLastLogin = objUser.lastLogin
        RegEx.Pattern = ".*S-1-5-21.*500"
        RegEx2.Pattern = ".*IUSR.*"
        if RegEx.Test(objItem.SID) Then
        Elseif RegEx2.Test(objItem.Name) Then
        ElseIf cLng(objUser.userFlags) And PWDREQUIRED Then
           finding="open"
           strMessage = strMessage & vbCRLF & "    " & objUser.Name & "'s account does not require a password"
        End If
      Next
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)  & strMessage
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      End If 


   ElseIf(StrComp(arrType(intCount), "ApplicationPassword", vbTextCompare)) = 0 Then
      finding = "closed"
      strMessage = ""
      For Each objItem in colItems
        On Error Resume Next
        Set objUser = GetObject("WinNT://" & strComputer & "/" & objItem.Name & ",user")
        dtmLastLogin = objUser.lastLogin
        RegEx.Pattern = ".*S-1-5-21.*50[0|1]"
        RegEx2.Pattern = ".*IUSR.*"
        if RegEx.Test(objItem.SID) Then
        Elseif RegEx2.Test(objItem.Name) Then
        ElseIf cLng(objUser.PasswordAge) > 31536000 Then
           finding="open"
           strMessage = strMessage & vbCRLF & "    " & objUser.Name & "'s password is older than 1 year"
        End If
      Next
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)  & strMessage
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      End If


   ElseIf(StrComp(arrType(intCount), "RenameUser", vbTextCompare)) = 0 Then
      finding = "closed"
      RegEx.Pattern = ".*S-1-5-21.*50[1|0]"
      For Each objItem in colItems
         If RegEx.Test(objItem.SID) Then
           If objItem.Name = arrExpectedValue(intCount) Then
              finding = "open"
              Exit For
           End If
         End If
      Next
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      End If 


   Elseif (StrComp(arrType(intCount), "TriString", vbTextCompare)) = 0 Then
      finding = "closed"
      objReg.GetStringValue arrHive(intCount), arrKeyPath(intCount), arrSubkey(intCount), arrValue(intCount)
      objReg.GetStringValue arrHive(intCount), arrKeyPath(intCount), arrSubkey2(intCount), arrValue2(intCount)
      objReg.GetStringValue arrHive(intCount), arrKeyPath(intCount), arrSubkey3(intCount), arrValue3(intCount)
      intCompare = StrComp(arrValue(intCount), arrExpectedValue(intCount), vbTextCompare)
      If intCompare = 0 Then
      Else
         finding = "open"
      End If 
      intCompare = StrComp(arrValue2(intCount), arrExpectedValue2(intCount), vbTextCompare)
      If intCompare = 0 Then
      Else
         finding = "open"
      End If
      intCompare = StrComp(arrValue3(intCount), arrExpectedValue3(intCount), vbTextCompare)
      If intCompare = 0 Then
      Else
         finding = "open"
      End If

      If intCompare = 0 Then 
      Else
         finding = "open"
      End If
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      End If 


   ElseIf(StrComp(arrType(intCount), "PasswordFilter", vbTextCompare)) = 0 Then
      finding = "open"
      objReg.GetMultiStringValue arrHive(intCount), arrKeyPath(intCount), arrSubkey(intCount), arrMultiReg
      RegEx.Pattern = ".*EnPasFltv2x64.*"
      RegEx.IgnoreCase = True
      For i = 0 to Ubound(arrMultiReg)
         If RegEx.Test(arrMultiReg(i)) Then
            If objFSO.FileExists("C:\Windows\System32\EnPasFltv2x64.dll") Then
               finding = "closed"
               Exit For
            End If
         End If
      Next
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If


   ElseIf(StrComp(arrType(intCount), "WinReg", vbTextCompare)) = 0 Then
      finding = "open"
      objReg.GetSecurityDescriptor HKLM, arrSubkey(intCount), objSD1
      If Not isArray(objSD1.DACL) Then
      Else
         For Each objACE In objSD1.DACL
            If objACE.Trustee.Name <> "Administrators" And ObjACE.Trustee.Name <> "LOCAL SERVICE" And objACE.Trustee.Name <> "Backup Operators" Then
		Exit For
            End If
            If (objACE.Trustee.Name = "LOCAL SERVICE" And (objACE.AccessMask XOr 131097) = 0 And objACE.AceType = 0) Then
               finding = "closed"
               Exit For
            End If
            If (objACE.Trustee.Name = "LOCAL SERVICE" And (CLng(objACE.AccessMask) XOr -2147483648) = 0 And objACE.AceType = 0) Then
               finding = "closed"
               Exit For
            End If
         Next
         If finding = "closed" Then
            finding = "open"
            For Each objACE In objSD1.DACL
              If objACE.Trustee.Name = "Backup Operators" And (objAce.AccessMask Xor 131097) = 0 And objACE.AceType = 0 Then
                 finding = "closed"
                 Exit For
              End If
            Next
         End If  
      End If
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      End If


   Elseif (StrComp(arrType(intCount), "StringRange", vbTextCompare)) = 0 Then
      finding = "open"
      objReg.GetStringValue arrHive(intCount), arrKeyPath(intCount), arrSubkey(intCount), arrValue(intCount)
      If arrValue(intCount) >= arrExpectedValue(intCount) And arrValue(intCount) <= arrExpectedValue2(intCount) Then
         finding = "closed"
      End If
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If


   ElseIf(StrComp(arrType(intCount), "BackupOperators", vbTextCompare)) = 0 Then
      finding = "closed" 
      Set objGroup = GetObject("WinNT://./Backup Operators, group")
      For Each objUser in objGroup.Members
         If objUser.Name > "" Then
            finding = "open"
            Exit For
         End If
      Next
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      End If 


   ElseIf(StrComp(arrType(intCount), "RegMultiBlank", vbTextCompare)) = 0 Then
      finding = "open"
      objReg.GetMultiStringValue arrHive(intCount), arrKeyPath(intCount), arrSubkey(intCount), arrMultiReg
      If arrMultiReg(0) <= "" Then
         finding = "closed"
      End If
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If


   ElseIf(StrComp(arrType(intCount), "RegRemotePath", vbTextCompare)) = 0 Then
      finding = "open"
      objReg.GetMultiStringValue arrHive(intCount), arrKeyPath(intCount), arrSubkey(intCount), arrMultiReg
      If arrMultiReg(0) <= "" Then
         finding = "closed"
      End If
      If finding = "open" Then
         For i = 0 To Ubound(arrMultiReg)
            If arrMultiReg(i) = arrExpectedValue(intCount) Then
               finding = "closed"
            ElseIf arrMultiReg(i) = arrExpectedValue2(intCount) Then
               finding = "closed"
	    ElseIf arrMultiReg(i) = arrExpectedValue3(intCount) Then
               finding = "closed"
            Else
               finding = "open"
            End If
         Next
      End If
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If


   ElseIf(StrComp(arrType(intCount), "RemoteReg", vbTextCompare)) = 0 Then
      finding = "open"
      strMessage = ""
      objReg.GetMultiStringValue arrHive(intCount), arrKeyPath(intCount), arrSubkey(intCount), arrMultiReg
      If arrMultiReg(0) <= "" Then
         finding = "closed"
      End If
      If finding = "open" Then
         For i = 0 To Ubound(arrMultiReg)
            finding="open"
            For j = 0 To Ubound(arrRemoteReg)
               if StrComp(arrMultiReg(i), arrRemoteReg(j), vbTextCompare) = 0 Then
                  finding = "closed"
               End If
            Next
            If finding = "open" Then
               strMessage = strMessage & vbCRLF & "    " & arrMultiReg(i) & " does not match STIG policy"
            End If
         Next
      End If         
      If StrComp(strMessage, "", vbTextCompare) <> 0 Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)  & strMessage
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If


   ElseIf (StrComp(arrType(intCount), "ValueExReg", vbTextCompare)) = 0 Then
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
      RegEx.Pattern = arrExpectedValue(intCount)
      objReg.GetStringValue arrHive(intCount), arrKeyPath(intCount), arrSubkey(intCount), arrValue(intCount)
      If boolResult = false Then
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      ElseIf RegEx.Test(arrValue(intCount)) Then
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


   ElseIf (StrComp(arrType(intCount), "QuadValueEx", vbTextCompare)) = 0 Then
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
      objReg.GetDWORDValue arrHive(intCount), arrKeyPath(intCount), arrSubkey2(intCount), arrValue(intCount)
      If finding = "closed" Then
         If CLng(arrvalue(intCount)) = CLng(arrExpectedValue(intCount)) Then
            finding = "closed"
         Else 
            finding = "open"
         End If
      End If
      objReg.GetDWORDValue arrHive(intCount), arrKeyPath(intCount), arrSubkey3(intCount), arrValue(intCount)
      If finding = "closed" Then
         If CLng(arrvalue(intCount)) = CLng(arrExpectedValue(intCount)) Then
            finding = "closed"
         Else 
            finding = "open"
         End If
      End If
      objReg.GetDWORDValue arrHive(intCount), arrKeyPath(intCount), arrFile1(intCount), arrValue(intCount)
      If finding = "closed" Then
         If CLng(arrvalue(intCount)) = CLng(arrExpectedValue(intCount)) Then
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

   ElseIf (StrComp(arrType(intCount), "QuintValueEx", vbTextCompare)) = 0 Then
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
         For Each strSubkey In arrKeyList
            If (StrComp(arrSubkey2(intCount),strSubkey,vbTextCompare)) = 0 Then
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
      objReg.GetDWORDValue arrHive(intCount), arrKeyPath(intCount), arrSubkey2(intCount), arrValue(intCount)
      If finding = "closed" Then
         If CLng(arrvalue(intCount)) = CLng(arrExpectedValue(intCount)) Then
            finding = "closed"
         Else 
            finding = "open"
         End If
      End If
      objReg.GetDWORDValue arrHive(intCount), arrKeyPath(intCount), arrSubkey3(intCount), arrValue(intCount)
      If finding = "closed" Then
         If CLng(arrvalue(intCount)) = CLng(arrExpectedValue(intCount)) Then
            finding = "closed"
         Else 
            finding = "open"
         End If
      End If
      objReg.GetDWORDValue arrHive(intCount), arrKeyPath(intCount), arrFile1(intCount), arrValue(intCount)
      If finding = "closed" Then
         If CLng(arrvalue(intCount)) = CLng(arrExpectedValue(intCount)) Then
            finding = "closed"
         Else 
            finding = "open"
         End If
      End If
      objReg.GetDWORDValue arrHive(intCount), arrKeyPath(intCount), arrFile2(intCount), arrValue(intCount)
      If finding = "closed" Then
         If CLng(arrvalue(intCount)) = CLng(arrExpectedValue(intCount)) Then
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


   ElseIf (StrComp(arrType(intCount), "BinaryEx", vbTextCompare)) = 0 Then
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
      objReg.GetBinaryValue arrHive(intCount), arrKeyPath(intCount), arrSubkey(intCount), arrBinary
      If boolResult = false Then
         WScript.Echo arrFinding(intCount) & " is Open" 
      ElseIf arrBinary(arrExpectedValue2(intCount)) = CInt(arrExpectedValue(intCount)) Then
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)  
      Else
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If

     
   ElseIf (StrComp(arrType(intCount), "StringValueExLess", vbTextCompare)) = 0 Then
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
      If boolResult = false Then
         WScript.Echo arrFinding(intCount) & " is Open -- " & arrTitle(intCount) 
      ElseIf isNull(arrValue(intCount)) Or arrValue(intCount) = "" Then
         WScript.Echo arrFinding(intCount) & " is Open -- " & arrTitle(intCount)
      ElseIf CLng(arrvalue(intCount)) <= CLng(arrExpectedValue(intCount)) Then
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)  
      Else
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If


   ElseIf (StrComp(arrType(intCount), "ServiceState", vbTextCompare)) = 0 Then
        finding = "closed"
        strMessage = ""
      	For Each objService in colServiceList
           For i = 0 to Ubound(arrServices, 1)
              If LCase(objService.Name) = LCase(arrServices(i, 0)) Then
                 If LCase(objService.StartMode) = LCase(arrServices(i,1)) Then
                 Else
                    strMessage = strMessage & vbCRLF & "    " & objService.Name & " starts as " & objService.StartMode & " and should be " & arrServices(i,1)
                    finding = "open"
                 End If
              End If
           Next
	Next
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) & strMessage
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      End If


   ElseIf (StrComp(arrType(intCount), "HbssClient", vbTextCompare)) = 0 Then
        finding = "open"
      	For Each objService in colServiceList
           if StrComp(objService.Name, arrExpectedValue(intCount),vbTextCompare) = 0 Then
              finding = "closed"
              Exit For
           End If
	Next
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If


   ElseIf(StrComp(arrType(intCount), "DirPFX", vbTextCompare)) = 0 Then
      finding = "closed"    
      RegEx.Pattern = arrExpectedValue(intCount)
      RegEx2.Pattern = arrExpectedValue2(intCount)
      RegEx.IgnoreCase = True
      RegEx2.IgnoreCase = True

      If pfxLines(0) <= "" Or isNull(pfxLines) Or isEmpty(pfxLines) Then
      Else

      For i = 0 to Ubound(pfxLines)
         If RegEx.Test(pfxLines(i)) Or RegEx2.Test(pfxLines(i)) Then
            finding = "open"
            Exit For
         End If
      Next

      End If
      
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
         WScript.Echo "    See pfx_results.txt for a list of files found"
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If

 
   ElseIf(StrComp(arrType(intCount), "WinLogon", vbTextCompare)) = 0 Then
      finding = "closed"
      objReg.GetSecurityDescriptor HKLM, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", objSD1
      For Each objACE In objSD1.DACL
         If objACE.Trustee.Name = "Users" Or objACE.Trustee.Name = "Authenticated Users" Or objACE.Trustee.Name = "Everyone" Then
            If (objACE.AccessMask XOr 131097) = 0 Then
            ElseIf (objACE.AccessMask XOr -2147483648) = 0 Then
            Else
               finding = "open"
               Exit For
            End If
         End If
      Next
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      End If

   ElseIf(StrComp(arrType(intCount), "DualRegPerms", vbTextCompare)) = 0 Then
      finding = "closed"
      objReg.GetSecurityDescriptor HKLM, arrKeyPath(x), objSD1
      For Each objACE In objSD1.DACL
         If objACE.Trustee.Name = "Users" Or objACE.Trustee.Name = "Authenticated Users" Or objACE.Trustee.Name = "Everyone" Then
            If (objACE.AccessMask XOr 131097) = 0 Then
            ElseIf (objACE.AccessMask XOr -2147483648) = 0 Then
            Else
               finding = "open"
               Exit For
            End If
         ElseIf objACE.Trustee.Name = "Users" Or objACE.Trustee.Name = "Authenticated Users" Then
            If (objACE.AccessMask XOr 131097) = 0 Then
            ElseIf (objACE.AccessMask XOr -2147483648) = 0 Then
            Else
               finding = "open"
               Exit For
            End If
         ElseIf objACE.Trustee.Name = "Administrators" Or objACE.Trustee.Name = "SYSTEM" Or objACE.Trustee.Name = "CREATOR OWNER" Then
            If (objACE.AccessMask XOr 983103) = 0 Then
            ElseIf (objACE.AccessMask XOr 268435456) = 0 Then
            Else
               finding = "open"
               Exit For
            End If
         Else
            finding = "open"  
         End If
      Next
      If finding = "closed" Then
         objReg.GetSecurityDescriptor HKLM, arrSubKey(x), objSD1
         For Each objACE In objSD1.DACL
            If objACE.Trustee.Name = "Users" Or objACE.Trustee.Name = "Authenticated Users" Or objACE.Trustee.Name = "Everyone" Then
               If (objACE.AccessMask XOr 131097) = 0 Then
               ElseIf (objACE.AccessMask XOr -2147483648) = 0 Then
               Else
                  finding = "open"
                  Exit For
               End If
            ElseIf objACE.Trustee.Name = "Users" Or objACE.Trustee.Name = "Authenticated Users" Then
               If (objACE.AccessMask XOr 131097) = 0 Then
               ElseIf (objACE.AccessMask XOr -2147483648) = 0 Then
               Else
                  finding = "open"
                  Exit For
               End If
            ElseIf objACE.Trustee.Name = "Administrators" Or objACE.Trustee.Name = "SYSTEM" Or objACE.Trustee.Name = "CREATOR OWNER" Then
               If (objACE.AccessMask XOr 983103) = 0 Then
               ElseIf (objACE.AccessMask XOr 268435456) = 0 Then
               Else
                  finding = "open"
                  Exit For
               End If
            Else
               finding = "open"  
            End If
         Next
      End If
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      End If


   ElseIf(StrComp(arrType(intCount), "AuditRegExp", vbTextCompare)) = 0 Then
      finding = "open"
      RegEx.Pattern = arrExpectedValue(intCount) 
      RegEx.IgnoreCase = True 
      For i = 0 to UBound(auditLines)
         if RegEx.Test(auditLines(i)) Then
            finding = "closed"
         End If
      Next
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If 


   ElseIf (StrComp(arrType(intCount), "ServiceDisabled", vbTextCompare)) = 0 Then
        finding = "closed"
        RegEx.Pattern = arrExpectedValue(intCount)
        RegEx.IgnoreCase = True
      	For Each objService in colServiceList
           If RegEx.Test(objService.Name) Then
              If LCase(objService.StartMode) <> LCase(arrExpectedValue2(intCount)) Then
                 strMessage = vbCrLf & "    " & objService.Name & " " & objService.StartMode
                 finding = "open"
                 Exit For
              End If
           End If
	Next
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)  & strMessage
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If
 

   ElseIf (StrComp(arrType(intCount), "FileSharePerms", vbTextCompare)) = 0 Then
      finding = "closed"
      strMessage = ""
      Set colShares = objWMIService.ExecQuery("SELECT * FROM Win32_LogicalShareSecuritySetting")
      For Each objShare in colShares
        i = objShare.GetSecurityDescriptor(objSD1)
        For Each objACE in objSD1.DACL
           If objACE.Trustee.Name = arrExpectedValue(intCount) Then
              finding = "open"
              strMessage = strMessage & vbCRLF & "    " & objACE.Trustee.Name & " has access to share " & objShare.Name
              Exit For
           End If
        Next
      Next
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)  & strMessage
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If

  
   ElseIf (StrComp(arrType(intCount), "FileVer", vbTextCompare)) = 0 Then
      finding = "closed"
      If objFSO.FileExists(arrFile1(intCount)) Then
         fileVer = objFSO.GetFileVersion(arrFile1(intCount))
         testValue1 = split(arrExpectedValue(intCount), ".")
         actualValue = split(fileVer, ".")
         If UBound(testValue1) >= UBound(actualValue) Then
            h = UBound(actualValue)
         Else
            h = UBound(testValue1)
         End If
         for i = 0 to h
            If actualValue(i) < testValue1(i) Then
               finding = "open"
               testValue2 = split(arrFile1(intCount), "\")
               strMessage = vbCrLf & "    " & testValue2(UBound(testValue2)) & " " & _
                            "is version: " & fileVer & " Expected: " & arrExpectedValue(intCount)
               Exit For
            ElseIf actualValue(i) > testValue1(i) Then
               Exit For
            End If
         Next
      End If
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) & strMessage
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If


   ElseIf (StrComp(arrType(intCount), "DualFileVer", vbTextCompare)) = 0 Then
      finding = "closed"
      strMessage = ""
      If objFSO.FileExists(arrFile1(intCount)) Then
         fileVer = objFSO.GetFileVersion(arrFile1(intCount))
         testValue1 = split(arrExpectedValue(intCount), ".")
         actualValue = split(fileVer, ".")
         
         If UBound(testValue1) >= UBound(actualValue) Then
            h = UBound(actualValue)
         Else
            h = UBound(testValue1)
         End If
         for i = 0 to h
            If actualValue(i) < testValue1(i) Then
               finding = "open"
               testValue1 = split(arrFile1(intCount), "\")
               strMessage = vbCrLf & "    " & testValue1(UBound(testValue1)) & " " & _
                            "is version: " & fileVer & " Expected: " & arrExpectedValue(intCount)
               Exit For
            Elseif actualValue(i) > testValue1(i) Then
               Exit For
            End If
         Next
      End If
      If finding = "closed" Then
         actualValue = ""
         If objFSO.FileExists(arrFile2(intCount)) Then
            fileVer = objFSO.GetFileVersion(arrFile2(intCount))
            testValue2 = split(arrExpectedValue2(intCount), ".")
            actualValue = split(fileVer, ".")
            If CLng(UBound(testValue2)) >= CLng(UBound(actualValue)) Then
               h = UBound(actualValue)
            Else
               h = UBound(testValue2)
            End If
            for i = 0 to h
               If CLng(actualValue(i)) < CLng(testValue2(i)) Then
                  finding = "open"
                  testValue2 = split(arrFile2(intCount), "\")
                  strMessage = strMessage & vbCrLf & "    " & testValue2(UBound(testValue2)) & " " & _
                              "is version: " & fileVer & " Expected: " & arrExpectedValue2(intCount)
                  Exit For
               Elseif CLng(actualValue(i)) > CLng(testValue2(i)) Then
                  Exit For
               End If
            Next
         End If
      End If
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) & strMessage 
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If

   ElseIf (StrComp(arrType(intCount), "TriFileVer", vbTextCompare)) = 0 Then
      finding = "closed"
      strMessage = ""
      If objFSO.FileExists(arrFile1(intCount)) Then
         fileVer = objFSO.GetFileVersion(arrFile1(intCount))
         testValue1 = split(arrExpectedValue(intCount), ".")
         actualValue = split(fileVer, ".")
         
         If UBound(testValue1) >= UBound(actualValue) Then
            h = UBound(actualValue)
         Else
            h = UBound(testValue1)
         End If
         for i = 0 to h
            If CLng(actualValue(i)) < CLng(testValue1(i)) Then
               finding = "open"
               testValue1 = split(arrFile1(intCount), "\")
               strMessage = vbCrLf & "    " & testValue1(UBound(testValue1)) & " " & _
                            "is version: " & fileVer & " Expected: " & arrExpectedValue(intCount)
               Exit For
            Elseif actualValue(i) > testValue1(i) Then
               Exit For
            End If
         Next
      End If
      If finding = "closed" Then
         actualValue = ""
         If objFSO.FileExists(arrFile2(intCount)) Then
            fileVer = objFSO.GetFileVersion(arrFile2(intCount))
            testValue2 = split(arrExpectedValue2(intCount), ".")
            actualValue = split(fileVer, ".")
            If UBound(testValue2) >= UBound(actualValue) Then
               h = UBound(actualValue)
            Else
               h = UBound(testValue2)
            End If
            for i = 0 to h
               If CLng(actualValue(i)) < CLng(testValue2(i)) Then
                  finding = "open"
                  testValue2 = split(arrFile2(intCount), "\")
                  strMessage = strMessage & vbCrLf & "    " & testValue2(UBound(testValue2)) & " " & _
                              "is version: " & fileVer & " Expected: " & arrExpectedValue2(intCount)
                  Exit For
               Elseif CLng(actualValue(i)) > CLng(testValue2(i)) Then
                  Exit For
               End If
            Next
         End If
      End If
      If finding = "closed" Then
         actualValue = ""
         If objFSO.FileExists(arrFile3(intCount)) Then
            fileVer = objFSO.GetFileVersion(arrFile3(intCount))
            testValue3 = split(arrExpectedValue3(intCount), ".")
            actualValue = split(fileVer, ".")
            If UBound(testValue3) >= UBound(actualValue) Then
               h = UBound(actualValue)
            Else
               h = UBound(testValue3)
            End If
            for i = 0 to h
               If CLng(actualValue(i)) < CLng(testValue3(i)) Then
                  finding = "open"
                  testValue3 = split(arrFile3(intCount), "\")
                  strMessage = strMessage & vbCrLf & "    " & testValue3(UBound(testValue3)) & " " & _
                              "is version: " & fileVer & " Expected: " & arrExpectedValue3(intCount)
                  Exit For
               Elseif CLng(actualValue(i)) > CLng(testValue3(i)) Then
                  Exit For
               End If
            Next
         End If
      End If
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) & strMessage 
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If

   ElseIf (StrComp(arrType(intCount), "NetTriFileVer", vbTextCompare)) = 0 Then
      finding = "closed"
      strMessage = ""
      If objFSO.FileExists(arrFile1(intCount)) Then
         fileVer = objFSO.GetFileVersion(arrFile1(intCount))
         testValue1 = split(arrExpectedValue(intCount), ".")
         actualValue = split(fileVer, ".")
         
         If UBound(testValue1) >= UBound(actualValue) Then
            h = UBound(actualValue)
         Else
            h = UBound(testValue1)
         End If
         for i = 0 to h
            If CLng(actualValue(i)) < CLng(testValue1(i)) Then
               finding = "open"
               testValue1 = split(arrFile1(intCount), "\")
               strMessage = vbCrLf & "    " & testValue1(UBound(testValue1)) & " " & _
                            "is version: " & fileVer & " Expected: " & arrExpectedValue(intCount)
               Exit For
            Elseif actualValue(i) > testValue1(i) Then
               Exit For
            End If
         Next
      End If
      If finding = "closed" Then
         actualValue = ""
         If objFSO.FileExists(arrFile2(intCount)) Then
            fileVer = objFSO.GetFileVersion(arrFile2(intCount))
            testValue2 = split(arrExpectedValue2(intCount), ".")
            actualValue = split(fileVer, ".")
            If UBound(testValue2) >= UBound(actualValue) Then
               h = UBound(actualValue)
            Else
               h = UBound(testValue2)
            End If
            for i = 0 to h
               If CLng(actualValue(i)) < CLng(testValue2(i)) Then
                  finding = "open"
                  testValue2 = split(arrFile2(intCount), "\")
                  strMessage = strMessage & vbCrLf & "    " & testValue2(UBound(testValue2)) & " " & _
                              "is version: " & fileVer & " Expected: " & arrExpectedValue2(intCount)
                  Exit For
               Elseif CLng(actualValue(i)) > CLng(testValue2(i)) Then
                  Exit For
               End If
            Next
         End If
      End If
      If finding = "closed" Then
         actualValue = ""
         If objFSO.FileExists(arrFile3(intCount)) Then
            fileVer = objFSO.GetFileVersion(arrFile3(intCount))
            testValue3 = split(arrExpectedValue3(intCount), ".")
            actualValue = split(fileVer, ".")
            If UBound(testValue3) >= UBound(actualValue) Then
               h = UBound(actualValue)
            Else
               h = UBound(testValue3)
            End If
            for i = 0 to h
               If CLng(actualValue(UBound(actualValue))) < CLng(arrSubkey3(intCount)) Then
                   finding = "closed"
                   Exit For
               End If
               If CLng(actualValue(i)) < CLng(testValue3(i)) Then
                  finding = "open"
                  testValue3 = split(arrFile3(intCount), "\")
                  strMessage = strMessage & vbCrLf & "    " & testValue3(UBound(testValue3)) & " " & _
                              "is version: " & fileVer & " Expected: " & arrExpectedValue3(intCount)
                  Exit For
               Elseif CLng(actualValue(i)) > CLng(testValue3(i)) Then
                  Exit For
               End If
            Next
         End If
      End If
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) & strMessage 
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If

   ElseIf (StrComp(arrType(intCount), "NetQuadFileVer", vbTextCompare)) = 0 Then
      finding = "closed"
      strMessage = ""
      If objFSO.FileExists(arrFile1(intCount)) Then
         fileVer = objFSO.GetFileVersion(arrFile1(intCount))
         testValue1 = split(arrExpectedValue(intCount), ".")
         actualValue = split(fileVer, ".")
         
         If UBound(testValue1) >= UBound(actualValue) Then
            h = UBound(actualValue)
         Else
            h = UBound(testValue1)
         End If
         for i = 0 to h
            If CLng(actualValue(i)) < CLng(testValue1(i)) Then
               finding = "open"
               testValue1 = split(arrFile1(intCount), "\")
               strMessage = vbCrLf & "    " & testValue1(UBound(testValue1)) & " " & _
                            "is version: " & fileVer & " Expected: " & arrExpectedValue(intCount)
               Exit For
            Elseif actualValue(i) > testValue1(i) Then
               Exit For
            End If
         Next
      End If
      If finding = "closed" Then
         actualValue = ""
         If objFSO.FileExists(arrFile2(intCount)) Then
            fileVer = objFSO.GetFileVersion(arrFile2(intCount))
            testValue2 = split(arrExpectedValue2(intCount), ".")
            actualValue = split(fileVer, ".")
            If UBound(testValue2) >= UBound(actualValue) Then
               h = UBound(actualValue)
            Else
               h = UBound(testValue2)
            End If
            for i = 0 to h
               If CLng(actualValue(i)) < CLng(testValue2(i)) Then
                  finding = "open"
                  testValue2 = split(arrFile2(intCount), "\")
                  strMessage = strMessage & vbCrLf & "    " & testValue2(UBound(testValue2)) & " " & _
                              "is version: " & fileVer & " Expected: " & arrExpectedValue2(intCount)
                  Exit For
               Elseif CLng(actualValue(i)) > CLng(testValue2(i)) Then
                  Exit For
               End If
            Next
         End If
      End If
      If finding = "closed" Then
         actualValue = ""
         If objFSO.FileExists(arrFile3(intCount)) Then
            fileVer = objFSO.GetFileVersion(arrFile3(intCount))
            testValue3 = split(arrExpectedValue3(intCount), ".")
            actualValue = split(fileVer, ".")
            If UBound(testValue3) >= UBound(actualValue) Then
               h = UBound(actualValue)
            Else
               h = UBound(testValue3)
            End If
            for i = 0 to h
               If CLng(actualValue(i)) < CLng(testValue3(i)) Then
                  finding = "open"
                  testValue3 = split(arrFile3(intCount), "\")
                  strMessage = strMessage & vbCrLf & "    " & testValue3(UBound(testValue3)) & " " & _
                              "is version: " & fileVer & " Expected: " & arrExpectedValue3(intCount)
                  Exit For
               Elseif CLng(actualValue(i)) > CLng(testValue3(i)) Then
                  Exit For
               End If
            Next
         End If
      End If
      If finding = "closed" Then
         actualValue = ""
         If objFSO.FileExists(arrKeyPath(intCount)) Then
            fileVer = objFSO.GetFileVersion(arrKeyPath(intCount))
            testValue3 = split(arrSubKey(intCount), ".")
            actualValue = split(fileVer, ".")
            If UBound(testValue3) >= UBound(actualValue) Then
               h = UBound(actualValue)
            Else
               h = UBound(testValue3)
            End If
            for i = 0 to h
               If CLng(actualValue(UBound(actualValue))) < CLng(arrSubkey3(intCount)) Then
                  finding = "closed"
                  Exit For
               End If
               If CLng(actualValue(i)) < CLng(testValue3(i)) Then
                  finding = "open"
                  testValue3 = split(arrKeyPath(intCount), "\")
                  strMessage = strMessage & vbCrLf & "    " & testValue3(UBound(testValue3)) & " " & _
                              "is version: " & fileVer & " Expected: " & arrSubkey(intCount)
                  Exit For
               Elseif CLng(actualValue(i)) > CLng(testValue3(i)) Then
                  Exit For
               End If
            Next
         End If
      End If
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) & strMessage 
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If

   ElseIf (StrComp(arrType(intCount), "OctFileVer", vbTextCompare)) = 0 Then
      finding = "open"
      strMessage = ""
      Set objFolder = objFSO.GetFolder(arrKeyPath(intCount))
      Set colSubFolders = objFolder.SubFolders
      For Each i in colSubFolders
         If objFSO.FileExists(arrKeyPath(intCount) & i.name & "/" & arrSubKey(intCount)) Then
            fileVer = objFSO.GetFileVersion(arrKeyPath(intCount) & i.name & "/" & arrSubKey(intCount))
            testValue1 = split(arrSubKey2(intCount), ".")
            actualValue = split(fileVer, ".")
         
            If UBound(testValue1) >= UBound(actualValue) Then
               h = UBound(actualValue)
            Else
               h = UBound(testValue1)
            End If
            for j = 0 to h
               If CLng(actualValue(j)) < CLng(testValue1(j)) Then
                  finding = "open"
                  testValue2 = split(arrFile1(intCount), "\")
                  strMessage = vbCrLf & "    " & arrSubKey(intCount) & " " & _
                            "is version: " & fileVer & " Expected: " & arrSubKey2(intCount)
               Elseif CLng(actualValue(j)) >= CLng(testValue1(j)) Then
                  finding = "closed"
               End If
            Next
            If finding = "closed" Then
               Exit For
            End If 
         End If
      Next
      If objFSO.FileExists(arrFile1(intCount)) Then
         fileVer = objFSO.GetFileVersion(arrFile1(intCount))
         testValue1 = split(arrExpectedValue(intCount), ".")
         actualValue = split(fileVer, ".")
        
         If UBound(testValue1) >= UBound(actualValue) Then
            h = UBound(actualValue)
         Else
            h = UBound(testValue1)
         End If
         for i = 0 to h
            If CLng(actualValue(i)) < CLng(testValue1(i)) Then
               finding = "open"
               testValue1 = split(arrFile1(intCount), "\")
               strMessage = strMessage & vbCrLf & "    " & testValue1(UBound(testValue1)) & " " & _
                            "is version: " & fileVer & " Expected: " & arrExpectedValue(intCount)
               Exit For
            Elseif CLng(actualValue(i)) > CLng(testValue1(i)) Then
               Exit For
            End If
         Next
      End If
      actualValue = ""
      If objFSO.FileExists(arrFile2(intCount)) Then
         fileVer = objFSO.GetFileVersion(arrFile2(intCount))
         testValue2 = split(arrExpectedValue2(intCount), ".")
         actualValue = split(fileVer, ".")
         If CLng(UBound(testValue2)) >= CLng(UBound(actualValue)) Then
            h = UBound(actualValue)
         Else
            h = UBound(testValue2)
         End If
         for i = 0 to h
            If CLng(actualValue(i)) < CLng(testValue2(i)) Then
               finding = "open"
               testValue2 = split(arrFile2(intCount), "\")
               strMessage = strMessage & vbCrLf & "    " & testValue2(UBound(testValue2)) & " " & _
                           "is version: " & fileVer & " Expected: " & arrExpectedValue2(intCount)
               Exit For
            Elseif CLng(actualValue(i)) > CLng(testValue2(i)) Then
               Exit For
            End If
         Next
      End If
      testValue2 = ""
      actualValue = ""
      If objFSO.FileExists(arrFile3(intCount)) Then
         fileVer = objFSO.GetFileVersion(arrFile3(intCount))
         testValue2 = split(arrExpectedValue3(intCount), ".")
         actualValue = split(fileVer, ".")
         If CLng(UBound(testValue2)) >= CLng(UBound(actualValue)) Then
            h = UBound(actualValue)
         Else
            h = UBound(testValue2)
         End If
         for i = 0 to h
            If CLng(actualValue(i)) < CLng(testValue2(i)) Then
               finding = "open"
               testValue2 = split(arrFile3(intCount), "\")
               strMessage = strMessage & vbCrLf & "    " & testValue2(UBound(testValue2)) & " " & _
                           "is version: " & fileVer & " Expected: " & arrExpectedValue2(intCount)
               Exit For
            Elseif CLng(actualValue(i)) > CLng(testValue2(i)) Then
               Exit For
            End If
         Next
      End If
      testValue2 = ""
      actualValue = ""
      If objFSO.FileExists(arrSubKey3(intCount)) Then
         fileVer = objFSO.GetFileVersion(arrSubKey3(intCount))
         testValue2 = split(arrHive(intCount), ".")
         actualValue = split(fileVer, ".")
         If CLng(UBound(testValue2)) >= CLng(UBound(actualValue)) Then
            h = UBound(actualValue)
         Else
            h = UBound(testValue2)
         End If
         for i = 0 to h
            If CLng(actualValue(i)) < CLng(testValue2(i)) Then
               finding = "open"
               testValue2 = split(arrSubKey3(intCount), "\")
               strMessage = strMessage & vbCrLf & "    " & testValue2(UBound(testValue2)) & " " & _
                           "is version: " & fileVer & " Expected: " & arrHive(intCount)
               Exit For
            Elseif CLng(actualValue(i)) > CLng(testValue2(i)) Then
               Exit For
            End If
         Next
      End If
      testValue2 = ""
      actualValue = ""
      If objFSO.FileExists("C:\Program Files\Microsoft Silverlight\sllauncher.exe") Then
         fileVer = objFSO.GetFileVersion("C:\Program Files\Microsoft Silverlight\sllauncher.exe")
         testValue2 = split("5.1.10411", ".")
         actualValue = split(fileVer, ".")
         If CLng(UBound(testValue2)) >= CLng(UBound(actualValue)) Then
            h = UBound(actualValue)
         Else
            h = UBound(testValue2)
         End If
         for i = 0 to h
            If CLng(actualValue(i)) < CLng(testValue2(i)) Then
               finding = "open"
               testValue2 = split("C:\Program Files\Microsoft Silverlight\sllauncher.exe", "\")
               strMessage = strMessage & vbCrLf & "    " & testValue2(UBound(testValue2)) & " " & _
                           "is version: " & fileVer & " Expected: " & "5.1.10411"
               Exit For
            Elseif CLng(actualValue(i)) > CLng(testValue2(i)) Then
               Exit For
            End If
         Next
      End If
      testValue2 = ""
      actualValue = ""
      If objFSO.FileExists("C:\Program Files\Common Files\Microsoft Shared\OFFICE12\OGL.DLL") Then
         fileVer = objFSO.GetFileVersion("C:\Program Files\Common Files\Microsoft Shared\OFFICE12\OGL.DLL")
         testValue2 = split("14.0.6117.5001", ".")
         actualValue = split(fileVer, ".")
         If CLng(UBound(testValue2)) >= CLng(UBound(actualValue)) Then
            h = UBound(actualValue)
         Else
            h = UBound(testValue2)
         End If
         for i = 0 to h
            If CLng(actualValue(i)) < CLng(testValue2(i)) Then
               finding = "open"
               testValue2 = split("C:\Program Files\Common Files\Microsoft Shared\OFFICE12\OGL.DLL", "\")
               strMessage = strMessage & vbCrLf & "    " & testValue2(UBound(testValue2)) & " " & _
                           "is version: " & fileVer & " Expected: " & "14.0.6117.5001"
               Exit For
            Elseif CLng(actualValue(i)) > CLng(testValue2(i)) Then
               Exit For
            End If
         Next
      End If
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) & strMessage
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If

   ElseIf (StrComp(arrType(intCount), "WPFFileVer", vbTextCompare)) = 0 Then
      finding = "open"
      strMessage = ""
      Set objFolder = objFSO.GetFolder(arrFile1(intCount))
      Set colSubFolders = objFolder.SubFolders
      For Each i in colSubFolders
         If objFSO.FileExists(arrFile1(intCount) & i.name & "/" & arrSubKey(intCount)) Then
            fileVer = objFSO.GetFileVersion(arrFile1(intCount) & i.name & "/" & arrSubKey(intCount))
            testValue1 = split(arrExpectedValue(intCount), ".")
            actualValue = split(fileVer, ".")
         
            If UBound(testValue1) >= UBound(actualValue) Then
               h = UBound(actualValue)
            Else
               h = UBound(testValue1)
            End If
            for j = 0 to h
               If CLng(actualValue(j)) < CLng(testValue1(j)) Then
                  finding = "open"
                  testValue2 = split(arrFile1(intCount), "\")
                  strMessage = vbCrLf & "    " & arrSubKey(intCount) & " " & _
                            "is version: " & fileVer & " Expected: " & arrExpectedValue(intCount)
                  Exit For
               Elseif CLng(actualValue(j)) > CLng(testValue1(j)) Then
                  finding = "closed"
                  Exit For
               Elseif CLng(actualValue(j)) = CLng(testValue1(j)) Then
                  finding = "closed"
               End If
            Next
            If finding = "closed" Then
               Exit For
            End If 
         End If
      Next
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) & strMessage
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If

   ElseIf (StrComp(arrType(intCount), "ServicePack", vbTextCompare)) = 0 Then
      finding = "closed"
      If StrComp(ServicePack, arrExpectedValue(intCount), vbTextCompare) < 0 Then
         finding = "open"
      End If
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If


   ElseIf (StrComp(arrType(intCount), "Hotfix", vbTextCompare)) = 0 Then
      finding = "open"
      strMessage = ""
      For Each objKB in colKB
         If objKB.HotfixID = arrExpectedValue(intCount) Then
            finding = "closed"
            Exit For
         End If
      Next
      strMessage = vbCRLF & "    " & arrExpectedValue(intCount) & " is not installed" 
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)  & strMessage
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      End If


   ElseIf (StrComp(arrType(intCount), "DualRoleHotfix", vbTextCompare)) = 0 Then
      finding = "closed"
      strMessage = ""
      For Each objRole in colRoles
         If objRole.Name = arrSubKey(intCount) Or objRole.Name = arrSubKey2(intCount) Then
            finding = "open" 
            For Each objKB in colKB
               If objKB.HotfixID = arrExpectedValue(intCount) Then
                  finding = "closed"
                  Exit For
               End If
            Next
         End If
      Next
      strMessage = vbCRLF & "    " & arrExpectedValue(intCount) & " is not installed" 
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) & strMessage
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If


   ElseIf (StrComp(arrType(intCount), "RoleHotfix", vbTextCompare)) = 0 Then
      finding = "closed"
      strMessage = ""
      For Each objRole in colRoles
         If objRole.Name = arrSubKey(intCount) Then
            finding = "open" 
            For Each objKB in colKB
               If objKB.HotfixID = arrExpectedValue(intCount) Then
                  finding = "closed"
                  Exit For
               End If
            Next
         End If
      Next
      strMessage = vbCRLF & "    " & arrExpectedValue(intCount) & " is not installed" 
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) & strMessage
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If

   ElseIf (StrComp(arrType(intCount), "AntiVirus", vbTextCompare)) = 0 Then
      finding = "open"
      strMessage = vbCRLF & "    No AV detected"
      For Each objService in colServiceList
         If objService.Name = arrExpectedValue(intCount) Then
            objReg.GetStringValue HKLM, "Software\Symantec\Symantec Endpoint Protection\SMC", "Version", arrValue(intCount)
            If arrValue(intCount) >= "11.0" Then
               objReg.GetBinaryValue HKLM, arrKeyPath(intCount), arrSubkey2(intCount), arrBinary
               intYear = 1970 + arrBinary(0)
               intMonth = arrBinary(1) + 1
               intDay = arrBinary(2)
               strDate = intMonth & "/" & intDay & "/" & intYear
               datAVDate = CDate(strDate)
               intDay = Abs(DateDiff("d", datAvDate, Date))
               if intDay > 7 Then
                  strMessage = vbCRLF & "    Symantec Virus Definitions are " & intDay & " days old"
                  Exit For
               Else
                  finding = "closed"
                  Exit For   
               End If
            Else
               strMessage = vbCRLF & "    Symantec is not a compliant version: Version " & arrValue(intCount) & "is installed"
               Exit For 
            End If  
         ElseIf objService.Name = arrExpectedValue2(intCount) Then
            objReg.GetStringValue HKLM, "Software\Wow6432Node\Mcafee\DesktopProtection", "szProductVer", arrValue(intCount)
            If arrValue(intCount) >= "8.0" Then
               objReg.GetStringValue HKLM, arrSubkey(intCount), arrSubkey3(intCount), arrValue(intCount)
               datAVDate = CDate(arrValue(intCount))
               intDay = Abs(DateDiff("d", datAvDate, Date))
               if intDay > 7 Then
                  strMessage = vbCRLF & "    Mcafee Virus Definitions are " & intDay & " days old "
                  Exit For
               Else
                  finding = "closed"
                  Exit For   
               End If
            Else
               strMessage = vbCRLF & "    Mcafee is not a compliant version: Version " & arrValue(intCount) & "is installed"
               Exit For
            End If
         End If
      Next
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) & strMessage
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If


   ElseIf (StrComp(arrType(intCount), "PrinterShares", vbTextCompare)) = 0 Then
      finding = "closed"
      strMessage = ""
      If Not colPrinters.Count = 0 Then
         For Each objPrinter In colPrinters
            i = objPrinter.GetSecurityDescriptor(objSD1)
            RegEx.Pattern = ".*Admin.*"
            RegEx.IgnoreCase = True
            RegEx2.Pattern = ".*creator.*owner.*"
            RegEx2.IgnoreCase = True
            For Each objACE in objSD1.DACL
               If RegEx.Test(objACE.Trustee.Name) Then
               Elseif RegEx2.Test(objACE.Trustee.Name) Then
               Else
                  If (objACE.AccessMask XOr 131080) = 0 Or (objACE.AccessMask Xor 536870912) = 0 Then
                  Else
                     finding = "open"
                     strMessage = vbCRLF & "    " & objPrinter.Name & " is shared and has incorrect permissions"
                     Exit For
                  End If  
               End If
            Next
            If finding = "open" Then
               Exit For
            End If
         Next
      End If 
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)  & strMessage
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If


   ElseIf (StrComp(arrType(intCount), "ProgramHotfix", vbTextCompare)) = 0 Then
      finding = "closed"
      strMessage = ""
      For Each objProduct in colProducts
         WScript.Echo objProject.Name
         If objProduct.Name = arrSubKey(intCount) Then
            finding = "open" 
            If Not colKB.Count = 0 Then
               For Each objKB in colKB
                  If objKB.HotfixID = arrExpectedValue(intCount) Then
                     finding = "closed"
                     Exit For
                  End If
               Next
            End If
         End If
      Next
      strMessage = vbCRLF & "    " & arrExpectedValue(intCount) & " is not installed" 
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) & strMessage
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If

  ElseIf (StrComp(arrType(intCount), "IE11FileVer", vbTextCompare)) = 0 Then
      finding = "closed"
      If objFSO.FileExists(arrFile1(intCount)) Then
         fileVer = objFSO.GetFileVersion(arrFile1(intCount))
         actualValue = split(fileVer, ".")
         If CInt(actualValue(0)) = 8 Then
            testValue1 = split(arrExpectedValue(intCount), ".")
            If UBound(testValue1) >= UBound(actualValue) Then
               h = UBound(actualValue)
            Else
               h = UBound(testValue1)
            End If
            for i = 0 to h
               If actualValue(i) < testValue1(i) Then
                  finding = "open"
                  testValue2 = split(arrFile1(intCount), "\")
                  strMessage = vbCrLf & "    " & testValue2(UBound(testValue2)) & " " & _
                               "is version: " & fileVer & " Expected: " & arrExpectedValue(intCount)
                  Exit For
               ElseIf actualValue(i) > testValue1(i) Then
                  Exit For
               End If
            Next
         ElseIf CInt(actualValue(0)) = 9 Then
            testValue1 = split(arrExpectedValue2(intCount), ".")
            If UBound(testValue1) >= UBound(actualValue) Then
               h = UBound(actualValue)
            Else
               h = UBound(testValue1)
            End If
            for i = 0 to h
               If actualValue(i) < testValue1(i) Then
                  finding = "open"
                  testValue2 = split(arrFile1(intCount), "\")
                  strMessage = vbCrLf & "    " & testValue2(UBound(testValue2)) & " " & _
                               "is version: " & fileVer & " Expected: " & arrExpectedValue2(intCount)
                  Exit For
               ElseIf actualValue(i) > testValue1(i) Then
                  Exit For
               End If
            Next
         ElseIf CInt(actualValue(0)) = 10 Then
            testValue1 = split(arrExpectedValue3(intCount), ".")
            If UBound(testValue1) >= UBound(actualValue) Then
               h = UBound(actualValue)
            Else
               h = UBound(testValue1)
            End If
            for i = 0 to h
               If actualValue(i) < testValue1(i) Then
                  finding = "open"
                  testValue2 = split(arrFile1(intCount), "\")
                  strMessage = vbCrLf & "    " & testValue2(UBound(testValue2)) & " " & _
                               "is version: " & fileVer & " Expected: " & arrExpectedValue3(intCount)
                  Exit For
               ElseIf actualValue(i) > testValue1(i) Then
                  Exit For
               End If
            Next
         ElseIf CInt(actualValue(0)) = 11 Then
            testValue1 = split(arrFile2(intCount), ".")
            If UBound(testValue1) >= UBound(actualValue) Then
               h = UBound(actualValue)
            Else
               h = UBound(testValue1)
            End If
            for i = 0 to h
               If actualValue(i) < testValue1(i) Then
                  finding = "open"
                  testValue2 = split(arrFile1(intCount), "\")
                  strMessage = vbCrLf & "    " & testValue2(UBound(testValue2)) & " " & _
                               "is version: " & fileVer & " Expected: " & arrFile2(intCount)
                  Exit For
               ElseIf actualValue(i) > testValue1(i) Then
                  Exit For
               End If
            Next
         End If 
      End If
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) & strMessage
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If

   ElseIf (StrComp(arrType(intCount), "IE10FileVer", vbTextCompare)) = 0 Then
      finding = "closed"
      If objFSO.FileExists(arrFile1(intCount)) Then
         fileVer = objFSO.GetFileVersion(arrFile1(intCount))
         actualValue = split(fileVer, ".")
         If CInt(actualValue(0)) = 8 Then
            testValue1 = split(arrExpectedValue(intCount), ".")
            If UBound(testValue1) >= UBound(actualValue) Then
               h = UBound(actualValue)
            Else
               h = UBound(testValue1)
            End If
            for i = 0 to h
               If actualValue(i) < testValue1(i) Then
                  finding = "open"
                  testValue2 = split(arrFile1(intCount), "\")
                  strMessage = vbCrLf & "    " & testValue2(UBound(testValue2)) & " " & _
                               "is version: " & fileVer & " Expected: " & arrExpectedValue(intCount)
                  Exit For
               ElseIf actualValue(i) > testValue1(i) Then
                  Exit For
               End If
            Next
         ElseIf CInt(actualValue(0)) = 9 Then
            testValue1 = split(arrExpectedValue2(intCount), ".")
            If UBound(testValue1) >= UBound(actualValue) Then
               h = UBound(actualValue)
            Else
               h = UBound(testValue1)
            End If
            for i = 0 to h
               If actualValue(i) < testValue1(i) Then
                  finding = "open"
                  testValue2 = split(arrFile1(intCount), "\")
                  strMessage = vbCrLf & "    " & testValue2(UBound(testValue2)) & " " & _
                               "is version: " & fileVer & " Expected: " & arrExpectedValue2(intCount)
                  Exit For
               ElseIf actualValue(i) > testValue1(i) Then
                  Exit For
               End If
            Next
		 ElseIf CInt(actualValue(0)) = 10 Then
            testValue1 = split(arrExpectedValue3(intCount), ".")
            If UBound(testValue1) >= UBound(actualValue) Then
               h = UBound(actualValue)
            Else
               h = UBound(testValue1)
            End If
            for i = 0 to h
               If actualValue(i) < testValue1(i) Then
                  finding = "open"
                  testValue2 = split(arrFile1(intCount), "\")
                  strMessage = vbCrLf & "    " & testValue2(UBound(testValue2)) & " " & _
                               "is version: " & fileVer & " Expected: " & arrExpectedValue3(intCount)
                  Exit For
               ElseIf actualValue(i) > testValue1(i) Then
                  Exit For
               End If
            Next
         End If 
      End If
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) & strMessage
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If
	  
      ElseIf (StrComp(arrType(intCount), "IE9FileVer", vbTextCompare)) = 0 Then
      finding = "closed"
      If objFSO.FileExists(arrFile1(intCount)) Then
         fileVer = objFSO.GetFileVersion(arrFile1(intCount))
         actualValue = split(fileVer, ".")
         If CInt(actualValue(0)) = 8 Then
            testValue1 = split(arrExpectedValue(intCount), ".")
            If UBound(testValue1) >= UBound(actualValue) Then
               h = UBound(actualValue)
            Else
               h = UBound(testValue1)
            End If
            for i = 0 to h
               If actualValue(i) < testValue1(i) Then
                  finding = "open"
                  testValue2 = split(arrFile1(intCount), "\")
                  strMessage = vbCrLf & "    " & testValue2(UBound(testValue2)) & " " & _
                               "is version: " & fileVer & " Expected: " & arrExpectedValue(intCount)
                  Exit For
               ElseIf actualValue(i) > testValue1(i) Then
                  Exit For
               End If
            Next
         ElseIf CInt(actualValue(0)) = 9 Then
            testValue1 = split(arrExpectedValue2(intCount), ".")
            If UBound(testValue1) >= UBound(actualValue) Then
               h = UBound(actualValue)
            Else
               h = UBound(testValue1)
            End If
            for i = 0 to h
               If actualValue(i) < testValue1(i) Then
                  finding = "open"
                  testValue2 = split(arrFile1(intCount), "\")
                  strMessage = vbCrLf & "    " & testValue2(UBound(testValue2)) & " " & _
                               "is version: " & fileVer & " Expected: " & arrExpectedValue2(intCount)
                  Exit For
               ElseIf actualValue(i) > testValue1(i) Then
                  Exit For
               End If
            Next
         End If 
      End If
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) & strMessage
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If

   ElseIf (StrComp(arrType(intCount), "SymantecValueLessEx", vbTextCompare)) = 0 Then
      finding = "closed"
      boolResult = false
      objReg.EnumValues arrHive(intCount), arrKeyPath(intCount), arrKeyList
      For Each strSubkey In arrKeyList
         If (StrComp(arrSubkey(intCount),strSubkey,vbTextCompare)) = 0 Then
            boolResult = true
            Exit For
         End If
      Next
      objReg.GetStringValue arrHive(intCount), arrKeyPath(intCount), arrSubkey(intCount), arrValue(intCount)
      If boolResult = false then
         finding = "closed"
      ElseIf arrvalue(intCount) <= arrExpectedValue(intCount) Then
         finding = "open" 
      Else
         finding = "closed"
      End If
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      End If


   ElseIf (StrComp(arrType(intCount), "FlashStringLess", vbTextCompare)) = 0 Then
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
      If boolResult Then
         objReg.GetStringValue arrHive(intCount), arrKeyPath(intCount), arrSubkey(intCount), arrValue(intCount)
         testValue1 = split(arrExpectedValue(intCount), ".")
         actualValue = split(arrValue(intCount), ",")
         If UBound(testValue1) >= UBound(actualValue) Then
            h = UBound(actualValue)
         Else
            h = UBound(testValue1)
         End If
         If CInt(actualValue(0)) < CInt(testValue1(0)) Then
            finding = "open"
         ElseIf CInt(actualValue(0)) = CInt(testValue1(0)) Then
            For i = 0 to h
               If Cint(actualValue(i)) < Cint(testValue1(i)) Then
                  finding = "open"
                  Exit For
               Elseif Cint(actualValue(i)) > Cint(testValue1(i)) Then
                  Exit For
               End If
            Next
         ElseIf CInt(actualValue(0)) > CInt(testValue1(0)) Then
            finding = "closed"
         End If
      End If  
      If finding="open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      ElseIf finding="closed" Then
         Wscript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      End If

   ElseIf (StrComp(arrType(intCount), "FlashDualStringLess", vbTextCompare)) = 0 Then
      finding = "closed"
      strMessage = ""
      boolResult = "false"
      objReg.EnumValues arrHive(intCount), arrKeyPath(intCount), arrKeyList
      If(isArray(arrKeyList)) Then
         For Each strSubkey In arrKeyList
            If (StrComp(arrSubkey(intCount),strSubkey,vbTextCompare)) = 0 Then
               boolResult = true
               Exit For
            End If
         Next
      End If
      If boolResult Then
         objReg.GetStringValue arrHive(intCount), arrKeyPath(intCount), arrSubkey(intCount), arrValue(intCount)
         testValue1 = split(arrExpectedValue(intCount), ".")
         testValue2 = split(arrExpectedValue2(intCount), ".")
         actualValue = split(arrValue(intCount), ",")
         If UBound(testValue1) >= UBound(actualValue) Then
            h = UBound(actualValue)
         Else
            h = UBound(testValue1)
         End If
         If CLng(UBound(testValue2)) >= CLng(UBound(actualValue)) Then
            l = UBound(actualValue)
         Else
            l = UBound(testValue2)
         End If
         If CInt(actualValue(0)) < CInt(testValue1(0)) Then
            finding = "open"
         ElseIf CInt(actualValue(0)) = CInt(testValue1(0)) Then
            For i = 0 to h
               If Cint(actualValue(i)) < Cint(testValue1(i)) Then
                  finding = "open"
                  strMessage = vbCrLf & "    FlashPlayer is version: " & strMessage & " Expected version: " & arrExpectedValue(intCount)
                  Exit For
               Elseif Cint(actualValue(i)) > Cint(testValue1(i)) Then
                  Exit For
               End If
            Next
         ElseIf Cint(actualValue(0)) = CInt(testValue2(0)) Then
            For i = 0 to l
               If Cint(actualValue(i)) < Cint(testValue2(i)) Then
                  finding = "open"
                  strMessage = vbCrLf & "    FlashPlayer is version: " & strMessage & " Expected version: " & arrExpectedValue(intCount)
                  Exit For
               Elseif Cint(actualValue(i)) > Cint(testValue2(i)) Then
                  Exit For
               End If
            Next
         ElseIf CInt(actualValue(0)) > CInt(testValue2(0)) Then
            finding = "closed"
         End If
      End If  
      If finding="open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      ElseIf finding="closed" Then
         Wscript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      End If
	  
   ElseIf (StrComp(arrType(intCount), "FileFlashDualStringLess", vbTextCompare)) = 0 Then
      finding = "closed"
      boolResult = "false"
	  strMessage = ""
      objReg.EnumValues arrHive(intCount), arrKeyPath(intCount), arrKeyList
      If(isArray(arrKeyList)) Then
         For Each strSubkey In arrKeyList
            If (StrComp(arrSubkey(intCount),strSubkey,vbTextCompare)) = 0 Then
               boolResult = true
               Exit For
            End If
         Next
      End If
      If boolResult Then
         objReg.GetStringValue arrHive(intCount), arrKeyPath(intCount), arrSubkey(intCount), arrValue(intCount)
         testValue1 = split(arrExpectedValue(intCount), ".")
         testValue2 = split(arrExpectedValue2(intCount), ".")
         actualValue = split(arrValue(intCount), ",")
		 For i = 0 to UBound(actualValue)
		    If i < UBound(actualValue) Then
			   strMessage = strMessage & actualvalue(i) & "."
			Else
			   strMessage = strMessage & actualValue(i)
			End If
	     Next
         If UBound(testValue1) >= UBound(actualValue) Then
            h = UBound(actualValue)
         Else
            h = UBound(testValue1)
         End If
         If CLng(UBound(testValue2)) >= CLng(UBound(actualValue)) Then
            l = UBound(actualValue)
         Else
            l = UBound(testValue2)
         End If
         If CInt(actualValue(0)) < CInt(testValue1(0)) Then
            finding = "open"
			WScript.Echo "Test"
         ElseIf CInt(actualValue(0)) = CInt(testValue1(0)) Then
            For i = 0 to h
               If Cint(actualValue(i)) < Cint(testValue1(i)) Then
                  finding = "open"
	          strMessage = vbCrLf & "    FlashPlayer is version: " & strMessage & " Expected version: " & arrExpectedValue(intCount)
                  Exit For
               Elseif Cint(actualValue(i)) > Cint(testValue1(i)) Then
                  Exit For
               End If
            Next
         ElseIf Cint(actualValue(0)) = CInt(testValue2(0)) Then
            For i = 0 to l
               If Cint(actualValue(i)) < Cint(testValue2(i)) Then
                  finding = "open"
				  strMessage = vbCrLf & "    FlashPlayer is version: " & strMessage & " Expected version: " & arrExpectedValue2(intCount)
                  Exit For
               Elseif Cint(actualValue(i)) > Cint(testValue2(i)) Then
                  Exit For
               End If
            Next
         ElseIf CInt(actualValue(0)) > CInt(testValue2(0)) Then
            finding = "closed"
         End If
      End If 
	  If finding = "closed" Then
         If objFSO.FileExists(arrFile1(intCount)) Then
            fileVer = objFSO.GetFileVersion(arrFile1(intCount))
            testValue1 = split(arrExpectedValue3(intCount), ".")
            actualValue = split(fileVer, ".")
            If UBound(testValue1) >= UBound(actualValue) Then
               h = UBound(actualValue)
            Else
               h = UBound(testValue1)
            End If
            for i = 0 to h
               If actualValue(i) < testValue1(i) Then
                  finding = "open"
                  testValue2 = split(arrFile1(intCount), "\")
                  strMessage = vbCrLf & "    " & testValue2(UBound(testValue2)) & " " & _
                               "is version: " & fileVer & " Expected: " & arrExpectedValue3(intCount)
                  Exit For
               ElseIf actualValue(i) > testValue1(i) Then
                  Exit For
               End If
            Next
         End If
	  End If
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) & strMessage
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If	  

   ElseIf(StrComp(arrType(intCount), "UntrustedCertificates", vbTextCompare)) = 0 Then
      finding = "open"
      RegEx.Pattern = arrExpectedValue(intCount) 
      RegEx.IgnoreCase = True 
      For i = 0 to UBound(certLines)
         if RegEx.Test(certLines(i)) Then
            finding = "closed"
         End If
      Next
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If 

   ElseIf(StrComp(arrType(intCount), "TriUntrustedCertificates", vbTextCompare)) = 0 Then
      finding = "open"
      RegEx.Pattern = arrExpectedValue(intCount)
      RegEx2.Pattern = arrExpectedValue2(intCount) 
      RegEx3.Pattern = arrExpectedValue3(intCount)  
      RegEx.IgnoreCase = True
      RegEx2.IgnoreCase = True
      RegEx3.IgnoreCase = True 
      For i = 0 to UBound(certLines)
         if RegEx.Test(certLines(i)) Then
            finding = "closed"
         End If
      Next
      If finding = "closed" Then
         finding = "open"
         For i = 0 to UBound(certLines)
            if RegEx.Test(certLines(i)) Then
               finding = "closed"
            End If
         Next
      End If
      If finding = "closed" Then
         finding = "open"
         For i = 0 to UBound(certLines)
            if RegEx.Test(certLines(i)) Then
               finding = "closed"
            End If
         Next
      End If		
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If 

   ElseIf(StrComp(arrType(intCount), "DoDCertificates", vbTextCompare)) = 0 Then
      finding = "open"
      RegEx.Pattern = arrExpectedValue(intCount) 
      RegEx.IgnoreCase = True 
      For i = 0 to UBound(dodLines)
         if RegEx.Test(dodLines(i)) Then
            finding = "closed"
         End If
      Next
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If 

   ElseIf(StrComp(arrType(intCount), "ManualReview", vbTextCompare)) = 0 Then
      WScript.Echo arrFinding(intCount) & " Manual Review -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)

   ElseIf(StrComp(arrType(intCount), "Closed", vbTextCompare)) = 0 Then
      WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)

   ElseIf(StrComp(arrType(intCount), "LocalAccounts", vbTextCompare)) = 0 Then
      finding = "closed"
      strMessage = ""
      For Each objComputerSystem in colComputerSystems
         If objComputerSystem.DomainRole = 1 Then
            For Each objItem in colItems
               On Error Resume Next
               Set objUser = GetObject("WinNT://" & strComputer & "/" & objItem.Name & ",user")
               dtmLastLogin = objUser.lastLogin
               RegEx.Pattern = ".*S-1-5-21.*500"
               RegEx2.Pattern = ".*S-1-5-21.*501"
               if RegEx.Test(objItem.SID) or RegEx2.Test(objItem.SID) Then
               Else
                 finding = "open"
                 strMessage = vbCrLf & "    Local user " & objItem.Name & " exists and this workstation is on a domain"       
              End If
            Next 
         End If
      Next
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)  & strMessage
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      End If 

   ElseIf(StrComp(arrType(intCount), "BuiltinAccount", vbTextCompare)) = 0 Then
      finding = "closed"
      RegEx.Pattern = ".*S-1-5-21.*500"
      For Each objItem in colItems
         If RegEx.Test(objItem.SID) Then
            If objItem.Disabled Then
            Else
               finding = "open"
               Exit For
            End If
         End If
      Next
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)  & strMessage
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      End If 

   ElseIf(StrComp(arrType(intCount), "FirewallRule", vbTextCompare)) = 0 Then
      finding = "open"
      objReg.EnumValues arrHive(intCount), arrKeyPath(intCount), arrKeyList
      For each strSubKey in arrKeyList
         objReg.GetStringValue arrHive(intCount), arrKeyPath(intCount), strSubKey, arrValue(intCount)
         RegEx.Pattern = arrSubKey(intCount)
         RegEx.IgnoreCase = True
         RegEx2.Pattern = arrSubkey2(intCount)
         RegEx2.IgnoreCase = True
         RegEx3.Pattern = arrSubkey3(intCount)
         RegEx3.IgnoreCase = True
         If RegEx.Test(arrValue(intCount)) And RegEx2.Test(arrValue(intCount)) And RegEx3.Test(arrValue(intCount)) Then
            RegEx.Pattern = arrFile1(intCount)
            RegEx.IgnoreCase = True
            RegEx2.Pattern = arrFile2(intCount)
            RegEx2.IgnoreCase = True
            RegEx3.Pattern = arrFile3(intCount)
            RegEx3.IgnoreCase = True 
            If RegEx.Test(arrValue(intCount)) And RegEx2.Test(arrValue(intCount)) And Not RegEx3.Test(arrValue(intCount)) Then 
               RegEx.Pattern = arrExpectedValue(intCount)
               RegEx.IgnoreCase = True
               RegEx2.Pattern = arrExpectedValue2(intCount)
               RegEx2.IgnoreCase = True
               If Not RegEx.Test(arrValue(intCount)) And Not RegEx2.Test(arrValue(intCount)) Then
                  finding = "closed"
                  Exit For
               End If 
            End If        
         End If
      Next
      If finding = "open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)  & strMessage
      Else
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      End If 

   ElseIf (StrComp(arrType(intCount), "MSMalwareStringLess", vbTextCompare)) = 0 Then
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

      If boolResult Then
         objReg.GetStringValue arrHive(intCount), arrKeyPath(intCount), arrSubkey(intCount), arrValue(intCount)
         testValue1 = split(arrExpectedValue(intCount), ".")
         actualValue = split(arrValue(intCount), ".")
         If UBound(testValue1) >= UBound(actualValue) Then
            h = UBound(actualValue)
         Else
            h = UBound(testValue1)
         End If
         If CInt(actualValue(0)) < CInt(testValue1(0)) Then
            finding = "open"
         ElseIf CInt(actualValue(0)) = CInt(testValue1(0)) Then
            For i = 0 to h
               If Cint(actualValue(i)) < Cint(testValue1(i)) Then
                  finding = "open"
                  Exit For
               Elseif Cint(actualValue(i)) > Cint(testValue1(i)) Then
                  Exit For
               End If
            Next
         ElseIf CInt(actualValue(0)) > CInt(testValue1(0)) Then
            finding = "closed"
        End If  
      End If
      If finding="open" Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      ElseIf finding="closed" Then
         Wscript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount) 
      End If
   End If  
Next

if objFSO.FileExists("audit_results.txt") Then
   objFSO.DeleteFile("audit_results.txt")
End If
if objFSO.FileExists("stig.db") Then
   objFSO.DeleteFile("stig.db")
End If
if objFSO.FileExists("secedit_results.txt") Then
   objFSO.DeleteFile("secedit_results.txt")
End If
if objFSO.FileExists("untrusted_certificates_results.txt") Then
   objFSO.DeleteFile("untrusted_certificates_results.txt")
End If
if objFSO.FileExists("dir_results.txt") Then
   objFSO.DeleteFile("dir_results.txt")
End If
if objFSO.FileExists("pfx_results.txt") Then
   objFSO.DeleteFile("pfx_results.txt")
End If
if objFSO.FileExists("dod_certificates_results.txt") Then
   objFSO.DeleteFile("dod_certificates_results.txt")
End If