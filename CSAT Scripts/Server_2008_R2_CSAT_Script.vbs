Option Explicit
'On Error Resume Next
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
Dim actualValue
Dim h
Dim l

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

arrFinding(x) = "V-1070"
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
arrTitle(x) = "Physical Security"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1072"
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
arrTitle(x) = "Shared User Accounts"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1073"
arrKeyPath(x) = "Software\Microsoft\Windows NT\CurrentVersion"
arrSubKey(x) = "CurrentVersion"
arrSubKey2(x) = "CurrentBuild"
arrSubkey3(x) = "x"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "6.1"
arrExpectedValue2(x) = "7600"
arrExpectedValue3(x) = "0"
arrType(x) = "OS_Version"
arrTitle(x) = "Unsupported Service Packs"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1074"
arrKeyPath(x) = "SOFTWARE\Wow6432node\Symantec\Symantec Endpoint Protection\AV"
arrSubKey(x) = "SOFTWARE\Wow6432Node\McAfee\AVEngine"
arrSubKey2(x) = "PatternFileDate"
arrSubkey3(x) = "AVDatDate"
arrFile1(x) = "8.0.0.0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "SepMasterService"
arrExpectedValue2(x) = "McShield"
arrExpectedValue3(x) = "0"
arrType(x) = "Antivirus"
arrTitle(x) = "Approved DoD Virus Scan Program"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-1075"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
arrSubKey(x) = "ShutdownWithoutLogon"
arrSubKey2(x) = "x"
arrSubkey3(x) = "x"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "Value"
arrTitle(x) = "Display Shutdown Button"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1076"
arrKeyPath(x) = "Default"
arrSubKey(x) = "x"
arrSubKey2(x) = "x"
arrSubkey3(x) = "x"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "Manual Review"
arrTitle(x) = "System Recovery Backups"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-1077"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "x"
arrSubkey3(x) = "x"
arrFile1(x) = "C:\\Windows\System32\Winevt\Logs\Application.evtx"
arrFile2(x) = "C:\\Windows\System32\Winevt\Logs\Security.evtx"
arrFile3(x) = "C:\\Windows\System32\Winevt\Logs\System.evtx"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "LogPerms"
arrTitle(x) = "Incorrect ACLs for event logs"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1080"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "x"
arrSubkey3(x) = "x"
arrFile1(x) = "x"
arrFile2(x) = "x"
arrFile3(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrType(x) = "VolAudit"
arrTitle(x) = "File Auditing Configuration"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1081"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "x"
arrSubkey3(x) = "x"
arrFile1(x) = "x"
arrFile2(x) = "x"
arrFile3(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileSystem"
arrTitle(x) = "NTFS Requirement"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-1088"
arrKeyPath(x) = "x"
arrSubKey(x) = "Software"
arrSubKey2(x) = "System"
arrSubkey3(x) = "x"
arrFile1(x) = "x"
arrFile2(x) = "x"
arrFile3(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "RegAudit"
arrTitle(x) = "Registry Key Auditing"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1089"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
arrSubKey(x) = "LegalNoticeText"
arrSubKey2(x) = "System"
arrSubkey3(x) = "x"
arrFile1(x) = "x"
arrFile2(x) = "x"
arrFile3(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "RegStringRegExp"
arrTitle(x) = "Legal Notice Display"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1090"
arrKeyPath(x) = "Software\Microsoft\Windows NT\CurrentVersion\Winlogon"
arrSubKey(x) = "CachedLogonsCount"
arrSubKey2(x) = "x"
arrSubkey3(x) = "x"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "2"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "String"
arrTitle(x) = "Caching of Logon Credentials"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1093"
arrKeyPath(x) = "System\CurrentControlSet\Control\Lsa"
arrSubKey(x) = "RestrictAnonymous"
arrSubKey2(x) = "x"
arrSubkey3(x) = "x"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Restrict Anonymous Network Shares"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-1097"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "x"
arrSubkey3(x) = "x"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*LockoutBadCount.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SeceditRegExp"
arrTitle(x) = "Bad Logon Attempts"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1098"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "x"
arrSubkey3(x) = "x"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*ResetLockoutCount.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SeceditRegExp"
arrTitle(x) = "Bad Logon Counter Reset"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1099"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "x"
arrSubkey3(x) = "x"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*LockoutDuration.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SeceditRegExp"
arrTitle(x) = "Lockout Duration"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1102"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "x"
arrSubkey3(x) = "x"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeTcbPrivilege.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SeceditRegExp"
arrTitle(x) = "User Right - Act as Part of OS"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-1104"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "x"
arrSubkey3(x) = "x"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*MaximumPasswordAge.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SeceditRegExp"
arrTitle(x) = "Maximum Password Age"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1105"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "x"
arrSubkey3(x) = "x"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*MinimumPasswordAge.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SeceditRegExp"
arrTitle(x) = "Minimum Password Age"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1107"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "x"
arrSubkey3(x) = "x"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*PasswordHistorySize.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SeceditRegExp"
arrTitle(x) = "Password Uniqueness"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1112"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "x"
arrSubKey3(x) = "x"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "StaleUsers"
arrTitle(x) = "Dormant Accounts"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1113"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "x"
arrSubkey3(x) = "x"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*EnableGuestAccount.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SeceditRegExp"
arrTitle(x) = "Disable Guest Account"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1114"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "x"
arrSubkey3(x) = "x"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "Guest"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "RenameUser"
arrTitle(x) = "Rename Built-in Guest Account"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1115"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "x"
arrSubkey3(x) = "x"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "Administrator"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "RenameUser"
arrTitle(x) = "Rename Built-in Admin Account"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1119"
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
arrTitle(x) = "Booting into Multiple Operating Systems"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1120"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*ftpsvc.*"
arrExpectedValue2(x) = "Disabled"
arrExpectedValue3(x) = "0"
arrType(x) = "ServiceDisabled"
arrTitle(x) = "Prohibited FTP Logins"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1121"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*ftpsvc.*"
arrExpectedValue2(x) = "Disabled"
arrExpectedValue3(x) = "0"
arrType(x) = "ServiceDisabled"
arrTitle(x) = "FTP System File Access"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-1122"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\Control Panel\Desktop"
arrSubKey(x) = "ScreenSaveActive"
arrSubKey2(x) = "ScreenSaverIsSecure"
arrSubkey3(x) = "ScreenSaveTimeout"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "1"
arrExpectedValue3(x) = "900"
arrType(x) = "TriString"
arrTitle(x) = "Password Protected Screen Saver"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1127"
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
arrTitle(x) = "Restricted Administrator Group Membership"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1128"
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
arrTitle(x) = "Security Configuration Tools"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-1130"
arrKeyPath(x) = "System\CurrentControlSet\Control\Lsa"
arrSubKey(x) = "EveryoneIncludesAnonymous"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "Value"
arrTitle(x) = "System File ACLs"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1131"
arrKeyPath(x) = "System\CurrentControlSet\Control\Lsa"
arrSubKey(x) = "Notification Packages"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "PasswordFilter"
arrTitle(x) = "Strong Password Filtering"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1135"
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
arrType(x) = "PrinterShares"
arrTitle(x) = "Printer Share Permissions"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-1136"
arrKeyPath(x) = "System\CurrentControlSet\Services\LanManServer\Parameters"
arrSubKey(x) = "EnableForcedLogoff"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "Value"
arrTitle(x) = "Forcibly Disconnect when Logon Hours Expire"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-1137"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeSecurityPrivilege.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Access Restrictions to Logs"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1140"
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
arrTitle(x) = "Users with Administrative Privileges"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1141"
arrKeyPath(x) = "System\CurrentControlSet\Services\LanmanWorkstation\Parameters"
arrSubKey(x) = "EnablePlainTextPassword"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "Value"
arrTitle(x) = "Unencrypted Password is Sent to SMB Server"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1145"
arrKeyPath(x) = "Software\Microsoft\Windows NT\CurrentVersion\Winlogon"
arrSubKey(x) = "AutoAdminLogon"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "String"
arrTitle(x) = "Disable Administrator Automatic Logon"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-1150"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*PasswordComplexity.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Microsoft Strong Password Filtering"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-1151"
arrKeyPath(x) = "System\CurrentControlSet\Control\Print\Providers\LanMan Print Services\Servers"
arrSubKey(x) = "AddPrinterDrivers"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "Value"
arrTitle(x) = "Secure Print Driver Installation"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1152"
arrKeyPath(x) = "Default"
arrSubKey(x) = "System\CurrentControlSet\Control\SecurePipeServers\Winreg"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "WinReg"
arrTitle(x) = "Anonymous Access to the Registry"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-1153"
arrKeyPath(x) = "System\CurrentControlSet\Control\Lsa"
arrSubKey(x) = "LmCompatibilityLevel"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "5"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "Value"
arrTitle(x) = "LanMan Authentication Level"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-1154"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
arrSubKey(x) = "DisableCAD"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "Value"
arrTitle(x) = "Ctrl+Alt+Del Security Attention Sequence"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1155"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeDenyNetworkLogonRight.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Deny Access from the Network"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1157"
arrKeyPath(x) = "Software\Microsoft\Windows NT\CurrentVersion\Winlogon"
arrSubKey(x) = "SCRemoveOption"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "2"
arrExpectedValue3(x) = "0"
arrType(x) = "StringRange"
arrTitle(x) = "Smart Card Removal Option"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1158"
arrKeyPath(x) = "Software\Microsoft\Windows Nt\CurrentVersion\Setup\RecoveryConsole"
arrSubKey(x) = "SetCommand"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "Value"
arrTitle(x) = "Recovery Console - SET Command"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-1159"
arrKeyPath(x) = "Software\Microsoft\Windows Nt\CurrentVersion\Setup\RecoveryConsole"
arrSubKey(x) = "SecurityLevel"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "Value"
arrTitle(x) = "Recovery Console - Automatic Logon"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-1162"
arrKeyPath(x) = "System\CurrentControlSet\Services\LanManServer\Parameters"
arrSubKey(x) = "EnableSecuritySignature"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "Value"
arrTitle(x) = "SMB Server Packet Signing (if client agrees)"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1163"
arrKeyPath(x) = "System\CurrentControlSet\Services\Netlogon\Parameters"
arrSubKey(x) = "SealSecureChannel"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "Value"
arrTitle(x) = "Encryption of Secure Channel Traffic"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1164"
arrKeyPath(x) = "System\CurrentControlSet\Services\Netlogon\Parameters"
arrSubKey(x) = "SignSecureChannel"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "Value"
arrTitle(x) = "Signing of Secure Channel Traffic"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1165"
arrKeyPath(x) = "System\CurrentControlSet\Services\Netlogon\Parameters"
arrSubKey(x) = "DisablePasswordChange"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "Value"
arrTitle(x) = "Computer Account Password Reset"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-1166"
arrKeyPath(x) = "System\CurrentControlSet\Services\LanManWorkstation\Parameters"
arrSubKey(x) = "EnableSecuritySignature"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "Value"
arrTitle(x) = "SMB Client Packet Signing (if server agrees)"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1168"
arrKeyPath(x) = "Software"
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
arrType(x) = "BackupOperators"
arrTitle(x) = "Members of the Backup Operators Group"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1171"
arrKeyPath(x) = "Software\Microsoft\Windows NT\CurrentVersion\Winlogon"
arrSubKey(x) = "AllocateDASD"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "String"
arrTitle(x) = "Format and Eject Removable Media"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-1172"
arrKeyPath(x) = "Software\Microsoft\Windows NT\CurrentVersion\Winlogon"
arrSubKey(x) = "PasswordExpiryWarning"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "14"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueGreater"
arrTitle(x) = "Password Expiration Warning"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-1173"
arrKeyPath(x) = "System\CurrentControlSet\Control\Session Manager"
arrSubKey(x) = "ProtectionMode"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "Value"
arrTitle(x) = "Global System Objects Permission Strength"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-1174"
arrKeyPath(x) = "System\CurrentControlSet\Services\LanManServer\Parameters"
arrSubKey(x) = "AutoDisconnect"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "15"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueLess"
arrTitle(x) = "Idle Time Before Suspending a Session"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-2372"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*ClearTextPassword.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Reversible Password Encryption"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-2374"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
arrSubKey(x) = "NoDriveTypeAutorun"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "255"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "Value"
arrTitle(x) = "Disable Media Autoplay"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-2907"
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
arrTitle(x) = "System File Changes"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-2908"
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
arrTitle(x) = "Unencrypted Remote Access"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-3245"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "Everyone"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileSharePerms"
arrTitle(x) = "File Share ACLs"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-3289"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "McAfeeFramework"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "HbssClient"
arrTitle(x) = "Intrusion Detection System"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-3337"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*LSAAnonymousNameLookup.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Anonymous SID/Name Translation"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-3338"
arrKeyPath(x) = "System\CurrentControlSet\Services\LanManServer\Parameters"
arrSubKey(x) = "NullSessionPipes"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ""
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "RegMultiBlank"
arrTitle(x) = "Anonymous Access to Named Pipes"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-3339"
arrKeyPath(x) = "System\CurrentControlSet\Control\SecurePipeServers\Winreg\AllowedExactPaths"
arrSubKey(x) = "Machine"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "System\CurrentControlSet\Control\ProductOptions"
arrExpectedValue2(x) = "System\CurrentControlSet\Control\Server Applications"
arrExpectedValue3(x) = "Software\Microsoft\Windows NT\CurrentVersion"
arrType(x) = "RegRemotePath"
arrTitle(x) = "Remotely Access Registry Paths"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-3340"
arrKeyPath(x) = "System\CurrentControlSet\Services\LanManServer\Parameters"
arrSubKey(x) = "NullSessionShares"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "RegMultiBlank"
arrTitle(x) = "Anonymous Access to Network Shares"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-3343"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows NT\Terminal Services"
arrSubKey(x) = "fAllowToGetHelp"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Remote Assistance - Solicit Remoet Assistance"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-3344"
arrKeyPath(x) = "System\CurrentControlSet\Control\Lsa"
arrSubKey(x) = "LimitBlankPasswordUse"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "Value"
arrTitle(x) = "Limit Blank Passwords"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-3372"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
arrSubKey(x) = "UndockWithoutLogon"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "Value"
arrTitle(x) = "Undock Without Logging On"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-3373"
arrKeyPath(x) = "System\CurrentControlSet\Services\Netlogon\Parameters"
arrSubKey(x) = "MaximumPasswordAge"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "30"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueLess"
arrTitle(x) = "Maximum Machine Account Password Age"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-3374"
arrKeyPath(x) = "System\CurrentControlSet\Services\Netlogon\Parameters"
arrSubKey(x) = "RequireStrongKey"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "Value"
arrTitle(x) = "Strong Session Key"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-3376"
arrKeyPath(x) = "System\CurrentControlSet\Control\Lsa"
arrSubKey(x) = "DisableDomainCreds"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "Value"
arrTitle(x) = "Strong of Passwords and Credentials"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-3377"
arrKeyPath(x) = "System\CurrentControlSet\Control\Lsa"
arrSubKey(x) = "EveryoneIncludesAnonymous"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "Value"
arrTitle(x) = "Everyone Permissions Apply to Anonymous"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-3378"
arrKeyPath(x) = "System\CurrentControlSet\Control\Lsa"
arrSubKey(x) = "ForceGuest"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "Value"
arrTitle(x) = "Sharing and Security Model for Local Accounts"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-3379"
arrKeyPath(x) = "System\CurrentControlSet\Control\Lsa"
arrSubKey(x) = "NoLMHash"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "Value"
arrTitle(x) = "LAN Manager Hash Value Stored"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-3380"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*ForceLogOffWhenHourExpire.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Force Logoff When Logon Hours Expire"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-3381"
arrKeyPath(x) = "System\CurrentControlSet\Services\LDAP"
arrSubKey(x) = "LDAPClientIntegrity"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueGreater"
arrTitle(x) = "LDAP Client Signing"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-3382"
arrKeyPath(x) = "System\CurrentControlSet\Control\Lsa\MSV1_0"
arrSubKey(x) = "NTLMMinClientSec"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "537395200"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "Value"
arrTitle(x) = "Session Security for NTLM SSP Based Clients"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-3383"
arrKeyPath(x) = "System\CurrentControlSet\Control\Lsa\FIPSAlgorithmPolicy"
arrSubKey(x) = "Enabled"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "Value"
arrTitle(x) = "FIPS Compliant Algorithms"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-3385"
arrKeyPath(x) = "System\CurrentControlSet\Control\Session Manager\Kernel"
arrSubKey(x) = "ObCaseInsensitive"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "Value"
arrTitle(x) = "Case Insensitivity for Non-Windows"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-3449"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows NT\Terminal Services"
arrSubKey(x) = "fSingleSessionPerUser"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "TS/RDS - Session Limit"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-3453"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows NT\Terminal Services"
arrSubKey(x) = "fPromptForPassword"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "TS/RDS - Password Prompting"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-3454"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows NT\Terminal Services"
arrSubKey(x) = "MinEncryptionLevel"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "TS/RDS - Set Encryption Level"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-3455"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows NT\Terminal Services"
arrSubKey(x) = "PerSessionTempDir"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "TS/RDS - Do Not Use Temp Folders"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-3456"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows NT\Terminal Services"
arrSubKey(x) = "DeleteTempDirsOnExit"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "TS/RDS - Delete Temp Folders"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-3457"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows NT\Terminal Services"
arrSubKey(x) = "MaxDisconnectionTime"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "60000"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueLessEx"
arrTitle(x) = "TS/RDS - Time Limit for Disc."
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-3458"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows NT\Terminal Services"
arrSubKey(x) = "MaxIdleTime"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "900000"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueLessEx"
arrTitle(x) = "TS/RDS - Time Limit for Idle Session"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-3469"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
arrSubKey(x) = "DisableBkGndGroupPolicy"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "Value"
arrTitle(x) = "Group Policy - Do Not Turn Off Background Refresh"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-3470"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows NT\Terminal Services"
arrSubKey(x) = "fAllowUnsolicited"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Remote Assistance - Offer Remote Assistance"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-3471"
arrKeyPath(x) = "Software\Policies\Microsoft\PCHealth\ErrorReporting"
arrSubKey(x) = "DoReport"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Error Reporting - Repot Errors"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-3472"
arrKeyPath(x) = "Software\Policies\Microsoft\W32Time\Parameters"
arrSubKey(x) = "Ntpserver"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*time.*windows.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueExReg"
arrTitle(x) = "Windows Time Service - Configure NTP Client"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-3479"
arrKeyPath(x) = "System\CurrentControlSet\Control\Session Manager"
arrSubKey(x) = "SafeDllSearchMode"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Safe DLL Search Mode"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-3480"
arrKeyPath(x) = "Software\Policies\Microsoft\WindowsMediaPlayer"
arrSubKey(x) = "DisableAutoUpdate"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Media Player - Disable Automatic Updates"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-3481"
arrKeyPath(x) = "Software\Policies\Microsoft\WindowsMediaPlayer"
arrSubKey(x) = "PreventCodecDownload"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Media Player - Prevent Codec Download"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-3487"
arrKeyPath(x) = "Software"
arrSubKey(x) = "Default"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ServiceState"
arrTitle(x) = "Unnecessary Services"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-3491"
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
arrTitle(x) = "Reviewing Audit Logs"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-3666"
arrKeyPath(x) = "System\CurrentControlSet\Control\Lsa\MSV1_0"
arrSubKey(x) = "NTLMMinServerSec"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "537395200"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "Value"
arrTitle(x) = "Session Security for NTLM SSP based Servers"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-3828"
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
arrTitle(x) = "Security-Related Software Patches"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-4108"
arrKeyPath(x) = "System\CurrentControlSet\Services\Eventlog\Security"
arrSubKey(x) = "WarningLevel"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "90"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueLessExZero"
arrTitle(x) = "Audit Log Warning Level"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-4110"
arrKeyPath(x) = "System\CurrentControlSet\Services\Tcpip\Parameters"
arrSubKey(x) = "DisableIPSourceRouting"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "2"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Disable IP Source Routing"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-4111"
arrKeyPath(x) = "System\CurrentControlSet\Services\Tcpip\Parameters"
arrSubKey(x) = "EnableICMPRedirect"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Disable ICMP Redirect"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-4112"
arrKeyPath(x) = "System\CurrentControlSet\Services\Tcpip\Parameters"
arrSubKey(x) = "PerformRouterDiscovery"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Disable Router Discovery"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-4113"
arrKeyPath(x) = "System\CurrentControlSet\Services\Tcpip\Parameters"
arrSubKey(x) = "KeepAliveTime"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "300000"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueLessExZero"
arrTitle(x) = "TCP Connection Keep-Alive Time"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-4116"
arrKeyPath(x) = "System\CurrentControlSet\Services\Netbt\Parameters"
arrSubKey(x) = "NoNameReleaseOnDemand"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Name-Release Attacks"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-4438"
arrKeyPath(x) = "System\CurrentControlSet\Services\Tcpip\Parameters"
arrSubKey(x) = "TcpMaxDataRetransmissions"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "TCP Data Retransmissions"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-4442"
arrKeyPath(x) = "Software\Microsoft\Windows NT\CurrentVersion\Winlogon"
arrSubKey(x) = "ScreenSaverGracePeriod"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "5"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "StringValueExLess"
arrTitle(x) = "Screen Saver Grace Period"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-4443"
arrKeyPath(x) = "System\CurrentControlSet\Control\SecurePipeServers\Winreg\AllowedPaths"
arrSubKey(x) = "Machine"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "5"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "RemoteReg"
arrTitle(x) = "Remotely Accessible Registry Paths and Sub-Paths"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-4444"
arrKeyPath(x) = "Software\Policies\Microsoft\Cryptography"
arrSubKey(x) = "ForceKeyProtection"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "2"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Strong Key Protection"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-4445"
arrKeyPath(x) = "System\CurrentControlSet\Control\Session Manager\Subsystems"
arrSubKey(x) = "Optional"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "2"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "RegMultiBlank"
arrTitle(x) = "Optional Subsystems"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-4446"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers"
arrSubKey(x) = "AuthenticodeEnabled"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Software Restriction Policies"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-4447"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows NT\Terminal Services"
arrSubKey(x) = "fEncryptRPCTraffic"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "TS/RDS - Secure RPC Connection"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-4448"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\Group Policy\{35378EAC-683F-11D2-A89A-00C04FBBCFA2}"
arrSubKey(x) = "NoGPOListChanges"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Group Policy - Registry Policy Processing"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-6831"
arrKeyPath(x) = "System\CurrentControlSet\Services\Netlogon\Parameters"
arrSubKey(x) = "RequireSignOrSeal"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Encrypting and Signing of Secure Channel Traffic"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-6832"
arrKeyPath(x) = "System\CurrentControlSet\Services\LanmanWorkstation\Parameters"
arrSubKey(x) = "RequireSecuritySignature"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "SMB Client Packet Signing"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-6833"
arrKeyPath(x) = "System\CurrentControlSet\Services\LanmanServer\Parameters"
arrSubKey(x) = "RequireSecuritySignature"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "SMB Server Packet Signing (Always)"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-6834"
arrKeyPath(x) = "System\CurrentControlSet\Services\LanmanServer\Parameters"
arrSubKey(x) = "RestrictNullSessAccess"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Anonymous Access to Named Pipes and Shares"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-6836"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*MinimumPasswordLength.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Minimum Password Length"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-6840"
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
arrType(x) = "AccountExpires"
arrTitle(x) = "Password Expiration"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-7002"
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
arrType(x) = "PasswordRequired"
arrTitle(x) = "Password Requirement"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-11806"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
arrSubKey(x) = "DontDisplayLastUserName"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Display of Last User Name"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-14224"
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
arrTitle(x) = "Backup Administrator Account"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-14225"
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
arrTitle(x) = "Administrator Account Passsword Changes"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-14226"
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
arrTitle(x) = "Archiving Audit Logs"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-14228"
arrKeyPath(x) = "System\CurrentControlSet\Control\Lsa"
arrSubKey(x) = "AuditBaseObjects"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Audit Access to Global System Objects"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-14229"
arrKeyPath(x) = "System\CurrentControlSet\Control\Lsa"
arrSubKey(x) = "FullPrivilegeAuditing"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "BinaryEx"
arrTitle(x) = "Audit Backup and Restore Privileges"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-14230"
arrKeyPath(x) = "System\CurrentControlSet\Control\Lsa"
arrSubKey(x) = "SCENoApplyLegacyAuditPolicy"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Audit Policy Subcategory Setting"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-14232"
arrKeyPath(x) = "System\CurrentControlSet\Services\IPSEC"
arrSubKey(x) = "NoDefaultExempt"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "IPSec Exemptions"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-14234"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
arrSubKey(x) = "FilterAdministratorToken"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "UAC - Admin Approval Mode"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-14235"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
arrSubKey(x) = "ConsentPromptBehaviorAdmin"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "4"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueLessExZero"
arrTItle(x) = "UAC - Admin Elevation Prompt"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-14236"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
arrSubKey(x) = "ConsentPromptBehaviorUser"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "UAC - User Elevation Prompt"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-14237"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
arrSubKey(x) = "EnableInstallerDetection"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "UAC - Application Installations"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-14239"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
arrSubKey(x) = "EnableSecureUIAPaths"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "UAC - UIAccess Application Elevation"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-14240"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
arrSubKey(x) = "EnableLUA"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "UAC - All Admin Approval Mode"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-14241"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
arrSubKey(x) = "PromptOnSecureDesktop"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "UAC - Secure Desktop Mode"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-14242"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
arrSubKey(x) = "EnableVirtualization"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "UAC - Non UAC Compliant Application Virtualization"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-14243"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\CredUI"
arrSubKey(x) = "EnumerateAdministrators"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Enumerate Administrator Accounts on Elevation"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-14247"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows NT\Terminal Services"
arrSubKey(x) = "DisablePasswordSaving"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "TS/RDS - Prevent Password Saving"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-14249"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows NT\Terminal Services"
arrSubKey(x) = "fDisableCdm"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "TS/RDS - Drive Redirection"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-14250"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\WindowsUpdate\AU"
arrSubKey(x) = "NoAutoUpdate"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Configure Automatic Updates"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-14253"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows NT\Rpc"
arrSubKey(x) = "RestrictRemoteClients"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "RPC - Unauthenticated RPC Clients"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-14254"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows NT\Rpc"
arrSubKey(x) = "EnableAuthEpResolution"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "RPC - Endpoint Mapper"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-14255"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
arrSubKey(x) = "NoPublishingWizard"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Publish to Web"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-14256"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
arrSubKey(x) = "NoWebServices"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Internet Downlaod / Online Ordering"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-14257"
arrKeyPath(x) = "Software\Policies\Microsoft\Messenger\Client"
arrSubKey(x) = "CEIP"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "2"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Windows Messenger Experience Improvement"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-14258"
arrKeyPath(x) = "Software\Policies\Microsoft\SearchCompanion"
arrSubKey(x) = "DisableContentFileUpdates"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Search Companion Content File Updates"
arrCategory(x) = "II"
arrTitle(x) = "Printing over HTTP will be prevented"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-14259"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows NT\Printers"
arrSubKey(x) = "DisableHTTPPrinting"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Printing Over HTTP"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-14260"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows NT\Printers"
arrSubKey(x) = "DisableWebPnPDownload"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "HTTP Printer Drivers"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-14261"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\DriverSearching"
arrSubKey(x) = "DontSearchWindowsUpdate"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Windows Update Device Drive Searching"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-14262"
arrKeyPath(x) = "System\CurrentControlSet\Services\Tcpip6\Parameters"
arrSubKey(x) = "DisabledComponents"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "-1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "IPv6 Transition"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-14268"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\Attachments"
arrSubKey(x) = "SaveZoneInformation"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKCU
arrExpectedValue(x) = "2"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Attachment Mgr - Preserve Zone Info"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-14269"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\Attachments"
arrSubKey(x) = "HideZoneInfoOnProperties"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Attachment Mgr - Hide Mech to Remove Zone Info"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-14270"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\Attachments"
arrSubKey(x) = "ScanWithAntiVirus"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKCU
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Attachment Mgr - Scan with Antivirus"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-14271"
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
arrType(x) = "ApplicationPassword"
arrTitle(x) = "Application Account Passwords"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-15505"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "McAfeeFramework"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "HbssClient"
arrTitle(x) = "HBSS McAfee Agent"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-15666"
arrKeyPath(x) = "Software\Policies\Microsoft\Peernet"
arrSubKey(x) = "Disabled"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Windows Peer to Peer Networking"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-15667"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\Network Connections"
arrSubKey(x) = "NC_AllowNetBridge_NLA"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Prohibit Network Bridge"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-15671"
arrKeyPath(x) = "Software\Policies\Microsoft\SystemCertificates\AuthRoot"
arrSubKey(x) = "DisableRootAutoUpdate"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Root Certificate Update"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-15672"
arrKeyPath(x) = "Software\Policies\Microsoft\EventViewer"
arrSubKey(x) = "MicrosoftEventVwrDisableLinks"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Event Viewer Events.asp Links"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-15673"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\Internet Connection Wizard"
arrSubKey(x) = "ExitOnMSICW"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Internet Connection Wizard ISP Downloads"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-15674"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
arrSubKey(x) = "NoInternetOpenWith"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Internet File Association Service"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-15675"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\Registration Wizard Control"
arrSubKey(x) = "NoRegistration"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Windows Registration Wizard"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-15676"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
arrSubKey(x) = "NoOnlinePrintsWizard"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Order Prints Online"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-15680"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\system"
arrSubKey(x) = "LogonType"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Classic Logon"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-15682"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Feeds"
arrSubKey(x) = "DisableEnclosureDownload"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "RSS Attachment Downloads"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-15683"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
arrSubKey(x) = "PreXPSP2ShellProtocolBehavior"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Window Explorer - Shell Protocol Protected Mode"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-15684"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\Installer"
arrSubKey(x) = "SafeForScripting"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "WIndows Installer - IE Security Prompt"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-15685"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\Installer"
arrSubKey(x) = "EnableUserControl"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Windows Installer - User Control"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-15686"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\Installer"
arrSubKey(x) = "DisableLUAPatching"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Windows Installer - Vendor Signed Updates"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-15687"
arrKeyPath(x) = "Software\Policies\Microsoft\WindowsMediaPlayer"
arrSubKey(x) = "GroupPrivacyAcceptance"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Media Player - First Use Dialog Boxes"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-15696"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\LLTD"
arrSubKey(x) = "AllowLLTDIOOndomain"
arrSubKey2(x) = "AllowLLTDIOOnPublicNet"
arrSubkey3(x) = "EnableLLTDIO"
arrFile1(x) = "ProhibitLLTDIOOnPrivateNet"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "QuadValueEx"
arrTitle(x) = "Network - Mapper I/O Driver"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-15697"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\LLTD"
arrSubKey(x) = "AllowRspndrOndomain"
arrSubKey2(x) = "AllowRspndrOnPublicNet"
arrSubkey3(x) = "EnableRspndr"
arrFile1(x) = "ProhibitRspndrOnPrivateNet"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "QuadValueEx"
arrTitle(x) = "Network - Responder Driver"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-15698"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\WCN\Registrars"
arrSubKey(x) = "DisableFlashConfigRegistrar"
arrSubKey2(x) = "DisableInBand802DOT11Registrar"
arrSubkey3(x) = "DisableUPnPRegistrar"
arrFile1(x) = "DisableWPDRegistrar"
arrFile2(x) = "EnableRegistrars"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "QuintValueEx"
arrTitle(x) = "Network - WCN Wireless Configuration"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-15699"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\WCN\UI"
arrSubKey(x) = "DisableWcnUi"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Network - Windows Connect Now wizards will be disabled"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-15700"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\DeviceInstall\Settings"
arrSubKey(x) = "AllowRemoteRPC"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Device install - PnP Interface remote Access"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-15701"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\DeviceInstall\Settings"
arrSubKey(x) = "DisableSystemRestore"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Device Install - Drivers System Restore Point"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-15702"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\DeviceInstall\Settings"
arrSubKey(x) = "DisableSendGenericDriverNotFoundToWER"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Device Install - Generic Driver Error Report"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-15703"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\DriverSearching"
arrSubKey(x) = "DontPromptForWindowsUpdate"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Driver Install - Device Driver"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-15704"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\HandwritingErrorReports"
arrSubKey(x) = "PreventHandwritingErrorReports"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Handwriting Recognition Error Reporting"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-15705"
arrKeyPath(x) = "Software\Policies\Microsoft\Power\PowerSettings\0e796bdb-100d-47d6-a2d5-f7d2daa51f51"
arrSubKey(x) = "DCSettingIndex"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Power Mgmt - Password Wake on Battery"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-15706"
arrKeyPath(x) = "Software\Policies\Microsoft\Power\PowerSettings\0e796bdb-100d-47d6-a2d5-f7d2daa51f51"
arrSubKey(x) = "ACSettingIndex"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Power Mgmt - Password Wkae When Plugged In"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-15707"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows NT\Terminal Services"
arrSubKey(x) = "LoggingEnabled"
arrSubKey2(x) = "0"
arrSubkey3(x) = "0"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Remote Assistance - Session Logging"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-15709"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\GameUX"
arrSubKey(x) = "DownloadGameInfo"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Game Explorer Information Downloads"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-15713"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows Defender\SpyNet"
arrSubKey(x) = "SpyNetReporting"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "Value"
arrTitle(x) = "Defender - SpyNet Reporting"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-15714"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\Windows Error Reporting"
arrSubKey(x) = "LoggingDisabled"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Error Reporting - Logging"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-15715"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\Windows Error Reporting"
arrSubKey(x) = "Disabled"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Error Reporting - Windows Error Reporting"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-15717"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\Windows Error Reporting"
arrSubKey(x) = "DontSendAdditionalData"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Error Reporting - Additional Data"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-15718"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\Explorer"
arrSubKey(x) = "NoHeapTerminationOnCorruption"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Windows Explorer - Heap Termination"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-15719"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
arrSubKey(x) = "ReportControllerMissing"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Logon - Report Logon Server"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-15722"
arrKeyPath(x) = "Software\Policies\Microsoft\WMDRM"
arrSubKey(x) = "DisableOnline"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Media DRM - Internet Access"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-15727"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
arrSubKey(x) = "NoInPlaceSharing"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "User Network Sharing"
arrCategory(x) = "II"

if globalFind = "no" Then
x = x + 1
arrFinding(x) = "V-15823"
arrKeyPath(x) = "Default"
arrSubKey(x) = "0"
arrSubKey2(x) = "0"
arrSubkey3(x) = "0"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*\.pfx.*"
arrExpectedValue2(x) = ".*\.p12.*"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Software Certificate Installation Files"
arrCategory(x) = "II"
Else
x = x + 1
arrFinding(x) = "V-15823"
arrKeyPath(x) = "Default"
arrSubKey(x) = "0"
arrSubKey2(x) = "0"
arrSubkey3(x) = "0"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*\.pfx.*"
arrExpectedValue2(x) = ".*\.p12.*"
arrExpectedValue3(x) = "0"
arrType(x) = "DirPFX"
arrTitle(x) = "Software Certificate Installation Files"
arrCategory(x) = "II"
End If

x = x + 1
arrFinding(x) = "V-15991"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
arrSubKey(x) = "EnableUIADesktopToggle"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "UAC - UIAccess Secure Desktop"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-15996"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows NT\Terminal Services"
arrSubKey(x) = "fDisableClip"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "TS/RDS - Clipboard Redirection"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-15997"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows NT\Terminal Services"
arrSubKey(x) = "fDisableCcm"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "TS/RDS - COM Port Redirection"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-15998"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows NT\Terminal Services"
arrSubKey(x) = "fDisableLPT"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "TS/RDS - LPT Port Redirecton"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-15999"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows NT\Terminal Services"
arrSubKey(x) = "fDisablePNPRedir"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "TS/RDS - PNP Device Redirection"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-16000"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows NT\Terminal Services"
arrSubKey(x) = "fEnableSmartCard"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "TS/RDS - Smart Card Device Redirection"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-16001"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows NT\Terminal Services"
arrSubKey(x) = "RedirectOnlyDefaultClientPrinter"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "TS/RDS - Printer Redirection"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-16005"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
arrSubKey(x) = "NoDisconnect"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "TS/RDS - Remove Disconnect Option"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-16006"
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
arrTitle(x) = "Unnecessary Features Installed"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-16008"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
arrSubKey(x) = "ValidateAdminCodeSignatures"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "UAC - Appliation Elevations"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-16020"
arrKeyPath(x) = "Software\Policies\Microsoft\SQMClient\Windows"
arrSubKey(x) = "CEIPEnable"
arrSubKey2(x) = "0"
arrSubkey3(x) = "0"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Windows Customer Experience Improvement Program"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-16021"
arrKeyPath(x) = "Software\Policies\Microsoft\Assistance\Client\1.0"
arrSubKey(x) = "NoImplicitFeedback"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Help Experience Improvement Programm"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-16048"
arrKeyPath(x) = "Software\Policies\Microsoft\Assistance\Client\1.0"
arrSubKey(x) = "NoExplicitFeedback"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Help Ratings"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-18010"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeDebugPrivilege.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "User Right - Debug Programs"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-21950"
arrKeyPath(x) = "System\CurrentControlSet\Services\LanmanServer\Parameters"
arrSubKey(x) = "SmbServerNameHardeningLevel"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "SPN Target Name Validation Level"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-21951"
arrKeyPath(x) = "System\CurrentControlSet\Control\LSA"
arrSubKey(x) = "UseMachineId"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Computer Identity Authentication for NTLM"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-21952"
arrKeyPath(x) = "System\CurrentControlSet\Control\LSA\MSV1_0"
arrSubKey(x) = "allownullsessionfallback"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "NTML Null Session Fallback"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-21953"
arrKeyPath(x) = "System\CurrentControlSet\Control\LSA\pku2u"
arrSubKey(x) = "AllowOnlineID"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "PKU2U Online Identities Authentication"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-21954"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\System\Kerberos\Parameters"
arrSubKey(x) = "SupportedEncryptionTypes"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "2147483644"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Kerberos Encryption Types"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-21955"
arrKeyPath(x) = "System\CurrentControlSet\Services\Tcpip6\Parameters"
arrSubKey(x) = "DisableIpSourceRouting"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "2"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "IPv6 Source Routing"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-21956"
arrKeyPath(x) = "System\CurrentControlSet\Services\Tcpip6\Parameters"
arrSubKey(x) = "TcpMaxDataRetransmissions"
arrSubKey2(x) = "0"
arrSubkey3(x) = "0"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueLessEx"
arrTitle(x) = "IPv6 TCP Data Retransmissions"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-21960"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\Network Connections"
arrSubKey(x) = "NC_StdDomainUserSetLocation"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Elevate When Setting a Network's Location"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-21961"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\TCPIP\v6Transition"
arrSubKey(x) = "Force_Tunneling"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "Enabled"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "StringEx"
arrTitle(x) = "Direct Access - Route Through Internal Networks"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-21963"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows NT\Printers"
arrSubKey(x) = "DoNotInstallCompatibleDriverFromWindowsUpdate"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Windows Updated Point and Print Driver Search"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-21964"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\Device Metadata"
arrSubKey(x) = "PreventDeviceMetadataFromNetwork"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Prevent Device Metadata Retrieval from Internet"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-21965"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\DriverSearching"
arrSubKey(x) = "SearchOrderConfig"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Prevent Windows Update for Device Driver Search"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-21967"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\ScriptedDiagnosticsProvider\Policy"
arrSubKey(x) = "DisableQueryRemoteServer"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "MSDT Interactive Communication"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-21969"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\ScriptedDiagnosticsProvider\Policy"
arrSubKey(x) = "EnableQueryRemoteServer"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Windows Online Troubleshooting Service"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-21970"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\WDI\{9c5a40da-b965-4fc3-8781-88dd50a6299d}"
arrSubKey(x) = "ScenarioExecutionEnabled"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Disable PerfTrack"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-21971"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\AppCompat"
arrSubKey(x) = "DisableInventory"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Application Compatibility Program Inventory"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-21973"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\Explorer"
arrSubKey(x) = "NoAutoplayfornonVolume"
arrSubKey2(x) = "0"
arrSubkey3(x) = "0"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Autoplay for non-volume devices"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-21974"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\GameUX"
arrSubKey(x) = "GameUpdateOptions"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Turn Off Game Updates"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-21975"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\Homegroup"
arrSubKey(x) = "DisableHomeGroup"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Prevent Joining Homegroup"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-21978"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\WAU"
arrSubKey(x) = "Disabled"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Windows Anytime Upgrade"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-21980"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\Explorer"
arrSubKey(x) = "NoDataExecutionPrevention"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Explorer Data Execution Prevention"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-22692"
arrKeyPath(x) = "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer"
arrSubKey(x) = "NoAutorun"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Default Autorun Behavior"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-26070"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\ScriptedDiagnosticsProvider\Policy"
arrSubKey(x) = "EnableQueryRemoteServer"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "WinLogon"
arrTitle(x) = "Winlogon Registry Permissions"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-26283"
arrKeyPath(x) = "System\CurrentControlSet\Control\Lsa"
arrSubKey(x) = "RestrictAnonymousSAM"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Restrict Anonymous SAM Enumeration"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-26359"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
arrSubKey(x) = "LegalNoticeCaption"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "US Department of Defense Warning Statement"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "String"
arrTitle(x) = "Legal Banner Dialog Box Title"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-26395 - IAVA 2011-B-0041"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*login\.live\.com.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "UntrustedCertificates"
arrTitle(x) = "Microsoft Windows Fraudulent Digital Certificate Vulnerability"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26469"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "0"
arrSubkey3(x) = "0"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeTrustedCredManAccessPrivilege.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Access Credential Manager as a trusted caller"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26470"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeDenyNetworkLogonRight.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Access this comptuer from the network"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26471"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeIncreaseQuotaPrivilege.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Adjust Memory Quotas for a Process"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26472"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeInteractiveLogonRight.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Allow Log on Locally"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26473"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeRemoteInteractiveLogonRight.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Allow log on through Remote Desktop Services"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26474"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeBackupPrivilege.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Back up files and directories"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26475"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeChangeNotifyPrivilege.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Bypass Traverse Checking"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-26476"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeSystemtimePrivilege.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Change the System Time"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26477"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeTimeZonePrivilege.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Change the Time Zone"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-26478"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeCreatePagefilePrivilege.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Create a Pagefile"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26479"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "0"
arrSubkey3(x) = "0"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeCreateTokenPrivilege.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Create a Token Object"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-26480"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeCreateGlobalPrivilege.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Create Global Objects"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26481"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeCreatePermanentPrivilege.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Create Permanent Shared Objects"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26482"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeCreateSymbolicLinkPrivilege.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Create Symbolic Links"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26483"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeDenyBatchLogonRight.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Deny Log on as a Batch Job"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26484"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeDenyServiceLogonRight.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ManualReview"
arrTitle(x) = "Deny Log on as a Service"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26485"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeDenyInteractiveLogonRight.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Deny Log on Locally"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26486"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeDenyRemoteInteractiveLogonRight.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Deny Log on Through Remote Desktop Services"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26487"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeEnableDelegationPrivilege.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Enable Accounts to be Trusted for Delegation"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26488"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeRemoteShutdownPrivilege.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Force Shutdown from a Remote System"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26489"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "0"
arrSubkey3(x) = "0"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeAuditPrivilege.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Generate Security Audits"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26490"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeImpersonatePrivilege.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Impersonate a Client After Authentication"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26491"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeIncreaseWorkingSetPrivilege.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Increase a Process Working Set"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26492"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeIncreaseBasePriorityPrivilege.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Increase Scheduling Priority"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26493"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeLoadDriverPrivilege.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Load and Unload Device Drivers"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26494"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeLoadDriverPrivilege.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Lock Pages in Memory"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26495"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeBatchLogonRight.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Log on as a Batch Job"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26496"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeSecurityPrivilege.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Manage Auditing and Security Logs"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26497"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeRelabelPrivilege.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Modify an Object Label"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26498"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeSystemEnvironmentPrivilege.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Modify Firmware Environment Values"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26499"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "0"
arrSubkey3(x) = "0"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeManageVolumePrivilege.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Perform Volume Maintenance Tasks"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26500"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeProfileSingleProcessPrivilege.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Profile Single Process"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26501"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeSystemProfilePrivilege.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Profile System Performance"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26502"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeUndockPrivilege.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Remove Computer from Docking Station"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-26503"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeAssignPrimaryTokenPrivilege.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Replace a Process Level Token"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26504"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeRestorePrivilege.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Restore Files and Directories"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26505"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeShutdownPrivilege.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Shut Down the System"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26506"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Mismatch.*SeTakeOwnershipPrivilege.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "SecEditRegExp"
arrTitle(x) = "Take Ownership of Files or Other Objects"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26529"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Credential.*Validation.*Success.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "AuditRegExp"
arrTitle(x) = "Audit - Credential Validation - Success"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26530"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Credential.*Validation.*Failure.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "AuditRegExp"
arrTitle(x) = "Audit - Credential Validation - Failure"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26531"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "0"
arrSubkey3(x) = "0"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Computer.*Account.*Management.*Success.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "AuditRegExp"
arrTitle(x) = "Audit - Computer Account Management - Success"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26532"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Computer.*Account.*Management.*Failure.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "AuditRegExp"
arrTitle(x) = "Audit - Computer Account Management - Failure"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26533"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Other.*Account.*Management.*Events.*Success.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "AuditRegExp"
arrTitle(x) = "Audit - Other Account Management Events - Success"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26534"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Other.*Account.*Management.*Events.*Failure.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "AuditRegExp"
arrTitle(x) = "Audit - Other Account Management Events - Failure"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26535"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Security.*Group.*Management.*Success.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "AuditRegExp"
arrTitle(x) = "Audit - Security Group Management - Success"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26536"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Security.*Group.*Management.*Failure.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "AuditRegExp"
arrTitle(x) = "Audit - Security Group Management - Failure"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26537"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*User.*Account.*Management.*Success.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "AuditRegExp"
arrTitle(x) = "Audit - User Account Management - Success"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26538"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*User.*Account.*Management.*Failure.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "AuditRegExp"
arrTitle(x) = "Audit - User Account Management - Failure"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26539"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Process.*Creation.*Success.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "AuditRegExp"
arrTitle(x) = "Audit - Process Creation - Success"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26540"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Logoff.*Success.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "AuditRegExp"
arrTitle(x) = "Audit - Logoff - Success"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26541"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "0"
arrSubkey3(x) = "0"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Logon.*Success.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "AuditRegExp"
arrTitle(x) = "Audit - Logon - Success"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26542"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Logon.*Failure.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "AuditRegExp"
arrTitle(x) = "Audit - Logon - Failure"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26543"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Special.*Logon.*Success.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "AuditRegExp"
arrTitle(x) = "Audit - Special Logon - Success"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26544"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*File.*System.*Failure.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "AuditRegExp"
arrTitle(x) = "Audit - File System - Failure"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26545"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Registry.*Failure.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "AuditRegExp"
arrTitle(x) = "Audit - Registry - Failure"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26546"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Audit.*Policy.*Change.*Success.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "AuditRegExp"
arrTitle(x) = "Audit - Audit Policy Change - Success"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26547"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Audit.*Policy.*Change.*Failure.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "AuditRegExp"
arrTitle(x) = "Audit - Audit Policy Change - Failure"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26548"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Authentication.*Policy.*Change.*Success.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "AuditRegExp"
arrTitle(x) = "Audit - Authentication Policy Change - Success"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26549"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Sensitive.*Privilege.*Use.*Success.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "AuditRegExp"
arrTitle(x) = "Audit - Sensitive Privilege Use - Success"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26550"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Sensitive.*Privilege.*Use.*Failure.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "AuditRegExp"
arrTitle(x) = "Audit - Sensitive Privilege Use - Failure"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26551"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "0"
arrSubkey3(x) = "0"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*IPSec.*Driver.*Success.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "AuditRegExp"
arrTitle(x) = "Audit - IPSec Driver - Success"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26552"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*IPSec.*Driver.*Failure.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "AuditRegExp"
arrTitle(x) = "Audit - IPSec Driver- Fialure"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26553"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Security.*State.*Change.*Success.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "AuditRegExp"
arrTitle(x) = "Audit - Security State Change - Success"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26554"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Security.*State.*Change.*Failure.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "AuditRegExp"
arrTitle(x) = "Audit - Security State Change - Failure"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26555"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Security.*System.*Extension.*Success.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "AuditRegExp"
arrTitle(x) = "Audit - Security System Extension - Success"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26556"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Security.*System.*Extension.*Failure.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "AuditRegExp"
arrTitle(x) = "Audit - Security System Extention - Failure"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26557"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*System.*Integrity.*Success.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "AuditRegExp"
arrTitle(x) = "Audit - System Integrity - Success"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26558"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*System.*Integrity.*Failure.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "AuditRegExp"
arrTitle(x) = "Audit - System Integrity - Failure"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26575"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\TCPIP\v6Transition"
arrSubKey(x) = "6to4_State"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "Disabled"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "StringEx"
arrTitle(x) = "6to4 State"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26576"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\TCPIP\v6Transition\IPHTTPS\IPHTTPSInterface"
arrSubKey(x) = "IPHTTPS_ClientState"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "IP-HTTPS State"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26577"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\TCPIP\v6Transition"
arrSubKey(x) = "ISATAP_State"
arrSubKey2(x) = "0"
arrSubkey3(x) = "0"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "Disabled"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "StringEx"
arrTitle(x) = "ISATAP State"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26578"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\TCPIP\v6Transition"
arrSubKey(x) = "Teredo_State"
arrSubKey2(x) = "0"
arrSubkey3(x) = "0"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "Disabled"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "StringEx"
arrTitle(x) = "Teredo State"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26579"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\EventLog\Application"
arrSubKey(x) = "MaxSize"
arrSubKey2(x) = "0"
arrSubkey3(x) = "0"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "32768"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueGreaterEx"
arrTitle(x) = "Maximum Log Size - Application"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26580"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\EventLog\Security"
arrSubKey(x) = "MaxSize"
arrSubKey2(x) = "0"
arrSubkey3(x) = "0"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "196608"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueGreaterEx"
arrTitle(x) = "Maximum Log Size - Security"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26581"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\EventLog\Setup"
arrSubKey(x) = "MaxSize"
arrSubKey2(x) = "0"
arrSubkey3(x) = "0"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "32768"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueGreaterEx"
arrTitle(x) = "Maximum Log Size - Setup"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26582"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\EventLog\System"
arrSubKey(x) = "MaxSize"
arrSubKey2(x) = "0"
arrSubkey3(x) = "0"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "32768"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueGreaterEx"
arrTitle(x) = "Maximum Log Size - System"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26600"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Fax.*"
arrExpectedValue2(x) = "Disabled"
arrExpectedValue3(x) = "0"
arrType(x) = "ServiceDisabled"
arrTitle(x) = "Fax Service Disabled"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26602"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*ftpsvc.*"
arrExpectedValue2(x) = "Disabled"
arrExpectedValue3(x) = "0"
arrType(x) = "ServiceDisabled"
arrTitle(x) = "Microsoft FTP Service Disabled"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26604"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*p2pimsvc.*"
arrExpectedValue2(x) = "Disabled"
arrExpectedValue3(x) = "0"
arrType(x) = "ServiceDisabled"
arrTitle(x) = "Peer Network Identity Manager Service Disabled"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26605"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*simptcp.*"
arrExpectedValue2(x) = "Disabled"
arrExpectedValue3(x) = "0"
arrType(x) = "ServiceDisabled"
arrTitle(x) = "Simple TCP/IP Services Disabled"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-26606"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*TlntSvr.*"
arrExpectedValue2(x) = "Disabled"
arrExpectedValue3(x) = "0"
arrType(x) = "ServiceDisabled"
arrTitle(x) = "Telnet Service Disabled"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-28504"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\DeviceInstall\Settings"
arrSubKey(x) = "DisableSendRequestAdditionalSoftwareToWER"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Device Install Software Request Error Report"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V-32272"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "Cert\s+Hash\(sha1\)\:\s+8c\s+94\s+1b\s+34\s+ea\s+1e\s+a6\s+ed\s+9a\s+e2\s+bc\s+54\s+cf\s+68\s+72\s+52\s+b4\s+c9\s+b5\s+61"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "DoDCertificates"
arrTitle(x) = "The DoD Root Certificate must be installed"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-32274"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "Cert\s+Hash\(sha1\)\:\s+67\s+5d\s+2d\s+78\s+5a\s+f7\s+ee\s+66\s+50\s+b1\s+a0\s+56\s+cb\s+3f\s+82\s+fd\s+ae\s+a8\s+67\s+3e"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "UntrustedCertificates"
arrTitle(x) = "The Exeternal CA Root Certificate must be installed"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-32282"
arrKeyPath(x) = "Software\Microsoft\Active Setup\Installed Components"
arrSubKey(x) = "Software\Wow6432Node\Microsoft\Active Setup\Installed Components"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "DualRegPerms"
arrTitle(x) = "Active Setup\Installed Components Reg Perms"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-34974"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\Installer"
arrSubKey(x) = "AlwaysInstallElevated"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "The Windows Installer Always install with elevated privileges must be disabled"
arrCategory(x) = "I"

x = x + 1
arrFinding(x) = "V-36439"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
arrSubKey(x) = "LocalAccountTokenFilterPolicy"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Local admin accounts filtered token policy enabled on domain systems"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-36451"
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
arrTitle(x) = "Accounts with Administrative Privileges Internet Access"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-36663"
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
arrTitle(x) = "System BIOS or System controllers must have administrator accounts/passwords configured"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-36664"
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
arrTitle(x) = "The system must not use removable media as the boot loader"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-36669"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = ".*Handle.*Manipulation.*Failure.*"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "AuditRegExp"
arrTitle(x) = "Audit - Handle Manipulation - Failure"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-36701"
arrKeyPath(x) = "Software\Policies\Microsoft\EMET\SysSettings"
arrSubKey(x) = "ASLR"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Enhanced Mitigation Experience Tookit (EMET) Address Space Layout Randomization (ASLR)"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-36702"
arrKeyPath(x) = "Software\Policies\Microsoft\EMET\Defaults"
arrSubKey(x) = "IE"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "*\Internet Explorer\iexplore.exe"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "StringEx"
arrTitle(x) = "Enhanced Mitigation Experience Tookit (EMET) Default Protections for IE"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-36703"
arrKeyPath(x) = "Software\Policies\Microsoft\EMET\Defaults"
arrSubKey(x) = "Access"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "*\OFFICE1*\MSACCESS.EXE"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "StringEx"
arrTitle(x) = "Enhanced Mitigation Experience Tookit (EMET) Default Protections for Recommended Software"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-36704"
arrKeyPath(x) = "Software\Policies\Microsoft\EMET\Defaults"
arrSubKey(x) = "7z"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "*\7-Zip\7z.exe -EAF"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "StringEx"
arrTitle(x) = "Enhanced Mitigation Experience Tookit (EMET) Default Protections for Popular Software"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-36705"
arrKeyPath(x) = "Software\Policies\Microsoft\EMET\SysSettings"
arrSubKey(x) = "DEP"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "2"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Enhanced Mitigation Experience Tookit (EMET) DEP Application Opt Out"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-36706"
arrKeyPath(x) = "Software\Policies\Microsoft\EMET\SysSettings"
arrSubKey(x) = "SEHOP"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "2"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "ValueEx"
arrTitle(x) = "Enhanced Mitigation Experience Tookit (EMET) Structured Exception Handler Overwrite Protection (SEHOP)"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-39137"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "C:\Program Files (x86)\EMET 4.1\EMET.dll"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "FileVer"
arrTitle(x) = "The Enhanced Mitigation Experience Tookit (EMET) must be installed"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-40195"
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
arrTitle(x) = "System BIOS or system controllers must not allow user-level access"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V-40237"
arrKeyPath(x) = "Default"
arrSubKey(x) = "X"
arrSubKey2(x) = "X"
arrSubkey3(x) = "X"
arrFile1(x) = "0"
arrFile2(x) = "0"
arrFile3(x) = "0"
arrHive(x) = HKLM
arrExpectedValue(x) = "Cert\s+Hash\(sha1\)\:\s+7d\s+a8\s+e8\s+42\s+96\s+ee\s+23\s+88\s+18\s+ee\s+42\s+72\s+87\s+77\s+45\s+08\s+b2\s+6d\s+09\s+4a"
arrExpectedValue2(x) = "0"
arrExpectedValue3(x) = "0"
arrType(x) = "UntrustedCertificates"
arrTitle(x) = "The Exeternal CA Root Certificate must be installed"
arrCategory(x) = "II"

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
               strMessage = vbCrLf & "    " & testValue2(UBound(testValue1)) & " " & _
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
               If actualValue(i) < testValue2(i) Then
                  finding = "open"
                  testValue2 = split(arrFile2(intCount), "\")
                  strMessage = strMessage & vbCrLf & "    " & testValue2(UBound(testValue2)) & " " & _
                              "is version: " & fileVer & " Expected: " & arrExpectedValue2(intCount)
                  Exit For
               Elseif actualValue(i) > testValue2(i) Then
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
               If actualValue(j) < testValue1(j) Then
                  finding = "open"
                  testValue1 = split(arrFile1(intCount), "\")
                  strMessage = vbCrLf & "    " & arrSubKey(intCount) & " " & _
                            "is version: " & fileVer & " Expected: " & arrSubKey2(intCount)
               Elseif actualValue(j) >= testValue1(i) Then
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
            If actualValue(i) < testValue1(i) Then
               finding = "open"
               testValue1 = split(arrFile1(intCount), "\")
               strMessage = strMessage & vbCrLf & "    " & testValue1(UBound(testValue1)) & " " & _
                            "is version: " & fileVer & " Expected: " & arrExpectedValue(intCount)
               Exit For
            Elseif actualValue(i) > testValue1(i) Then
               Exit For
            End If
         Next
      End If
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
            If actualValue(i) < testValue2(i) Then
               finding = "open"
               testValue2 = split(arrFile2(intCount), "\")
               strMessage = strMessage & vbCrLf & "    " & testValue2(UBound(testValue2)) & " " & _
                           "is version: " & fileVer & " Expected: " & arrExpectedValue2(intCount)
               Exit For
            Elseif actualValue(i) > testValue2(i) Then
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
         If UBound(testValue2) >= UBound(actualValue) Then
            h = UBound(actualValue)
         Else
            h = UBound(testValue2)
         End If
         for i = 0 to h
            If actualValue(i) < testValue2(i) Then
               finding = "open"
               testValue2 = split(arrFile3(intCount), "\")
               strMessage = strMessage & vbCrLf & "    " & testValue2(UBound(testValue2)) & " " & _
                           "is version: " & fileVer & " Expected: " & arrExpectedValue2(intCount)
               Exit For
            Elseif actualValue(i) > testValue2(i) Then
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
         If UBound(testValue2) >= UBound(actualValue) Then
            h = UBound(actualValue)
         Else
            h = UBound(testValue2)
         End If
         for i = 0 to h
            If actualValue(i) < testValue2(i) Then
               finding = "open"
               testValue2 = split(arrSubKey3(intCount), "\")
               strMessage = strMessage & vbCrLf & "    " & testValue2(UBound(testValue2)) & " " & _
                           "is version: " & fileVer & " Expected: " & arrHive(intCount)
               Exit For
            Elseif actualValue(i) > testValue2(i) Then
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
         If UBound(testValue2) >= UBound(actualValue) Then
            h = UBound(actualValue)
         Else
            h = UBound(testValue2)
         End If
         for i = 0 to h
            If actualValue(i) < testValue2(i) Then
               finding = "open"
               testValue2 = split("C:\Program Files\Microsoft Silverlight\sllauncher.exe", "\")
               strMessage = strMessage & vbCrLf & "    " & testValue2(UBound(testValue2)) & " " & _
                           "is version: " & fileVer & " Expected: " & "5.1.10411"
               Exit For
            Elseif actualValue(i) > testValue2(i) Then
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
         If UBound(testValue2) >= UBound(actualValue) Then
            h = UBound(actualValue)
         Else
            h = UBound(testValue2)
         End If
         for i = 0 to h
            If actualValue(i) < testValue2(i) Then
               finding = "open"
               testValue2 = split("C:\Program Files\Common Files\Microsoft Shared\OFFICE12\OGL.DLL", "\")
               strMessage = strMessage & vbCrLf & "    " & testValue2(UBound(testValue2)) & " " & _
                           "is version: " & fileVer & " Expected: " & "14.0.6117.5001"
               Exit For
            Elseif actualValue(i) > testValue2(i) Then
               Exit For
            End If
         Next
      End If
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
         WScript.Echo "Test"
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
         If UBound(testValue2) >= UBound(actualValue) Then
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
                  Exit For
               Elseif Cint(actualValue(i)) > Cint(testValue1(i)) Then
                  Exit For
               End If
            Next
         ElseIf Cint(actualValue(0)) = CInt(testValue2(0)) Then
            For i = 0 to l
               If Cint(actualValue(i)) < Cint(testValue2(i)) Then
                  finding = "open"
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