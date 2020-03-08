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
arrFinding(x) = "V6228-DTBI001"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings"
arrKeyPath2(x) = "x"
arrSubKey(x) = "ProxySettingsPerUser"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ManualReview"
arrTitle(x) = "IE - THe IE home page is not set correctly"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V17296-DTBI010"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Main"
arrKeyPath2(x) = "x"
arrSubKey(x) = "DisableFirstRunCustomize"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Prevent Performance of First-Run Customization"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V6238-DTBI014"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings"
arrKeyPath2(x) = "x"
arrSubKey(x) = "SecureProtocols"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "160"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "IE SSL/TLS Settings"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V6239-DTBI015"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings"
arrKeyPath2(x) = "x"
arrSubKey(x) = "WarnOnBadCertRecving"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Publishers Certificate Revocation"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V32808-DTBI018"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\WinTrust\Trust Providers\Software Publishing"
arrKeyPath2(x) = "x"
arrSubKey(x) = "State"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "512"
arrExpectedValue2(x) = "0"
arrType(x) = "BitMaskValueEx"
arrTitle(x) = "Publishers Certificate Revocation"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V6243-DTBI022"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1001"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Download Signed Active X Controls - Internet Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V6244-DTBI023"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1004"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Download Unsigned ActiveX Controls - Internet Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V6245-DTBI024"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1201"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Initialize and Script ActiveX Controls"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V6248-DTBI030"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1604"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Font Download Control - Internet Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V6249-DTBI031"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1C00"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Java Permissions Not Set for Internet Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V6250-DTBI032"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1406"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Access Data Sources Across Domains - Internet Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V6253-DTBI036"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1802"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Drag and Drop or Copy and Paste - Internet Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V6254-DTBI037"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1800"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Installation of Desktop Items - Internet Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V6255-DTBI038"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1804"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Lauchning Programs and Files in IFRAME - Internet Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V6256-DTBI039"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1607"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Navigate Sub-Frames Across Domains - Internet Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V6259-DTBI042"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1606"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Userdata Persistence - Internet Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V6260-DTBI044"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1407"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Allow Paste Operations via Script - Internet Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V6262-DTBI046"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1A00"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "65536"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "User Authentication-Logon - Internet Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V6267-DTBI061"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1C00"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "65536"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Java Permissions Not Set - Local Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V6281-DTBI091"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1C00"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "65536"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Java Permissions Not Set - Trusted Sites Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V22687-DTBI1010"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_RESTRICT_ACTIVEXINSTALL"
arrKeyPath2(x) = "x"
arrSubKey(x) = "explorer.exe"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "StringEx"
arrTitle(x) = "Process Restrict ActiveX Install (Explorer)"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V22688-DTBI1020"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_RESTRICT_ACTIVEXINSTALL"
arrKeyPath2(x) = "x"
arrSubKey(x) = "iexplore.exe"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "StringEx"
arrTitle(x) = "IE Process Restrict ActiveX Install (IExplore)"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V6289-DTBI112"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1001"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Download Signed Active X - Restricted Sites Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V6290-DTBI113"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1004"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Download Unsigned ActiveX - Restricted Sites Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V6291-DTBI114"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1201"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Initialize and Script ActiveX - Restricted Sites Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V6292-DTBI115"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1200"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Run ActiveX Controls and Plugins - Restricted Sites Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V6293-DTBI116"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1405"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Script ActiveX Controls Marked Safe for Scripting - Restricted Sites Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V6294-DTBI119"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1803"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "File Download Control - Restricted Sites Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V6295-DTBI120"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1604"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Font Downloads Control - Restricted Sites Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V7007-DTBI121"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1C00"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Java permissions Not Set - Restricted Sites Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V6297-DTBI122"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1406"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Access Data Sources - Restricted Sites Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V6298-DTBI123"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1608"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Allow Meta Refresh - Restricted Sites Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V6301-DTBI126"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1802"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Drag and Drop or Copy and Paste - Restricted Sites Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V6302-DTBI127"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1800"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Installation of Desktop Items - Restricted Sites Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V6303-DTBI128"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1804"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Launching of Programs and Files in IFrame - Restricted Sites Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V6304-DTBI129"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1607"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Navigate Sub-Frames Across Domains - Restricted Sites Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V6307-DTBI132"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1606"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Userdata Persistence - Restricted Sites Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V6308-DTBI133"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1400"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Active Scripting - Restricted Sites Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V6309-DTBI134"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1407"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Allow Paste Operations via Scripts - Restricted Sites Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V6311-DTBI136"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1A00"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "196608"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "User Authentication-Logon - Restricted Sites Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V21887-DTBI300"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Control Panel"
arrKeyPath2(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Url History"
arrSubKey(x) = "History"
arrSubKey2(x) = "DaysToKeep"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "40"
arrType(x) = "DualValueEx"
arrTitle(x) = "History Setting is not Enabled or set to 40 days"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15490-DTBI305"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Control Panel"
arrKeyPath2(x) = "x"
arrSubKey(x) = "Autoconfig"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Automatic Configuration Not Disabled"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15492-DTBI315"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\SQM"
arrKeyPath2(x) = "x"
arrSubKey(x) = "DisableCustomerImprovementProgram"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Customer Experience Improvement Pgm"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V3429-DTBI318"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings"
arrKeyPath2(x) = "x"
arrSubKey(x) = "Security_zones_map_edit"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "IE - Zones:Do Not Allow Users to Add/Delete Sites"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V3428-DTBI319"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings"
arrKeyPath2(x) = "x"
arrSubKey(x) = "Security_options_edit"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "IE - Zones: Do Not Allow Users to Change Policies"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V3428 - DTBI319"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings"
arrKeyPath2(x) = "x"
arrSubKey(x) = "Security_options_edit"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "IE - Zones: Do Not Allow Users to Change Policies"
arrCategory(x) = "II"

arrFinding(x) = "V3427-DTBI320"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings"
arrKeyPath2(x) = "x"
arrSubKey(x) = "Security_HKLM_only"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "IE - Zones: Use Only Machine Settings"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15494-DTBI325"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "DisableSecuritySettingsCheck"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Turn of Security Settings Check Feature"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15497-DTBI340"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_LOCALMACHINE_LOCKDOWN\Settings"
arrKeyPath2(x) = "x"
arrSubKey(x) = "LOCALMACHINE_CD_UNLOCK"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Allow Active Content From CDs to Run"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15499-DTBI350"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Download"
arrKeyPath2(x) = "x"
arrSubKey(x) = "RunInvalidSignatures"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Allow Software to Run or Install with Invalid Signature"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15500-DTBI355"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Main"
arrKeyPath2(x) = "x"
arrSubKey(x) = "Enable Browser Extensions"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "no"
arrExpectedValue2(x) = "x"
arrType(x) = "StringEx"
arrTitle(x) = "Allow Third-Party Browser Extensions"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15502-DTBI365"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings"
arrKeyPath2(x) = "x"
arrSubKey(x) = "CertificateRevocation"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Check for Certificate Revocation"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V3430-DTBI367"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings"
arrKeyPath2(x) = "x"
arrSubKey(x) = "ProxySettingsPerUser"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "IE - Make Proxy Settings Per Machine"
arrCategory(x) = "III"

x = x + 1
arrFinding(x) = "V15503-DTBI370"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Download"
arrKeyPath2(x) = "x"
arrSubKey(x) = "CheckExeSignatures"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "yes"
arrExpectedValue2(x) = "x"
arrType(x) = "StringEx"
arrTitle(x) = "Check for Signautres on Downloaded Programs"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15504-DTBI375"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap"
arrKeyPath2(x) = "x"
arrSubKey(x) = "UNCAsIntranet"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Intranet Sites: Include All Networks Path"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15507-DTBI385"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3"
arrKeyPath2(x) = "x"
arrSubKey(x) = "2102"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Allow Script Intiated Windows Without Size or Position Contraints - Internet Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15508-DTBI390"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3"
arrKeyPath2(x) = "x"
arrSubKey(x) = "2102"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Allow Script Intiated Windows Without Size or Position Contraints - Restricted Sites Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15509-DTBI395"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1209"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Allow Scriptlets are Not Disabled - Internet Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15513-DTBI415"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3"
arrKeyPath2(x) = "x"
arrSubKey(x) = "2200"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Automatic Prompting for File Downloads - Internet Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15515-DTBI425"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\0"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1C00"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Java Permissions Not Set - Local Machine Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15516-DTBI430"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Lockdown_Zones\0"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1C00"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Java Permissions Not Set - Locked Down Local Machine Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15517-DTBI435"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Lockdown_Zones\1"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1C00"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Java Permissions Not Set - Locked Down Intranet Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15518-DTBI440"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Lockdown_Zones\2"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1C00"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Java Permissions Not Set - Locked Down Trusted Site Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15519-DTBI445"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Lockdown_Zones\3"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1C00"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Java Permissions Not Set - Locked Down Internet Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15520-DTBI450"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Lockdown_Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1C00"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Java Permissions Not Set - Locked Down Restricted Sites Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15521-DTBI455"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3"
arrKeyPath2(x) = "x"
arrSubKey(x) = "2402"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Loose or Uncompiled XAML files - Internent Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15522-DTBI460"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "2402"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Loose or Uncompiled XAML files - Restricted Sites Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15523-DTBI465"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3"
arrKeyPath2(x) = "x"
arrSubKey(x) = "2100"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Open Files Based on Content, not File Extension - Internet Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15524-DTBI470"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "2100"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Open Files Based on Content, no File Extension - Restricted Files Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15525-DTBI475"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1208"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Turn Off First-Run Opt-In - Internent Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15526-DTBI480"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1208"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Turn Off First-Run Opt-In - Restricted Sites Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15527-DTBI485"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3"
arrKeyPath2(x) = "x"
arrSubKey(x) = "2500"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Turn on Protected Mode - Internet Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15528-DTBI490"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "2500"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Turn On Protected Mode - Restricted Sites Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15529-DTBI495"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1809"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Use Pop-up Blocker - Internet Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V1530-DTBI500"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1809"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Use Pop-up Blocker - Restricted Sites Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15533-DTBI515"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3"
arrKeyPath2(x) = "x"
arrSubKey(x) = "2101"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Web Sites in Less Privileged Zones Can Navigate Up - Internet Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15534-DTBI520"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "2101"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Web Sites in Less Privildged Zones Can Navigate Up - Restricted Sites Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15545-DTBI575"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "2000"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Allow Binary and Script Behaviors - Restricted Sites Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15546-DTBI580"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "2200"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Automatic Prompting For File Downloads - Restricted Sites Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15548-DTBI590"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_MIME_HANDLING"
arrKeyPath2(x) = "x"
arrSubKey(x) = "(Reserved)"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "StringEx"
arrTitle(x) = "Internet Explorer Processes For MIME Handling (Reserved)"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15565-DTBI592"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_MIME_HANDLING"
arrKeyPath2(x) = "x"
arrSubKey(x) = "explorer.exe"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "StringEx"
arrTitle(x) = "Internet Explorer Processes for MIME Handling (Explorer)"
arrCategorY(x) = "II"

x = x + 1
arrFinding(x) = "V15566-DTBI594"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_MIME_HANDLING"
arrKeyPath2(x) = "x"
arrSubKey(x) = "iexplore.exe"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "StringEx"
arrTitle(x) = "Internet Explorer PRocesses for MIME handling (IExplore)"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15549-DTBI595"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_MIME_SNIFFING"
arrKeyPath2(x) = "x"
arrSubKey(x) = "(Reserved)"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "StringEx"
arrTitle(x) = "Internet Explorer Processes for Mime Sniffing (Reserved)"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15603-DTBI596"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_MIME_SNIFFING"
arrKeyPath2(x) = "x"
arrSubKey(x) = "explorer.exe"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "StringEx"
arrTitle(x) = "IE Processes for MIME sniffing is not enabled (Explorer)"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15604-DTBI597"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_MIME_SNIFFING"
arrKeyPath2(x) = "x"
arrSubKey(x) = "iexplore.exe"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "StringEx"
arrTitle(x) = "Internet Explorer Processes for MIME Sniffing (IExplore)"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15568-DTBI599"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_DISABLE_MK_PROTOCOL"
arrKeyPath2(x) = "x"
arrSubKey(x) = "(Reserved)"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "StringEx"
arrTitle(x) = "Internet Explorer Processes for MK Protocol (Reserved)"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15550-DTBI600"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_DISABLE_MK_PROTOCOL"
arrKeyPath2(x) = "x"
arrSubKey(x) = "explorer.exe"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "StringEx"
arrTitle(x) = "Internet Explorer Processes for MK Protocol (Explorer)"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15551-DTBI605"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_DISABLE_MK_PROTOCOL"
arrKeyPath2(x) = "x"
arrSubKey(x) = "iexplore.exe"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "StringEx"
arrTitle(x) = "Internet Explorer Processes for MK Protocol (IExplore)"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15552-DTBI610"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_ZONE_ELEVATION"
arrKeyPath2(x) = "x"
arrSubKey(x) = "(Reserved)"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "StringEx"
arrTitle(x) = "Internet Explorer Process for Zone Elevation (Reserved)"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15569-DTBI612"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_ZONE_ELEVATION"
arrKeyPath2(x) = "x"
arrSubKey(x) = "explorer.exe"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "StringEx"
arrTitle(x) = "Internet Explorer Processes for Zone Elevation (Explorer)"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15570-DTBI614"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_ZONE_ELEVATION"
arrKeyPath2(x) = "x"
arrSubKey(x) = "iexplore.exe"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "StringEx"
arrTitle(x) = "Internet Explorer Processes for Zone Elevation (IExplore)"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15556-DTBI630"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_RESTRICT_FILEDOWNLOAD"
arrKeyPath2(x) = "x"
arrSubKey(x) = "(Reserved)"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "StringEx"
arrTitle(x) = "Internet Explorer Processes for Download (Reserved)"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15557-DTBI635"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_RESTRICT_FILEDOWNLOAD"
arrKeyPath2(x) = "x"
arrSubKey(x) = "explorer.exe"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "StringEx"
arrTitle(x) = "Internet Explorer Processes for Download (Explorer)"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15558-DTBI640"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_RESTRICT_FILEDOWNLOAD"
arrKeyPath2(x) = "x"
arrSubKey(x) = "iexplore.exe"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "StringEx"
arrTitle(x) = "Internet Explorer Processes for Download (IExplorer)"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15559-DTBI645"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_WINDOW_RESTRICTIONS"
arrKeyPath2(x) = "x"
arrSubKey(x) = "(Reserved)"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "StringEx"
arrTitle(x) = "Internet Explorer Processes for Restricting Popups (Reserved)"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15571-DTBI647"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_WINDOW_RESTRICTIONS"
arrKeyPath2(x) = "x"
arrSubKey(x) = "explorer.exe"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "StringEx"
arrTitle(x) = "Internet Explorer Processes for Restricting Popups (Explorer)"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15572-DTBI649"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_WINDOW_RESTRICTIONS"
arrKeyPath2(x) = "x"
arrSubKey(x) = "iexplore.exe"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "StringEx"
arrTitle(x) = "Internet Explorer Processes for Restricting Popups (IExplore)"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15560-DTBI650"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "2004"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Run .NET Framework-reliant Components not Signed with Authenticode - Restricted Sites Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15561-DTBI655"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "2001"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Run .NET Framework-reliant Components Signed with Authenticode - Restricted Sites Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15562-DTBI670"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1402"
arrHive(x) = HKLM
arrSubKey2(x) = "x"
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Scripting of Java Applets - Restricted Sites Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15563-DTBI675"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Main"
arrKeyPath2(x) = "x"
arrSubKey(x) = "Update_Check_Page"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = ""
arrExpectedValue2(x) = "x"
arrType(x) = "StringEx"
arrTitle(x) = "Turn Off Changing the URL to be Displayed"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15564-DTBI680"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Main"
arrKeyPath2(x) = "x"
arrSubKey(x) = "Update_Check_Interval"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "30"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Turn Off Configuring the Update Check Interval"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15574-DTBI690"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Control Panel"
arrKeyPath2(x) = "x"
arrSubKey(x) = "FormSuggest"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Disable AutoComplete for Forms is Not Enabled"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15575-DTBI695"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Restrictions"
arrKeyPath2(x) = "x"
arrSubKey(x) = "NoExternalBranding"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Disable External Branding of Internet Explorer"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V14245-DTBI697"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Restrictions"
arrKeyPath2(x) = "x"
arrSubKey(x) = "NoExtensionManagement"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueNotEx"
arrTitle(x) = "IE - Users Enable or Disable Add-ons" 
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15579-DTBI715"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Restrictions"
arrKeyPath2(x) = "x"
arrSubKey(x) = "NoCrashDetection"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Turn Off Crash Detection"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15580-DTBI720"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Main"
arrKeyPath2(x) = "x"
arrSubKey(x) = "Page_Transitions"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Turn Off Page Transitions"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V15581-DTBI725"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Main"
arrKeyPath2(x) = "Software\Policies\Microsoft\Internet Explorer\Control Panel"
arrSubKey(x) = "FormSuggest Passwords"
arrSubKey2(x) = "FormSuggest Password"
arrHive(x) = HKCU
arrExpectedValue(x) = "no"
arrExpectedValue2(x) = "1"
arrType(x) = "DualValueOr"
arrTitle(x) = "Turn On the Auto-Complete Feature is not Turned Off"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V22108-DTBI740"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\PhishingFilter"
arrKeyPath2(x) = "x"
arrSubKey(x) = "EnabledV8"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Turn Off Managing SmartScreen Filter"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V22147-DTBI750"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\BrowserEmulation"
arrKeyPath2(x) = "x"
arrSubKey(x) = "MSCompatibilityMode"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Include Updated Web Sites Lists from Microsoft"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V22148-DTBI760"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Privacy"
arrKeyPath2(x) = "x"
arrSubKey(x) = "ClearBrowsingHistoryOnExit"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Configure Delete Browsing History on Exit"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V30776-DTBI765"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Suggested Sites"
arrKeyPath2(x) = "x"
arrSubKey(x) = "Enabled"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Suggested Sites Functionality"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V22149-DTBI770"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Privacy"
arrKeyPath2(x) = "x"
arrSubKey(x) = "CleanHistory"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Prevent Deleting Web Sites that the User has Visisted"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V30777-DTBI775"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Main"
arrKeyPath2(x) = "x"
arrSubKey(x) = "NoUpdateCheck"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Internet Explorer Update Checking"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V22150-DTBI780"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Privacy"
arrKeyPath2(x) = "x"
arrSubKey(x) = "EnableInPrivateBrowsing"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Turn Off InPrivate Browsing"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V22152-DTBI800"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1206"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Allow Scripting of IE Web Browser"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V30778-DTBI805"
arrKeyPath(x) = "Software\Microsoft\Windows\CurrentVersion\Policies\Ext"
arrKeyPath2(x) = "x"
arrSubKey(x) = "NoFirsttimeprompt"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Opt-In Prompts for ActiveX"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V22153-DTBI810"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3"
arrKeyPath2(x) = "x"
arrSubKey(x) = "160A"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Include Local Directory Path when Uploaded"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V30779-DTBI815"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_SECURITYBAND"
arrKeyPath2(x) = "x"
arrSubKey(x) = "(Reserved)"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "StringEx"
arrTitle(x) = "Notification Bar Process - Reserved"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V22154-DTBI820"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1806"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Launching Programs and Unsafe Files Property"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V30780-DTBI825"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_SECURITYBAND"
arrKeyPath2(x) = "x"
arrSubKey(x) = "explorer.exe"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "StringEx"
arrTitle(x) = "Notification Bar Process - Explorer"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V22155-DTBI830"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3"
arrKeyPath2(x) = "x"
arrSubKey(x) = "120B"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Only Allow Approved Domains to use ActiveX Control Without Prompt"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V30781-DTBI835"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_SECURITYBAND"
arrKeyPath2(x) = "x"
arrSubKey(x) = "iexplore.exe"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "StringEx"
arrTitle(x) = "Notification Bar Process - IExplore"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V22156-DTBI840"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1409"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Turn on Cross-Site Scripting (XSS) Filter"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V22157-DTBI850"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1206"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Allow Scripting of IE Web Browser Control"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V22158-DTBI860"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "160A"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Include Local Directory Path When Uploaded"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V22159-DTBI870"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1806"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Lauchning Programs and Unsafe Files"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V22160-DTBI880"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "120B"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Only Allow Approved Domains to use ActiveX Controls Without Prompt"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V22161-DTBI890"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1409"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Turn on Cross-Site Scripting (XSS) Filter"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V22171-DTBI900"
arrKeyPath(x) = "Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_RESTRICT_ACTIVEXINSTALL"
arrKeyPath2(x) = "x"
arrSubKey(x) = "(Reserved)"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "StringEx"
arrTitle(x) = "IE Processes Restrict ActiveX Install (Reserved)"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V22634-DTBI910"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3"
arrKeyPath2(x) = "x"
arrSubKey(x) = "2103"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Allow Status Bar Updates Via Script - Internet Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V22635-DTBI920"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3"
arrKeyPath2(x) = "x"
arrSubKey(x) = "2004"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Run .NET Framework-reliant Components not Signed with Authenticode - Internet Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V22636-DTBI930"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3"
arrKeyPath2(x) = "x"
arrSubKey(x) = "2001"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Run .NET Framework-reliant Components Signed with Authenticode - Internet Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V22637-DTBI940"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "1209"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Allow Scripts Property - Restricted Sites Zone"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "V22638-DTBI950"
arrKeyPath(x) = "Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4"
arrKeyPath2(x) = "x"
arrSubKey(x) = "2103"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Allow Status Bar Updates via Script - Restricted Sites Zone"
arrCategory(x) = "II"

WScript.Echo "Performing " & x + 1 & " security checks on Internet Explorer 8" & vbCrLf

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

    ElseIf (StrComp(arrType(intCount), "BitMaskValueEx", vbTextCompare)) = 0 Then
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
      intCompare = arrValue(intCount) And arrExpectedValue(intCount)
      If boolResult = false Or intCompare Then
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)   
      ElseIf CDbl(intCompare) = CDbl(arrExpectedValue2(intCount)) Then
         WScript.Echo arrFinding(intCount) & " is Closed -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)  
      Else
         WScript.Echo arrFinding(intCount) & " is Open -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
      End If

   ElseIf(StrComp(arrType(intCount), "ManualReview", vbTextCompare)) = 0 Then
      WScript.Echo arrFinding(intCount) & " Manual Review -- CAT " & arrCategory(intCount) & " -- " & arrTitle(intCount)
   End If
Next
