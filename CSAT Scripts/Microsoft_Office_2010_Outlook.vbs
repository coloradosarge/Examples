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
arrFinding(x) = "DTOO104"
arrKeyPath(x) = "Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_HTTP_USERNAME_PASSWORD_DISABLE"
arrKeyPath2(x) = "x"
arrSubKey(x) = "outlook.exe"
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
arrSubKey(x) = "outlook.exe"
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
arrSubKey(x) = "outlook.exe"
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
arrSubKey(x) = "outlook.exe"
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
arrSubKey(x) = "outlook.exe"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Block Pop-Ups"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO272"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\options\mail"
arrKeyPath2(x) = "x"
arrSubKey(x) = "UnblockSafeZone"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Content download from safe zones"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO219"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\options\pubcal"
arrKeyPath2(x) = "x"
arrSubKey(x) = "RestrictedAccessOnly"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Access to Published Calendars"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO224"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\options\mail"
arrKeyPath2(x) = "x"
arrSubKey(x) = "JunkMailTrustOutgoingRecipients"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Email Recipient to Safe Sender List"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO234"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "AllowActiveXOneOffForms"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Active X One-Off Forms"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO246"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "EnableOneOffFormScripts"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Scripts in One-Off Forms"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO273"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\options\mail"
arrKeyPath2(x) = "x"
arrSubKey(x) = "TrustedZone"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Block Trusted Zones"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO236"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "AddinTrust"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Add-In Trust Level"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO250"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "PromptOOMAddressBookAccess"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Object Model Prompt for Address Book"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO241"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "AllowUsersToLowerAttachments"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Demote Attachments to Level 2"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO254"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "PromptOOMFormulaAccess"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Object Model Prompt for Formula Property"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO253"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "PromptOOMSaveAs"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Object Model Prompt for Save As"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO251"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "PromptOOMAddressInformationAccess"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Object Model Prompt for Reading Address"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO252"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "PromptOOMMeetingTaskRequestResponse"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Object Model Prompt for Meeting Response"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO249"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "PromptOOMSend"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Object Model Prmpt for auto email send"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO256"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\Outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "trustedaddins"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "x"
arrExpectedValue2(x) = "x"
arrType(x) = "KeyExists"
arrTitle(x) = "Trusted Add-Ins Security"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO226"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\options\mail"
arrKeyPath2(x) = "x"
arrSubKey(x) = "Hangup after Spool"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Dial-up Options "
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO225"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\options\mail"
arrKeyPath2(x) = "x"
arrSubKey(x) = "warn on dialup"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Warn before Switching Dial-up"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO237"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "EnableRememberPwd"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Disable remember password on eMail Accts"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO243"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "DontPromptLevel1AttachClose"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Level 1 Attachment prompt"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO242"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "DontPromptLevel1AttachSend"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Level 1 Attachment Prompt on sending"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO283"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\options\rss"
arrKeyPath2(x) = "x"
arrSubKey(x) = "EnableFullTextHTML"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Download articles as HTML Att. - Outloook"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO277"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\options\mail"
arrKeyPath2(x) = "x"
arrSubKey(x) = "JunkMailEnableLinks"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Enable links in Email Messages - Outlook"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO279"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\rpc"
arrKeyPath2(x) = "x"
arrSubKey(x) = "EnableRPCEncryption"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Enable RPC Encryption"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO221"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook"
arrKeyPath2(x) = "x"
arrSubKey(x) = "DisableAntiSpam"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Hide Junk Mail UI - Outlook"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO274"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\options\mail"
arrKeyPath2(x) = "x"
arrSubKey(x) = "Internet"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Incl. Internet w/Safe Zones - Outlook"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO275"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\options\mail"
arrKeyPath2(x) = "x"
arrSubKey(x) = "Intranet"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Incl. Intranet w/ Safe Zone - Outlook"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO240"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "ShowLevel1Attach"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Display Level 1 Att. - Outlook"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO270"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\options\mail"
arrKeyPath2(x) = "x"
arrSubKey(x) = "BlockExtContent"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "External Pictures & content "
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO227"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\mailsettings"
arrKeyPath2(x) = "x"
arrSubKey(x) = "DisableSignatures"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Digital Signature handling"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO230"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "NonDefaultStoreScript"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "No fldr home pages / non-default stores "
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO233"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "PublicFolderScript"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "No OOM scripts for pub fldrs - Outlook"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO232"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "SharedFolderScript"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "OOM scripts for Shared Folders"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO285"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\options\webcal"
arrKeyPath2(x) = "x"
arrSubKey(x) = "Disable"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Internet Calendar Integration"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO269"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "OutlookSecureTempFolder"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "65536"
arrExpectedValue2(x) = "x"
arrType(x) = "Exists"
arrTitle(x) = "Att. Secure Temporary Folder - Outlook"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO280"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "AuthenticationService"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "9"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Authentication w/Exchange Svr"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO278"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\autodiscover"
arrKeyPath2(x) = "x"
arrSubKey(x) = "ZeroConfigExchange"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Auto Cfg. profile based on AD - Outlook"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO284"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\options\webcal"
arrKeyPath2(x) = "x"
arrSubKey(x) = "EnableAttachments"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Auto download attachments Internet Cal"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO271"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\options\mail"
arrKeyPath2(x) = "x"
arrSubKey(x) = "UnblockSpecificSenders"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Auto Download from Safe lists "
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO229"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\options\general"
arrKeyPath2(x) = "x"
arrSubKey(x) = "Check Default Client"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Make Outlook the default - Outlook"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO260"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "MsgFormats"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Message formats / SMime - Outlook" 
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO268"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "SigStatusNoTrustDecision"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "2"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Missing Root Certificates "
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO239"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "AdminSecurityMode"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "3"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Outlook Security Mode "
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO228"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\options\mail"
arrKeyPath2(x) = "x"
arrSubKey(x) = "Message Plain Format Mime"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Plain Text Options"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO217"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\options\pubcal"
arrKeyPath2(x) = "x"
arrSubKey(x) = "DisableDav"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Prevent publishing to DAV Servers"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO216"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\options\pubcal"
arrKeyPath2(x) = "x"
arrSubKey(x) = "DisableOfficeOnline"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Prevent Publshg to Ofc.Online - Outlook"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO238"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook"
arrKeyPath2(x) = "x"
arrSubKey(x) = "DisallowAttachmentCustomization"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Prev't users customizing security set"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO214"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\options\mail"
arrKeyPath2(x) = "x"
arrSubKey(x) = "ReadAsPlain"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Read EMail as plain text "
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO215"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\options\mail"
arrKeyPath2(x) = "x"
arrSubKey(x) = "ReadSignedAsPlain"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Read signed EMail as plain text"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO244"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "FileExtensionsRemoveLevel1"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "x"
arrExpectedValue2(x) = "x"
arrType(x) = "Exists"
arrTitle(x) = "Lvl 1 File extensions"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO245"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security\"
arrKeyPath2(x) = "x"
arrSubKey(x) = "FileExtensionsRemoveLevel2"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "x"
arrExpectedValue2(x) = "x"
arrType(x) = "Exists"
arrTitle(x) = "Lvl 2 File Extensions"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO218"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\options\pubcal"
arrKeyPath2(x) = "x"
arrSubKey(x) = "PublishCalendarDetailsPolicy"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "16384"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Calendar details published by users"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO220"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\options\pubcal"
arrKeyPath2(x) = "x"
arrSubKey(x) = "SingleUploadOnly"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Restrict Upload Method - Outlook"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO267"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "UseCRLChasing"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Retrieving CRL Data"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO262"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "FIPSMode"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "FIPS compliant mode"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO257"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "ExternalSMime"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "No S/Mime interop w/ external clients"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO266"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "RespondToReceiptRequests"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "2"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "S/Mime receipt requests "
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO276"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "Level"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "2"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Security settings for macros"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO264"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "ClearSign"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Clear signed messages"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO247"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "PromptOOMCustomAction"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Custom OOM Action Exe. Prompt "
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO265"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "WarnAboutInvalid"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Signature Warning - Outlook"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO281"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\options\rss"
arrKeyPath2(x) = "x"
arrSubKey(x) = "SyncToSysCFL"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Sync RSS Feeds w/Common List"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO223"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\options\mail"
arrKeyPath2(x) = "x"
arrSubKey(x) = "JunkMailTrustContacts"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Trust EMail from Contacts"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO282"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\options\rss"
arrKeyPath2(x) = "x"
arrSubKey(x) = "Disable"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "RSS Feeds"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO286"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\meetings\profile"
arrKeyPath2(x) = "x"
arrSubKey(x) = "ServerUI"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Disable User Entries to Server list"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO126"
arrKeyPath(x) = "Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_ADDON_MANAGEMENT"
arrKeyPath2(x) = "x"
arrSubKey(x) = "outlook.exe"
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
arrSubKey(x) = "outlook.exe"
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
arrSubKey(x) = "outlook.exe"
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
arrSubKey(x) = "outlook.exe"
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
arrSubKey(x) = "outlook.exe"
arrSubKey2(x) = "x"
arrHive(x) = HKLM
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Scripted Window Security"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO128"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
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
arrFinding(x) = "DTOO305"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\common\toolbars\outlook"
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
arrFinding(x) = "DTOO313"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\options\rss"
arrKeyPath2(x) = "x"
arrSubKey(x) = "EnableAttachments"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Automatically download enclosures"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO344"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\options\mail"
arrKeyPath2(x) = "x"
arrSubKey(x) = "Message RTF Format"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Outlook Rich Text options"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO314"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\options\mail"
arrKeyPath2(x) = "x"
arrSubKey(x) = "EditorPreference"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "65536"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Set message format"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO315"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "ForceDefaultProfile"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "0"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Outlook Security settings"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO316"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "MinEncKey"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "168"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Minimum encryption settings"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO317"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "NoCheckOnSessionSecurity"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Signed/encrypted messages"
arrCategory(x) = "II"

x = x + 1
arrFinding(x) = "DTOO320"
arrKeyPath(x) = "Software\Policies\Microsoft\Office\14.0\outlook\security"
arrKeyPath2(x) = "x"
arrSubKey(x) = "SupressNameChecks"
arrSubKey2(x) = "x"
arrHive(x) = HKCU
arrExpectedValue(x) = "1"
arrExpectedValue2(x) = "x"
arrType(x) = "ValueEx"
arrTitle(x) = "Check e-mail address against certificate"
arrCategory(x) = "II"

WScript.Echo "Performing " & x & " security checks on Microsoft Office 2010 Outlook" & vbCrLf

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

   Elseif (StrComp(arrType(intCount), "KeyExists", vbTextCompare)) = 0 Then
      objReg.EnumKey arrHive(intCount), arrKeyPath(intCount), arrKeyList
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
