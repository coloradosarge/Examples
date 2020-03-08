# Setup constants necessary to make MC API calls.
$User = ""
$APIKey = ""
$MChost = ""
$creds = "$($User):$($APIKey)"
$encodeCreds = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($creds))
$authValue = "Basic $encodeCreds"
$Headers = @{ Authorization = $authValue }
$listID = ""
$users = @()
$global:progressPreference = 'silentlyContinue'
$ErrorActionPreference = 'Continue'

# Variables to capture the Campaign IDs
$NewAccountID = ""
$RecentVisitID = ""
$SixMonthVisitID = ""
$FunMoneyID = ""
$AnniversaryID = ""

# Variables to define subject names
$NewAccountSubject = "Welcome to the Fun Club!"
$RecentVisitSubject = "Thanks for stopping by!"
$SixMonthVisitSubject = "WE MISS YOU!!!"
$FunMoneySubject = "You've Got Money Waiting!"
$AnniversarySubject = "HAPPY ANNIVERSARY!"

# Variables that define Tag IDs
$NewAccountTagID = 22585
$RecentVisitTagID = 22589
$SixMonthVisitTagID = 22593
$FunMoneyTagID = 22601
$AnniversaryTagID = 22605

# Set the Template IDs for each campaign
$NewAccountTemplateID = 90137
$RecentVisitTemplateID = 101065
$SixMonthVisitTemplateID = 101061
$FunMoneyTemplateID = 101073
$AnniversaryTemplateID = 101069

# Get today's Date
$today = Get-Date -format "MM-dd-yyyy"

# Name the daily campaigns
$NewAccountCampaign = "New Accounts " + $today
$RecentVisitCampaign = "Recently Visted " + $today
$SixMonthVisitCampaign = "Six Month Visit " + $today
$FunMoneyCampaign = "You Have Fun Money " + $today
$AnniversaryCampaign = "Anniversary " + $today

# Tag Names
$NewAccountTag = "New Account"
$RecentVisitTag = "Recent Visit"
$SixMonthVisitTag = "Six Month Visit"
$FunMoneyTag = "Fun Money Balance"
$AnniversaryTag = "Anniversary"

# Setup queries needed to process fun club members
$dbUser = "mailchimp"
$dbPass = '$$MCchimp11!!'
$ServerInstance = "MANGELSENSDB\SQL2008R2"

#$queryTest = "SELECT [EMailAddress] FROM [rms_mangelsens].[dbo].[Customer] WHERE DATALENGTH(EmailAddress) > 3 AND EmailAddress
# LIKE '%@%'"

$queryUpdateFunMoney = "SELECT [EmailAddress], [CustomNumber1] As FunMoney, [FirstName], [LastName], [Address], [City], 
 [State], [Zip], [AccountOpened], [LastVisit], [AccountNumber], [PhoneNumber], [TotalVisits]
  FROM [rms_mangelsens].[dbo].[Customer]
  WHERE ((YEAR(LastVisit) = YEAR(GETDATE()) 
  AND MONTH(LastVisit) = MONTH(GETDATE()) 
  AND DAY(LastVisit) = DAY(GETDATE())) OR 
  (YEAR(LastUpdated) = YEAR(GETDATE()) 
  AND MONTH(LastUpdated) = MONTH(GETDATE()) 
  AND DAY(LastUpdated) = DAY(GETDATE())))
  AND (DATALENGTH(EmailAddress) > 0 AND EmailAddress LIKE '%@%'
  AND AccountTypeID = 2)"

$queryNewAccounts = "SELECT [EmailAddress], [AccountNumber], [CustomNumber1] AS FunMoney, [FirstName], [LastName], [Address], 
  [City], [State], [Zip], [AccountOpened], [LastVisit], [PhoneNumber], [TotalVisits]
  FROM [rms_mangelsens].[dbo].[Customer]
  WHERE YEAR(AccountOpened) = YEAR(GETDATE()) 
  AND MONTH(AccountOpened) = MONTH(GETDATE()) 
  AND DAY(AccountOpened) = DAY(GETDATE())
  AND DATALENGTH(EmailAddress) > 0 AND EmailAddress LIKE '%@%'
  AND AccountTypeID = 2"

  $queryNewAccountsLast = "SELECT [EmailAddress], [AccountNumber], [CustomNumber1] AS FunMoney, [FirstName], [LastName], [Address], 
  [City], [State], [Zip], [AccountOpened], [LastVisit], [PhoneNumber], [TotalVisits]
  FROM [rms_mangelsens].[dbo].[Customer]
  WHERE YEAR(AccountOpened) = YEAR(GETDATE()) 
  AND MONTH(AccountOpened) = MONTH(GETDATE()) 
  AND DAY(AccountOpened) = DAY(GETDATE())-1
  AND DATALENGTH(EmailAddress) > 0 AND EmailAddress LIKE '%@%'
  AND AccountTypeID = 2"

 #$queryNewAccounts = "SELECT [EmailAddress], [AccountNumber], [CustomNumber1] AS FunMoney, [FirstName], [LastName], [Address], 
 # [City], [State], [Zip], [AccountOpened], [LastVisit], [PhoneNumber], [TotalVisits]
 # FROM [rms_mangelsens].[dbo].[Customer]
 # WHERE DATALENGTH(EmailAddress) > 0 AND EmailAddress LIKE '%@%'
 # AND AccountTypeID = 2"

$queryRecentVisit = "SELECT [EmailAddress]
  FROM [rms_mangelsens].[dbo].[Customer]
  WHERE YEAR(LastVisit) = YEAR(GETDATE()) 
  AND MONTH(LastVisit) = MONTH(GETDATE()) 
  AND DAY(LastVisit) = DAY(GETDATE())-1 
  AND NOT AccountOpened = LastVisit
  AND DATALENGTH(EmailAddress) > 0 AND EmailAddress LIKE '%@%'
  AND AccountTypeID = 2"

 $queryRecentVisitLast = "SELECT [EmailAddress]
  FROM [rms_mangelsens].[dbo].[Customer]
  WHERE YEAR(LastVisit) = YEAR(GETDATE()) 
  AND MONTH(LastVisit) = MONTH(GETDATE()) 
  AND DAY(LastVisit) = DAY(GETDATE())-2 
  AND NOT AccountOpened = LastVisit
  AND DATALENGTH(EmailAddress) > 0 AND EmailAddress LIKE '%@%'
  AND AccountTypeID = 2"

$querySixMonthVisit = "SELECT [EmailAddress]
  FROM [rms_mangelsens].[dbo].[Customer]
  WHERE YEAR(LastVisit) = YEAR(DATEADD(MONTH, -6, GETDATE()))
  AND MONTH(LastVisit) = MONTH(DATEADD(MONTH,-6, GETDATE()))
  AND DAY(LastVisit) = DAY(DATEADD(MONTH, -6, GETDATE()))
  AND DATALENGTH(EmailAddress) > 0 AND EmailAddress LIKE '%@%'
  AND AccountTypeID = 2"

  $querySixMonthVisitLast = "SELECT [EmailAddress]
  FROM [rms_mangelsens].[dbo].[Customer]
  WHERE YEAR(LastVisit) = YEAR(DATEADD(MONTH, -6, GETDATE()))
  AND MONTH(LastVisit) = MONTH(DATEADD(MONTH,-6, GETDATE()))
  AND DAY(LastVisit) = DAY(DATEADD(MONTH, -6, GETDATE())) - 1
  AND DATALENGTH(EmailAddress) > 0 AND EmailAddress LIKE '%@%'
  AND AccountTypeID = 2"

 $queryFunMoney = "SELECT [EmailAddress]
  FROM [rms_mangelsens].[dbo].[Customer]
  WHERE CustomNumber1 >= 100
  AND DATEDIFF(DAY, LastVisit, GETDATE()) % 60 = 0
  AND DATEDIFF(DAY, LastVisit, GETDATE()) < 900
  AND DATEDIFF(DAY, LastVisit, GETDATE()) > 0
  AND DATALENGTH(EmailAddress) > 0 AND EmailAddress LIKE '%@%'
  AND AccountTypeID = 2"

  $queryFunMoneyLast = "SELECT [EmailAddress]
  FROM [rms_mangelsens].[dbo].[Customer]
  WHERE CustomNumber1 >= 100
  AND DATEDIFF(DAY, LastVisit, DATEADD(day, -1, cast(GETDATE() AS DATE))) % 60 = 0
  AND DATEDIFF(DAY, LastVisit, DATEADD(day, -1, cast(GETDATE() AS DATE))) < 900
  AND DATEDIFF(DAY, LastVisit, DATEADD(day, -1, cast(GETDATE() AS DATE))) > 0
  AND DATALENGTH(EmailAddress) > 0 AND EmailAddress LIKE '%@%'
  AND AccountTypeID = 2"

 $queryAnniversary = "SELECT [EmailAddress]
  FROM [rms_mangelsens].[dbo].[Customer]
  WHERE DAY(AccountOpened) = DAY(GETDATE())
  AND MONTH(AccountOpened) = MONTH(GETDATE())
  AND NOT AccountOpened = LastVisit
  AND DATALENGTH(EmailAddress) > 0 AND EmailAddress LIKE '%@%'
  AND AccountTypeID = 2"

  $queryAnniversaryLast = "SELECT [EmailAddress]
  FROM [rms_mangelsens].[dbo].[Customer]
  WHERE DAY(AccountOpened) = DAY(GETDATE())
  AND MONTH(AccountOpened) = MONTH(GETDATE())-1
  AND NOT AccountOpened = LastVisit
  AND DATALENGTH(EmailAddress) > 0 AND EmailAddress LIKE '%@%'
  AND AccountTypeID = 2"


# Run the queries against the RMS database
$UpdateFunMoney = Invoke-Sqlcmd -Query $queryUpdateFunMoney -ServerInstance $ServerInstance -Username $dbUser -Password $dbPass
$NewAccounts = Invoke-Sqlcmd -Query $queryNewAccounts -ServerInstance $ServerInstance -Username $dbUser -Password $dbPass
$RecentVisit = Invoke-Sqlcmd -Query $queryRecentVisit -ServerInstance $ServerInstance -Username $dbUser -Password $dbPass
$SixMonthVisit = Invoke-Sqlcmd -Query $querySixMonthVisit -ServerInstance $ServerInstance -Username $dbUser -Password $dbPass
$FunMoney = Invoke-Sqlcmd -Query $queryFunMoney -ServerInstance $ServerInstance -Username $dbUser -Password $dbPass
$Anniversary = Invoke-Sqlcmd -Query $queryAnniversary -ServerInstance $ServerInstance -Username $dbUser -Password $dbPass
$NewAccountsLast = Invoke-Sqlcmd -Query $queryNewAccountsLast -ServerInstance $ServerInstance -Username $dbUser -Password $dbPass
$RecentVisitLast = Invoke-Sqlcmd -Query $queryRecentVisitLast -ServerInstance $ServerInstance -Username $dbUser -Password $dbPass
$SixMonthVisitLast = Invoke-Sqlcmd -Query $querySixMonthVisitLast -ServerInstance $ServerInstance -Username $dbUser -Password $dbPass
$FunMoneyLast = Invoke-Sqlcmd -Query $queryFunMoneyLast -ServerInstance $ServerInstance -Username $dbUser -Password $dbPass
$AnniversaryLast = Invoke-Sqlcmd -Query $queryAnniversaryLast -ServerInstance $ServerInstance -Username $dbUser -Password $dbPass
#$Test = Invoke-Sqlcmd -Query $queryTest -ServerInstance $ServerInstance -Username $dbUser -Password $dbPass

# This function adds a campaign to mailchimp.  It takes the host, list id, name, subject, title, segment id, and template id
# Example invocation: Add-Campaign -request_host $MChost -subject "Test 44" -list_id "dd60692a57" -title "Test 44" -segment_id 74063 -template_id 60163
# Returns: [string] campaign_id
function Add-Campaign($request_host, $list_id, $subject, $title, $segment_id, $template_id) {

    # Build JSON request from PSObject
    $requestURL = "https://" + $request_host + "/3.0/campaigns"
    $requestCampaign = New-Object -TypeName PSObject
    $recipients = New-Object -TypeName PSObject
    $settings = New-Object -TypeName PSObject
    $segment_opts = New-Object -TypeName PSObject
    $cond_opts = @{"condition_type" = "StaticSegment";
      "field" = "static_segment";
      "op" = "static_is"
      "value" = $segment_id
    }
    $conditions = @($cond_opts)
    $variate_settings = New-Object -TypeName PSObject
    $tracking = New-Object -TypeName PSObject
    $salesforce = New-Object -TypeName PSObject
    $capsule = New-Object -TypeName PSObject
    $rss_opts = New-Object -TypeName PSObject
    $schedule = New-Object -TypeName PSObject
    $daily_send = New-Object -TypeName PSObject
    $ab_split_opts = New-Object -TypeName PSObject
    $social_card = New-Object -TypeName PSObject
    Add-Member -InputObject $requestCampaign -MemberType NoteProperty -Name type -Value "regular"
    Add-Member -InputObject $requestCampaign -MemberType NoteProperty -Name recipients -Value $recipients
    Add-Member -InputObject $recipients -MemberType NoteProperty -Name list_id -Value $list_id
    Add-Member -InputObject $recipients -MemberType NoteProperty -Name segment_opts -Value $segment_opts
    Add-Member -InputObject $segment_opts -MemberType NoteProperty -Name saved_segment_id -Value $segment_id
    #Add-Member -InputObject $segment_opts -MemberType NoteProperty -Name prebuilt_segment_id -Value ""
    Add-Member -InputObject $segment_opts -MemberType NoteProperty -Name match -Value "any"
    Add-Member -InputObject $segment_opts -MemberType NoteProperty -Name conditions -Value $conditions
    Add-Member -InputObject $requestCampaign -MemberType NoteProperty -Name settings -Value $settings
    Add-Member -InputObject $settings -MemberType NoteProperty -Name subject_line -Value $subject
    #Add-Member -InputObject $settings -MemberType NoteProperty -Name preview_text -Value ""
    Add-Member -InputObject $settings -MemberType NoteProperty -Name title -Value $title
    Add-Member -InputObject $settings -MemberType NoteProperty -Name from_name -Value "Mangelsen's"
    Add-Member -InputObject $settings -MemberType NoteProperty -Name reply_to -Value "contact@mangelsens.com"
    Add-Member -InputObject $settings -MemberType NoteProperty -Name use_conversation -Value $false
    Add-Member -InputObject $settings -MemberType NoteProperty -Name to_name -Value "*|FNAME|*"
    Add-Member -InputObject $settings -MemberType NoteProperty -Name folder_id -Value ""
    Add-Member -InputObject $settings -MemberType NoteProperty -Name authenticate -Value $true
    Add-Member -InputObject $settings -MemberType NoteProperty -Name auto_footer -Value $false
    Add-Member -InputObject $settings -MemberType NoteProperty -Name inline_css -Value $false
    Add-Member -InputObject $settings -MemberType NoteProperty -Name auto_tweet -Value $false
    #Add-Member -InputObject $settings -MemberType NoteProperty -Name auto_fb_post -Value ""
    Add-Member -InputObject $settings -MemberType NoteProperty -Name fb_comments -Value $true
    Add-Member -InputObject $settings -MemberType NoteProperty -Name template_id -Value $template_id
    #Add-Member -InputObject $request -MemberType NoteProperty -Name variate_settings -Value $variate_settings
    #Add-Member -InputObject $variate_settings -MemberType NoteProperty -Name winner_criteria -Value ""
    #Add-Member -InputObject $variate_settings -MemberType NoteProperty -Name wait_time -Value ""
    #Add-Member -InputObject $variate_settings -MemberType NoteProperty -Name test_size -Value ""
    #Add-Member -InputObject $variate_settings -MemberType NoteProperty -Name subject_lines -Value ""
    #Add-Member -InputObject $variate_settings -MemberType NoteProperty -Name send_times -Value ""
    #Add-Member -InputObject $variate_settings -MemberType NoteProperty -Name from_names -Value ""
    #Add-Member -InputObject $variate_settings -MemberType NoteProperty -Name reply_to_addresses -Value ""
    Add-Member -InputObject $requestCampaign -MemberType NoteProperty -Name tracking -Value $tracking
    Add-Member -InputObject $tracking -MemberType NoteProperty -Name opens -Value $true
    Add-Member -InputObject $tracking -MemberType NoteProperty -Name html_clicks -Value $true
    Add-Member -InputObject $tracking -MemberType NoteProperty -Name goal_tracking -Value $false
    Add-Member -InputObject $tracking -MemberType NoteProperty -Name ecomm360 -Value $false
    Add-Member -InputObject $tracking -MemberType NoteProperty -Name google_analytics -Value ""
    Add-Member -InputObject $tracking -MemberType NoteProperty -Name clicktale -Value ""
    Add-Member -InputObject $tracking -MemberType NoteProperty -Name salesforce -Value $salesforce
    #Add-Member -InputObject $salesforce -MemberType NoteProperty -Name campaign -Value ""
    #Add-Member -InputObject $salesforce -MemberType NoteProperty -Name notes -Value ""
    #Add-Member -InputObject $tracking -MemberType NoteProperty -Name capsule -Value $capsule
    #Add-Member -InputObject $capsule -MemberType NoteProperty -Name notes -Value ""
    #Add-Member -InputObject $request -MemberType NoteProperty -Name rss_opts -Value $rss_opts
    #Add-Member -InputObject $rss_opts -MemberType NoteProperty -Name feed_rul -Value ""
    #Add-Member -InputObject $rss_opts -MemberType NoteProperty -Name frequency -Value ""
    #Add-Member -InputObject $rss_opts -MemberType NoteProperty -Name schedule -Value $schedule
    #Add-Member -InputObject $schedule -MemberType NoteProperty -Name hour -Value ""
    #Add-Member -InputObject $schedule -MemberType NoteProperty -Name weekly_send_day -Value ""
    #Add-Member -InputObject $schedule -MemberType NoteProperty -Name monthy_send_date -Value ""
    #Add-Member -InputObject $schedule -MemberType NoteProperty -Name daily_send -Value $daily_send
    #Add-Member -InputObject $daily_send -MemberType NoteProperty -Name sunday -Value $false
    #Add-Member -InputObject $daily_send -MemberType NoteProperty -Name monday -Value $false
    #Add-Member -InputObject $daily_send -MemberType NoteProperty -Name tuesday -Value $false
    #Add-Member -InputObject $daily_send -MemberType NoteProperty -Name wednesday -Value $false
    #Add-Member -InputObject $daily_send -MemberType NoteProperty -Name thursday -Value $false
    #Add-Member -InputObject $daily_send -MemberType NoteProperty -Name friday -Value $false
    #Add-Member -InputObject $daily_send -MemberType NoteProperty -Name saturday -Value $false
    #Add-Member -InputObject $rss_opts -MemberType NoteProperty -Name contrain_rss_img -Value ""
    #Add-Member -InputObject $request -MemberType NoteProperty -Name ab_split_opts -Value $ab_split_opts
    #Add-Member -InputObject $ab_split_opts -MemberType NoteProperty -Name split_test -Value ""
    #Add-Member -InputObject $ab_split_opts -MemberType NoteProperty -Name pick_winner -Value ""
    #Add-Member -InputObject $ab_split_opts -MemberType NoteProperty -Name wait_units -Value ""
    #Add-Member -InputObject $ab_split_opts -MemberType NoteProperty -Name wait_time -Value ""
    #Add-Member -InputObject $ab_split_opts -MemberType NoteProperty -Name split_size -Value ""
    #Add-Member -InputObject $ab_split_opts -MemberType NoteProperty -Name from_name_a -Value ""
    #Add-Member -InputObject $ab_split_opts -MemberType NoteProperty -Name from_name_b -Value ""
    #Add-Member -InputObject $ab_split_opts -MemberType NoteProperty -Name reply_email_a -Value ""
    #Add-Member -InputObject $ab_split_opts -MemberType NoteProperty -Name reply_email_b -Value ""
    #Add-Member -InputObject $ab_split_opts -MemberType NoteProperty -Name subject_a -Value ""
    #Add-Member -InputObject $ab_split_opts -MemberType NoteProperty -Name subject_b -Value ""
    #Add-Member -InputObject $ab_split_opts -MemberType NoteProperty -Name send_time_a -Value ""
    #Add-Member -InputObject $ab_split_opts -MemberType NoteProperty -Name send_time_b -Value ""
    #Add-Member -InputObject $ab_split_opts -MemberType NoteProperty -Name send_time_winner -Value ""
    #Add-Member -InputObject $request -MemberType NoteProperty -Name social_card -Value $social_card
    #Add-Member -InputObject $social_card -MemberType NoteProperty -Name image_url -Value ""
    #Add-Member -InputObject $social_card -MemberType NoteProperty -Name description -Value ""
    #Add-Member -InputObject $social_card -MemberType NoteProperty -Name title -Value ""
    $requestJson = ConvertTo-Json -InputObject $requestCampaign -Depth 10
    $response = Invoke-WebRequest -URI $requestURL -Headers $Headers -Method POST -Body $requestJson | ConvertFrom-Json
    #$requestJson
    $response.id
}

# This function sends a campaign
# Example invocation: Delete-Tag -request_host "host" -list_id "list id" -tag_name "tag name"
# Returns: Void
function Campaign-send ($request_host, $campaign_id) {
    $requestURL = "https://" + $request_host + "/3.0/campaigns/" + $campaign_id + "/actions/send"
    Invoke-WebRequest -URI $requestURL -Headers $Headers -Method POST 
}


# This function is used to convert a string to an MD5 hash.
# Example invocation: MD5-Hash -plainString "stringToConvert"
function MD5-Hash ($plainString) {
     $md5 = New-Object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
     $utf8 = New-Object -TypeName System.Text.UTF8Encoding
     $string = $plainString.replace(' ','')
     $hash = [System.BitConverter]::ToString($md5.ComputeHash($utf8.GetBytes($string)))
     $hash.replace('-','').ToLower()
}

# This functions gets a segment id in a list for a specific tag name
# Example invocation: Match-Tags -request_hsot host -list_id "list id" -tag_name "Test"
# Returns: [int] tag_id
function Match-Tag ($request_host, $list_id, $tag_name) {
    $requestURL = "https://" + $request_host + "/3.0/lists/" + $list_id + "/segments"
    $response = Invoke-WebRequest -URI $requestURL -Headers $Headers -Method GET | ConvertFrom-Json
    Write-Output $response
    foreach($segment in $response.segments) {
        Write-Output $segment.name
        if($segment.name -eq $tag_name) {
            return $segment.id
        }
    }
    
}

# This functions adds a tag to a list member
# Example invocation: Match-Tags -request_host host -list_id "list id" -tag_name "Test" -user $user
# Returns: N/A
function AddUser-Tag($request_host, $list_id, $tag_name, $user) {
    $md5Email = MD5-Hash -plainString $user.EmailAddress.ToLower()
    $requestURL = "https://" + $request_host + "/3.0/lists/" + $list_id + "/members/$md5Email/tags"
    $requestTag = New-Object -TypeName PSObject
    $Tag = @{name=$tag_name;status="active"}
    Add-Member -InputObject $requestTag -MemberType NoteProperty -Name tags -Value @($Tag)
    $requestJson = ConvertTo-Json -InputObject $requestTag
    #Write-Output $requestJson
    #Write-Output $requestURL
    $response = Invoke-WebRequest -URI $requestURL -Headers $Headers -Method POST -Body $requestJson | ConvertFrom-Json
    return $response.id
  }

# This functions removes a tag from a list member
# Example invocation: Match-Tags -request_host host -list_id "list id" -tag_name "Test" -user $user
# Returns: N/A
function RemoveUser-Tag($request_host, $list_id, $tag_name, $user) {
    $md5Email = MD5-Hash -plainString $user.EmailAddress.ToLower()
    $requestURL = "https://" + $request_host + "/3.0/lists/" + $list_id + "/members/$md5Email/tags"
    $requestTag = New-Object -TypeName PSObject
    $Tag = @{name=$tag_name;status="inactive"}
    Add-Member -InputObject $requestTag -MemberType NoteProperty -Name tags -Value @($Tag)
    $requestJson = ConvertTo-Json -InputObject $requestTag
    #Write-Output $requestJson
    #Write-Output $requestURL
    $response = Invoke-WebRequest -URI $requestURL -Headers $Headers -Method POST -Body $requestJson | ConvertFrom-Json
    return $response.id
  }

# This function creates a new segment/tag
# Example invocation: Create-Tag -request_host "host" -list_id "list id" -tag_name "tag name"
# Returns: [int] tag_id
function Create-Tag ($request_host, $list_id, $tag_name, $users) {
    $requestURL = "https://" + $request_host + "/3.0/lists/" + $list_id + "/segments"
    $requestTag = New-Object -TypeName PSObject
    $options = New-Object -TypeName PSObject
    Add-Member -InputObject $requestTag -MemberType NoteProperty -Name name -Value $tag_name
    Add-Member -InputObject $requestTag -MemberType NoteProperty -Name static_segment -Value $users
    #Add-Member -InputObject $requestTag -MemberType NoteProperty -Name options -Value $options
    #Add-Member -InputObject $options -MemberType NoteProperty -Name match -Value "any"
    #Add-Member -InputObject $options -MemberType NoteProperty -Name conditions -Value @()
    $requestJson = ConvertTo-Json -InputObject $requestTag
    #Write-Output $requestJson
    $response = Invoke-WebRequest -URI $requestURL -Headers $Headers -Method POST -Body $requestJson | ConvertFrom-Json
    return $response.id
  }

# This function deletes a tag by name
# Example invocation: Delete-Tag -request_host "host" -list_id "list id" -tag_name "tag name"
# Returns: Void
function Delete-Tag ($request_host, $list_id, $tag_name) {
    $tag_id = Match-Tag -request_host $request_host -list_id  $list_id -tag_name  $tag_name
    $requestURL = "https://" + $request_host + "/3.0/lists/" + $list_id + "/segments/" + $tag_id
    #Invoke-WebRequest -URI $requestURL -Headers $Headers -Method DELETE 
}

# This function builds a member array
# Example invocation:
# Return: Member object
function Add-Customer($list_id, $request_host, $user) {
    # Merge Fields
    # Email Address = EMAIL or MERGE0 text
    # First Name = FNANE or MERGE1 text
    # Last Name = LNAME or MERGE2 text
    # Fun Club Number = MMERGE3 or MERGE3 text
    # Address = MMERGE4 or MERGE4 text
    # City = MMERGE5 or MERGE5 text
    # State/Province/Region = MMERGE6 or MERGE6 text
    # Postal/Zip Code = MMERGE7 or MERGE7 zip code
    # Opened = MMERGE8 or MERGE8 date
    # Last Visit = MMERGE9 or MERGE9 date
    # Old List = MMERGE10 or MERGE10 date
    # Total Visits = MMERGE11 or MERGE11 number
    # Reward Points = MMERGE12 or MERGE12 number
    # Fun Money = MMERGE13 or MERGE13 number
    # $ Amount Til Fun Money = MMERGE14 or MERGE14 number
    # Segment = MMERGE15 or MERGE15 text
    # Non Member = MMERGE16 or MERGE16 text
    # Account Number = MMERGE17 or MERGE17 text
    # Phone Number = MMERGE18 or MERGE18 text
    # phone = MMERGE19 or MERGE 19 text
    # Full Customer Name = MMERGE20 or MERGE20 text
    # Fun Money 12/22 = MMERGE31 or MERGE21 text
    # Customer.EmailAddress = MMERGE22 or MERGE22
    # Email = MMERGE23 or MERGE23

    # Setup variables for Request
    $requestURL = "https://" + $request_host + "/3.0/lists/" + $list_id + "/members"
    $userObject = New-Object -TypeName PSObject
    $addressObject = New-Object -TypeName PSObject
    $mergeFields = New-Object -TypeName PSObject
    Add-Member -InputObject $userObject -MemberType NoteProperty -Name email_address -Value $user.EmailAddress # Email Address
    Add-Member -InputObject $userObject -MemberType NoteProperty -Name email_type -Value "html" # Email Type
    Add-Member -InputObject $userObject -MemberType NoteProperty -Name status -Value "subscribed" # Subscription Status
    Add-Member -InputObject $userObject -MemberType NoteProperty -Name merge_fields -Value $mergeFields 

    # First Name
    if([string]::IsNullOrEmpty($user.FirstName)) {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name FNAME -Value ""
    } else {
        $fname = (Get-Culture).TextInfo.ToTitleCase($user.FirstName.ToLower())
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name FNAME -Value $fname
    }

    # Last Name
    if([string]::IsNullOrEmpty($user.LastName)) {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name LNAME -Value "" 
    } else {
        $lname = (Get-Culture).TextInfo.ToTitleCase($user.LastName.ToLower())
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name LNAME -Value $lname
    }

    # Account Number
    if([string]::IsNullOrEmpty($user.AccountNumber)) {
            
    } else {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE3 -Value $user.AccountNumber
    }

    # Address Object
    Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE4 -Value $addressObject
      
    # Address: Street Address
    if([string]::IsNullOrEmpty($user.Address)) {
        Add-Member -InputObject $addressObject -MemberType NoteProperty -Name addr1 -Value "No Address"
    } else {
        Add-Member -InputObject $addressObject -MemberType NoteProperty -Name addr1 -Value $user.Address
    }
    Add-Member -InputObject $addressObject -MemberType NoteProperty -Name addr2 -Value ""
      
    # Address: City
    if([string]::IsNullOrEmpty($user.City)) {
        Add-Member -InputObject $addressObject -MemberType NoteProperty -Name city -Value "No City"
    } else {
        $city = (Get-Culture).TextInfo.ToTitleCase($user.City.ToLower())
        Add-Member -InputObject $addressObject -MemberType NoteProperty -Name city -Value $city
    }

    # Address: State
    if([string]::IsNullOrEmpty($user.State)) {
        Add-Member -InputObject $addressObject -MemberType NoteProperty -Name state -Value "No State"
    } else {
        if($user.State.length -gt 2) {
            Write-Output "Long State"
            $state = (Get-Culture).TextInfo.ToTitleCase($user.State)
        } else {
            $state = $user.State.ToUpper()
        }
        Add-Member -InputObject $addressObject -MemberType NoteProperty -Name state -Value $state
    }

    # Address: Zip
    if([string]::IsNullOrEmpty($user.Zip) -or $user.Zip.length -gt 5 -or $user.Zip.length -lt 5) {
        Add-Member -InputObject $addressObject -MemberType NoteProperty -Name zip -Value "00000"
    } else {
        Add-Member -InputObject $addressObject -MemberType NoteProperty -Name zip -Value $user.Zip
    }

    # Address: Country
    Add-Member -InputObject $addressObject -MemberType NoteProperty -Name country -Value "US"
      
    # City
    if([string]::IsNullOrEmpty($user.City)) {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE5 -Value "No City"
    } else {
        $city = (Get-Culture).TextInfo.ToTitleCase($user.City.ToLower())
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE5 -Value $city
    }

    # State
    if([string]::IsNullOrEmpty($user.State)) {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE6 -Value "No State"
    } else {
        if($user.State.length -gt 2) {
            Write-Output "Long State"
            $state = (Get-Culture).TextInfo.ToTitleCase($user.State)
        } else {
            $state = $user.State.ToUpper()
        }
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE6 -Value $state
    }

    # Zip
    if([string]::IsNullOrEmpty($user.Zip) -or $user.zip.length -gt 5 -or $user.zip.length -lt 5) {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE7 -Value "00000"
    } else {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE7 -Value $user.Zip
    }

    # Account Opened Date
    if([string]::IsNullOrEmpty($user.AccountOpened)) {
            
    } else {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE8 -Value $user.AccountOpened.ToString("MM/dd/yyyy")
    }

    # Last Vist
    if([string]::IsNullOrEmpty($user.LastVisit)) {
            
    } else {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE9 -Value $user.LastVisit.ToString("MM/dd/yyyy")
    }

    # Last Visit - Old
    if([string]::IsNullOrEmpty($user.LastVisit)) {
            
    } else {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE10 -Value $user.LastVisit.ToString("MM/dd/yyyy")
    }
      
    # Fun Money
    if([string]::IsNullOrEmpty($user.TotalVisits)) {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE11 -Value 0
    } else {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE11 -Value $user.TotalVisits
    }

    Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE12 -Value 0 # Reward Points
      
    # Reward Points
    if([string]::IsNullOrEmpty($user.FunMoney)) {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE12 -Value 0
    } else {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE12 -Value $user.FunMoney
    }

    # Fun Money
    if([string]::IsNullOrEmpty($user.FunMoney)) {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE13 -Value 0
    } else {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE13 -Value ($user.FunMoney / 10)
    }

    # Amount Til Fun Money
    if([string]::IsNullOrEmpty($user.FunMoney)) {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE14 -Value 0
    } else {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE14 -Value (100 - $user.FunMoney)
    }


    Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE14 -Value 0 # Amount Til Fun Money
    Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE15 -Value "" # Segment
    Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE16 -Value "" # Non Member
      
    # Account Number - Fun Club Number
    if([string]::IsNullOrEmpty($user.AccountNumber)) {
            
    } else {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE17 -Value $user.AccountNumber
    }
      
    # Phone Number
    if([string]::IsNullOrEmpty($user.PhoneNumber)) {
            
    } else {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE18 -Value $user.PhoneNumber
    }

    # Phone
    if([string]::IsNullOrEmpty($user.PhoneNumber)) {
            
    } else {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE19 -Value $user.PhoneNumber
    }

    Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE20 -Value "" # Full Customer Name
    Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE21 -Value "" # Fun Money 12/22
    Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE22 -Value "" # Customer.EmailAddress
    Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE23 -Value "" # Email
    $requestJson = ConvertTo-Json -InputObject $userObject -Depth 10
    Write-Output $requestJson
    $response = Invoke-WebRequest -URI $requestURL -Headers $Headers -Method POST -Body $requestJson -UseBasicParsing | ConvertFrom-Json
    return $response.id
}

# This function updates customer attributes (Fun Money)
# Example invocation:
# Return: Customer ID
function FunMoney-Update($list_id, $request_host, $user) {

    # Setup variables for requests
    $md5Email = MD5-Hash -plainString $user.EmailAddress.ToLower()
    $requestURL = "https://" + $request_host + "/3.0/lists/" + $list_id + "/members/$md5Email"
      
    # Build Request
    $userObject = New-Object -TypeName PSObject
    $mergeFields = New-Object -TypeName PSObject
    $addressObject = New-Object -TypeName PSObject
    Add-Member -InputObject $userObject -MemberType NoteProperty -Name email_address -Value $user.EmailAddress
    Add-Member -InputObject $userObject -MemberType NoteProperty -Name merge_fields -Value $mergeFields
      
    # First Name
    if([string]::IsNullOrEmpty($user.FirstName)) {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name FNAME -Value ""
    } else {
        $fname = (Get-Culture).TextInfo.ToTitleCase($user.FirstName.ToLower())
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name FNAME -Value $fname
    }

    # Last Name
    if([string]::IsNullOrEmpty($user.LastName)) {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name LNAME -Value "" 
    } else {
        $lname = (Get-Culture).TextInfo.ToTitleCase($user.LastName.ToLower())
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name LNAME -Value $lname
    }

    # Account Number
    if([string]::IsNullOrEmpty($user.AccountNumber)) {
            
    } else {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE3 -Value $user.AccountNumber
    }

    # Address Object
    Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE4 -Value $addressObject
      
    # Address: Street Address
    if([string]::IsNullOrEmpty($user.Address)) {
        Add-Member -InputObject $addressObject -MemberType NoteProperty -Name addr1 -Value "No Address"
    } else {
        Add-Member -InputObject $addressObject -MemberType NoteProperty -Name addr1 -Value $user.Address
    }
    Add-Member -InputObject $addressObject -MemberType NoteProperty -Name addr2 -Value ""
      
    # Address: City
    if([string]::IsNullOrEmpty($user.City)) {
        Add-Member -InputObject $addressObject -MemberType NoteProperty -Name city -Value "No City"
    } else {
        $city = (Get-Culture).TextInfo.ToTitleCase($user.City.ToLower())
        Add-Member -InputObject $addressObject -MemberType NoteProperty -Name city -Value $city
    }

    # Address: State
    if([string]::IsNullOrEmpty($user.State)) {
        Add-Member -InputObject $addressObject -MemberType NoteProperty -Name state -Value "No State"
    } else {
        if($user.State.length -gt 2) {
            Write-Output "Long State"
            $state = (Get-Culture).TextInfo.ToTitleCase($user.State)
        } else {
            $state = $user.State.ToUpper()
        }
        Add-Member -InputObject $addressObject -MemberType NoteProperty -Name state -Value $state
    }

    # Address: Zip
    if([string]::IsNullOrEmpty($user.Zip) -or $user.Zip.length -gt 5 -or $user.Zip.length -lt 5) {
        Add-Member -InputObject $addressObject -MemberType NoteProperty -Name zip -Value "00000"
    } else {
        Add-Member -InputObject $addressObject -MemberType NoteProperty -Name zip -Value $user.Zip
    }

    # Address: Country
    Add-Member -InputObject $addressObject -MemberType NoteProperty -Name country -Value "US"
      
    # City
    if([string]::IsNullOrEmpty($user.City)) {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE5 -Value "No City"
    } else {
        $city = (Get-Culture).TextInfo.ToTitleCase($user.City.ToLower())
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE5 -Value $city
    }

    # State
    if([string]::IsNullOrEmpty($user.State)) {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE6 -Value "No State"
    } else {
        if($user.State.length -gt 2) {
            Write-Output "Long State"
            $state = (Get-Culture).TextInfo.ToTitleCase($user.State)
        } else {
            $state = $user.State.ToUpper()
        }
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE6 -Value $state
    }

    # Zip
    if([string]::IsNullOrEmpty($user.Zip) -or $user.Zip.length -gt 5 -or $user.Zip.length -lt 5 ) {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE7 -Value "68124"
    } else {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE7 -Value $user.Zip
    }

    # Account Opened Date
    if([string]::IsNullOrEmpty($user.AccountOpened)) {
            
    } else {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE8 -Value $user.AccountOpened.ToString("MM/dd/yyyy")
    }

    # Last Vist
    if([string]::IsNullOrEmpty($user.LastVisit)) {
            
    } else {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE9 -Value $user.LastVisit.ToString("MM/dd/yyyy")
    }

    # Last Visit - Old
    if([string]::IsNullOrEmpty($user.LastVisit)) {
            
    } else {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE10 -Value $user.LastVisit.ToString("MM/dd/yyyy")
    }

    # Reward Points
    if([string]::IsNullOrEmpty($user.FunMoney)) {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE12 -Value 0
    } else {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE12 -Value $user.FunMoney
    }

    # Fun Money
    if([string]::IsNullOrEmpty($user.FunMoney)) {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE13 -Value 0
    } else {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE13 -Value ($user.FunMoney / 10)
    }

    # Amount Til Fun Money
    if([string]::IsNullOrEmpty($user.FunMoney)) {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE14 -Value 0
    } else {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE14 -Value (100 - $user.FunMoney)
    }

    # Account Number - Fun Club Number
    if([string]::IsNullOrEmpty($user.AccountNumber)) {
            
    } else {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE17 -Value $user.AccountNumber
    }

    # Phone Number
    if([string]::IsNullOrEmpty($user.PhoneNumber)) {
            
    } else {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE18 -Value $user.PhoneNumber
    }

    # Phone
    if([string]::IsNullOrEmpty($user.PhoneNumber)) {
            
    } else {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE19 -Value $user.PhoneNumber
    }

    # Fun Money
    if([string]::IsNullOrEmpty($user.TotalVisits)) {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE11 -Value 0
    } else {
        Add-Member -InputObject $mergeFields -MemberType NoteProperty -Name MMERGE11 -Value $user.TotalVisits
    }
    $requestJson = ConvertTo-Json -InputObject $userObject -Depth 10
    $response = Invoke-WebRequest -URI $requestURL -Headers $Headers -Method PUT -Body $requestJson | ConvertFrom-Json
    return $response.id
}

#$count = 0
# Add all new customers for the last day
foreach ($row in $NewAccounts) {
     #$count += 1
     #Write-Output $count $row.EmailAddress
    Add-Customer -list_id $listID -request_host $MChost -user $row
}


#$count = 0
#Update Address, FunMoney and Dates for all customers
foreach ($row in $updateFunMoney) {
    #$count += 1
    #Write-Output $count $row.EmailAddress
    FunMoney-Update -list_id $listID -request_host $MChost -user $row
}

#foreach ($row in $NewAccounts) {
#    $users += $row.EmailAddress
#}
#Create-Tag -request_host $MChost -list_id $listID -tag_name $NewAccountTag -users $users

#$users = @()

#foreach ($row in $RecentVisit) {
#    $users += $row.EmailAddress
#}
#Create-Tag -request_host $MChost -list_id $listID -tag_name $RecentVisitTag -users $users

#$users = @()

#foreach ($row in $SixMonthVisit) {
#    $users += $row.EmailAddress
#}
#Create-Tag -request_host $MChost -list_id $listID -tag_name $SixMonthVisitTag -users $users

#$users = @()

#foreach ($row in $FunMoney) {
#    $users += $row.EmailAddress
#}
#Create-Tag -request_host $MChost -list_id $listID -tag_name $FunMoneyTag -users $users

#$users = @()

#foreach ($row in $Anniversary) {
#    $users += $row.EmailAddress
#}
#Create-Tag -request_host $MChost -list_id $listID -tag_name $AnniversaryTag -users $users

#foreach ($row in $Test) {
#    if($row.EmailAddress -match ".*@[\d\w\\.-]+[\d\w]") {
#        #Write-Ouput '-'
#    } else {
#     Write-Output $row.EmailAddress
#    }
#}

foreach ($row in $NewAccountsLast) {
    RemoveUser-Tag -request_host $MChost -list_id $listID -tag_name $NewAccountTag -user $row
}

foreach ($row in $NewAccounts) {
    AddUser-Tag -request_host $MChost -list_id $listID -tag_name $NewAccountTag -user $row
}

foreach ($row in $RecentVisitLast) {
    RemoveUser-Tag -request_host $MChost -list_id $listID -tag_name $RecentVisitTag -user $row
}

foreach ($row in $RecentVisit) {
    AddUser-Tag -request_host $MChost -list_id $listID -tag_name $RecentVisitTag -user $row
}

foreach ($row in $SixMonthVisitLast) {
   RemoveUser-Tag -request_host $MChost -list_id $listID -tag_name $SixMonthVisitTag -user $row
}

foreach ($row in $SixMonthVisit) {
    AddUser-Tag -request_host $MChost -list_id $listID -tag_name $SixMonthVisitTag -user $row
}

foreach ($row in $FunMoneyLast) {
    RemoveUser-Tag -request_host $MChost -list_id $listID -tag_name $FunMoneyTag -user $row
}

foreach ($row in $FunMoney) {
   AddUser-Tag -request_host $MChost -list_id $listID -tag_name $FunMoneyTag -user $row
}

foreach ($row in $AnniversaryLast) {
    RemoveUser-Tag -request_host $MChost -list_id $listID -tag_name $AnniversaryTag -user $row
}

foreach ($row in $Anniversary) {
    AddUser-Tag -request_host $MChost -list_id $listID -tag_name $AnniversaryTag -user $row
}

$NewAccoundID = Add-Campaign -request_host $MChost -subject $NewAccountSubject -list_id $listID -title $NewAccountCampaign -segment_id $NewAccountTagID -template_id $NewAccountTemplateID
$RecentVisitID = Add-Campaign -request_host $MChost -subject $RecentVisitSubject -list_id $listID -title $RecentVisitCampaign -segment_id $RecentVisitTagID -template_id $RecentVisitTemplateID
$SixMonthVisitID = Add-Campaign -request_host $MChost -subject $SixMonthVisitSubject -list_id $listID -title $SixMonthVisitCampaign -segment_id $SixMonthVisitTagID -template_id $SixMonthVisitTemplateID
$FunMoneyID = Add-Campaign -request_host $MChost -subject $FunMoneySubject -list_id $listID -title $FunMoneyCampaign -segment_id $FunMoneyTagID -template_id $FunMoneyTemplateID
$AnniversaryID = Add-Campaign -request_host $MChost -subject $AnniversarySubject -list_id $listID -title $AnniversaryCampaign -segment_id $AnniversaryTagID -template_id $AnniversaryTemplateID

#Campaign-send -request_host $MChost -campaign_id $NewAccountID
#Campaign-send -request_host $MChost -campaign_id $RecentVisitID
#Campaign-send -request_host $MChost -campaign_id $SixMonthVisitID
#Campaign-send -request_host $MChost -campaign_id $FunMoneyID
#Campaign-send -request_host $MChost -campaign_id $AnniversaryID