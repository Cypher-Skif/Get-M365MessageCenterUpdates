[CmdletBinding()]
param ()

Function Get-ApiToken {

    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True)]
        [String]
        $AppId, $AppSecret, $TenantID
    )

    $AuthUrl = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"
    $Scope = "https://manage.office.com/.default"

    $Body = @{
        client_id = $AppId
        client_secret = $AppSecret
        scope = $Scope
        grant_type = 'client_credentials'
    }

    $PostSplat = @{
        ContentType = 'application/x-www-form-urlencoded'
        Method = 'POST'
        Body = $Body
        Uri = $AuthUrl
    }

    try {
        Invoke-RestMethod @PostSplat -ErrorAction Stop
    }
    catch {
        Write-Warning "$(Get-Date): Exception was caught: $($_.Exception.Message)" 
    }
}


Function Get-ApiRequestResult {

    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True)]
        [String]
        $Url, $Method, $Token
    )
 
    $Header = @{
        Authorization = "$($Token.token_type) $($Token.access_token)"
    }

    $PostSplat = @{
        ContentType = 'application/json'
        Method = $Method
        Header = $Header
        Uri = $Url
    }

    try {
        Invoke-RestMethod @PostSplat -ErrorAction Stop
    }
    catch {
        $Ex = $_.Exception
        $ErrorResponse = $ex.Response.GetResponseStream()
        $Reader = New-Object System.IO.StreamReader($errorResponse)
        $Reader.BaseStream.Position = 0
        $Reader.DiscardBufferedData()
        $ResponseBody = $Reader.ReadToEnd();
        Write-Output "$(Get-Date): Response content:`n$responseBody" -f Red
        throw Write-Error "$(Get-Date): Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
    }
}


Function Get-MCMessages {

    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True)]
        $APIUrl, $TenantId
    )

    $ApiVersion = "v1.0"
    $MS_resource = "ServiceComms/Messages?&`$filter=MessageType%20eq%20'MessageCenter'"
    $Uri = "$APIUrl/$ApiVersion/$($TenantId)/$MS_resource"
    
    $Method = "GET"

    try {
        Get-ApiRequestResult -Url $Uri -Token $Token -Method $Method -ErrorAction Stop
        Write-Output "$(Get-Date): New messages successfully collected"
    }
    catch {
        $Ex = $_.Exception
        $ErrorResponse = $ex.Response.GetResponseStream()
        $Reader = New-Object System.IO.StreamReader($errorResponse)
        $Reader.BaseStream.Position = 0
        $Reader.DiscardBufferedData()
        $ResponseBody = $Reader.ReadToEnd();
        Write-Output "$(Get-Date): Response content:`n$responseBody" -f Red
        throw Write-Error "$(Get-Date): Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
    }
}


Function Remove-HtmlTags {

    param (
        $Text
    )

    $SimpleTags = @(
        'p',
        'i',
        'span',
        'div',
        'ul',
        'ol',
        'h1',
        'h2',
        'h3',
        'div'
    )

    $TagsToRemove = (
        "\<\/?font[^>]*\>",
        '\<br\s?\/?\>',
        '\&rarr',
        ' style=""',
        ' target\=\"_blank\"'
    )

    $TagsToReplace = @(
        @('\[','<b>'),
        @('\]','</b>'),
        @('\<A','<a'),
        @('\<\/A\>','</a>'),
        @('\<img[^>]*\>','[There was an image]'),
        @('&nbsp;',' '),
        @('\<li\>',' -'),
        @('\<\/li\>',"`n")
    )

    foreach($Tag in $SimpleTags){
        $Pattern = "\<\/?$tag\>"
        $Text = $Text -replace $Pattern
    }

    foreach($Tag in $TagsToRemove){
        $Text = $Text -replace $Tag
    }

    foreach($Tag in $TagsToReplace){
        $Text = $Text -replace $Tag

    }
    
    foreach($Tag in $SimpleTags){
        $Pattern = "\<\/?$Tag\>"
        $Text = $Text -replace $Pattern
    }

    $Text
    
}


function Send-TelegramMessage {

    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [ValidateSet("html","markdown")]
        [String]
        $ParsingType,

        [Parameter(Mandatory=$true)]
        [String]
        $MessageText, $TokenTelegram, $ChatID
    )

    $URL_set = "https://api.telegram.org/bot$tokenTelegram/sendMessage"
    
    $Body = @{
        text = $MessageText
        parse_mode = $ParsingType
        chat_id = $ChatID
    }

    $MessageJson = $Body | ConvertTo-Json

    try {
        Invoke-RestMethod $URL_set -Method Post -ContentType 'application/json; charset=utf-8' -Body $messageJson -ErrorAction Stop
        Write-Output "$(Get-Date): Message has been sent"
    }
    catch {
        Write-Error "$(Get-Date): Can't sent message"
        Write-Error "$(Get-Date): $($_.Exception.Message)"
        Write-Output "$(Get-Date): StatusCode:" $_.Exception.Response.StatusCode.value__ 
        Write-Output "$(Get-Date): StatusDescription:" $_.Exception.Response.StatusDescription
        throw
    }
        
}

Write-Output "$(Get-Date): Getting Graph API credentials"
try {
    $GraphClientCreds = Get-AutomationPSCredential -Name 'APIMaster' -ErrorAction Stop
    Write-Output "$(Get-Date): Graph API credentials successfully collected"
}
catch {
    write-error "$(Get-Date): Can't get Graph API credentials"
    write-error "$(Get-Date): $($_.Exception.Message)"
    break
}

$ClientId = "ec286b42-4052-b74b-8dab-d5580212f4c5"
$TenantId = "58ca8e5f-fdbc-fb56-4ea5-70de6a380052"
$ClientSecret = $GraphClientCreds.GetNetworkCredential().Password
$APIUrl = "https://manage.office.com/api"

Write-Output "$(Get-Date): Getting Telegram credentials"
try{
    $TelegramClientCreds = Get-AutomationPSCredential -Name 'telegramToken' -ErrorAction Stop
    Write-Output "$(Get-Date): Telegram credentials successfully collected"
}
catch{
    write-error "$(Get-Date): Can't get telegram credentials"
    write-error "$(Get-Date): $($_.Exception.Message)"
    break
}


$TokenTelegram = $TelegramClientCreds.GetNetworkCredential().Password
[string]$ChatID = '-1000000001'
[string]$ErrorsHandlerChatId = '1'


Write-Output "$(Get-Date): Getting the M365 Graph Token"
try {
    $Token = Get-ApiToken -AppId $ClientId -AppSecret $ClientSecret -TenantID $TenantId -ErrorAction Stop
    Write-Output "$(Get-Date): Token successfully issued"
}
catch {
    Write-Error "$(Get-Date): Can't get the token!"
    write-error "$(Get-Date): $($_.Exception.Message)"
    break
}

Write-Output "$(Get-Date): Collecting Messages"
try {
    $Messages = Get-MCmessages -APIUrl $APIUrl -TenantId $TenantId -ErrorAction Stop
    Write-Output "$(Get-Date): MC messages successfully collected"
}
catch {
    Write-Error "$(Get-Date): Can't collect the messages"
    write-error "$(Get-Date): $($_.Exception.Message)"
    break
}

Write-Output "$(Get-Date): Checking current time"
$CurrentTime = Get-Date
Write-Output "$(Get-Date): Current time: $CurrentTime"
$СontrolTime = ($CurrentTime).AddMinutes(-60) 
Write-Output "$(Get-Date): Control time is: $СontrolTime"

write-Output "$(Get-Date): Checking new messages"
$NewMessages = $Messages.value | Where-Object {$(Get-date $_.LastUpdatedTime) -gt $(Get-date $СontrolTime)} 
$NewMessagesCount = $NewMessages.id.count

if ($NewMessagesCount -gt 0) {
    Write-Output "$(Get-Date): There are $NewMessagesCount new messages"
}else {
    Write-Output "$(Get-Date): There is no new messages"
    break
}


if ($NewMessagesCount -gt 0) {

    foreach ($NewMessage in $NewMessages){

        $MessagePreview = $NewMessage.Messages.MessageText
        $messageID = $NewMessage.Id
        $MessageTitle = $NewMessage.Title
        $MessageType = $NewMessage.Actiontype
        $PublishedTime = Get-date $($NewMessage.Messages.publishedTime)
        $UpdatedTime = Get-Date $($NewMessage.LastUpdatedTime)
        $MessageActionRequiredByDate = $(Get-date $($NewMessage.ActionRequiredByDate) -ErrorAction SilentlyContinue) 
        $MessageAdditionalInformation = $NewMessage.ExternalLink
        $MessageBlogLink = $NewMessage.BlogLink
        
        #adding emoji to Message Type
        if ($MessageType -eq 'Awareness') {
            $MessageType = '&#129000; ' + $MessageType
        }elseif ($MessageType -eq 'Action') {
            $MessageType = '&#128997; ' + $MessageType                
        }elseif ($MessageType -eq 'Opportunity') {
            $MessageType = '&#129001; ' + $MessageType            
        }
        ###

        $MessageTextWithHtmlString = $MessagePreview -split ('\<\/p\>') 
        
        $FormattedMesssageText = $($(Remove-HtmlTags $MessageTextWithHtmlString) -creplace '(?m)^\s*\r?\n','') -join "`n"
        
        $BoldMessageTitle = "<b>$MessageTitle</b>"
        $MessageDescription = "$MessageType Message $messageID"
        $PublishingInfo = "Published: $PublishedTime `nUpdated: $UpdatedTime"
        $TgmMessage = "$BoldMessageTitle `n$MessageDescription `n$PublishingInfo `n$FormattedMesssageText"

        if($MessageActionRequiredByDate){
            $TgmMessage += "`n<b>Action required by date: </b> $MessageActionRequiredByDate"
        }elseif ($MessageAdditionalInformation) {
            $TgmMessage += "`n<a href='$MessageAdditionalInformation'>Additional info</a>"
        }elseif ($MessageBlogLink) {
            $TgmMessage += "`n<a href='$MessageBlogLink'>Blog</a>"
        }

        try {
            Send-TelegramMessage -MessageText $TgmMessage -TokenTelegram $TokenTelegram -ChatID $chatID -ParsingType 'html' -ErrorAction Stop
            Send-TelegramTextMessage -BotToken $TokenTelegram -ChatID $ChatID -Message $TgmMessage
            Write-Output "$(Get-Date): Message $messageID has been successfully sent"
        }
        catch {
            Write-Error "$(Get-Date): There is issue with sending message: $MessageId `nPublished timestamp: $PublishedTime `nUpdated: $UpdatedTime"
            Write-Error "$(Get-Date): $($_.Exception.Message)"
            $ErrorMessage = "Message send error to 'M365 Message Center Updates': `nMessageID: $messageID `nTimeStamp: $TimeStamp `nMessage text:`n$TgmMessage `nErrorMessage: $($_.Exception.Message)"
            Send-TelegramMessage -MessageText $ErrorMessage -TokenTelegram $TokenTelegram -ChatID $ErrorsHandlerChatId -ParsingType 'markdown'
        }

        Write-Output "$(Get-Date): Has been sent message $messageID with published time: $PublishedTime"
    }
}




