$scirptPath = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$configPath = Get-Content -Raw -Path "$scirptPath\config.json" | ConvertFrom-Json 


#tenant configuration variables
$clientId = $configPath.Configuration.tenantConfiguration.appID
$tenantId = $configPath.Configuration.tenantConfiguration.tenantId
$clientSecret = $configPath.Configuration.tenantConfiguration.clientSecret
$graphUrl = $configPath.Configuration.tenantConfiguration.graphUrl

#telegram bot configuration variables
$tokenTelegram = $configPath.Configuration.TelegramConfig.telegramToken
[string]$chatID = $configPath.Configuration.TelegramConfig.chatId

#functions block
Function Get-GraphResult ($Url, $Token, $Method) {

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
        $ex = $_.Exception
    
        $errorResponse = $ex.Response.GetResponseStream()
    
        $reader = New-Object System.IO.StreamReader($errorResponse)
    
        $reader.BaseStream.Position = 0
    
        $reader.DiscardBufferedData()
    
        $responseBody = $reader.ReadToEnd();
    
        Write-Error "Response content:`n$responseBody" -f Red
        
        Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
    
        break
    }
}

Function Get-GraphToken ($AppId, $AppSecret, $TenantID) {
        #defice resources' URLs
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
            Write-Warning "Exception was caught: $($_.Exception.Message)" 
        }
}

Function Get-MCmessages(){

    $graphApiVersion = "v1.0"
    $MC_resource = "ServiceComms/Messages?&`$filter=MessageType%20eq%20'MessageCenter'" 
    $uri = "$graphUrl/$graphApiVersion/$($tenantId)/$MC_resource"
    $Method = "GET"

    try {
            Get-GraphResult -Url $uri -Token $Token -Method $Method
            Write-Verbose "New messages successfully collected"
    }
    catch {
            $ex = $_.Exception
            $errorResponse = $ex.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($errorResponse)
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            $responseBody = $reader.ReadToEnd();

            Write-Host "Response content:`n$responseBody" -f Red
            Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
            Write-Verbose "Can't get new messages"

            break
             
    }
}

function Remove-HtmlTags ($Message) {

        $message = $message -replace '^\<p\>',""
        $message = $message -replace '\<p\>',""
        $message = $message -replace "\<p style\=[^>]*\>",''
        $message = $message -replace "\<p class\=[^>]*\>",''
        $message = $message -replace '\<\/p\>'," `n "
        $message = $message -replace "\<li style\=[^>]*\>"
        $message = $message -replace '\<font[^>]*\>'
        $message = $message -replace '\<\/font\>'
        $message = $message -replace ' style=""'
        $message = $message -replace '&nbsp;',''
        $message = $message -replace '\<span\>'
        $message = $message -replace '\<\/span\>'
        $message = $message -replace '\<div\>'
        $message = $message -replace '\<\/div\>'
        $message = $message -replace '\<br\><\<br\>','<br>'
        $message = $message -replace '\<br\>',""
        $message = $message -replace '\<ul\>'
        $message = $message -replace '\<\/ul\>'
        $message = $message -replace '\<ol\>'
        $message = $message -replace '\<\/ol\>'
        $message = $message -replace '\<li\>'," - "
        $message = $message -replace '\<\/li\>'
        $message = $message -replace ' target\=\"_blank\"',''
        $message = $message -replace '\[','<b>'
        $message = $message -replace '\]','</b>'
        $message = $message -replace '\<A','<a'
        $message = $message -replace '\<\/A\>','</a>'
        $message = $message -replace "\<img[^>]*\>",'[There was image]'
        $message = $message -replace '\<o\:p\>'
        $message = $message -replace '\<\/o\:p\>'
        $message = $message -replace '\<Boolean\>'
        $message = $message -replace "`n `n `n `n","`n"
        $message = $message -replace "`n `n `n","`n"
        $message = $message -replace "`n `n","`n"

        return $message
}

function Send-TelegramMessage {

        [CmdletBinding()]

        param (
            [Parameter(Mandatory=$true)]
            [string]$messageText,
            [Parameter(Mandatory=$true)]
            [string]$tokenTelegram,
            [Parameter(Mandatory=$true)]
            [string]$chatID
        )

        $URL_set = "https://api.telegram.org/bot$tokenTelegram/sendMessage"
          
        $body = @{
            text = $messageText
            parse_mode = "html"
            chat_id = $chatID
        }
    
        $messageJson = $body | ConvertTo-Json
    
        try {
            Invoke-RestMethod $URL_set -Method Post -ContentType 'application/json; charset=utf-8' -Body $messageJson
            Write-Verbose "Message has been sent"
        }
        catch {
            Write-Error "Can't sent message"
        }
        
}

function RunScript {

    [CmdletBinding()]
    param()
    Write-Verbose "Getting the M365 Graph Token"

    try {
        $Token = Get-GraphToken -AppId $clientId -AppSecret $clientSecret -TenantID $tenantId -ErrorAction Stop
        Write-Verbose "Token successfully issued"
    }
    catch {
        Write-Error "Can't get the token!"
        break
    }

    Write-Verbose "Collecting Messages"
    try {
        $messages = Get-MCmessages -ErrorAction Stop
        Write-Verbose "MC messages successfully collected"
    }
    catch {
        Write-Error "Can't collect the messages"
        break
    }

    Write-Verbose "Checking current time"
    try {
        $controlTime = (Get-date).AddMinutes(-61) 
        Write-Verbose "Control time is: $controlTime"
        $CurrentTime = Get-Date
        Write-Verbose "Current time: $CurrentTime"
    }
    catch {
        Write-Error "Can't check control time"
        break
    }

    write-host "Checking new messages"
    try {
        $CheckingTime = $messages.value.Messages | Where-Object {$(Get-date $($_.publishedTime)) -gt $(Get-date($controlTime))} | Select-Object publishedTime
        $NewMessagesCount = $CheckingTime.publishedTime.count
        
        if ($NewMessagesCount -gt 0) {
            Write-Verbose "There are $NewMessagesCount new messages"
        }else {
            Write-Verbose "There is no new messages"
        }

    }
    catch {
        Write-Error "Can't check new messages"
    }
    
    if ($NewMessagesCount -gt 0) {

        foreach ($TimeStamp in $($CheckingTime.publishedTime)){
            
            $MessagePreview = $null
            $MessagePreview = $messages.value.Messages | Where-Object {$_.publishedTime -eq $TimeStamp} | Select-Object MessageText
            $messageID = ($messages.value | Where-Object {$_.Messages.publishedTime -eq $TimeStamp}).id
            $MessageTitle = ($messages.value | Where-Object {$_.Messages.publishedTime -eq $TimeStamp}).Title
            $MessageType = ($messages.value | Where-Object {$_.Messages.publishedTime -eq $TimeStamp}).actiontype
            
            #adding emoji to Message Type
            if ($MessageType -eq 'Awareness') {
                $MessageType = '&#129001; ' + $MessageType
            }elseif ($MessageType -eq 'Action') {
                $MessageType = '&#128997; ' + $MessageType                
            }
            ###
            
            $PublishedTime = ($messages.value.messages | Where-Object {$_.publishedTime -eq $TimeStamp}).publishedTime

            $MessageText = $MessagePreview.MessageText
            $FormattedMesssageText = Remove-HtmlTags -Message $MessageText
            $BoldMessageTitle = "<b>$MessageTitle</b>"
            $MessageDescription = "$MessageType Message $messageID"
            $FinalMessage = "$BoldMessageTitle `n$MessageDescription `n$FormattedMesssageText"

            Send-TelegramMessage -messageText $FinalMessage -tokenTelegram $tokenTelegram -chatID $chatID

            Write-Verbose "Has been sent message $messageID with published time: $PublishedTime"
        }
    }
}

RunScript -Verbose
