[CmdletBinding()]
param ()


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
            Write-Output "$(Get-Date): Response content:`n$responseBody" -f Red
            Write-Error "$(Get-Date): Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
            #break
        }
    }


    Function Get-GraphToken ($AppId, $AppSecret, $TenantID) {

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


    Function Get-MCmessages(){

        $graphApiVersion = "v1.0"
        $MC_resource = "ServiceComms/Messages?&`$filter=MessageType%20eq%20'MessageCenter'"
        $uri = "$global:graphUrl/$graphApiVersion/$($global:tenantId)/$MC_resource"
        
        $Method = "GET"

        try {
            Get-GraphResult -Url $uri -Token $Token -Method $Method
            Write-Output "$(Get-Date): New messages successfully collected"
        }
        catch {
            $ex = $_.Exception
            $errorResponse = $ex.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($errorResponse)
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            $responseBody = $reader.ReadToEnd();
            Write-Output "$(Get-Date): Response content:`n$responseBody" -f Red
            Write-Error "$(Get-Date): Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
            break
            Write-Output "$(Get-Date): Can't get new messaegs" 
        }
    }


    function Remove-HtmlTags {

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
            'li',
            'h1',
            'h2',
            'h3',
            'div'
        )

        foreach($tag in $SimpleTags){
            $Pattern = "\<\/?$tag\>"
            $Text = $Text -replace $Pattern
        }

        $Text = $Text -replace "\<\/?font[^>]*\>"
        $Text = $Text -replace '\<br\s?\/?\>'
        $Text = $Text -replace '\&rarr'
        $Text = $Text -replace ' style=""'
        $Text = $Text -replace '&nbsp;',''
        $Text = $Text -replace ' target\=\"_blank\"',''
        $Text = $Text -replace '\[','<b>'
        $Text = $Text -replace '\]','</b>'
        $Text = $Text -replace '\<A','<a'
        $Text = $Text -replace '\<\/A\>','</a>'
        $Text = $Text -replace "\<img[^>]*\>",'[There was an image]'
        $Text = $Text -replace ' - ',"`n- "

        $Text

    }


    function Send-TelegramMessage {

        [CmdletBinding()]
        param (
            [Parameter(Mandatory=$true)]
            [string]$MessageText,
            [Parameter(Mandatory=$true)]
            [string]$TokenTelegram,
            [Parameter(Mandatory=$true)]
            [string]$ChatID
        )

        $URL_set = "https://api.telegram.org/bot$tokenTelegram/sendMessage"
        
    
        $body = @{
            text = $MessageText
            parse_mode = "html"
            chat_id = $ChatID
        }
    

        $MessageJson = $body | ConvertTo-Json
    

        try {
            Invoke-RestMethod $URL_set -Method Post -ContentType 'application/json; charset=utf-8' -Body $MessageJson -ErrorAction Stop
            Write-Output "$(Get-Date): Message has been sent"
        }
        catch {
            Write-Error "$(Get-Date): Can't sent message"
            Write-Output "$(Get-Date): StatusCode:" $_.Exception.Response.StatusCode.value__ 
            Write-Output "$(Get-Date): StatusDescription:" $_.Exception.Response.StatusDescription
            throw
        }
            
    }

    $ScirptPath = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
    $ConfigPath = Get-Content -Raw -Path "$scirptPath\config.json" | ConvertFrom-Json 
    
    
    #tenant configuration variables
    $ClientId = $ConfigPath.Configuration.TenantConfiguration.AppID
    $TenantId = $ConfigPath.Configuration.TenantConfiguration.TenantId
    $ClientSecret = $ConfigPath.Configuration.TenantConfiguration.ClientSecret
    $GraphUrl = $ConfigPath.Configuration.TenantConfiguration.GraphUrl
    
    #telegram production configuration variables
    $TokenTelegram = $ConfigPath.Configuration.TelegramConfig.telegramToken
    [string]$ChatID = $ConfigPath.Configuration.TelegramConfig.chatId

    #telegram tests and logs configuration variables
    [string]$TestChatID = $ConfigPath.Configuration.TelegramTestConfig.ChatId


    Write-Output "$(Get-Date): Getting the M365 Graph Token"
    try {
        $Token = Get-GraphToken -AppId $ClientId -AppSecret $ClientSecret -TenantID $TenantId -ErrorAction Stop
        Write-Output "$(Get-Date): Graph API token successfully issued"
    }
    catch {
        Write-Error "$(Get-Date): Can't get the token!"
        break
    }

    Write-Output "$(Get-Date): Collecting Messages"
    try {
        $Messages = Get-MCmessages -ErrorAction Stop
        Write-Output "$(Get-Date): MC messages successfully collected"
    }
    catch {
        Write-Error "$(Get-Date): Can't collect the messages"
        break
    }

    Write-Output "$(Get-Date): Checking current time"
    $ControlTime = (Get-date).AddMinutes(-61) 
    Write-Output "$(Get-Date): Control time is: $ControlTime"
    $CurrentTime = Get-Date
    Write-Output "$(Get-Date): Current time: $CurrentTime"


    write-Output "$(Get-Date): Checking new messages"
    $CheckingTime = $Messages.value.Messages | Where-Object {$(Get-date $($_.publishedTime)) -gt $(Get-date($ControlTime))} | Select-Object publishedTime
    $NewMessagesCount = $CheckingTime.publishedTime.count

    if ($NewMessagesCount -gt 0) {
        Write-Output "$(Get-Date): There are $NewMessagesCount new messages"
    }else {
        Write-Output "$(Get-Date): There is no new messages"
        break
    }


    if ($NewMessagesCount -gt 0) {

        foreach ($TimeStamp in $($CheckingTime.publishedTime)){

            $MessagePreview = $null
            $MessagePreview = $Messages.value.Messages | Where-Object {$_.publishedTime -eq $TimeStamp} | Select-Object MessageText
            $MessageID = ($Messages.value | Where-Object {$_.Messages.publishedTime -eq $TimeStamp}).id
            $MessageTitle = ($Messages.value | Where-Object {$_.Messages.publishedTime -eq $TimeStamp}).Title
            $MessageType = ($Messages.value | Where-Object {$_.Messages.publishedTime -eq $TimeStamp}).actiontype
            
            #adding emoji to Message Type
            if ($MessageType -eq 'Awareness') {
                $MessageType = '&#129000; ' + $MessageType
            }elseif ($MessageType -eq 'Action') {
                $MessageType = '&#128997; ' + $MessageType                
            }elseif ($MessageType -eq 'Opportunity') {
                $MessageType = '&#129001; ' + $MessageType            
            }
            ###

            $PublishedTime = Get-date $($($Messages.value.messages | Where-Object {$_.publishedTime -eq $TimeStamp}).publishedTime)
            $MessageText = $MessagePreview.MessageText

            $FormattedMesssageText = $(Remove-HtmlTags $MessageText) -creplace '(?m)^\s*\r?\n',''
            $BoldMessageTitle = "<b>$MessageTitle</b>"
            $MessageDescription = "$MessageType Message $messageID"
            $FinalMessage = "$BoldMessageTitle `n$MessageDescription `n$FormattedMesssageText"
            
            try {
                Send-TelegramMessage -messageText $FinalMessage -tokenTelegram $tokenTelegram -chatID $chatID -ErrorAction Stop
                Write-Output "$(Get-Date): Message $MessageID has been successfully sent"
            }
            catch {
                Write-Error "$(Get-Date): There is issue with sending message: $MessageID"
                $ErrorMessage = "Message send error to 'M365 Message Center Updates': `nMessageID: `n$MessageID"
                Send-TelegramMessage -messageText $ErrorMessage -tokenTelegram $TokenTelegram -chatID $TestChatID
            }

            Write-Output "$(Get-Date): Has been sent message $MessageID with published time: $PublishedTime"
        }
    }




