# Get-M365MessageCenterUpdates
M365 Message Center monitoring bot for telegram channels. Adopted for running from Azure Automation  
Just register your bot using @BotFather and register Graph API app in Azure AD.  
![botFather](https://github.com/Cypher-Skif/PublicRepoPictures/blob/master/Get-M365MessageCenterUpdates/Readme_picture_botFather.png)   
Please do not forget to add your Azure AD app and Telegram credentials.  
You can specify your credentials from Azure Automation Secure Assets using these variables:  
$GraphClientCreds = Get-AutomationPSCredential -Name 'APIMaster'  
$TelegramClientCreds = Get-AutomationPSCredential -Name 'telegramToken'     

Also you need to specify the telegram ChatID for your messages and additional chat ID for errors logs.  
[string]$ChatID = '-1000000001' #Production chat ID  
[string]$ErrorsHandlerChatId = '1' #Chat for errors log  
![Config_Screen](https://github.com/Cypher-Skif/PublicRepoPictures/blob/master/Get-M365MessageCenterUpdates/Readme_picture_config.png)   

Link to the channel with messages: https://t.me/M365MessageCenter  
![TelegramExample](https://github.com/Cypher-Skif/PublicRepoPictures/blob/master/Get-M365MessageCenterUpdates/Readme_picture_telegramExample.png)  
