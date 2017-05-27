#Web Service Path
param(
$mailbox="sunil.chauhan@xyzdomain.com",
$userName="AdminUSER@xyzdomain.com",
$password="AdminPassword",
$EWSServicePath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
)

#Importing WebService DLL
Import-Module $EWSServicePath 

#Creating Service Object
$Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList $userName, $password
$EWSurl = "https://outlook.office365.com/EWS/Exchange.asmx"
$Service.URL = $EWSurl

#Setting up ImperSonated User
$ArgumentList = ([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SMTPAddress),$mailbox
$ImpUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId -ArgumentList $ArgumentList
$service.ImpersonatedUserId = $ImpUserId

#Setting up Email message Class
$message = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $service
#Creating and Saving Message in Draft Folder
$message.Subject = "This Message has been Created by EWS - Sunil Chauhan"
$message.From = "sunil.chauhan@xyzdomain.com"
$message.ToRecipients.Add("user2@xyzdomain.com")
$message.ToRecipients.Add("user1@xyzdomain.com")
$message.Body = "This is Test Message By Sunil Chauhan From EWS Client"
$message.SendAndSaveCopy()
