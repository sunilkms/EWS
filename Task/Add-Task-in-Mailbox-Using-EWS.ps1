#Author: Sunil Chauhan
#this script Add task to using mailbox using app Impersonation Rights

param(
$userName=$cred.UserName,
$password=$cred.GetNetworkCredential().password,
$impdUser="sunil.chauhan@xyz.com",
$EWSServicePath = "C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll",
$EWSurl = "https://outlook.office365.com/EWS/Exchange.asmx",
$duedate=$(get-date).adddays(2)
)

#Importing WebService DLL
Import-Module $EWSServicePath

#Creating Service Object
$Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
$ExchVer = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1
$Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($exchver)
$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList $userName, $password
$Service.URL = $EWSurl
$ArgumentList = ([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SMTPAddress),$impdUser
$ImpUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId -ArgumentList $ArgumentList
$service.ImpersonatedUserId = $ImpUserId

# Setting Task EWS Class and Crating a Test Task
$task=New-Object Microsoft.Exchange.WebServices.Data.task -ArgumentList $service
$task.subject="This is test task"
$task.body="This Test task was created By EWS Service"
$duedate=$(get-date).adddays(2)
$task.duedate=$duedate
$task.save()
"Adding Task to Mailbox"
"Done!"
