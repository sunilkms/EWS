#---------------------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------------------
# This script will fetch the meeting deails from the room mailboxes
# Author: Sunil chauhan <sunilkms@gmail.com> 
# Requirement : following requirement must be met before running the script.
#     EWS Managed API Ver 2.2 (https://www.nuget.org/packages/Microsoft.Exchange.WebServices/)
#     install msal ps module (https://www.powershellgallery.com/packages/MSAL.PS/4.37.0.0)
#     account used in auth must have full access permissons to all the mailbox you wish to fetch the details.
#     Register and azure app and add a redirect uri as "https://localhost"
#     Updated the $$tenantId and $clientID (app id) in the script
#     change the EWS api parth as per your system environment
#     
#     Example Run: 
#          1) - Fetch-Meeting-Details-Using-Oauth.ps1 -multipleMailboxes (gc users.txt) #file should contain upn
#          2) - Fetch-Meeting-Details-Using-Oauth.ps1 -Mailbox mailbox@domain.com  
#----------------------------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------------------------
param (
$Mailbox,
$multipleMailboxes,
$StartDate=(Get-Date),
$EndDate=(get-date).adddays(30)
)

$tenantId="tenant_id"
$clientID="client_id"
$RUri="https://localhost"
$Scope="https://outlook.office365.com/EWS.AccessAsUser.All"
$EWSAPIPath="$($(Get-Location).Path)\Microsoft.Exchange.WebServices.dll"

$TokenResponse=Get-MsalToken -ClientId $clientID -TenantId $tenantID -RedirectUri $RUri -Scopes $scope #-ForceRefresh
# for debugging check details of the access token uncomment below.
#$TokenResponse.AccessToken | Get-JWTDetails

if ($TokenResponse) {
Import-Module $EWSAPIPath
$Service = [Microsoft.Exchange.WebServices.Data.ExchangeService]::new()
$Service.Credentials = [Microsoft.Exchange.WebServices.Data.OAuthCredentials]$($TokenResponse.AccessToken)
$Service.Url = "https://outlook.office365.com/EWS/Exchange.asmx"

if ($multipleMailboxes){
foreach ($Mailbox in $multipleMailboxes) {
Write-host "Checking for $Mailbox"
$Service = [Microsoft.Exchange.WebServices.Data.ExchangeService]::new()
$Service.Credentials = [Microsoft.Exchange.WebServices.Data.OAuthCredentials]$($TokenResponse.AccessToken)
$Service.Url = "https://outlook.office365.com/EWS/Exchange.asmx"
$Folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId `
([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$mailbox)
$Service.HttpHeaders.add("X-AnchorMailbox",$mailbox)
$cal=[Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$Folderid)

#Define the calendar view
$itemsView=1000
$CalendarView=New-Object Microsoft.Exchange.WebServices.Data.CalendarView($StartDate,$EndDate,$itemsView)    
$findCalItems=$service.FindAppointments($Cal.Id,$CalendarView)
#$report = $findCalItems | Select Start,End,Duration,AppointmentType,Subject,Location,
#Organizer,DisplayTo,DisplayCC,HasAttachments,IsReminderSet,ReminderDueBy
$report=$findCalItems.Items | select DateTimeCreated,Subject,@{N="Organizer";E={$_.Organizer.Name}},Location,
DisplayCc,Displayto,IsCancelled,start,End,
@{
  N="Total_Rcpt";E={
$a=([int]($_.displayto.split(";").trim()).count)
  if($_.displaycc) { $b=([int]($_.displaycc.split(";").trim()).count)}
  if($_.displaycc){$a+$b} else {[int]$a}
  }
 }
$report }
} else {
Write-host "Checking for $Mailbox"
$Folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId `
([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$mailbox)
$Service.HttpHeaders.add("X-AnchorMailbox",$mailbox)
$cal=[Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$Folderid)

#Define the calendar view
$itemsView=1000
$CalendarView=New-Object Microsoft.Exchange.WebServices.Data.CalendarView($StartDate,$EndDate,$itemsView)    
$findCalItems=$service.FindAppointments($Cal.Id,$CalendarView)
#$report = $findCalItems | Select Start,End,Duration,AppointmentType,Subject,Location,
#Organizer,DisplayTo,DisplayCC,HasAttachments,IsReminderSet,ReminderDueBy
$report=$findCalItems.Items | select DateTimeCreated,Subject,@{N="Organizer";E={$_.Organizer.Name}},Location,
DisplayCc,Displayto,IsCancelled,start,End,
@{
  N="Total_Rcpt";E={
  $a=([int]($_.displayto.split(";").trim()).count)
  if($_.displaycc) { $b=([int]($_.displaycc.split(";").trim()).count)}
  if($_.displaycc){$a+$b} else {[int]$a}
  }
 }
$report
}
}
else {write-host "Failed to fetch the AccessToken"}
