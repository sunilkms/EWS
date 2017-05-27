param(
$mailbox="testuser@xyz.com",
$itemsView=1000,
$userName,
$password,
$StartDate =(Get-Date),
$EndDate =(get-date).AddDays(7) # find meeting upto next 7 days.
)
#set EWS URL and Web Service DLL file location path below.
$uri=[system.URI] "https://outlook.office365.com/ews/exchange.asmx"
$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
Import-Module $dllpath

#Set Exchange Version
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList $userName, $password
$service.url = $uri

$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId `
([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SMTPAddress,$mailbox);

$Folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId `
([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$mailbox)
$cal=[Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$Folderid)

#Define the calendar view
$CalendarView = New-Object Microsoft.Exchange.WebServices.Data.CalendarView($StartDate,$EndDate,$itemsView)    
$findCalItems = $service.FindAppointments($Cal.Id,$CalendarView)

$report = $findCalItems | Select Start,End,Duration,AppointmentType,Subject,Location,
Organizer,DisplayTo,DisplayCC,HasAttachments,IsReminderSet,ReminderDueBy

$report | Export-Csv $($mailbox + "-Meetings.csv") -NoTypeInformation
