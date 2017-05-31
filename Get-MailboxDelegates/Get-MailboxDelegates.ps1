#################################################################################
# Script : Get-MailboxDelegates
# Author : Sunil Chauhan
# Email help :sunilkms@hotmail.com 
# Details : this script will get you Delegates information.
# usage Example : 
# Get-mailboxdelegates -mailbox sunil@letsExchange.in -user userid -pass password
# ###############################################################################

function Get-MailboxDelegates {

                  [CmdletBinding(DefaultParameterSetName='mailbox', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$True,
                  HelpUri = 'http://www.LetsExchange.in/',
                  ConfirmImpact='High')]
                  [OutputType([String])]

param (
                   [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Mailbox')]
                   
                   [ValidateNotNull()]
                   [ValidateNotNullOrEmpty()]
                   [ValidateCount(0,100)]
                   [ValidatePattern("[@]*")]
                   $Mailbox,
                   $user= $cred.UserName,
                   $pass= $cred.GetNetworkCredential().Password                  
   ) 
Begin    
   {
$EWS = "C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll"
Import-Module $EWS

# Setting EWS Service Client
$Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService

# Setting up Credentials for service
$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList $user, $pass

# EWS url for the Service
$EWSurl = "https://outlook.office365.com/EWS/Exchange.asmx"
$Service.URL = $EWSurl
    }
    Process
    {

# Setting up Impersonated User account in Service.

$impdUser = $Mailbox
$ArgumentList = ([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SMTPAddress),$impdUser
$ImpUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId -ArgumentList $ArgumentList
$service.ImpersonatedUserId = $ImpUserId

$DelgDetails = @()

# Getting into Mailbox Delegates Object
$delegates = $service.getdelegates($Mailbox,$true)

foreach($Delegate in $delegates.DelegateUserResponses){
    $Obj = "" | select Mailbox,DelegateEmailAddress,Calendar,ReceiveMeetingCopies,Tasks,Inbox,Contacts,Notes,ViewPrivateItems,Journal  
    $Obj.Mailbox = $Mailbox
    $Obj.DelegateEmailAddress = $Delegate.DelegateUser.UserId.PrimarySmtpAddress
    $Obj.Calendar = $Delegate.DelegateUser.Permissions.CalendarFolderPermissionLevel
    $Obj.ReceiveMeetingCopies = $Delegate.DelegateUser.ReceiveCopiesOfMeetingMessages
    $Obj.Tasks = $Delegate.DelegateUser.Permissions.TasksFolderPermissionLevel  
    $Obj.Inbox = $Delegate.DelegateUser.Permissions.InboxFolderPermissionLevel
    $Obj.Contacts = $Delegate.DelegateUser.Permissions.ContactsFolderPermissionLevel
    $Obj.Notes = $Delegate.DelegateUser.Permissions.NotesFolderPermissionLevel  
    $Obj.ViewPrivateItems = $Delegate.DelegateUser.ViewPrivateItems
    $Obj.Journal = $Delegate.DelegateUser.Permissions.JournalFolderPermissionLevel
    $DelgDetails += $Obj
      }  
    }
    End
    {
      $DelgDetails
    }
}
