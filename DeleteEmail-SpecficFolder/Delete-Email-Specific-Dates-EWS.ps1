
#######################################################################################
#Author = Sunil Chauhan
#Email= Sunilkms@gmail.com
#Ver =https://sunil-chauhan.Blogspot.com
#Deleting Emails From a Specific Folder between specfic dates.#
########################################################################################

param (
$Adminuser=$cred.UserName,
$Pass=$cred.GetNetworkCredential().password,
$Folder="Deleted Items",
$ToDate="03/10/2017",              #DATE Format MM/DD/YYYY
$FromDate="03/09/2017",             #DATE Format MM/DD/YYYY
$Items=1000 ,
$MailboxToImpersonate="sunil.chauhan@xyz.com.com",
$Report
)

#Web Service Path
$EWSServicePath = "C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll"
Import-Module $EWSServicePath
$Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
$ExchVer = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1
$Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($exchver) 
$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList $Adminuser, $pass
 
#Setting up EWS URL
$EWSurl = "https://outlook.office365.com/EWS/Exchange.asmx"
$Service.URL = $EWSurl

$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId `
([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SMTPAddress,$MailboxToImpersonate);
 
# Defining Itemview depth
$ItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView($items) 
$MailboxRootid = new-object  Microsoft.Exchange.WebServices.Data.FolderId `
([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$MailboxToImpersonate)
$MailboxRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$MailboxRootid)
 
# Get Folder ID from Path

$View=New-Object  Microsoft.Exchange.WebServices.Data.FolderView(2,0);
$View.Traversal=[Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep;
$View.PropertySet=[Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly;
$SearchFilter=New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo `
([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$Folder);
$FolderResults=$MailboxRoot.FindFolders($SearchFilter, $View);                                                                                             
$findItemResults = $FolderResults.FindItems("System.Message.DateReceived:$fromDate..$todate",$ItemView)

$Deleted=0
if ($report) 
          {        
           if ($findItemResults) 
                 {
                  Write-Host "Folder:"$SearchFilter.value                 
                  Write-Host "Total Item Found:" $findItemResults.count -NoNewline
                  Write-Host " Between" $fromDate "and" $toDate
                  $findItemResults | Select Subject,DateTimeReceived,Sender          
                 }
          }
else {        
        Write-Host "Folder:"$SearchFilter.value                 
        Write-Host "Total Item Found:" $findItemResults.count -NoNewline
        Write-Host " Between" $fromDate "and" $toDate
        
        foreach ($item in $findItemResults) 
        {        
        try {
             $Deleted++
             cls
             ""
             Write-Host "Folder:"$SearchFilter.value                 
             Write-Host "Total Item Found:" $findItemResults.count -NoNewline
             Write-Host " Between" $fromDate "and" $toDate
             ""        
             Write-host "Deleting:"$Deleted
             [void]$item.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete)                       
            } 
        catch 
            {
             Write-host "Unable to delete item:$($item.subject)"
             Write-host "Error:$($Error[0].Exception.Message)"
            }
       }
       if ($Deleted -gt 0) { Write-host "$Deleted email items has been deleted from the Mailbox." }    
    }                              
