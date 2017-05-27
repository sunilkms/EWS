############################################################
# Author : Sunil Chauhan
# Email : Sunilkms@gmail.com
# Blog: sunil-chauhan.blogspot.in
#
# This script can be used where you wants to restore items from user mailbox purges folder, exports the items from users Purges folder,
# Items Once Deleted from Recover Deleted items, they are out of user control and can't be restored, this script can be used to export the items.
# This script will export each items in .eml format and the item name would there subject.
# 
# Eml file cane be simply copy and pasted into user outlook.
#
# Usage Example: Get messages details in Purges & Deletions Folders.
# >.\Export-Email-Ews.ps1 -folder purges -itemsView 10 -mailbox test@user.com -userName admin@user.com -password Pass | ft
# >.\Export-Email-Ews.ps1 -folder deletions -itemsView 10 -mailbox test@user.com -userName admin@user.com -password Pass | ft

# 
# Usage Example 1: to export items from Purges folder run the script with below paramiters.
# >.\Export-Email-Ews.ps1 -export $true -folder purges -itemsView 10 -mailbox test@user.com -dir "D:\Purges\" -userName admin@user.com -password Pass -export $true

# Usage Example 2: to export items from Purges folder run the script with below paramiters.
# >.\Export-Email-Ews.ps1 -export $true -folder Deletions -itemsView 10 -mailbox test@user.com -dir "D:\Deletions\" -userName admin@user.com -password Pass -export $true
############################################################

param (

$report=$true,
$export=$false,
$folder="Deletions",
$mailbox="sunil.chauhan@xyz.com" ,
$itemsView=1000 ,                                       #No of Items to Export from the Folder
$dir="D:\Deletions\" ,                                  # Folder Path to save items in
$userName=$cred.UserName ,
$password=$cred.GetNetworkCredential().password

)

$uri=[system.URI] "https://outlook.office365.com/ews/exchange.asmx"
$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
Import-Module $dllpath

#Setup EWS Service Client
$ExchangeVersion=[Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2
$service=New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList $userName, $password
$service.url = $uri

#Entering into Mailbox
$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId `
([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SMTPAddress,$Mailbox);  

#Binding to Recoverable Items Root
$MailboxRootid= new-object Microsoft.Exchange.WebServices.Data.FolderId `
([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Recoverableitemsroot,$mailbox)
$MailboxRoot=[Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$MailboxRootid)

#Loading All Folders Under Recoverable Items Root
$FolderList = new-object Microsoft.Exchange.WebServices.Data.FolderView(10)
$FolderList.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
$findFolderResults = $MailboxRoot.FindFolders($FolderList)
$folderID = ($findFolderResults | ? {$_.DisplayName -eq $folder}).ID

#Setup property set for email
$psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet `
([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
$psPropset.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::HTML

#Loading Email Items
$view = New-Object Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList $ItemsView
$view.PropertySet = $propertyset
$items=$service.FindItems($folderid,$View)
$items.load($psPropset)
if ($report) {

$items.items | select Subject,LastModifiedTime,DateTimeReceived,Sender,{$_.ToRecipients}

}

if ($export) {

        foreach ($item in $items.items) 
        
        {
        Write-Host "Writing Email to File:" $item.Subject
        $fileName = $Dir + $($item.Subject).replace(":","-") + ".eml"
        $psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.ItemSchema]::MimeContent)
        $item.load($psPropset)
        $Email = new-object System.IO.FileStream($fileName, [System.IO.FileMode]::Create)
        $Email.Write($Item.MimeContent.Content, 0,$Item.MimeContent.Content.Length)
        $Email.Close()
        }

 }
