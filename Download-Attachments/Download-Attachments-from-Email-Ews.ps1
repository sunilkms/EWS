#Download Attachment From User Mailbox Programmatically
#This script can be used to download attachments Programmatically from a user Mailbox

# Usage Example: 
# .\Download-Attachments-from-Email-Ews.ps1 -mailbox "test@xyz.com" -downloadDirectory "\\Lab-01\Downloads" `
# -folderName "Reports" -AdminuserName admin@xyz.com -password Mypass

param (
       $mailbox="sunil.chauhan@xyz.com",
		   $downloadDirectory="\\Lab-host\Attachments",
		   $folderName="Server Storage Report",
		   $AdminuserName=$cred.UserName,
		   $password=$cred.GetNetworkCredential().password
      )
	  
$itemsView=10
$uri=[system.URI] "https://outlook.office365.com/ews/exchange.asmx"
$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
Import-Module $dllpath

## Set Exchange Version
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList $AdminuserName, $password
$service.url = $uri

$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId `
([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SMTPAddress,$mailbox);

$Folderid=new-object Microsoft.Exchange.WebServices.Data.FolderId `
([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$mailbox)
$MailboxRoot=[Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)

$FolderList = new-object Microsoft.Exchange.WebServices.Data.FolderView(100)
$FolderList.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
$findFolderResults = $MailboxRoot.FindFolders($FolderList)
$allFolders=$findFolderResults | ? {$_.FolderClass -eq "IPF.Note"}
$FolderToSearchForAttachment=$allFolders | ? {$_.DisplayName -eq $folderName} 

$ItemsWithAttachments=new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo `
([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::HasAttachments, $true)

$ItemView=New-Object Microsoft.Exchange.WebServices.Data.ItemView($itemsView)

$ItemsWithAttachments = $FolderToSearchForAttachment.FindItems($ItemsWithAttachments,$ItemView)

Write-host "Downloading..."

foreach($MailItems in $ItemsWithAttachments.Items){
		$MailItems.Load()
		foreach($attach in $MailItems.Attachments){
		    $Name=("AD-" +(($attach.lAstModifiedTime).ToShortDateString()) + "-" + $attach.Name.ToString())
			$Name=$Name.replace("/","-")
     		write-host "Attachment saved to Path:"(($downloadDirectory + "\" + $name))
			$attach.Load()
			$File = new-object System.IO.FileStream(($downloadDirectory + "\" + $Name), [System.IO.FileMode]::Create)
			$File.Write($attach.Content, 0, $attach.Content.Length)
			$File.Close()			
	}
}
