#===================================================================================================
#Author = Sunil Chauhan
#Email= Sunilkms@gmail.com
#Blogs=Sunil.chauhan.blogspot.com
#=====================================================================================================
# this Script can be useful to Quickly Empty a specific Folder in user mailbox
# can also delete subfolder under a specific folder.
#Usage Example:
#Empty Recover Deleted Items Folder From user Mailbox "Testuser@xzy.com"
#>.\CleanUp-Mailbox-Purge-Deletions-Folders.ps1 -Mailbox "Testuser@xzy.com" -Deletions:$true -admin "admin@xyz.com" -pass "adminpass"

#Empty purges Folder From user Mailbox "Testuser@xzy.com"
#>.\CleanUp-Mailbox-Purge-Deletions-Folders.ps1 -Mailbox "Testuser@xzy.com" -Purges:$true -admin "admin@xyz.com" -pass "adminpass"
#=====================================================================================================

param (
		$Mailbox,
		$Deletions=$false,
		$Purges=$false,
		$admin,
		$pass
      )

#Impersonate Admin Account details
$AccountWithImpersonationRights=$cred.UserName
$password=$cred.GetNetworkCredential().password
#$AccountWithImpersonationRights=$admin
#$password=$pass

#folder which you wants to empty from user mailbox
$FolderToEmpty=$folder

#EWS url for your orgnization
$uri=[system.URI] "https://outlook.office365.com/ews/exchange.asmx"
#define EWS Dll file Path
$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll"
Import-Module $dllpath
#-------------------------------------------------------------------------------------
## Set Exchange Version
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList `
$AccountWithImpersonationRights, $password
$service.url = $uri
$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId `
([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SMTPAddress,$Mailbox);
$MailboxRootid= new-object Microsoft.Exchange.WebServices.Data.FolderId `
        ([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Recoverableitemsroot,$ImpersonatedMailboxName)
$MailboxRoot=[Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$MailboxRootid)
$FolderList = new-object Microsoft.Exchange.WebServices.Data.FolderView(100)	
$FolderList.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
$findFolderResults = $MailboxRoot.FindFolders($FolderList)

if ($Deletions) {

	$Deletions = $findFolderResults | ? {$_.DisplayName -eq "Deletions"}
	Write-host "Item will be deleted from folder:" $Deletions.DisplayName
	Write-host "No of Items in Folder:"$Deletions.TotalCount
	Read-host "Hit enter to Remove all the items..."
	#$Deletions.Empty([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete,$False)
	$Deletions.Empty([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete,$True)
    "Done"
}

if ($Purges) {

	$Purges=$findFolderResults | ? {$_.DisplayName -eq "Purges"}
	Write-host "Item will be deleted from folder:" $Purges.DisplayName
	Write-host "No of Items in Folder:"$Purges.TotalCount
	Read-host "Hit enter to Remove all the items..."
	#$Purges.Empty([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete,$False)
	$Purges.Empty([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete,$True)
    "Done"
}
