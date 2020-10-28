#author=sunil chauhan
#myblog:www.lab365.in

param (
        $sender="labadmin@lab365.in",
        $subject="test",
        $mailbox="sunil.chauhan@lab365.in",
        $userName="labadmin@lab365.in",
        $password="myadminPassword"
      )

$SQ = "From:`"$Sender`" AND Subject:`"$subject`""
$report=@()
$itemsView=1000

$uri=[system.URI] "https://outlook.office365.com/ews/exchange.asmx"
$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
Import-Module $dllpath

$pass=$password
$AccountWithImpersonationRights=$userName
$MailboxToImpersonate=$mailbox

## Set Exchange Version
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList $AccountWithImpersonationRights, $pass
$service.url = $uri

#Write-Host 'Using ' $AccountWithImpersonationRights ' Account to work in ' $MailboxToImpersonate
$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId `
([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SMTPAddress,$MailboxToImpersonate);

$Folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId `
([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$ImpersonatedMailboxName)
#$MailboxRoot=[Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$Folderid)

$view = New-Object Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList $ItemsView
$propertyset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet ([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
$view.PropertySet = $propertyset

$items = $service.FindItems($Folderid,$SQ,$view)

if ($items -ne $null)
    {
    $emailProps = New-Object Microsoft.Exchange.WebServices.Data.PropertySet ([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
    $mail = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($service, $items.ID, $emailProps) 
    $datam=$mail | Select @{N="Sender";E={$_.Sender.Address}},@{N="USER";E={$mailbox}},Isread,Subject,DateTimeReceived,@{n="Folder";e={"Inbox"}}
    $report+=$datam
    }
Else 
    {
    Write-Host "Mail Not Found in Inbox Folder for:"$mailbox -f Yellow -NoNewline
    Write-Host " Checking Deleted Item Folder"
    $MailboxRootid= new-object Microsoft.Exchange.WebServices.Data.FolderId `
    ([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$ImpersonatedMailboxName)
    $MailboxRoot=[Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$MailboxRootid)
    $FolderList = new-object Microsoft.Exchange.WebServices.Data.FolderView(100)
    $FolderList.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
    $findFolderResults = $MailboxRoot.FindFolders($FolderList)
    $allFolders=$findFolderResults | ? {$_.FolderClass -eq "IPF.Note"} | select ID,Displayname
    $DI = $allFolders | ? {$_.DisplayName -eq "Deleted Items"}
    $Folderid=$DI.ID
    $items = $service.FindItems($Folderid,$SQ,$view)
    
     if ($items.count -eq $null) 
     {
        write-host "Item not found in the Deleted item folder, Now Checking in the Recover Deleted items Folder"
        $itemsView=90000
        $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList $ItemsView
        $propertyset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet ([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
        $view.PropertySet = $propertyset    
        $MailboxRootid= new-object Microsoft.Exchange.WebServices.Data.FolderId `
        ([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Recoverableitemsroot,$ImpersonatedMailboxName)
        $MailboxRoot=[Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$MailboxRootid)
        $FolderList = new-object Microsoft.Exchange.WebServices.Data.FolderView(100)
        $FolderList.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
        $findFolderResults = $MailboxRoot.FindFolders($FolderList)
        $Deletions = $findFolderResults | ? {$_.DisplayName -eq "Deletions"}
        $Folderid=$Deletions.ID
        $items=$service.FindItems($Folderid,$SQ,$view)

            if ($items.count -eq $null){

            Write-Host "Item Not Found in the Dumpsters."
            Write-host "Checking in other folders"
                
    $MailboxRootid= new-object Microsoft.Exchange.WebServices.Data.FolderId `
    ([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$ImpersonatedMailboxName)
    $MailboxRoot=[Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$MailboxRootid)
    $FolderList = new-object Microsoft.Exchange.WebServices.Data.FolderView(100)
    $FolderList.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
    $findFolderResults = $MailboxRoot.FindFolders($FolderList)
    $allFolders=$findFolderResults | ? {$_.FolderClass -eq "IPF.Note"} | select ID,Displayname
    $allFolders=$allFolders | ? { `
    $_.DisplayName -notlike "Inbox" -and `
    $_.DisplayName -notlike "Deleted Items" -and `
    $_.DisplayName -notlike "Drafts" -and `
    $_.DisplayName -notlike "Sent Items" -and `
    $_.DisplayName -notlike "Outbox"}    
    $allfoldersCount=$allfolders.count
    $counter=0
    $itemFound=$false
    if ($allFolders) {
	do {
        Write-Host "Checking Email Item in Folder:"$allfolders[$counter].DisplayName     
        $folderID=$allfolders[$counter].ID     
        $items =$service.FindItems($Folderid,$SQ,$view)         
        if ($items.count -eq $null) {Write-Host "Item Was Not Found in Folder:"$allfolders[$counter].DisplayName -ForegroundColor Yellow}     
        else {
                Write-Host "Item Was Found in Folder:"$allfolders[$counter].DisplayName -ForegroundColor Green     
                $emailProps = New-Object Microsoft.Exchange.WebServices.Data.PropertySet ([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)     
                $mail = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($service, $items.ID, $emailProps)      
                $data=$mail | Select @{N="Sender";E={$_.Sender.Address}},@{N="USER";E={$mailbox}},Isread,Subject,DateTimeReceived,@{n="Folder";e={$allfolders[$counter].DisplayName}}     
                $report+=$data         
                $itemFound=$true 
             } 
        $counter++ 
        } 
        until ($counter -eq $allfoldersCount -or $itemFound -eq $true) }
        }             
else    {
        $emailProps = New-Object Microsoft.Exchange.WebServices.Data.PropertySet ([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)             
        $mail = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($service, $items.ID, $emailProps)
        $data=$mail | Select @{N="Sender";E={$_.Sender.Address}},@{N="USER";E={$mailbox}},Isread,Subject,DateTimeReceived,@{n="Folder";e={"Deletions"}}
        $report+=$data}}          
        else  
        {        
        $emailProps = New-Object Microsoft.Exchange.WebServices.Data.PropertySet ([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties) 
        $mail = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($service, $items.ID, $emailProps)        
        $data=$mail | Select @{N="Sender";E={$_.Sender.Address}},@{N="USER";E={$mailbox}},Isread,Subject,DateTimeReceived,@{n="Folder";e={$DI.DisplayName}}       
        $report+=$data}} 
         
 $report
