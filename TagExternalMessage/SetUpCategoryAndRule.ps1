#region - Information- -------------------------------------------------------------------------    
#--------------------------------------------------------------------
# Project Scope: Setup a new category in the mailbox and create rule, the rule will tag the messages
# messages received from the external source
#
# Author: sunil chauhan <sunilkms@gmail.com>
# 
# How to use this script
# Requirement: 
# 1) - Service account with application impersonation and Exchange admin rights to manage 
# inbox rules,
# 2) - update the service account the Environmental variable section
# 2) - Install EWS Managed API and update the url in Environmental variable section
# 
# USAGE Example:
# dot source the ps1 file like  . "SetUpCategoryandRule.ps1" and run the below cmd.
#
# SetUpCategoryandRule -mailbox sunil@lab365.in -Category "MSG FROM EXTERNAL SOURCE"
#  
#endregion
#------------------------------------------------------------------------------------------------

#region - Environmental Variable-----------------------------------------------------------------
$ServiceAccountUPN="admin@lab365.onmicrosoft.com"
$ServiceAcPassword="AdminPassword"
$EWS = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
$EWSurl = "https://outlook.office365.com/EWS/Exchange.asmx"
#endregion
#------------------------------------------------------------------------------------------------

#region - Create EWS Client----------------------------------------------------------------------
Import-Module $EWS
$Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList $ServiceAccountUPN, $ServiceAcPassword
$Service.URL = $EWSurl
#endregion
#-------------------------------------------------------------------------------------------------

#region - Main Functions--------------------------------------------------------------------------
function GetMailboxCategory {
param($mailbox="sunil.chauhan@lab365.in")
    if ($mailbox) {
Write-Host "Fetching Existing Category from the Mailbox:$mailbox" -ForegroundColor Cyan

#Connecting with the impersonated user mailbox
$ArgumentList=([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SMTPAddress),$mailbox
$ImpUserId=New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId -ArgumentList $ArgumentList
$service.ImpersonatedUserId=$ImpUserId

#Bind with the FAI (Folder Associated Items) items
$folderid=new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$mailbox)     

#Specify the Calendar folder where the FAI Item is  
$UsrConfig = [Microsoft.Exchange.WebServices.Data.UserConfiguration]::Bind($service, "CategoryList", $folderid, [Microsoft.Exchange.WebServices.Data.UserConfigurationProperties]::All)  

#Get the XML in String Format  
$CatXML = [System.Text.Encoding]::UTF8.GetString($UsrConfig.XmlData)
[XML]$CatagoryXML=$CatXML
$Allcategory=$CatagoryXML.categories.category
$Allcategory | select name, guid,color
}
}
function AddNewCategory {
param (
       $mailbox="sunil.chauhan@lab365.in",
       $NewCategoryName="MESSAGE FROM EXTERNAL SOURCE"
      )
if($mailbox){
Write-Host "Adding a new Category Name [" -NoNewline
Write-Host $NewCategoryName -NoNewline -ForegroundColor Cyan
Write-Host "] for Mailbox:" -NoNewline 
Write-Host "$mailbox" -ForegroundColor Cyan
#Connecting with the impersonated user mailbox
$ArgumentList=([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SMTPAddress),$mailbox
$ImpUserId=New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId -ArgumentList $ArgumentList
$service.ImpersonatedUserId=$ImpUserId
#Bind with the FAI (Folder Associated Items) items
$Error.Clear()
$folderid=new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$mailbox)     
#Specify the Calendar folder where the FAI Item is  
$UsrConfig = [Microsoft.Exchange.WebServices.Data.UserConfiguration]::Bind($service, "CategoryList", $folderid, [Microsoft.Exchange.WebServices.Data.UserConfigurationProperties]::All)  
if ($Error){ "something went wrong, script will break"; $Error.Exception ; break}
#Get the XML in String Format  
$CatXML = [System.Text.Encoding]::UTF8.GetString($UsrConfig.XmlData)
[XML]$CatagoryXML=$CatXML
#$Allcategory=$CatagoryXML.categories.category
#$Allcategory | select name, guid,color
#$CatagoryXML
$NewCategory=$CatagoryXML.categories.category[0].Clone()  
#Set properties  
$NewCategory.name=$NewCategoryName  
$NewCategory.color = "16"  
$NewCategory.keyboardShortcut = "0"  
$NewCategory.guid = "{" + [System.Guid]::NewGuid().ToString() + "}"  
#$NewCategory.renameOnFirstUse = "0"  
[Void]$CatagoryXML.categories.AppendChild($NewCategory)  
$UsrConfig.XmlData = [System.Text.Encoding]::UTF8.GetBytes($CatagoryXML.OuterXml)  
#Update Item  
$UsrConfig.Update()
}
}
Function ValidateCagegory{
param ($mailbox)
$Error.Clear()
try{$ExistingCategories=GetMailboxCategory -mailbox $mailbox}
catch{Write-Host "$mailbox is not found" -ForegroundColor Yellow;$Error.Exception ;Break}
if($ExistingCategories){
if ($ExistingCategories -match "MSG FROM") {
Write-Host "Cagetories in the mailbox already setup" -ForegroundColor Green
} 
Else {
Write-Host "Cagetories not found in the mailbox" -ForegroundColor Yellow
}
}
}
Function SetupInboxRule {
param($mailbox)
New-InboxRule -Mailbox $mailbox -ApplyCategory "MSG FROM EXTERNAL SOURCE" `
-HeaderContainsWords "X-MS-Exchange-Organization-AuthAs: Anonymous" -Name "TAG EXTERNAL MSG:Do not delete this rule, created by Administrator"
}
Function SetUpCategoryandRule {
param ($mailbox,$Category)
#step 1 - Validate if the mailbox exist                     #Completed
#Step 2 - validate if the tag already already available.    #Completed
#Step 3 - Create the tag if it doesn't exist.               #Completed
#Step 4 - Validate if the Category addition was successful  #Completed
#step 5 - Create the rule.                                  #completed
#step 6 - Validate the rule creation.                       #InPrg
#Step 7 - Update user status.                               #InPrg

Write-Host "checking if a mailbox exist for:$mailbox"
$mbx = Get-Mailbox $mailbox -ea silentlycontinue
$CatAdditionSuccess=$false
if ($mbx) {
#step 2
Write-Host "Mailbox found"
Write-Host "Fetching existing cagetories in the mailbox"
$mbxcat = GetMailboxCategory $mailbox
if ($mbxcat) {
                if ($mbxcat.name -match $Category) 
                    {"Category already exist"} 
                else{
                     "Category Needs to be added"
                      "adding Cagetory"
                      AddNewCategory -mailbox $mailbox -NewCategoryName $Category
                      "fetching the category again to check if the new cat exist now"
                      $mbxc = GetMailboxCategory $mailbox
                      if ($mbxc.name -match $Category) {
                      Write-Host "Category has been created successfully"
                      $catAdditionSuccess=$true
                      } else {"Cagegory not found"}
                    }
            }
if ($CatAdditionSuccess){
"Creating Mailbox rule"
$Rulename = $Category + " [Do not delete created by Administrator]"
New-InboxRule -Mailbox $mailbox -ApplyCategory $Category -HeaderContainsWords "X-MS-Exchange-Organization-AuthAs: Anonymous"`
 -Name $RuleName # -whatif
 }
}
else {"Mailbox doesn't exist $mailbox"}
}
#endregion
#--------------------------------------------------------------------------------------------------
