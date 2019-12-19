#region - Information- -------------------------------------------------------------------------    
#--------------------------------------------------------------------
# Project Scope: Setup a new category in the mailbox and create rule which will show warning to user
# when a message is received from the external source.
#
# Author: sunil chauhan <sunilkms@gmail.com>
# 
# How to use this script
# Requirement: 
# 1) - Service Account with application impersonation account and Exchange admin rights to mange 
# inbox rules update the service account the Environmental variable section
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
$ServiceAccountUPN="svclabadmin@lab365.onmicrosoft.com"
$ServiceAcPassword="Password"
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
param($mailbox)
    if ($mailbox) 
         {
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
$CatXML=[System.Text.Encoding]::UTF8.GetString($UsrConfig.XmlData)

try {
[XML]$CatagoryXML=$CatXML
} catch {
[XML]$CatagoryXML=$CatXML.SubString(1)
}

# $CatXML.categories.category # [xml]$CatXML.SubString(1)
$Allcategory=$CatagoryXML.categories.category
$Allcategory | select name, guid,color

} 
    else {"Mailbox has not been supplied."}
}
function AddNewCategory {
param (
       $mailbox="labuser@lab365.com",
       $NewCategoryName
      )
      
if($mailbox){
Write-Host "Adding a new category with Name [" -NoNewline
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

# Some mailbox may have some unsupported XML characters, so handle them using substring.

try {
[XML]$CatagoryXML=$CatXML
} catch {
[XML]$CatagoryXML=$CatXML.SubString(1)
}

# Clone and setup New category.
if ($CatagoryXML.categories) {
$NewCategory=$CatagoryXML.categories.category[0].Clone() 
$NewCategory.name=$NewCategoryName  
$NewCategory.color = "16"  
$NewCategory.keyboardShortcut = "0"  
$NewCategory.guid = "{" + [System.Guid]::NewGuid().ToString() + "}"  
#$NewCategory.renameOnFirstUse = "0"  
[Void]$CatagoryXML.categories.AppendChild($NewCategory)  
$UsrConfig.XmlData = [System.Text.Encoding]::UTF8.GetBytes($CatagoryXML.OuterXml)  
#Update Item  
$UsrConfig.Update() } else {"Something went wronge, Cagegories were missing."}
}
}
Function SetupInboxRule {
param($mailbox,$Category)

Write-Host "Creating a new inbox rule for $mailbox" -ForegroundColor Yellow
$Rulename = "Categorise emails from external source [Do not delete or disable created by Administrator]"
New-InboxRule -Mailbox $mailbox -ApplyCategory $Category -HeaderContainsWords "X-MS-Exchange-Organization-AuthAs: Anonymous"`
 -Name $RuleName # -whatif

}
Function SetUpExternalEmailCategoryForUser {

param (
       $mailbox,
       $Category="MSG FROM EXTERNAL SOURCE"
      )

# Project Tracking----------------------------------------------------
#step 1 - Validate if the mailbox exist                     #Completed
#Step 2 - validate if the tag already already available.    #Completed
#Step 3 - Create the tag if it doesn't exist.               #Completed
#Step 4 - Validate if the Category addition was successful  #Completed
#step 5 - Create the rule.                                  #completed
#step 6 - Validate the rule creation.                       #InPrg
#Step 7 - Update user status.                               #InPrg

Write-Host "Checking if a mailbox exist for:$mailbox" -ForegroundColor Cyan
$mbx = Get-Mailbox $mailbox -ea silentlycontinue
$CatAdditionSuccess=$false

if ($mbx) {

#step 2
Write-Host "Mailbox found"
Write-Host "Fetching existing cagetories in the mailbox"
$mbxcat=GetMailboxCategory $mailbox

if ($mbxcat) {
                if ($mbxcat.name -match $Category) 
                    {
                     Write-Host "Category with the name [$Category] already exist in the mailbox" -ForegroundColor Cyan } 
                else{
                     Write-Host "Requested category was not found in the mailbox" -ForegroundColor Cyan
                     Write-Host "Trying adding a new cagetory" -ForegroundColor Cyan                   
                     AddNewCategory -mailbox $mailbox -NewCategoryName $Category
                     Write-Host "Fetching the category again to check if the new category exist now" -ForegroundColor Cyan
                     $mbxc = GetMailboxCategory $mailbox
                     if ($mbxc.name -match $Category) 
                            {
                            Write-Host "Category has been created successfully" -ForegroundColor Green
                            $catAdditionSuccess=$true
                            }
                     else  {
                            Write-Host "Category was not found" -ForegroundColor Yellow
                           }
                    }
            }
if ($CatAdditionSuccess){ SetupInboxRule -mailbox $mailbox -Category $Category }
}
else {
       Write-Host "Mailbox doesn't exist $mailbox" -ForegroundColor Yellow
     }
}
#endregion
#--------------------------------------------------------------------------------------------------
