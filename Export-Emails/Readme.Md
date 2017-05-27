
This script can be used where you wants to restore items from user mailbox purges folder, exports the items from users Purges folder,
 Items Once Deleted from Recover Deleted items, they are out of user control and can't be restored, this script can be used to export the items.
 This script will export each items in .eml format and the item name would there subject.
 
 Eml file cane be simply copy and pasted into user outlook.

 Usage Example: Get messages details in Purges & Deletions Folders.
 >.\Export-Email-Ews.ps1 -folder purges -itemsView 10 -mailbox test@user.com -userName admin@user.com -password Pass | ft
 >.\Export-Email-Ews.ps1 -folder deletions -itemsView 10 -mailbox test@user.com -userName admin@user.com -password Pass | ft

 
 Usage Example 1: to export items from Purges folder run the script with below paramiters.
 >.\Export-Email-Ews.ps1 -export $true -folder purges -itemsView 10 -mailbox test@user.com -dir "D:\Purges\" -userName admin@user.com -password Pass -export $true

 Usage Example 2: to export items from Purges folder run the script with below paramiters.
 >.\Export-Email-Ews.ps1 -export $true -folder Deletions -itemsView 10 -mailbox test@user.com -dir "D:\Deletions\" -userName admin@user.com -password Pass -export $true

For More Details on this plz checkout post on this script here:
http://www.sunilchauhan.info/2017/02/exporting-emails-from-office-365-and.html
