Project Scope: Setup a new category in the mailbox and create rule
Author: sunil chauhan <sunilkms@gmail.com>
 
How to use this script
Requirement: 
1) - Service Account with application impersonation account and Exchange admin rights to mange 
Inbox rules update the service account the Environmental variable section
2) - Install EWS Managed API and update the url in Environmental variable section

USAGE Example:
Dot source the ps1 file like  . "SetUpCategoryandRule.ps1" and run the below cmd.

SetUpCategoryandRule -mailbox sunil@lab365.in -Category "MSG FROM EXTERNAL SOURCE"
