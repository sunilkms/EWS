##### Project Scope: Setup a new category in the mailbox and create rule that assign external message category to email received the external source

How to use this script:
##### This script has the following Requirement:

- Service Account with application impersonation account and Exchange admin rights to mange 
Inbox rules update the service account the Environmental variable section
- Install EWS Managed API and update the url in Environmental variable section

##### USAGE Example:
Dot source the ps1 file like  . "SetUpCategoryandRule.ps1" and run the below cmd.
##### SetUpCategoryandRule -mailbox sunil@lab365.in -Category "MSG FROM EXTERNAL SOURCE"
