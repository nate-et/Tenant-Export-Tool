This tools is a PowerShell based tool that was designed to report on the following: 

- User Mailboxes (with full delegation details) 
- Shared Mailboxes (with delegation and permissions)
- Distribution Lists (with complete membership)
- Security Groups (with descriptions and members)
- Office 365 and Teams Groups (with detailed membership) 
- Exchange Permissions and delegation mapping 

You must use PowerShell 7 in order to run this script. you can find the download here: 
https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.5

This script uses the Microsoft Graph API and Exchange Online PowerShell to gather this information and export it to an excel spreadsheet (saved to the location of your choosing). The script will check that you have the required modules in order to run and will attempt to install them for you. You do not need to run this script as an administrator. 

Before running this script you are advised to login on a separate profile to the tenant you want to query, making sure you have signed in and making sure that it's the last browser window you clicked as the script will automatically open the prompt to authenticate with that tenant. When asked, you want to check the box to consent on behalf of the organisation otherwise you will run into permission issues and the script will fail to run as intended. 

You will additionally be prompted to connect to Exchange Online so have the credentials to hand. During the course of the script running you may see red text telling you the operation could not be performed because object "mailbox name" could not be found. This is normal. 