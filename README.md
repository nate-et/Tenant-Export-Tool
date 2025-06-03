# 📊 PowerShell Office 365 Reporting Tool

This PowerShell 7-based script gathers and exports detailed reporting on the following Microsoft 365 resources:

## 🔍 What It Reports

- **User Mailboxes** – Includes full delegation details  
- **Shared Mailboxes** – Includes delegation and permission information  
- **Distribution Lists** – Complete membership breakdown  
- **Security Groups** – Member list and group description  
- **Office 365 & Teams Groups** – Detailed membership  
- **Exchange Permissions** – Delegation mapping across mailboxes  

## ⚙️ Prerequisites

- **PowerShell 7 is required**  
  👉 [Download PowerShell 7.5 here](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.5)

- **Microsoft Graph API & Exchange Online PowerShell modules**  
  The script checks for required modules and will attempt to install any that are missing.

> ⚠️ **Windows Defender Note**  
During module installation, Defender may block `pwsh.exe` from accessing controlled folders. Watch for toast notifications (bottom-right of your screen) and **allow access to your Documents folder** if prompted.

## 📁 Output

- All data is exported to an **Excel spreadsheet**  
- You’ll choose the save location during execution

## 🔐 Authentication Steps

Before running the script:

1. **Login to the target tenant** using a separate browser profile  
2. Make sure it’s the **last active browser window** before running the script  
3. When prompted:
   - Check **“Consent on behalf of the organization”**  
   - This is required to avoid permission-related failures

You will also be prompted to **connect to Exchange Online**, so have your credentials ready.

> 🔴 **Note:** During execution, you may see red text like:
>  
> `"The operation could not be performed because object 'mailbox name' could not be found."`  
>  
> This is **normal behavior** and can be safely ignored.

## ✅ Admin Rights

You **do not need to run** this script as an administrator.
