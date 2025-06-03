# Office 365 Tenant Report Generator with Exchange Online Integration
# This script generates a comprehensive Excel report of your Office 365 tenant
# Author: PowerShell Automation Script
# Requirements: PowerShell 7, Microsoft Graph PowerShell SDK, Exchange Online Management Module

# Clear the screen for better visibility
Clear-Host

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "Office 365 Tenant Report Generator" -ForegroundColor Cyan
Write-Host "Enhanced with Exchange Online Integration" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "This script will generate a comprehensive Excel report of your Office 365 tenant." -ForegroundColor Yellow
Write-Host "The report will include:" -ForegroundColor Yellow
Write-Host "- User Mailboxes (with full delegation details)" -ForegroundColor White
Write-Host "- Shared Mailboxes (with delegation and permissions)" -ForegroundColor White
Write-Host "- Distribution Lists (with complete membership)" -ForegroundColor White
Write-Host "- Security Groups (with descriptions and members)" -ForegroundColor White
Write-Host "- Office 365 and Teams Groups (with detailed membership)" -ForegroundColor White
Write-Host "- Exchange permissions and delegation mapping" -ForegroundColor White
Write-Host ""

# Function to check if running as administrator (recommended but not required)
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Function to install required modules
function Install-RequiredModules {
    Write-Host "Checking for required PowerShell modules..." -ForegroundColor Yellow
    
    $requiredModules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Groups',
        'Microsoft.Graph.Mail',
        'ExchangeOnlineManagement',
        'ImportExcel'
    )
    
    foreach ($module in $requiredModules) {
        Write-Host "Checking for module: $module" -ForegroundColor Gray
        
        if (!(Get-Module -ListAvailable -Name $module)) {
            Write-Host "Module $module not found. Installing..." -ForegroundColor Yellow
            try {
                Install-Module -Name $module -Force -AllowClobber -Scope CurrentUser
                Write-Host "Successfully installed $module" -ForegroundColor Green
            }
            catch {
                Write-Host "Failed to install $module. Error: $($_.Exception.Message)" -ForegroundColor Red
                Write-Host "Please run PowerShell as Administrator and try again." -ForegroundColor Red
                exit 1
            }
        }
        else {
            Write-Host "Module $module is already installed" -ForegroundColor Green
        }
    }
    Write-Host "All required modules are available!" -ForegroundColor Green
    Write-Host ""
}

# Function to connect to Microsoft Graph
function Connect-ToGraph {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
    Write-Host "A browser window will open for authentication." -ForegroundColor Cyan
    Write-Host "Please sign in with your Office 365 administrator account." -ForegroundColor Cyan
    
    try {
        # Define required scopes for the operations we need to perform
        $scopes = @(
            'User.Read.All',
            'Group.Read.All',
            'GroupMember.Read.All',
            'Directory.Read.All',
            'Team.ReadBasic.All',
            'MailboxSettings.Read'
        )
        
        Connect-MgGraph -Scopes $scopes -NoWelcome
        Write-Host "Successfully connected to Microsoft Graph!" -ForegroundColor Green
        
        # Get tenant information
        $context = Get-MgContext
        Write-Host "Connected to tenant: $($context.TenantId)" -ForegroundColor Cyan
        Write-Host ""
        
        return $true
    }
    catch {
        Write-Host "Failed to connect to Microsoft Graph. Error: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function to connect to Exchange Online
function Connect-ToExchangeOnline {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
    Write-Host "Please use the same credentials as Microsoft Graph." -ForegroundColor Cyan
    
    try {
        Connect-ExchangeOnline -ShowBanner:$false
        Write-Host "Successfully connected to Exchange Online!" -ForegroundColor Green
        Write-Host ""
        return $true
    }
    catch {
        Write-Host "Failed to connect to Exchange Online. Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Some delegation information may not be available." -ForegroundColor Yellow
        return $false
    }
}

# Function to get detailed mailbox permissions
function Get-DetailedMailboxPermissions {
    param(
        [string]$Identity,
        [string]$DisplayName
    )
    
    $permissions = @{
        'FullAccess' = @()
        'SendAs' = @()
        'SendOnBehalf' = @()
    }
    
    try {
        # Get Full Access permissions
        $fullAccessPerms = Get-MailboxPermission -Identity $Identity | Where-Object { 
            $_.User -notlike "NT AUTHORITY\*" -and 
            $_.User -notlike "S-1-*" -and 
            $_.IsInherited -eq $false -and
            $_.AccessRights -contains "FullAccess"
        }
        
        foreach ($perm in $fullAccessPerms) {
            $permissions['FullAccess'] += "$($perm.User) (Full Access)"
        }
        
        # Get Send As permissions
        $sendAsPerms = Get-RecipientPermission -Identity $Identity | Where-Object { 
            $_.Trustee -notlike "NT AUTHORITY\*" -and 
            $_.Trustee -notlike "S-1-*" -and
            $_.AccessRights -contains "SendAs"
        }
        
        foreach ($perm in $sendAsPerms) {
            $permissions['SendAs'] += "$($perm.Trustee) (Send As)"
        }
        
        # Get Send on Behalf permissions
        $mailbox = Get-Mailbox -Identity $Identity
        if ($mailbox.GrantSendOnBehalfTo) {
            foreach ($delegate in $mailbox.GrantSendOnBehalfTo) {
                $permissions['SendOnBehalf'] += "$($delegate) (Send on Behalf)"
            }
        }
    }
    catch {
        Write-Warning "Could not retrieve permissions for $DisplayName : $($_.Exception.Message)"
    }
    
    return $permissions
}

# Function to get user mailboxes with detailed delegation
function Get-UserMailboxes {
    Write-Host "Retrieving user mailboxes with delegation information..." -ForegroundColor Yellow
    
    try {
        $users = Get-MgUser -All -Property "DisplayName,UserPrincipalName,Mail,AccountEnabled,UserType,CreatedDateTime" | 
                 Where-Object { $_.UserType -eq "Member" -and $_.Mail -ne $null -and $_.AccountEnabled -eq $true }
        
        $userMailboxes = @()
        $counter = 0
        
        foreach ($user in $users) {
            $counter++
            Write-Progress -Activity "Processing User Mailboxes" -Status "Processing $($user.DisplayName) ($counter of $($users.Count))" -PercentComplete (($counter / $users.Count) * 100)
            
            # Get detailed mailbox permissions using Exchange Online
            $permissions = @{
                'FullAccess' = @()
                'SendAs' = @()
                'SendOnBehalf' = @()
            }
            
            try {
                $permissions = Get-DetailedMailboxPermissions -Identity $user.UserPrincipalName -DisplayName $user.DisplayName
            }
            catch {
                Write-Warning "Could not get Exchange permissions for $($user.DisplayName)"
            }
            
            $allDelegates = @()
            $allDelegates += $permissions['FullAccess']
            $allDelegates += $permissions['SendAs'] 
            $allDelegates += $permissions['SendOnBehalf']
            
            $userMailboxes += [PSCustomObject]@{
                'Display Name' = $user.DisplayName
                'Email Address' = $user.Mail
                'User Principal Name' = $user.UserPrincipalName
                'Account Enabled' = $user.AccountEnabled
                'Created Date' = $user.CreatedDateTime
                'Full Access Delegates' = ($permissions['FullAccess'] -join "; ")
                'Send As Delegates' = ($permissions['SendAs'] -join "; ")
                'Send On Behalf Delegates' = ($permissions['SendOnBehalf'] -join "; ")
                'All Delegates' = ($allDelegates -join "; ")
                'Delegate Count' = $allDelegates.Count
            }
        }
        
        Write-Progress -Activity "Processing User Mailboxes" -Completed
        Write-Host "Retrieved $($userMailboxes.Count) user mailboxes" -ForegroundColor Green
        return $userMailboxes
    }
    catch {
        Write-Host "Error retrieving user mailboxes: $($_.Exception.Message)" -ForegroundColor Red
        return @()
    }
}

# Function to get shared mailboxes with detailed permissions
function Get-SharedMailboxes {
    Write-Host "Retrieving shared mailboxes with delegation information..." -ForegroundColor Yellow
    
    try {
        # Get shared mailboxes from Exchange Online directly
        $sharedMailboxes = @()
        
        try {
            $sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited
        }
        catch {
            Write-Warning "Could not get shared mailboxes from Exchange Online, falling back to Graph API"
            # Fallback to Graph API method
            $sharedMailboxes = Get-MgUser -All -Property "DisplayName,UserPrincipalName,Mail,AccountEnabled,UserType,CreatedDateTime" | 
                              Where-Object { $_.UserType -eq "Member" -and $_.AccountEnabled -eq $false -and $_.Mail -ne $null }
        }
        
        $sharedMailboxData = @()
        $counter = 0
        
        foreach ($mailbox in $sharedMailboxes) {
            $counter++
            $displayName = if ($mailbox.DisplayName) { $mailbox.DisplayName } else { $mailbox.Name }
            $userPrincipalName = if ($mailbox.UserPrincipalName) { $mailbox.UserPrincipalName } else { $mailbox.PrimarySmtpAddress }
            $emailAddress = if ($mailbox.Mail) { $mailbox.Mail } else { $mailbox.PrimarySmtpAddress }
            
            Write-Progress -Activity "Processing Shared Mailboxes" -Status "Processing $displayName ($counter of $($sharedMailboxes.Count))" -PercentComplete (($counter / $sharedMailboxes.Count) * 100)
            
            # Get detailed permissions
            $permissions = Get-DetailedMailboxPermissions -Identity $userPrincipalName -DisplayName $displayName
            
            $allDelegates = @()
            $allDelegates += $permissions['FullAccess']
            $allDelegates += $permissions['SendAs']
            $allDelegates += $permissions['SendOnBehalf']
            
            $sharedMailboxData += [PSCustomObject]@{
                'Display Name' = $displayName
                'Email Address' = $emailAddress
                'User Principal Name' = $userPrincipalName
                'Created Date' = if ($mailbox.CreatedDateTime) { $mailbox.CreatedDateTime } else { $mailbox.WhenCreated }
                'Full Access Delegates' = ($permissions['FullAccess'] -join "; ")
                'Send As Delegates' = ($permissions['SendAs'] -join "; ")
                'Send On Behalf Delegates' = ($permissions['SendOnBehalf'] -join "; ")
                'All Delegates' = ($allDelegates -join "; ")
                'Delegate Count' = $allDelegates.Count
            }
        }
        
        Write-Progress -Activity "Processing Shared Mailboxes" -Completed
        Write-Host "Retrieved $($sharedMailboxData.Count) shared mailboxes" -ForegroundColor Green
        return $sharedMailboxData
    }
    catch {
        Write-Host "Error retrieving shared mailboxes: $($_.Exception.Message)" -ForegroundColor Red
        return @()
    }
}

# Function to get distribution lists with enhanced member details
function Get-DistributionLists {
    Write-Host "Retrieving distribution lists with detailed membership..." -ForegroundColor Yellow
    
    try {
        # Try to get from Exchange Online first for more accurate results
        $distributionLists = @()
        
        try {
            $distributionLists = Get-DistributionGroup -ResultSize Unlimited
            $useExchangeData = $true
        }
        catch {
            Write-Warning "Could not get distribution groups from Exchange Online, using Graph API"
            $distributionLists = Get-MgGroup -All -Property "DisplayName,Mail,Description,CreatedDateTime,GroupTypes" | 
                               Where-Object { $_.Mail -ne $null -and ($_.GroupTypes -eq $null -or $_.GroupTypes.Count -eq 0) }
            $useExchangeData = $false
        }
        
        $distributionListData = @()
        $counter = 0
        
        foreach ($dl in $distributionLists) {
            $counter++
            $displayName = if ($dl.DisplayName) { $dl.DisplayName } else { $dl.Name }
            Write-Progress -Activity "Processing Distribution Lists" -Status "Processing $displayName ($counter of $($distributionLists.Count))" -PercentComplete (($counter / $distributionLists.Count) * 100)
            
            # Get members with detailed information
            $members = @()
            $memberDetails = @()
            $memberTypes = @()
            
            try {
                if ($useExchangeData) {
                    # Use Exchange Online to get members
                    $groupMembers = Get-DistributionGroupMember -Identity $dl.Identity -ResultSize Unlimited
                    foreach ($member in $groupMembers) {
                        $members += $member.DisplayName
                        $memberTypes += $member.RecipientType
                        $memberDetails += "$($member.DisplayName) ($($member.PrimarySmtpAddress)) [$($member.RecipientType)]"
                    }
                }
                else {
                    # Use Graph API
                    $groupMembers = Get-MgGroupMember -GroupId $dl.Id -All
                    foreach ($member in $groupMembers) {
                        try {
                            $userInfo = Get-MgUser -UserId $member.Id -Property "DisplayName,UserPrincipalName,Mail" -ErrorAction SilentlyContinue
                            if ($userInfo) {
                                $members += $userInfo.DisplayName
                                $memberTypes += "User"
                                $memberDetails += "$($userInfo.DisplayName) ($($userInfo.UserPrincipalName)) [User]"
                            }
                            else {
                                $groupInfo = Get-MgGroup -GroupId $member.Id -Property "DisplayName,Mail" -ErrorAction SilentlyContinue
                                if ($groupInfo) {
                                    $members += $groupInfo.DisplayName
                                    $memberTypes += "Group"
                                    $memberDetails += "$($groupInfo.DisplayName) [Group]"
                                }
                            }
                        }
                        catch {
                            $members += "Unknown Member"
                            $memberTypes += "Unknown"
                            $memberDetails += "Unknown Member ($($member.Id))"
                        }
                    }
                }
            }
            catch {
                $members = @("Unable to retrieve members")
                $memberDetails = @("Unable to retrieve members - insufficient permissions")
                $memberTypes = @("Unknown")
            }
            
            $distributionListData += [PSCustomObject]@{
                'Display Name' = $displayName
                'Email Address' = if ($dl.Mail) { $dl.Mail } else { $dl.PrimarySmtpAddress }
                'Description' = $dl.Description
                'Created Date' = if ($dl.CreatedDateTime) { $dl.CreatedDateTime } else { $dl.WhenCreated }
                'Member Count' = $members.Count
                'Members' = ($members -join "; ")
                'Member Details' = ($memberDetails -join "; ")
                'Member Types' = ($memberTypes -join "; ")
                'Requires Sender Authentication' = if ($useExchangeData) { $dl.RequireSenderAuthenticationEnabled } else { "Unknown" }
            }
        }
        
        Write-Progress -Activity "Processing Distribution Lists" -Completed
        Write-Host "Retrieved $($distributionListData.Count) distribution lists" -ForegroundColor Green
        return $distributionListData
    }
    catch {
        Write-Host "Error retrieving distribution lists: $($_.Exception.Message)" -ForegroundColor Red
        return @()
    }
}

# Function to get security groups with enhanced membership details
function Get-SecurityGroups {
    Write-Host "Retrieving security groups with detailed membership..." -ForegroundColor Yellow
    
    try {
        $securityGroups = Get-MgGroup -All -Property "DisplayName,Description,CreatedDateTime,GroupTypes,SecurityEnabled,Mail" | 
                         Where-Object { $_.SecurityEnabled -eq $true }
        
        $securityGroupData = @()
        $counter = 0
        
        foreach ($group in $securityGroups) {
            $counter++
            Write-Progress -Activity "Processing Security Groups" -Status "Processing $($group.DisplayName) ($counter of $($securityGroups.Count))" -PercentComplete (($counter / $securityGroups.Count) * 100)
            
            # Get members and owners with detailed information
            $members = @()
            $memberDetails = @()
            $owners = @()
            $memberTypes = @()
            
            try {
                # Get members
                $groupMembers = Get-MgGroupMember -GroupId $group.Id -All
                foreach ($member in $groupMembers) {
                    try {
                        $userInfo = Get-MgUser -UserId $member.Id -Property "DisplayName,UserPrincipalName,Mail,UserType" -ErrorAction SilentlyContinue
                        if ($userInfo) {
                            $members += $userInfo.DisplayName
                            $memberTypes += $userInfo.UserType
                            $memberDetails += "$($userInfo.DisplayName) ($($userInfo.UserPrincipalName)) [$($userInfo.UserType)]"
                        }
                        else {
                            $groupInfo = Get-MgGroup -GroupId $member.Id -Property "DisplayName,Mail" -ErrorAction SilentlyContinue
                            if ($groupInfo) {
                                $members += $groupInfo.DisplayName
                                $memberTypes += "Group"
                                $memberDetails += "$($groupInfo.DisplayName) [Group]"
                            }
                            else {
                                $members += "Unknown Member"
                                $memberTypes += "Unknown"
                                $memberDetails += "Unknown Member ($($member.Id))"
                            }
                        }
                    }
                    catch {
                        $members += "Unable to retrieve member info"
                        $memberTypes += "Unknown"
                        $memberDetails += "Unable to retrieve member info ($($member.Id))"
                    }
                }
                
                # Get owners
                $groupOwners = Get-MgGroupOwner -GroupId $group.Id -All
                foreach ($owner in $groupOwners) {
                    try {
                        $ownerInfo = Get-MgUser -UserId $owner.Id -Property "DisplayName,UserPrincipalName" -ErrorAction SilentlyContinue
                        if ($ownerInfo) {
                            $owners += "$($ownerInfo.DisplayName) ($($ownerInfo.UserPrincipalName))"
                        }
                        else {
                            $owners += "Unknown Owner ($($owner.Id))"
                        }
                    }
                    catch {
                        $owners += "Unable to retrieve owner info ($($owner.Id))"
                    }
                }
            }
            catch {
                $members = @("Unable to retrieve members")
                $memberDetails = @("Unable to retrieve members - insufficient permissions")
                $owners = @("Unable to retrieve owners")
                $memberTypes = @("Unknown")
            }
            
            $securityGroupData += [PSCustomObject]@{
                'Display Name' = $group.DisplayName
                'Description' = $group.Description
                'Email Address' = $group.Mail
                'Created Date' = $group.CreatedDateTime
                'Group Types' = ($group.GroupTypes -join "; ")
                'Member Count' = $members.Count
                'Owner Count' = $owners.Count
                'Members' = ($members -join "; ")
                'Member Details' = ($memberDetails -join "; ")
                'Member Types' = ($memberTypes -join "; ")
                'Owners' = ($owners -join "; ")
            }
        }
        
        Write-Progress -Activity "Processing Security Groups" -Completed
        Write-Host "Retrieved $($securityGroupData.Count) security groups" -ForegroundColor Green
        return $securityGroupData
    }
    catch {
        Write-Host "Error retrieving security groups: $($_.Exception.Message)" -ForegroundColor Red
        return @()
    }
}

# Function to get Office 365 and Teams groups with enhanced details
function Get-NormalAndTeamsGroups {
    Write-Host "Retrieving Office 365 and Teams groups with detailed membership..." -ForegroundColor Yellow
    
    try {
        $groups = Get-MgGroup -All -Property "DisplayName,Mail,Description,CreatedDateTime,GroupTypes,Visibility" | 
                 Where-Object { $_.GroupTypes -contains "Unified" }
        
        $groupData = @()
        $counter = 0
        
        foreach ($group in $groups) {
            $counter++
            Write-Progress -Activity "Processing Office 365/Teams Groups" -Status "Processing $($group.DisplayName) ($counter of $($groups.Count))" -PercentComplete (($counter / $groups.Count) * 100)
            
            # Determine if it's a Teams group
            $isTeamsGroup = $false
            $teamsInfo = ""
            try {
                $team = Get-MgTeam -TeamId $group.Id -ErrorAction SilentlyContinue
                if ($team) {
                    $isTeamsGroup = $true
                    $teamsInfo = "Teams-enabled"
                }
            }
            catch {
                $isTeamsGroup = $false
                $teamsInfo = "Office 365 Group only"
            }
            
            # Get members with detailed information
            $members = @()
            $memberDetails = @()
            $owners = @()
            $memberTypes = @()
            
            try {
                # Get regular members
                $groupMembers = Get-MgGroupMember -GroupId $group.Id -All
                foreach ($member in $groupMembers) {
                    try {
                        $userInfo = Get-MgUser -UserId $member.Id -Property "DisplayName,UserPrincipalName,Mail,UserType" -ErrorAction SilentlyContinue
                        if ($userInfo) {
                            $members += $userInfo.DisplayName
                            $memberTypes += $userInfo.UserType
                            $memberDetails += "$($userInfo.DisplayName) ($($userInfo.UserPrincipalName)) [$($userInfo.UserType)]"
                        }
                        else {
                            $members += "Unknown Member"
                            $memberTypes += "Unknown"
                            $memberDetails += "Unknown Member ($($member.Id))"
                        }
                    }
                    catch {
                        $members += "Unable to retrieve member info"
                        $memberTypes += "Unknown"
                        $memberDetails += "Unable to retrieve member info ($($member.Id))"
                    }
                }
                
                # Get owners
                $groupOwners = Get-MgGroupOwner -GroupId $group.Id -All
                foreach ($owner in $groupOwners) {
                    try {
                        $ownerInfo = Get-MgUser -UserId $owner.Id -Property "DisplayName,UserPrincipalName" -ErrorAction SilentlyContinue
                        if ($ownerInfo) {
                            $owners += "$($ownerInfo.DisplayName) ($($ownerInfo.UserPrincipalName))"
                        }
                        else {
                            $owners += "Unknown Owner ($($owner.Id))"
                        }
                    }
                    catch {
                        $owners += "Unable to retrieve owner info ($($owner.Id))"
                    }
                }
            }
            catch {
                $members = @("Unable to retrieve members")
                $memberDetails = @("Unable to retrieve members - insufficient permissions")
                $owners = @("Unable to retrieve owners")
                $memberTypes = @("Unknown")
            }
            
            $groupData += [PSCustomObject]@{
                'Display Name' = $group.DisplayName
                'Email Address' = $group.Mail
                'Description' = $group.Description
                'Type' = if ($isTeamsGroup) { "Teams Group" } else { "Office 365 Group" }
                'Visibility' = $group.Visibility
                'Teams Status' = $teamsInfo
                'Created Date' = $group.CreatedDateTime
                'Member Count' = $members.Count
                'Owner Count' = $owners.Count
                'Members' = ($members -join "; ")
                'Member Details' = ($memberDetails -join "; ")
                'Member Types' = ($memberTypes -join "; ")
                'Owners' = ($owners -join "; ")
            }
        }
        
        Write-Progress -Activity "Processing Office 365/Teams Groups" -Completed
        Write-Host "Retrieved $($groupData.Count) Office 365/Teams groups" -ForegroundColor Green
        return $groupData
    }
    catch {
        Write-Host "Error retrieving Office 365/Teams groups: $($_.Exception.Message)" -ForegroundColor Red
        return @()
    }
}

# Function to export data to Excel with enhanced formatting
function Export-ToExcel {
    param(
        [hashtable]$Data,
        [string]$FilePath
    )
    
    Write-Host "Exporting data to Excel file: $FilePath" -ForegroundColor Yellow
    
    try {
        # Remove the file if it exists
        if (Test-Path $FilePath) {
            Remove-Item $FilePath -Force
        }
        
        # Export each dataset to a separate worksheet
        foreach ($sheetName in $Data.Keys) {
            if ($Data[$sheetName].Count -gt 0) {
                $Data[$sheetName] | Export-Excel -Path $FilePath -WorksheetName $sheetName -AutoSize -TableStyle Medium2 -FreezeTopRow
                Write-Host "Exported $($Data[$sheetName].Count) items to worksheet: $sheetName" -ForegroundColor Green
            }
            else {
                # Create empty worksheet with headers if no data
                @([PSCustomObject]@{'No Data' = 'No items found for this category'}) | 
                Export-Excel -Path $FilePath -WorksheetName $sheetName -AutoSize
                Write-Host "Created empty worksheet: $sheetName (no data found)" -ForegroundColor Yellow
            }
        }
        
        Write-Host "Excel file created successfully!" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Error creating Excel file: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Main execution
try {
    # Check PowerShell version
    if ($PSVersionTable.PSVersion.Major -lt 7) {
        Write-Host "This script requires PowerShell 7 or later. You are running version $($PSVersionTable.PSVersion)" -ForegroundColor Red
        Write-Host "Please install PowerShell 7 from: https://github.com/PowerShell/PowerShell/releases" -ForegroundColor Yellow
        exit 1
    }
    
    # Install required modules
    Install-RequiredModules
    
    # Connect to Microsoft Graph
    if (-not (Connect-ToGraph)) {
        Write-Host "Unable to connect to Microsoft Graph. Exiting." -ForegroundColor Red
        exit 1
    }
    
    # Connect to Exchange Online
    $exchangeConnected = Connect-ToExchangeOnline
    if (-not $exchangeConnected) {
        Write-Host "Warning: Exchange Online connection failed. Delegation details will be limited." -ForegroundColor Yellow
        Write-Host "Continuing with available data..." -ForegroundColor Yellow
        Write-Host ""
    }
    
    # Get save location from user
    Write-Host "Please choose where to save the report file..." -ForegroundColor Cyan
    
    # Generate default filename with current date/time
    $dateTime = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $defaultFileName = "Enhanced_Tenant_Report_$dateTime.xlsx"
    
    # Try to use GUI dialog first, with fallback to text input
    $exportPath = $null
    
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx"
        $saveDialog.Title = "Save Enhanced Tenant Report"
        $saveDialog.FileName = $defaultFileName
        $saveDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
        
        Write-Host "Opening file save dialog..." -ForegroundColor Gray
        $result = $saveDialog.ShowDialog()
        
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $exportPath = $saveDialog.FileName
            Write-Host "Selected location: $exportPath" -ForegroundColor Green
        }
        else {
            Write-Host "Dialog cancelled, using default location..." -ForegroundColor Yellow
            $exportPath = Join-Path -Path (Get-Location) -ChildPath $defaultFileName
        }
    }
    catch {
        Write-Host "GUI dialog not available, using text input method..." -ForegroundColor Yellow
        Write-Host ""
        
        # Fallback to text input
        Write-Host "Current directory: $(Get-Location)" -ForegroundColor Gray
        Write-Host "Default filename: $defaultFileName" -ForegroundColor Gray
        Write-Host ""
        Write-Host "Options:" -ForegroundColor Cyan
        Write-Host "1. Press ENTER to use current directory with default filename" -ForegroundColor White
        Write-Host "2. Type a full path (e.g., C:\Reports\MyReport.xlsx)" -ForegroundColor White
        Write-Host "3. Type just a filename to save in current directory" -ForegroundColor White
        Write-Host ""
        
        $userInput = Read-Host "Enter your choice or file path"
        
        if ([string]::IsNullOrWhiteSpace($userInput)) {
            # Use default location
            $exportPath = Join-Path -Path (Get-Location) -ChildPath $defaultFileName
        }
        elseif ($userInput -match '^[a-zA-Z]:\\' -or $userInput.StartsWith('\\')) {
            # Full path provided
            if (-not $userInput.EndsWith('.xlsx')) {
                $userInput += '.xlsx'
            }
            $exportPath = $userInput
        }
        else {
            # Just filename provided
            if (-not $userInput.EndsWith('.xlsx')) {
                $userInput += '.xlsx'
            }
            $exportPath = Join-Path -Path (Get-Location) -ChildPath $userInput
        }
    }
    
    # Ensure we have a valid path
    if ([string]::IsNullOrWhiteSpace($exportPath)) {
        $exportPath = Join-Path -Path (Get-Location) -ChildPath $defaultFileName
    }
    
    Write-Host "Report will be saved to: $exportPath" -ForegroundColor Cyan
    
    Write-Host ""
    Write-Host "Starting data collection..." -ForegroundColor Cyan
    Write-Host "This may take several minutes depending on your tenant size..." -ForegroundColor Cyan
    Write-Host "Enhanced with Exchange Online delegation details..." -ForegroundColor Cyan
    Write-Host ""
    
    # Collect all data
    $allData = @{
        'User Mailboxes' = Get-UserMailboxes
        'Shared Mailboxes' = Get-SharedMailboxes
        'Distribution Lists' = Get-DistributionLists
        'Security Groups' = Get-SecurityGroups
        'Office365 and Teams Groups' = Get-NormalAndTeamsGroups
    }
    
    # Export to Excel
    if (Export-ToExcel -Data $allData -FilePath $exportPath) {
        Write-Host ""
        Write-Host "============================================" -ForegroundColor Green
        Write-Host "Enhanced Report Generation Completed!" -ForegroundColor Green
        Write-Host "============================================" -ForegroundColor Green
        Write-Host "File saved to: $exportPath" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Report Summary:" -ForegroundColor Yellow
        foreach ($category in $allData.Keys) {
            Write-Host "- $category : $($allData[$category].Count) items" -ForegroundColor White
        }
        
        Write-Host ""
        Write-Host "Enhanced Features Included:" -ForegroundColor Yellow
        Write-Host "✓ Full Access delegation permissions" -ForegroundColor Green
        Write-Host "✓ Send As permissions" -ForegroundColor Green
        Write-Host "✓ Send On Behalf permissions" -ForegroundColor Green
        Write-Host "✓ Detailed group membership with user types" -ForegroundColor Green
        Write-Host "✓ Distribution list member types and details" -ForegroundColor Green
        Write-Host "✓ Teams vs Office 365 group identification" -ForegroundColor Green
        Write-Host "✓ Group owners and member counts" -ForegroundColor Green
        
        # Ask if user wants to open the file
        Write-Host ""
        $openFile = Read-Host "Would you like to open the Excel file now? (Y/N)"
        if ($openFile -match '^[Yy]') {
            Start-Process $exportPath
        }
    }
    else {
        Write-Host "Report generation failed. Please check the errors above." -ForegroundColor Red
    }
}
catch {
    Write-Host "An unexpected error occurred: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}
finally {
    # Disconnect from services
    try {
        Disconnect-MgGraph | Out-Null
        Write-Host "Disconnected from Microsoft Graph" -ForegroundColor Gray
    }
    catch {
        # Ignore disconnection errors
    }
    
    try {
        Disconnect-ExchangeOnline -Confirm:$false | Out-Null
        Write-Host "Disconnected from Exchange Online" -ForegroundColor Gray
    }
    catch {
        # Ignore disconnection errors
    }
    
    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}