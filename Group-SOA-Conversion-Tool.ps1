#Requires -Version 5.1

<#
        .SYNOPSIS
        Group SOA Conversion Tool (Cloud / On-Prem Source of Authority)

        .DESCRIPTION
        GUI tool to manage Group Source of Authority (SOA) conversion between cloud-managed and on-premises-managed
        by toggling the group property `isCloudManaged` via the Microsoft Graph onPremisesSyncBehavior API.
        
        The tool connects to Microsoft Graph, lists directory-synced mail-enabled security groups and distribution groups,
        detects nested group relationships, and supports converting groups in the correct bottom-up order.
        
        This tool is intended to support the approach described in:
        https://learn.microsoft.com/en-us/entra/identity/hybrid/how-to-group-source-of-authority-configure

        .PARAMETER TenantId
        Optional. The Entra ID (Azure AD) tenant ID (GUID) to connect to.
        If not specified, the tool connects to the signed-in user's home tenant.
        Recommended in multi-tenant or partner scenarios to ensure you connect to the correct tenant.

        .EXAMPLE
        C:\PS> .\Group-SOA-Conversion-Tool.ps1

        .EXAMPLE
        C:\PS> .\Group-SOA-Conversion-Tool.ps1 -TenantId "00000000-0000-0000-0000-000000000000"

        .NOTES
        Version: 1.0

        .COPYRIGHT
        MIT License, feel free to distribute and use as you like, please leave author information.

       .LINK
        https://learn.microsoft.com/en-us/entra/identity/hybrid/how-to-group-source-of-authority-configure
        
        BLOG: http://www.hcandersen.net
        Twitter: @dk_hcandersen
        LinkedIn: https://www.linkedin.com/in/hanschrandersen/

        .DISCLAIMER
        This script is provided AS-IS, with no warranty - Use at own risk.
    #>

param(
    [Parameter(Mandatory=$true, HelpMessage="Enter the Entra ID (Azure AD) tenant ID (GUID) to connect to.")]
    [string]$TenantId
)

$script:Version = "1.0"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:LogFile = Join-Path $script:ScriptPath "GroupSOAConversion_$(Get-Date -Format 'yyyyMMdd_HHmm').log"
$script:AllGroups = @()
$script:CurrentPage = 1
$script:PageSize = 100
$script:PermissionOk = $false
$script:NestingMap = @{}
$script:NestingDepth = @{}
$script:TenantId = $TenantId

function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet('INFO','WARNING','ERROR')]
        [string]$Level = 'INFO'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    $retries = 3
    for ($i = 0; $i -lt $retries; $i++) {
        try {
            [System.IO.File]::AppendAllText($script:LogFile, "$logEntry`r`n", [System.Text.Encoding]::UTF8)
            break
        }
        catch {
            if ($i -eq ($retries - 1)) { Write-Host "WARNING: Could not write to log file: $($_.Exception.Message)" }
            Start-Sleep -Milliseconds 50
        }
    }
    
    Write-Host $logEntry
}

function Update-GroupGrid {
    param(
        [Parameter(Mandatory=$false)]
        [array]$Groups = $script:AllGroups
    )
    
    $script:AllGroups = $Groups
    $totalGroups = $script:AllGroups.Count
    $totalPages = [Math]::Ceiling($totalGroups / $script:PageSize)
    
    if ($script:CurrentPage -gt $totalPages -and $totalPages -gt 0) {
        $script:CurrentPage = $totalPages
    }
    
    if ($script:CurrentPage -lt 1) {
        $script:CurrentPage = 1
    }
    
    $startIndex = ($script:CurrentPage - 1) * $script:PageSize
    $endIndex = [Math]::Min($startIndex + $script:PageSize - 1, $totalGroups - 1)
    
    $dataGridView.Rows.Clear()
    
    if ($totalGroups -gt 0) {
        for ($i = $startIndex; $i -le $endIndex; $i++) {
            $group = $script:AllGroups[$i]
            $cloudManagedStatus = if ($group.IsCloudManaged -eq "Unknown") { "Unknown" } elseif ($group.IsCloudManaged -eq $true) { "True" } else { "False" }
            $depth = if ($script:NestingDepth.ContainsKey($group.Id)) { $script:NestingDepth[$group.Id] } else { 0 }
            $dataGridView.Rows.Add($group.DisplayName, $group.Mail, $group.GroupType, $cloudManagedStatus, $depth, $group.Id)
        }
        
        $showingCount = $endIndex - $startIndex + 1
        $pageInfo.Text = "Page $($script:CurrentPage) of $totalPages - Showing $showingCount of $totalGroups groups"
    } else {
        $pageInfo.Text = "No groups to display"
    }
    
    $buttonPrevPage.Enabled = ($script:CurrentPage -gt 1)
    $buttonNextPage.Enabled = ($script:CurrentPage -lt $totalPages)
}

function Search-GraphModule {
    Write-Log "Checking for Microsoft Graph Groups module..."
    
    $module = Get-Module -ListAvailable -Name Microsoft.Graph.Groups
    
    if (-not $module) {
        Write-Log "Microsoft.Graph.Groups module not found. Attempting to install..." -Level WARNING
        
        try {
            [System.Windows.Forms.MessageBox]::Show(
                "Microsoft.Graph.Groups module is not installed.`n`nThe tool will now attempt to install it. This may take a few minutes.",
                "Module Installation Required",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            
            Install-Module -Name Microsoft.Graph.Groups -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Write-Log "Microsoft.Graph.Groups module installed successfully." -Level INFO
            
            [System.Windows.Forms.MessageBox]::Show(
                "Microsoft.Graph.Groups module has been installed successfully.",
                "Installation Complete",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            
            return $true
        }
        catch {
            Write-Log "Failed to install Microsoft.Graph.Groups module: $($_.Exception.Message)" -Level ERROR
            
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to install Microsoft.Graph.Groups module.`n`nError: $($_.Exception.Message)`n`nPlease install manually using:`nInstall-Module -Name Microsoft.Graph.Groups -Scope CurrentUser",
                "Installation Failed",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            
            return $false
        }
    }
    else {
        Write-Log "Microsoft.Graph.Groups module is already installed (Version: $($module.Version))."
        return $true
    }
}

function Test-GraphPermissions {
    Write-Log "Checking Graph permissions..."
    
    try {
        $context = Get-MgContext
        
        if (-not $context) {
            Write-Log "No Graph context found - not connected." -Level WARNING
            $script:PermissionOk = $false
            return $false
        }
        
        $scopes = $context.Scopes
        
        if ($scopes -contains 'Group-OnPremisesSyncBehavior.ReadWrite.All') {
            Write-Log "Required permission 'Group-OnPremisesSyncBehavior.ReadWrite.All' is present."
            $script:PermissionOk = $true
            return $true
        }
        else {
            Write-Log "Required permission 'Group-OnPremisesSyncBehavior.ReadWrite.All' is NOT present." -Level WARNING
            $script:PermissionOk = $false
            return $false
        }
    }
    catch {
        Write-Log "Error checking permissions: $($_.Exception.Message)" -Level ERROR
        $script:PermissionOk = $false
        return $false
    }
}

function Update-PermissionStatus {
    if ($script:PermissionOk) {
        $buttonVerifySetupPermissions.Text = "Graph Permissions OK"
        $buttonVerifySetupPermissions.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#107C10")
        $buttonVerifySetupPermissions.ForeColor = [System.Drawing.Color]::White
    }
    else {
        $buttonVerifySetupPermissions.Text = "Verify/Setup Graph Permissions"
        $buttonVerifySetupPermissions.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#FFB900")
        $buttonVerifySetupPermissions.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
    }
}

function Connect-GraphSession {
    Write-Log "Attempting to connect to Microsoft Graph (TenantId: $($script:TenantId))..."
    
    try {
        Import-Module Microsoft.Graph.Groups -ErrorAction Stop
        
        Connect-MgGraph -Scopes 'Group.ReadWrite.All','Group-OnPremisesSyncBehavior.ReadWrite.All' -TenantId $script:TenantId -ErrorAction Stop -NoWelcome
        
        $context = Get-MgContext
        Write-Log "Successfully connected to Microsoft Graph. TenantId: $($context.TenantId)"
        
        $form.Text = "Group SOA Conversion Tool - Tenant: $($context.TenantId)"
        
        Test-GraphPermissions | Out-Null
        Update-PermissionStatus
        
        return $true
    }
    catch {
        Write-Log "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -Level ERROR
        
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to connect to Microsoft Graph.`n`nError: $($_.Exception.Message)",
            "Connection Failed",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        
        return $false
    }
}

function Setup-GraphPermissions {
    Write-Log "Requesting consent for Group-OnPremisesSyncBehavior.ReadWrite.All..."
    
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        
        Connect-MgGraph -Scopes 'Group.ReadWrite.All','Group-OnPremisesSyncBehavior.ReadWrite.All' -TenantId $script:TenantId -ErrorAction Stop -NoWelcome
        
        Write-Log "Consent flow completed."
        
        if (Test-GraphPermissions) {
            Write-Log "Permission 'Group-OnPremisesSyncBehavior.ReadWrite.All' successfully granted."
            Update-PermissionStatus
            
            [System.Windows.Forms.MessageBox]::Show(
                "Permission 'Group-OnPremisesSyncBehavior.ReadWrite.All' has been successfully granted.",
                "Permission Granted",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            
            return $true
        }
        else {
            Write-Log "Permission was not granted after consent flow." -Level WARNING
            Update-PermissionStatus
            
            [System.Windows.Forms.MessageBox]::Show(
                "The permission was not granted. This may require admin consent.`n`nTo grant manually:`n1. Go to Entra admin center`n2. Navigate to Enterprise Applications`n3. Find 'Microsoft Graph Command Line Tools'`n4. Go to Permissions`n5. Click 'Grant admin consent'`n`nAlternatively, ask your Application Administrator or Cloud Application Administrator to grant consent.",
                "Permission Not Granted",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            
            return $false
        }
    }
    catch {
        Write-Log "Error during permission setup: $($_.Exception.Message)" -Level ERROR
        
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to set up permissions.`n`nError: $($_.Exception.Message)`n`nTo grant manually:`n1. Go to Entra admin center`n2. Navigate to Enterprise Applications`n3. Find 'Microsoft Graph Command Line Tools'`n4. Go to Permissions`n5. Click 'Grant admin consent'",
            "Permission Setup Failed",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        
        return $false
    }
}

function Get-ExchangeGroups {
    Write-Log "Retrieving Exchange-relevant groups from Microsoft Graph..."
    
    try {
        $allSyncedGroups = Get-MgGroup -All -Filter "onPremisesSyncEnabled eq true" -Property Id,DisplayName,Mail,MailEnabled,SecurityEnabled,GroupTypes,OnPremisesSyncEnabled -ErrorAction Stop
        
        Write-Log "Retrieved $($allSyncedGroups.Count) total on-premises synced groups."
        
        $exchangeGroups = $allSyncedGroups | Where-Object {
            $_.MailEnabled -eq $true -and
            ($_.GroupTypes -eq $null -or $_.GroupTypes.Count -eq 0 -or $_.GroupTypes -notcontains "Unified")
        }
        
        Write-Log "Filtered to $($exchangeGroups.Count) Exchange-relevant groups (Mail-Enabled Security Groups and Distribution Groups)."
        
        $groupResults = @()
        $totalCount = $exchangeGroups.Count
        $currentIndex = 0
        
        foreach ($group in $exchangeGroups) {
            $currentIndex++
            
            if ($currentIndex % 10 -eq 0 -or $currentIndex -eq $totalCount) {
                $statusLabel.Text = "Loading group $currentIndex of $totalCount..."
                [System.Windows.Forms.Application]::DoEvents()
            }
            
            $groupType = if ($group.SecurityEnabled -eq $true) {
                "Mail-Enabled Security Group"
            } else {
                "Distribution Group"
            }
            
            $isCloudManaged = $null
            $soaRetries = 3
            for ($r = 0; $r -lt $soaRetries; $r++) {
                try {
                    $uri = "https://graph.microsoft.com/v1.0/groups/$($group.Id)/onPremisesSyncBehavior?`$select=isCloudManaged"
                    $soaResponse = Invoke-MgGraphRequest -Uri $uri -Method GET -ErrorAction Stop
                    $isCloudManaged = if ($soaResponse.isCloudManaged -eq $true) { $true } else { $false }
                    break
                }
                catch {
                    if ($r -lt ($soaRetries - 1)) {
                        Start-Sleep -Milliseconds 500
                    }
                    else {
                        Write-Log "Could not retrieve SOA status for group '$($group.DisplayName)' ($($group.Id)) after $soaRetries attempts: $($_.Exception.Message)" -Level WARNING
                        $isCloudManaged = "Unknown"
                    }
                }
            }
            
            $groupResults += [PSCustomObject]@{
                Id               = $group.Id
                DisplayName      = $group.DisplayName
                Mail             = $group.Mail
                GroupType        = $groupType
                IsCloudManaged   = $isCloudManaged
                SecurityEnabled  = $group.SecurityEnabled
                MailEnabled      = $group.MailEnabled
            }
        }
        
        Write-Log "Finished loading SOA status for all groups."
        
        return $groupResults
    }
    catch {
        Write-Log "Failed to retrieve groups: $($_.Exception.Message)" -Level ERROR
        
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to retrieve groups.`n`nError: $($_.Exception.Message)",
            "Retrieval Failed",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        
        return $null
    }
}

function Build-NestingMap {
    param(
        [Parameter(Mandatory=$true)]
        [array]$Groups
    )
    
    Write-Log "Building nested group map..."
    
    $script:NestingMap = @{}
    $script:NestingDepth = @{}
    
    $groupIds = @{}
    foreach ($group in $Groups) {
        $groupIds[$group.Id] = $group
        $script:NestingMap[$group.Id] = @()
    }
    
    $totalCount = $Groups.Count
    $currentIndex = 0
    
    foreach ($group in $Groups) {
        $currentIndex++
        
        if ($currentIndex % 10 -eq 0 -or $currentIndex -eq $totalCount) {
            $statusLabel.Text = "Analyzing nesting $currentIndex of $totalCount..."
            [System.Windows.Forms.Application]::DoEvents()
        }
        
        try {
            $members = Get-MgGroupMember -GroupId $group.Id -All -Property Id -ErrorAction Stop
            
            foreach ($member in $members) {
                if ($groupIds.ContainsKey($member.Id)) {
                    $script:NestingMap[$group.Id] += $member.Id
                }
            }
        }
        catch {
            Write-Log "Could not retrieve members for group '$($group.DisplayName)' ($($group.Id)): $($_.Exception.Message)" -Level WARNING
        }
    }
    
    $childToParents = @{}
    foreach ($parentId in $script:NestingMap.Keys) {
        foreach ($childId in $script:NestingMap[$parentId]) {
            if (-not $childToParents.ContainsKey($childId)) {
                $childToParents[$childId] = @()
            }
            $childToParents[$childId] += $parentId
        }
    }
    
    foreach ($group in $Groups) {
        $script:NestingDepth[$group.Id] = 0
    }
    
    $changed = $true
    $maxIterations = 100
    $iteration = 0
    
    while ($changed -and $iteration -lt $maxIterations) {
        $changed = $false
        $iteration++
        
        foreach ($parentId in $script:NestingMap.Keys) {
            foreach ($childId in $script:NestingMap[$parentId]) {
                $expectedDepth = $script:NestingDepth[$childId] + 1
                if ($expectedDepth -gt $script:NestingDepth[$parentId]) {
                    $script:NestingDepth[$parentId] = $expectedDepth
                    $changed = $true
                }
            }
        }
    }
    
    $nestedCount = ($script:NestingMap.Values | Where-Object { $_.Count -gt 0 }).Count
    $maxDepth = ($script:NestingDepth.Values | Measure-Object -Maximum).Maximum
    Write-Log "Nesting analysis complete. Found $nestedCount groups with nested children. Max depth: $maxDepth."
}

function Get-UnconvertedChildren {
    param(
        [Parameter(Mandatory=$true)]
        [string]$GroupId
    )
    
    $unconverted = @()
    
    if ($script:NestingMap.ContainsKey($GroupId)) {
        foreach ($childId in $script:NestingMap[$GroupId]) {
            $childGroup = $script:AllGroups | Where-Object { $_.Id -eq $childId }
            if ($childGroup -and -not $childGroup.IsCloudManaged) {
                $unconverted += $childGroup
            }
        }
    }
    
    return $unconverted
}

function Convert-GroupToCloudManaged {
    param(
        [Parameter(Mandatory=$true)]
        $SelectedGroup
    )
    
    $groupId = $SelectedGroup.Id
    $displayName = $SelectedGroup.DisplayName
    
    Write-Log "Converting group '$displayName' ($groupId) to Cloud Managed..."
    
    try {
        $body = @{ isCloudManaged = $true } | ConvertTo-Json
        $uri = "https://graph.microsoft.com/v1.0/groups/$groupId/onPremisesSyncBehavior"
        Invoke-MgGraphRequest -Uri $uri -Method PATCH -Body $body -ContentType "application/json" -ErrorAction Stop
        
        Write-Log "Successfully converted group '$displayName' ($groupId) to Cloud Managed." -Level INFO
        
        return $true
    }
    catch {
        Write-Log "Failed to convert group '$displayName' ($groupId) to Cloud Managed. Error: $($_.Exception.Message)" -Level ERROR
        
        return $false
    }
}

function Convert-GroupToOnPremManaged {
    param(
        [Parameter(Mandatory=$true)]
        $SelectedGroup
    )
    
    $groupId = $SelectedGroup.Id
    $displayName = $SelectedGroup.DisplayName
    
    Write-Log "Rolling back group '$displayName' ($groupId) to On-Premises Managed..."
    
    try {
        $body = @{ isCloudManaged = $false } | ConvertTo-Json
        $uri = "https://graph.microsoft.com/v1.0/groups/$groupId/onPremisesSyncBehavior"
        Invoke-MgGraphRequest -Uri $uri -Method PATCH -Body $body -ContentType "application/json" -ErrorAction Stop
        
        Write-Log "Successfully rolled back group '$displayName' ($groupId) to On-Premises Managed." -Level INFO
        
        return $true
    }
    catch {
        Write-Log "Failed to roll back group '$displayName' ($groupId) to On-Premises Managed. Error: $($_.Exception.Message)" -Level ERROR
        
        return $false
    }
}

# ============================================================
# MAIN SCRIPT - GUI SETUP
# ============================================================

Write-Log "========================================" -Level INFO
Write-Log "Group SOA Conversion Tool Started" -Level INFO
Write-Log "Target TenantId: $($script:TenantId)" -Level INFO
Write-Log "========================================" -Level INFO

if (-not (Search-GraphModule)) {
    Write-Log "Cannot proceed without Microsoft Graph Groups module. Exiting." -Level ERROR
    exit 1
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Group SOA Conversion Tool"
$form.Size = New-Object System.Drawing.Size(1100, 750)
$form.MinimumSize = New-Object System.Drawing.Size(1100, 750)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$form.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#F3F3F3")
$form.ShowIcon = $false

# --- Header Panel ---
$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Location = New-Object System.Drawing.Point(0, 0)
$headerPanel.Size = New-Object System.Drawing.Size(1100, 90)
$headerPanel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#E8E8E8")
$headerPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($headerPanel)

$labelTitle = New-Object System.Windows.Forms.Label
$labelTitle.Location = New-Object System.Drawing.Point(20, 15)
$labelTitle.Size = New-Object System.Drawing.Size(900, 40)
$labelTitle.Text = "Group SOA Conversion Tool"
$labelTitle.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Regular)
$labelTitle.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$labelTitle.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
$labelTitle.BackColor = [System.Drawing.Color]::Transparent
$labelTitle.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$headerPanel.Controls.Add($labelTitle)

$pictureBoxLogo = New-Object System.Windows.Forms.PictureBox
$pictureBoxLogo.Location = New-Object System.Drawing.Point(975, 8)
$pictureBoxLogo.Size = New-Object System.Drawing.Size(100, 75)
$pictureBoxLogo.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$pictureBoxLogo.BackColor = [System.Drawing.Color]::Transparent
$pictureBoxLogo.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$logoPath = Join-Path $script:ScriptPath "logo.png"
if (Test-Path $logoPath) {
    try {
        $fileStream = [System.IO.File]::OpenRead($logoPath)
        $memoryStream = New-Object System.IO.MemoryStream
        $fileStream.CopyTo($memoryStream)
        $fileStream.Close()
        $fileStream.Dispose()
        $memoryStream.Position = 0
        $pictureBoxLogo.Image = [System.Drawing.Image]::FromStream($memoryStream)
    } catch {
        Write-Log "Failed to load logo image: $($_.Exception.Message)" -Level WARNING
    }
}
$headerPanel.Controls.Add($pictureBoxLogo)

$labelDescription = New-Object System.Windows.Forms.Label
$labelDescription.Location = New-Object System.Drawing.Point(20, 58)
$labelDescription.Size = New-Object System.Drawing.Size(900, 25)
$labelDescription.Text = "Convert Group Source of Authority for Exchange on-premises groups (Mail-Enabled Security Groups & Distribution Groups)"
$labelDescription.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$labelDescription.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#605E5C")
$labelDescription.BackColor = [System.Drawing.Color]::Transparent
$labelDescription.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$labelDescription.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$headerPanel.Controls.Add($labelDescription)

# --- Connection Buttons Row ---
$buttonConnect = New-Object System.Windows.Forms.Button
$buttonConnect.Location = New-Object System.Drawing.Point(20, 105)
$buttonConnect.Size = New-Object System.Drawing.Size(170, 40)
$buttonConnect.Text = "Connect to Graph"
$buttonConnect.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonConnect.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#0078D4")
$buttonConnect.ForeColor = [System.Drawing.Color]::White
$buttonConnect.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$buttonConnect.FlatAppearance.BorderSize = 0
$buttonConnect.Cursor = [System.Windows.Forms.Cursors]::Hand
$buttonConnect.Add_Click({
    $buttonConnect.Enabled = $false
    $buttonRefresh.Enabled = $false
    
    if (Connect-GraphSession) {
        $buttonConnect.Text = "Connected"
        $buttonConnect.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#107C10")
        $buttonConnect.Enabled = $false
        $buttonRefresh.Enabled = $true
        $buttonDisconnect.Enabled = $true
        
        $groups = Get-ExchangeGroups
        
        if ($groups) {
            Build-NestingMap -Groups $groups
            $script:CurrentPage = 1
            Update-GroupGrid -Groups $groups
            $statusLabel.Text = "Connected - $($groups.Count) groups loaded"
        }
    }
    else {
        $buttonConnect.Enabled = $true
    }
})
$form.Controls.Add($buttonConnect)

$buttonRefresh = New-Object System.Windows.Forms.Button
$buttonRefresh.Location = New-Object System.Drawing.Point(205, 105)
$buttonRefresh.Size = New-Object System.Drawing.Size(150, 40)
$buttonRefresh.Text = "Refresh Groups"
$buttonRefresh.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonRefresh.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#E1E1E1")
$buttonRefresh.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
$buttonRefresh.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$buttonRefresh.FlatAppearance.BorderSize = 0
$buttonRefresh.Cursor = [System.Windows.Forms.Cursors]::Hand
$buttonRefresh.Enabled = $false
$buttonRefresh.Add_Click({
    $buttonRefresh.Enabled = $false
    
    $groups = Get-ExchangeGroups
    
    if ($groups) {
        Build-NestingMap -Groups $groups
        $script:CurrentPage = 1
        Update-GroupGrid -Groups $groups
        $statusLabel.Text = "Refreshed - $($groups.Count) groups loaded"
    }
    
    $buttonRefresh.Enabled = $true
})
$form.Controls.Add($buttonRefresh)

$buttonDisconnect = New-Object System.Windows.Forms.Button
$buttonDisconnect.Location = New-Object System.Drawing.Point(370, 105)
$buttonDisconnect.Size = New-Object System.Drawing.Size(180, 40)
$buttonDisconnect.Text = "Disconnect from Graph"
$buttonDisconnect.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonDisconnect.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#E1E1E1")
$buttonDisconnect.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
$buttonDisconnect.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$buttonDisconnect.FlatAppearance.BorderSize = 0
$buttonDisconnect.Cursor = [System.Windows.Forms.Cursors]::Hand
$buttonDisconnect.Enabled = $false
$buttonDisconnect.Add_Click({
    Write-Log "Disconnecting from Microsoft Graph..."
    
    try {
        Disconnect-MgGraph -ErrorAction Stop
        Write-Log "Successfully disconnected from Microsoft Graph."
        
        $form.Text = "Group SOA Conversion Tool"
        $buttonConnect.Text = "Connect to Graph"
        $buttonConnect.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#0078D4")
        $buttonConnect.Enabled = $true
        $buttonRefresh.Enabled = $false
        $buttonDisconnect.Enabled = $false
        
        $script:PermissionOk = $false
        Update-PermissionStatus
        
        $dataGridView.Rows.Clear()
        
        $statusLabel.Text = "Disconnected"
        
        [System.Windows.Forms.MessageBox]::Show(
            "Successfully disconnected from Microsoft Graph.",
            "Disconnected",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
    catch {
        Write-Log "Error during disconnect: $($_.Exception.Message)" -Level WARNING
        
        [System.Windows.Forms.MessageBox]::Show(
            "An error occurred while disconnecting.`n`nError: $($_.Exception.Message)",
            "Disconnect Warning",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
    }
})
$form.Controls.Add($buttonDisconnect)

# --- Verify/Setup Graph Permissions Button ---
$buttonVerifySetupPermissions = New-Object System.Windows.Forms.Button
$buttonVerifySetupPermissions.Location = New-Object System.Drawing.Point(20, 150)
$buttonVerifySetupPermissions.Size = New-Object System.Drawing.Size(280, 35)
$buttonVerifySetupPermissions.Text = "Verify/Setup Graph Permissions"
$buttonVerifySetupPermissions.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonVerifySetupPermissions.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#FFB900")
$buttonVerifySetupPermissions.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
$buttonVerifySetupPermissions.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$buttonVerifySetupPermissions.FlatAppearance.BorderSize = 0
$buttonVerifySetupPermissions.Cursor = [System.Windows.Forms.Cursors]::Hand
$buttonVerifySetupPermissions.Add_Click({
    $buttonVerifySetupPermissions.Enabled = $false
    $statusLabel.Text = "Verifying Graph permissions..."
    [System.Windows.Forms.Application]::DoEvents()
    
    Write-Log "Verifying Graph permissions..."
    
    try {
        Import-Module Microsoft.Graph.Groups -ErrorAction Stop
        
        # Connect if not already connected
        $context = Get-MgContext
        if (-not $context) {
            Connect-MgGraph -Scopes 'Group.ReadWrite.All','Group-OnPremisesSyncBehavior.ReadWrite.All' -TenantId $script:TenantId -ErrorAction Stop -NoWelcome
            $context = Get-MgContext
            Write-Log "Connected to Microsoft Graph for permission verification. TenantId: $($context.TenantId)"
            $form.Text = "Group SOA Conversion Tool - Tenant: $($context.TenantId)"
        }
        
        # Check if permission is already consented
        Test-GraphPermissions | Out-Null
        
        if ($script:PermissionOk) {
            Write-Log "Permission 'Group-OnPremisesSyncBehavior.ReadWrite.All' is already consented."
            Update-PermissionStatus
            $statusLabel.Text = "Graph permissions verified - OK"
        }
        else {
            # Permission missing - trigger consent flow
            Write-Log "Permission missing. Triggering consent flow..."
            $statusLabel.Text = "Permission missing - requesting consent..."
            [System.Windows.Forms.Application]::DoEvents()
            
            Setup-GraphPermissions
            
            if ($script:PermissionOk) {
                Update-PermissionStatus
                $statusLabel.Text = "Graph permissions granted - OK"
            }
            else {
                Update-PermissionStatus
                $statusLabel.Text = "Graph permission consent was not granted"
            }
        }
        
        # Update connection buttons if we have an active session
        $context = Get-MgContext
        if ($context) {
            $buttonConnect.Text = "Connected"
            $buttonConnect.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#107C10")
            $buttonConnect.Enabled = $false
            $buttonRefresh.Enabled = $true
            $buttonDisconnect.Enabled = $true
        }
    }
    catch {
        Write-Log "Error during permission verify/setup: $($_.Exception.Message)" -Level ERROR
        $statusLabel.Text = "Permission verification failed"
        
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to verify/setup permissions.`n`nError: $($_.Exception.Message)",
            "Verification Failed",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
    
    $buttonVerifySetupPermissions.Enabled = $true
})
$form.Controls.Add($buttonVerifySetupPermissions)

# --- DataGridView ---
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(20, 195)
$dataGridView.Size = New-Object System.Drawing.Size(1060, 370)
$dataGridView.AllowUserToAddRows = $false
$dataGridView.AllowUserToDeleteRows = $false
$dataGridView.ReadOnly = $true
$dataGridView.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dataGridView.MultiSelect = $true
$dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$dataGridView.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$dataGridView.BackgroundColor = [System.Drawing.Color]::White
$dataGridView.GridColor = [System.Drawing.ColorTranslator]::FromHtml("#E1E1E1")
$dataGridView.DefaultCellStyle.BackColor = [System.Drawing.Color]::White
$dataGridView.DefaultCellStyle.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
$dataGridView.DefaultCellStyle.SelectionBackColor = [System.Drawing.ColorTranslator]::FromHtml("#0078D4")
$dataGridView.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
$dataGridView.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$dataGridView.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#F9F9F9")
$dataGridView.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#E1E1E1")
$dataGridView.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
$dataGridView.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$dataGridView.ColumnHeadersDefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleLeft
$dataGridView.ColumnHeadersDefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(5, 0, 0, 0)
$dataGridView.ColumnHeadersHeight = 40
$dataGridView.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
$dataGridView.EnableHeadersVisualStyles = $false
$dataGridView.RowHeadersVisible = $false
$dataGridView.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
$dataGridView.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

$colDisplayName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colDisplayName.Name = "DisplayName"
$colDisplayName.HeaderText = "Display Name"
$colDisplayName.FillWeight = 25
[void]$dataGridView.Columns.Add($colDisplayName)

$colEmail = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colEmail.Name = "Email"
$colEmail.HeaderText = "Email Address"
$colEmail.FillWeight = 25
[void]$dataGridView.Columns.Add($colEmail)

$colGroupType = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colGroupType.Name = "GroupType"
$colGroupType.HeaderText = "Group Type"
$colGroupType.FillWeight = 20
[void]$dataGridView.Columns.Add($colGroupType)

$colCloudManaged = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colCloudManaged.Name = "IsCloudManaged"
$colCloudManaged.HeaderText = "Cloud Managed"
$colCloudManaged.FillWeight = 12
[void]$dataGridView.Columns.Add($colCloudManaged)

$colNestingDepth = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colNestingDepth.Name = "NestingDepth"
$colNestingDepth.HeaderText = "Nesting Depth"
$colNestingDepth.FillWeight = 12
[void]$dataGridView.Columns.Add($colNestingDepth)

$colObjectId = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colObjectId.Name = "ObjectId"
$colObjectId.HeaderText = "Object ID"
$colObjectId.Visible = $false
[void]$dataGridView.Columns.Add($colObjectId)

$form.Controls.Add($dataGridView)

# --- Pagination ---
$buttonPrevPage = New-Object System.Windows.Forms.Button
$buttonPrevPage.Location = New-Object System.Drawing.Point(20, 570)
$buttonPrevPage.Size = New-Object System.Drawing.Size(100, 30)
$buttonPrevPage.Text = "< Previous"
$buttonPrevPage.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonPrevPage.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#E1E1E1")
$buttonPrevPage.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
$buttonPrevPage.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$buttonPrevPage.FlatAppearance.BorderSize = 0
$buttonPrevPage.Cursor = [System.Windows.Forms.Cursors]::Hand
$buttonPrevPage.Enabled = $false
$buttonPrevPage.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$buttonPrevPage.Add_Click({
    if ($script:CurrentPage -gt 1) {
        $script:CurrentPage--
        Update-GroupGrid
    }
})
$form.Controls.Add($buttonPrevPage)

$pageInfo = New-Object System.Windows.Forms.Label
$pageInfo.Location = New-Object System.Drawing.Point(130, 570)
$pageInfo.Size = New-Object System.Drawing.Size(650, 30)
$pageInfo.Text = "No groups to display"
$pageInfo.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$pageInfo.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#605E5C")
$pageInfo.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$pageInfo.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($pageInfo)

$buttonNextPage = New-Object System.Windows.Forms.Button
$buttonNextPage.Location = New-Object System.Drawing.Point(790, 570)
$buttonNextPage.Size = New-Object System.Drawing.Size(100, 30)
$buttonNextPage.Text = "Next >"
$buttonNextPage.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonNextPage.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#E1E1E1")
$buttonNextPage.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
$buttonNextPage.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$buttonNextPage.FlatAppearance.BorderSize = 0
$buttonNextPage.Cursor = [System.Windows.Forms.Cursors]::Hand
$buttonNextPage.Enabled = $false
$buttonNextPage.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$buttonNextPage.Add_Click({
    $totalPages = [Math]::Ceiling($script:AllGroups.Count / $script:PageSize)
    if ($script:CurrentPage -lt $totalPages) {
        $script:CurrentPage++
        Update-GroupGrid
    }
})
$form.Controls.Add($buttonNextPage)

# --- Action Buttons ---
$buttonConvertToCloud = New-Object System.Windows.Forms.Button
$buttonConvertToCloud.Location = New-Object System.Drawing.Point(20, 610)
$buttonConvertToCloud.Size = New-Object System.Drawing.Size(230, 45)
$buttonConvertToCloud.Text = "Convert to Cloud Managed"
$buttonConvertToCloud.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonConvertToCloud.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#E1E1E1")
$buttonConvertToCloud.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
$buttonConvertToCloud.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$buttonConvertToCloud.FlatAppearance.BorderSize = 0
$buttonConvertToCloud.Cursor = [System.Windows.Forms.Cursors]::Hand
$buttonConvertToCloud.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$buttonConvertToCloud.Add_Click({
    if (-not $script:PermissionOk) {
        [System.Windows.Forms.MessageBox]::Show(
            "The required permission 'Group-OnPremisesSyncBehavior.ReadWrite.All' is not consented.`n`nPlease click 'Setup Graph Permissions' first to grant the required permission.",
            "Permission Required",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    if ($dataGridView.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select at least one group from the list.",
            "No Group Selected",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    $selectedGroups = @()
    foreach ($selectedRow in $dataGridView.SelectedRows) {
        $groupId = $selectedRow.Cells["ObjectId"].Value
        $group = $script:AllGroups | Where-Object { $_.Id -eq $groupId }
        if ($group) {
            $selectedGroups += $group
        }
    }
    
    $allWarnings = @()
    foreach ($group in $selectedGroups) {
        $unconvertedChildren = Get-UnconvertedChildren -GroupId $group.Id
        $childrenNotInSelection = $unconvertedChildren | Where-Object { $_.Id -notin ($selectedGroups | ForEach-Object { $_.Id }) }
        
        if ($childrenNotInSelection.Count -gt 0) {
            $childNames = ($childrenNotInSelection | ForEach-Object { $_.DisplayName }) -join ", "
            $allWarnings += "Group '$($group.DisplayName)' has unconverted nested children not in your selection: $childNames"
        }
    }
    
    if ($allWarnings.Count -gt 0) {
        $warningText = "WARNING: Nested group ordering issue detected!`n`n"
        $warningText += ($allWarnings -join "`n`n")
        $warningText += "`n`nMicrosoft recommends converting nested (child) groups before their parent groups.`n`nDo you want to continue anyway?"
        
        $warningResult = [System.Windows.Forms.MessageBox]::Show(
            $warningText,
            "Nested Group Warning",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        
        if ($warningResult -ne [System.Windows.Forms.DialogResult]::Yes) {
            return
        }
    }
    
    $sortedGroups = $selectedGroups | Sort-Object { 
        if ($script:NestingDepth.ContainsKey($_.Id)) { $script:NestingDepth[$_.Id] } else { 0 }
    }
    
    $selectedCount = $sortedGroups.Count
    $groupList = ($sortedGroups | ForEach-Object { "$($_.DisplayName) (Depth: $(if ($script:NestingDepth.ContainsKey($_.Id)) { $script:NestingDepth[$_.Id] } else { 0 }))" }) -join "`n"
    
    $confirmMessage = if ($selectedCount -eq 1) {
        "Are you sure you want to convert group '$($sortedGroups[0].DisplayName)' to Cloud Managed?"
    } else {
        "Are you sure you want to convert $selectedCount groups to Cloud Managed?`n`nGroups will be converted in this order (bottom-up):`n$groupList"
    }
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        $confirmMessage,
        "Confirm Conversion",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $successCount = 0
        $failCount = 0
        
        foreach ($group in $sortedGroups) {
            $statusLabel.Text = "Converting '$($group.DisplayName)'..."
            [System.Windows.Forms.Application]::DoEvents()
            
            if (Convert-GroupToCloudManaged -SelectedGroup $group) {
                $group.IsCloudManaged = $true
                
                foreach ($row in $dataGridView.Rows) {
                    if ($row.Cells["ObjectId"].Value -eq $group.Id) {
                        $row.Cells["IsCloudManaged"].Value = "True"
                        break
                    }
                }
                
                $successCount++
            } else {
                $failCount++
            }
        }
        
        $summaryMessage = "Batch conversion completed.`n`nSuccessful: $successCount`nFailed: $failCount"
        
        [System.Windows.Forms.MessageBox]::Show(
            $summaryMessage,
            "Batch Conversion Summary",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        
        $statusLabel.Text = "Conversion complete - Success: $successCount, Failed: $failCount"
        Write-Log "Batch conversion to Cloud Managed completed. Success: $successCount, Failed: $failCount"
    }
})
$form.Controls.Add($buttonConvertToCloud)

$buttonConvertToOnPrem = New-Object System.Windows.Forms.Button
$buttonConvertToOnPrem.Location = New-Object System.Drawing.Point(265, 610)
$buttonConvertToOnPrem.Size = New-Object System.Drawing.Size(230, 45)
$buttonConvertToOnPrem.Text = "Roll Back to On-Prem"
$buttonConvertToOnPrem.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonConvertToOnPrem.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#E1E1E1")
$buttonConvertToOnPrem.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
$buttonConvertToOnPrem.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$buttonConvertToOnPrem.FlatAppearance.BorderSize = 0
$buttonConvertToOnPrem.Cursor = [System.Windows.Forms.Cursors]::Hand
$buttonConvertToOnPrem.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$buttonConvertToOnPrem.Add_Click({
    if (-not $script:PermissionOk) {
        [System.Windows.Forms.MessageBox]::Show(
            "The required permission 'Group-OnPremisesSyncBehavior.ReadWrite.All' is not consented.`n`nPlease click 'Setup Graph Permissions' first to grant the required permission.",
            "Permission Required",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    if ($dataGridView.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select at least one group from the list.",
            "No Group Selected",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    $selectedGroups = @()
    foreach ($selectedRow in $dataGridView.SelectedRows) {
        $groupId = $selectedRow.Cells["ObjectId"].Value
        $group = $script:AllGroups | Where-Object { $_.Id -eq $groupId }
        if ($group) {
            $selectedGroups += $group
        }
    }
    
    $sortedGroups = $selectedGroups | Sort-Object { 
        if ($script:NestingDepth.ContainsKey($_.Id)) { $script:NestingDepth[$_.Id] } else { 0 }
    } -Descending
    
    $selectedCount = $sortedGroups.Count
    $groupList = ($sortedGroups | ForEach-Object { "$($_.DisplayName) (Depth: $(if ($script:NestingDepth.ContainsKey($_.Id)) { $script:NestingDepth[$_.Id] } else { 0 }))" }) -join "`n"
    
    $confirmMessage = if ($selectedCount -eq 1) {
        "Are you sure you want to roll back group '$($sortedGroups[0].DisplayName)' to On-Premises Managed?`n`nIMPORTANT: Make sure to remove cloud users from the group and remove the group from access packages before rolling back."
    } else {
        "Are you sure you want to roll back $selectedCount groups to On-Premises Managed?`n`nGroups will be rolled back in this order (top-down):`n$groupList`n`nIMPORTANT: Make sure to remove cloud users from the groups and remove the groups from access packages before rolling back."
    }
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        $confirmMessage,
        "Confirm Rollback",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $successCount = 0
        $failCount = 0
        
        foreach ($group in $sortedGroups) {
            $statusLabel.Text = "Rolling back '$($group.DisplayName)'..."
            [System.Windows.Forms.Application]::DoEvents()
            
            if (Convert-GroupToOnPremManaged -SelectedGroup $group) {
                $group.IsCloudManaged = $false
                
                foreach ($row in $dataGridView.Rows) {
                    if ($row.Cells["ObjectId"].Value -eq $group.Id) {
                        $row.Cells["IsCloudManaged"].Value = "False"
                        break
                    }
                }
                
                $successCount++
            } else {
                $failCount++
            }
        }
        
        $summaryMessage = "Batch rollback completed.`n`nSuccessful: $successCount`nFailed: $failCount`n`nNote: The rollback is only complete after the next scheduled or forced run of Connect Sync."
        
        [System.Windows.Forms.MessageBox]::Show(
            $summaryMessage,
            "Batch Rollback Summary",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        
        $statusLabel.Text = "Rollback complete - Success: $successCount, Failed: $failCount"
        Write-Log "Batch rollback to On-Premises Managed completed. Success: $successCount, Failed: $failCount"
    }
})
$form.Controls.Add($buttonConvertToOnPrem)

$buttonOpenLog = New-Object System.Windows.Forms.Button
$buttonOpenLog.Location = New-Object System.Drawing.Point(510, 610)
$buttonOpenLog.Size = New-Object System.Drawing.Size(160, 45)
$buttonOpenLog.Text = "Open Log File"
$buttonOpenLog.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonOpenLog.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#E1E1E1")
$buttonOpenLog.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
$buttonOpenLog.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$buttonOpenLog.FlatAppearance.BorderSize = 0
$buttonOpenLog.Cursor = [System.Windows.Forms.Cursors]::Hand
$buttonOpenLog.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$buttonOpenLog.Add_Click({
    if (Test-Path $script:LogFile) {
        try {
            Start-Process notepad.exe -ArgumentList $script:LogFile
            Write-Log "Log file opened: $script:LogFile"
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to open log file.`n`nError: $($_.Exception.Message)",
                "Error Opening Log",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    }
    else {
        [System.Windows.Forms.MessageBox]::Show(
            "Log file does not exist yet.`n`nPath: $script:LogFile",
            "Log File Not Found",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
})
$form.Controls.Add($buttonOpenLog)

# --- Status and Version Labels ---
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(685, 610)
$statusLabel.Size = New-Object System.Drawing.Size(395, 20)
$statusLabel.Text = "Not connected"
$statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$statusLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#605E5C")
$statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$statusLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($statusLabel)

$versionLabel = New-Object System.Windows.Forms.Label
$versionLabel.Location = New-Object System.Drawing.Point(685, 635)
$versionLabel.Size = New-Object System.Drawing.Size(395, 20)
$versionLabel.Text = "Version $script:Version"
$versionLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$versionLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#808080")
$versionLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$versionLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($versionLabel)

# --- Resize Handler ---
$form.Add_Resize({
    $buttonY = $form.ClientSize.Height - 105
    $buttonConvertToCloud.Location = New-Object System.Drawing.Point(20, $buttonY)
    $buttonConvertToOnPrem.Location = New-Object System.Drawing.Point(265, $buttonY)
    $buttonOpenLog.Location = New-Object System.Drawing.Point(510, $buttonY)
    
    $statusY = $form.ClientSize.Height - 105
    $statusLabel.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 415), $statusY)
    
    $versionY = $form.ClientSize.Height - 80
    $versionLabel.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 415), $versionY)
    
    $paginationY = $buttonY - 40
    $buttonPrevPage.Location = New-Object System.Drawing.Point(20, $paginationY)
    $pageInfo.Location = New-Object System.Drawing.Point(130, $paginationY)
    $pageInfo.Size = New-Object System.Drawing.Size(($form.ClientSize.Width - 270), 30)
    $buttonNextPage.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 120), $paginationY)
    
    $gridHeight = $paginationY - 205
    $dataGridView.Size = New-Object System.Drawing.Size(($form.ClientSize.Width - 40), $gridHeight)
})

# --- Form Closing Handler ---
$form.Add_FormClosing({
    Write-Log "Group SOA Conversion Tool Closing" -Level INFO
    
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Write-Log "Disconnected from Microsoft Graph." -Level INFO
    }
    catch {
        Write-Log "Error during disconnect: $($_.Exception.Message)" -Level WARNING
    }
    
    if ($pictureBoxLogo.Image) {
        $pictureBoxLogo.Image.Dispose()
    }
})

[void]$form.ShowDialog()

Write-Log "========================================" -Level INFO
Write-Log "Group SOA Conversion Tool Ended" -Level INFO
Write-Log "========================================" -Level INFO
