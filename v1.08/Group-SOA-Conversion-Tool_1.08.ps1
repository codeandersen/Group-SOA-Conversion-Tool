#Requires -Version 5.1

<#
        .SYNOPSIS
        Group SOA Conversion Tool (Cloud / On-Prem Source of Authority)

        .DESCRIPTION
        GUI tool to manage Group Source of Authority (SOA) conversion between cloud-managed and on-premises-managed
        by toggling the group property `isCloudManaged` via the Microsoft Graph onPremisesSyncBehavior API.

        The tool connects to Microsoft Graph, lists the on-premises synced groups (Security Groups, Mail-Enabled Security
        Groups and Distribution Groups) together with the groups whose SOA has already been converted to the cloud,
        detects nested group relationships, and supports converting groups in the correct bottom-up order.
        It is designed for large tenants (15K+ synced groups) and uses Microsoft Graph JSON batching with throttling handling.

        Features:
        - Automatic permission consent flow for required Graph API permissions
        - Display on-premises synced and already converted groups with their current SOA status (Cloud Managed: True/False)
        - Search scope by group type, free-text and group-type filters, and CSV-driven discovery / filtering
        - Nested group detection to ensure proper conversion order (optionally on demand)
        - Convert / roll back in batches per nesting depth level, with Success / Failed / Unverified reporting
        - Per-run CSV report, progress with ETA and a "Cancel job" button
        - Pagination support for large group lists
        - Export the current group list to a CSV file in the script directory

        This tool is intended to support the approach described in:
        https://learn.microsoft.com/en-us/entra/identity/hybrid/how-to-group-source-of-authority-configure

        .PARAMETER TenantId
        Optional. The Entra ID (Azure AD) tenant ID (GUID) to connect to.
        If not specified, the tool connects to the signed-in user's home tenant.
        Recommended in multi-tenant or partner scenarios to ensure you connect to the correct tenant.

        .EXAMPLE
        C:\PS> .\Group-SOA-Conversion-Tool_1.08.ps1

        .EXAMPLE
        C:\PS> .\Group-SOA-Conversion-Tool_1.08.ps1 -TenantId "00000000-0000-0000-0000-000000000000"

        .NOTES
        Version: 1.08

        Versions 1.04 - 1.07 (enterprise rework: batching, CSV waves, reports, cancel) were contributed by a community contributor.

        CHANGELOG:
        v1.08
        - Discovery also returns groups whose SOA is already converted (onPremisesSyncBehavior/isCloudManaged eq true): after conversion
          onPremisesSyncEnabled becomes null, so the 1.07 filter (onPremisesSyncEnabled eq true) hid converted groups and blocked rollback
        - CSV discovery accepts converted groups, applies the search scope tick boxes and no longer uses a broken fallback URL
        - Import CSV when groups are already loaded: choose to load the CSV groups fresh from Graph or to filter the current list
        - Disconnect now clears all loaded groups, nesting data and filters
        - "Select all" is unticked automatically when a row is clicked; the confirmation states that all filtered groups are processed
        - Search box is debounced (applied 400 ms after typing stops, or immediately on Enter)
        - Retry-After is honoured for single requests and whole batch calls; retries stop when a job is cancelled
        - Buttons that need Graph are only re-enabled after a job when still connected
        - Removed unused single-group convert / rollback functions, translated remaining comments to English, fixed an error message
        - Script saved as UTF-8 with BOM so special characters render correctly in Windows PowerShell 5.1

        v1.07
        - Search scope: three tick boxes (Security / Mail-enabled security / Distribution lists) applied on Search; none ticked = all synced groups, one or two ticked = only those types (replaces "Security groups only")
        - Discovery is limited to on-premises synced groups: server-side filter plus a client-side safety net that discards anything not synced
        - "Select all" tick box: selects every group in the list (all pages) for a mass Convert / Roll back (replaces "Apply to ALL filtered groups")
        - Removed the "Hide Converted Groups" option
        - Fixed import CSV after first discovery. now the discovery could be possible from specific csv list

        v1.06 (enterprise mode, tenants with 15K+ synced groups)
        - Fixed empty group id/name in the SOA batch lookup (range slicing on a generic List returned empty items, producing HTTP 403 on malformed URLs)
        - Discovery reworked: only on-premises synced groups are retrieved (server-side filter onPremisesSyncEnabled eq true); cloud-only groups are never queried
        - SOA status and nesting lookups use Graph JSON batching (20 requests per call) with Retry-After aware throttling handling: ~20x fewer calls, no more ThreadPool errors
        - Connect to Graph no longer loads groups automatically: use "Search Groups" (replaces "Refresh Groups")
        - "Security groups only" tick box (applied on Search): skips distribution and mail-enabled groups before any SOA lookup
        - "Skip nesting analysis" tick box: nesting is then analyzed on demand for the groups selected for convert / rollback
        - Convert / Roll back run in batches per nesting depth level; failed items are verified and reported as Success / Failed / Unverified (an unverified group is no longer counted as success)
        - "Apply to all filtered groups" tick box: operate on every filtered group across all pages (e.g. a CSV-driven wave) instead of only the selected rows
        - Per-run CSV report in the "Reports" folder; progress with ETA; "Cancel job" button; the window cannot be closed while a job runs
        - Raised ThreadPool minimum threads at startup

        v1.05
        - Log files are now written to a dedicated "Logs" subfolder of the script directory (created automatically; falls back to the script folder if it cannot be created)
        - Added discovery filters: group type, free-text search (display name / email) and import of a group list from CSV
        - CSV import accepts ObjectId / Mail / DisplayName columns (also the tool's own CSV export) or a plain single-column list; ',' and ';' delimiters
        - Nested-children warning now checks all loaded groups, so it is not skipped when a filter hides a child group
        - Added "Exit" button to the main window (closes the tool and disconnects from Graph)
        - Extended to on-premises synced Security Groups (not mail-enabled), in addition to Mail-Enabled Security Groups and Distribution Groups
        - Non mail-enabled security groups are listed only if they originate from on-premises AD (OnPremisesSecurityIdentifier / OnPremisesSyncEnabled), so cloud-only groups are not queried

        v1.04
        - Fixed 'no valid module file was found' after auto-install when the CurrentUser module folder is not in PSModulePath (e.g. Exchange servers, redirected Documents)
        - Verify module is discoverable after installation and log PSModulePath diagnostics

        v1.03
        - Added "Export List to CSV" button to export the current group list to the script directory

        v1.02 (2026-05-08)
        - Added scrollable confirmation dialog for large group selections
        - Prevents dialog overflow when converting/rolling back 100+ groups
        - Yes/No buttons now always visible and accessible

        v1.01
        - Bug fixes and improvements

        v1.0 (2026-03-04)
        - Initial release
        - Automatic permission consent flow
        - Display all Exchange-relevant groups with SOA status
        - Optional filter to hide converted groups
        - Nested group detection and smart ordering

        REQUIREMENTS:
        - Microsoft.Graph.Groups PowerShell module
        - Consent to the 'Group-OnPremisesSyncBehavior.ReadWrite.All' permission in Microsoft Graph
          (The tool will automatically prompt for consent when you connect)

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
    [Parameter(Mandatory=$false, HelpMessage="Enter the Entra ID (Azure AD) tenant ID (GUID) to connect to.")]
    [string]$TenantId
)

$script:Version = "1.08"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:LogFolder = Join-Path $script:ScriptPath "Logs"
if (-not (Test-Path -Path $script:LogFolder)) {
    try {
        [void](New-Item -Path $script:LogFolder -ItemType Directory -Force -ErrorAction Stop)
    }
    catch {
        Write-Host "WARNING: Could not create log folder '$($script:LogFolder)': $($_.Exception.Message). Falling back to the script folder."
        $script:LogFolder = $script:ScriptPath
    }
}
$script:LogFile = Join-Path $script:LogFolder "GroupSOAConversion_$(Get-Date -Format 'yyyyMMdd_HHmm').log"
$script:ReportFolder = Join-Path $script:ScriptPath "Reports"
try {
    if (-not (Test-Path -Path $script:ReportFolder)) {
        [void](New-Item -Path $script:ReportFolder -ItemType Directory -Force -ErrorAction Stop)
    }
}
catch {
    Write-Host "WARNING: Could not create report folder '$($script:ReportFolder)': $($_.Exception.Message). Falling back to the script folder."
    $script:ReportFolder = $script:ScriptPath
}
$script:AllGroups = @()
$script:AllGroupsUnfiltered = @()
$script:CurrentPage = 1
$script:PageSize = 100
$script:PermissionOk = $false
$script:NestingMap = @{}
$script:NestingDepth = @{}
$script:TenantId = $TenantId
$script:SortColumn = $null
$script:SortAscending = $true
$script:GroupTypeFilter = "All"
$script:BusyDepth = 0
$script:CancelRequested = $false
$script:NestingAnalyzed = @{}
$script:GroupById = @{}
$script:SearchText = ""
$script:CsvFilterIds = $null
$script:CsvFilterName = $null

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

function Show-ConfirmationDialog {
    param(
        [string]$Title,
        [string]$Message,
        [string]$GroupList,
        [string]$Icon = "Question"
    )

    $dialogForm = New-Object System.Windows.Forms.Form
    $dialogForm.Text = $Title
    $dialogForm.Size = New-Object System.Drawing.Size(650, 600)
    $dialogForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $dialogForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dialogForm.MaximizeBox = $false
    $dialogForm.MinimizeBox = $false
    $dialogForm.BackColor = [System.Drawing.Color]::White

    $iconLabel = New-Object System.Windows.Forms.Label
    $iconLabel.Location = New-Object System.Drawing.Point(20, 20)
    $iconLabel.Size = New-Object System.Drawing.Size(40, 40)
    $iconLabel.Font = New-Object System.Drawing.Font("Segoe UI", 24)
    $iconLabel.Text = if ($Icon -eq "Warning") { "⚠" } else { "?" }
    $iconLabel.ForeColor = if ($Icon -eq "Warning") { [System.Drawing.Color]::Orange } else { [System.Drawing.ColorTranslator]::FromHtml("#0078D4") }
    $dialogForm.Controls.Add($iconLabel)

    $messageLabel = New-Object System.Windows.Forms.Label
    $messageLabel.Location = New-Object System.Drawing.Point(70, 20)
    $messageLabel.Size = New-Object System.Drawing.Size(550, 115)
    $messageLabel.Text = $Message
    $messageLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $messageLabel.AutoSize = $false
    $dialogForm.Controls.Add($messageLabel)

    $groupListLabel = New-Object System.Windows.Forms.Label
    $groupListLabel.Location = New-Object System.Drawing.Point(20, 140)
    $groupListLabel.Size = New-Object System.Drawing.Size(600, 20)
    $groupListLabel.Text = "Groups will be processed in this order:"
    $groupListLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $dialogForm.Controls.Add($groupListLabel)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(20, 165)
    $textBox.Size = New-Object System.Drawing.Size(600, 320)
    $textBox.Multiline = $true
    $textBox.ReadOnly = $true
    $textBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $textBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $textBox.Text = $GroupList
    $textBox.BackColor = [System.Drawing.Color]::WhiteSmoke
    $textBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $dialogForm.Controls.Add($textBox)

    $buttonYes = New-Object System.Windows.Forms.Button
    $buttonYes.Location = New-Object System.Drawing.Point(370, 510)
    $buttonYes.Size = New-Object System.Drawing.Size(120, 35)
    $buttonYes.Text = "Yes"
    $buttonYes.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $buttonYes.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $buttonYes.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#0078D4")
    $buttonYes.ForeColor = [System.Drawing.Color]::White
    $buttonYes.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $buttonYes.FlatAppearance.BorderSize = 0
    $buttonYes.Cursor = [System.Windows.Forms.Cursors]::Hand
    $dialogForm.Controls.Add($buttonYes)

    $buttonNo = New-Object System.Windows.Forms.Button
    $buttonNo.Location = New-Object System.Drawing.Point(500, 510)
    $buttonNo.Size = New-Object System.Drawing.Size(120, 35)
    $buttonNo.Text = "No"
    $buttonNo.DialogResult = [System.Windows.Forms.DialogResult]::No
    $buttonNo.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $buttonNo.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#E1E1E1")
    $buttonNo.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
    $buttonNo.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $buttonNo.FlatAppearance.BorderSize = 0
    $buttonNo.Cursor = [System.Windows.Forms.Cursors]::Hand
    $dialogForm.Controls.Add($buttonNo)

    $dialogForm.AcceptButton = $buttonYes
    $dialogForm.CancelButton = $buttonNo

    return $dialogForm.ShowDialog()
}

function Update-GroupGrid {
    param(
        [Parameter(Mandatory=$false)]
        [array]$Groups = $null
    )

    if ($null -ne $Groups) {
        $script:AllGroupsUnfiltered = $Groups
    }

    $script:AllGroups = $script:AllGroupsUnfiltered

    # Apply discovery filters (group type, CSV list, free-text search)
    if ($script:GroupTypeFilter -ne "All") {
        $script:AllGroups = @($script:AllGroups | Where-Object { $_.GroupType -eq $script:GroupTypeFilter })
    }

    if ($null -ne $script:CsvFilterIds) {
        $script:AllGroups = @($script:AllGroups | Where-Object { $script:CsvFilterIds.ContainsKey($_.Id) })
    }

    if (-not [string]::IsNullOrEmpty($script:SearchText)) {
        $script:AllGroups = @($script:AllGroups | Where-Object {
            ("$($_.DisplayName)".IndexOf($script:SearchText, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) -or
            ("$($_.Mail)".IndexOf($script:SearchText, [System.StringComparison]::OrdinalIgnoreCase) -ge 0)
        })
    }

    # Apply sorting if a column is selected
    if ($null -ne $script:SortColumn) {
        $script:AllGroups = switch ($script:SortColumn) {
            "DisplayName" {
                if ($script:SortAscending) {
                    $script:AllGroups | Sort-Object { $_.DisplayName }
                } else {
                    $script:AllGroups | Sort-Object { $_.DisplayName } -Descending
                }
            }
            "Email" {
                if ($script:SortAscending) {
                    $script:AllGroups | Sort-Object { $_.Mail }
                } else {
                    $script:AllGroups | Sort-Object { $_.Mail } -Descending
                }
            }
            "GroupType" {
                if ($script:SortAscending) {
                    $script:AllGroups | Sort-Object { $_.GroupType }
                } else {
                    $script:AllGroups | Sort-Object { $_.GroupType } -Descending
                }
            }
            "IsCloudManaged" {
                if ($script:SortAscending) {
                    $script:AllGroups | Sort-Object { $_.IsCloudManaged }
                } else {
                    $script:AllGroups | Sort-Object { $_.IsCloudManaged } -Descending
                }
            }
            "NestingDepth" {
                if ($script:SortAscending) {
                    $script:AllGroups | Sort-Object { if ($script:NestingDepth.ContainsKey($_.Id)) { $script:NestingDepth[$_.Id] } else { 0 } }
                } else {
                    $script:AllGroups | Sort-Object { if ($script:NestingDepth.ContainsKey($_.Id)) { $script:NestingDepth[$_.Id] } else { 0 } } -Descending
                }
            }
            default { $script:AllGroups }
        }
    }

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

            $cloudManagedStatus = if ($group.IsCloudManaged -is [string] -and $group.IsCloudManaged -eq "Unknown") { 
                "Unknown" 
            } elseif ($group.IsCloudManaged -eq $true) { 
                "True" 
            } else { 
                "False" 
            }

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

    # Keep "Select all" applied when the page, sort or filters change
    if ($checkboxAllFiltered -and $checkboxAllFiltered.Checked) {
        $dataGridView.SelectAll()
    }
}

function Add-GraphModulePathToSession {
    $candidates = @(Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'WindowsPowerShell\Modules')
    $installed = Get-InstalledModule -Name Microsoft.Graph.Groups -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($installed -and $installed.InstalledLocation) {
        Write-Log "Microsoft.Graph.Groups installed location: $($installed.InstalledLocation)"
        $candidates += Split-Path (Split-Path $installed.InstalledLocation -Parent) -Parent
    }
    $current = $env:PSModulePath -split ';' | ForEach-Object { $_.TrimEnd('\') }
    foreach ($path in ($candidates | Select-Object -Unique)) {
        if ((Test-Path $path) -and ($current -notcontains $path.TrimEnd('\'))) {
            $env:PSModulePath = "$path;$env:PSModulePath"
            Write-Log "Added '$path' to PSModulePath for this session." -Level WARNING
        }
    }
}

function Search-GraphModule {
    Write-Log "Checking for Microsoft Graph Groups module..."

    Add-GraphModulePathToSession

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

            Add-GraphModulePathToSession
            if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Groups)) {
                Write-Log "PSModulePath: $env:PSModulePath" -Level ERROR
                throw "Module was installed but cannot be found in any PSModulePath directory."
            }

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

function Connect-GraphSession {
    if ([string]::IsNullOrEmpty($script:TenantId)) {
        Write-Log "Attempting to connect to Microsoft Graph (default tenant)..."
    } else {
        Write-Log "Attempting to connect to Microsoft Graph (TenantId: $($script:TenantId))..."
    }

    try {
        Add-GraphModulePathToSession
        Import-Module Microsoft.Graph.Groups -ErrorAction Stop

        if ([string]::IsNullOrEmpty($script:TenantId)) {
            Connect-MgGraph -Scopes 'Group.ReadWrite.All','Group-OnPremisesSyncBehavior.ReadWrite.All' -ErrorAction Stop -NoWelcome
        } else {
            Connect-MgGraph -Scopes 'Group.ReadWrite.All','Group-OnPremisesSyncBehavior.ReadWrite.All' -TenantId $script:TenantId -ErrorAction Stop -NoWelcome
        }

        $context = Get-MgContext
        Write-Log "Successfully connected to Microsoft Graph. TenantId: $($context.TenantId)"

        $form.Text = "Group SOA Conversion Tool - Tenant: $($context.TenantId)"

        if (Test-GraphPermissions) {
            Write-Log "Permission 'Group-OnPremisesSyncBehavior.ReadWrite.All' is already consented."
            $script:PermissionOk = $true
        }
        else {
            Write-Log "Permission missing. Triggering consent flow..." -Level WARNING

            Disconnect-MgGraph -ErrorAction SilentlyContinue

            if ([string]::IsNullOrEmpty($script:TenantId)) {
                Connect-MgGraph -Scopes 'Group.ReadWrite.All','Group-OnPremisesSyncBehavior.ReadWrite.All' -ErrorAction Stop -NoWelcome
            } else {
                Connect-MgGraph -Scopes 'Group.ReadWrite.All','Group-OnPremisesSyncBehavior.ReadWrite.All' -TenantId $script:TenantId -ErrorAction Stop -NoWelcome
            }

            Write-Log "Consent flow completed."

            if (Test-GraphPermissions) {
                Write-Log "Permission 'Group-OnPremisesSyncBehavior.ReadWrite.All' successfully granted."
                $script:PermissionOk = $true

                [System.Windows.Forms.MessageBox]::Show(
                    "Permission 'Group-OnPremisesSyncBehavior.ReadWrite.All' has been successfully granted.",
                    "Permission Granted",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
            }
            else {
                Write-Log "Permission was not granted after consent flow." -Level WARNING
                $script:PermissionOk = $false

                [System.Windows.Forms.MessageBox]::Show(
                    "The permission was not granted. This may require admin consent.`n`nTo grant manually:`n1. Go to Entra admin center`n2. Navigate to Enterprise Applications`n3. Find 'Microsoft Graph Command Line Tools'`n4. Go to Permissions`n5. Click 'Grant admin consent'`n`nAlternatively, ask your Application Administrator or Cloud Application Administrator to grant consent.",
                    "Permission Not Granted",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
            }
        }

        return $true
    }
    catch {
        Write-Log "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -Level ERROR
        Write-Log "PSModulePath: $env:PSModulePath" -Level ERROR

        [System.Windows.Forms.MessageBox]::Show(
            "Failed to connect to Microsoft Graph.`n`nError: $($_.Exception.Message)",
            "Connection Failed",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )

        return $false
    }
}

# ============================================================
# ENTERPRISE HELPERS (busy state, throttling, JSON batching)
# ============================================================

function Start-Busy {
    $script:BusyDepth++

    if ($script:BusyDepth -eq 1) {
        $script:CancelRequested = $false

        foreach ($control in @($buttonRefresh, $buttonDisconnect, $buttonConvertToCloud, $buttonConvertToOnPrem, $buttonImportCsv, $buttonClearFilters, $buttonExportCsv, $buttonExit)) {
            if ($control) { $control.Enabled = $false }
        }

        if ($buttonCancel) { $buttonCancel.Enabled = $true }
    }
}

function Stop-Busy {
    if ($script:BusyDepth -gt 0) { $script:BusyDepth-- }

    if ($script:BusyDepth -eq 0) {
        $connected = $false
        try { $connected = ($null -ne (Get-MgContext)) } catch { $connected = $false }

        foreach ($control in @($buttonClearFilters, $buttonExit)) {
            if ($control) { $control.Enabled = $true }
        }

        foreach ($control in @($buttonConvertToCloud, $buttonConvertToOnPrem, $buttonImportCsv, $buttonExportCsv, $buttonRefresh, $buttonDisconnect)) {
            if ($control) { $control.Enabled = $connected }
        }
        if ($buttonCancel) { $buttonCancel.Enabled = $false }
    }
}

function Wait-Interruptible {
    param([double]$Seconds)

    $end = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $end -and -not $script:CancelRequested) {
        Start-Sleep -Milliseconds 200
        [System.Windows.Forms.Application]::DoEvents()
    }
}

function Format-Eta {
    param(
        [System.Diagnostics.Stopwatch]$Watch,
        [int]$Done,
        [int]$Total
    )

    if ($Done -le 0) { return "calculating..." }

    $remaining = [TimeSpan]::FromSeconds($Watch.Elapsed.TotalSeconds / $Done * ($Total - $Done))
    return $remaining.ToString('hh\:mm\:ss')
}

function Get-HttpStatusCode {
    param($ErrorRecord)

    try { return [int]$ErrorRecord.Exception.Response.StatusCode } catch { return 0 }
}

function Get-RetryAfterSeconds {
    param($ErrorRecord)

    try {
        $retryAfter = $ErrorRecord.Exception.Response.Headers.RetryAfter
        if ($null -eq $retryAfter) { return 0 }
        if ($null -ne $retryAfter.Delta) { return [int][Math]::Ceiling($retryAfter.Delta.TotalSeconds) }
        if ($null -ne $retryAfter.Date) { return [int][Math]::Max(0, [Math]::Ceiling(($retryAfter.Date - [DateTimeOffset]::UtcNow).TotalSeconds)) }
    }
    catch { return 0 }

    return 0
}

function Invoke-GraphGetWithRetry {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Uri,

        [hashtable]$Headers = $null,

        [int]$MaxAttempts = 6
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            if ($Headers) {
                return (Invoke-MgGraphRequest -Uri $Uri -Method GET -Headers $Headers -ErrorAction Stop)
            }

            return (Invoke-MgGraphRequest -Uri $Uri -Method GET -ErrorAction Stop)
        }
        catch {
            $err = $_
            $code = Get-HttpStatusCode -ErrorRecord $err
            $retriable = ($code -eq 0 -or $code -eq 429 -or $code -ge 500)

            if (-not $retriable -or $attempt -ge $MaxAttempts) { throw $err }

            if ($script:CancelRequested) { throw "Cancelled by user" }

            $delay = Get-RetryAfterSeconds -ErrorRecord $err
            if ($delay -le 0) { $delay = [Math]::Min(60, [Math]::Pow(2, $attempt)) }
            Write-Log "Graph GET failed (HTTP $code), retry $attempt of $($MaxAttempts - 1) in $delay s: $($err.Exception.Message)" -Level WARNING
            Wait-Interruptible -Seconds $delay

            if ($script:CancelRequested) { throw "Cancelled by user" }
        }
    }
}

function Invoke-GraphBatch {
    <#
        Sends sub-requests through Graph JSON batching (max 20 per call) and returns a hashtable id -> sub-response.
        Throttled / transient sub-responses (429, 502, 503, 504) are retried, honouring Retry-After.
        Sub-requests that never succeed get a synthetic response with status 0.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [array]$Requests,

        [int]$MaxAttempts = 6
    )

    $results = @{}
    $pending = @{}
    foreach ($request in $Requests) { $pending[[string]$request.id] = $request }

    $attempt = 0
    $fatalMessage = $null
    $cancelled = $false

    while ($pending.Count -gt 0 -and $attempt -lt $MaxAttempts) {
        # The first attempt always runs; once a cancel is requested no further retries are made
        if ($attempt -gt 0 -and $script:CancelRequested) {
            $cancelled = $true
            break
        }

        $attempt++
        $waitSeconds = 0

        try {
            $payload = @{ requests = @($pending.Values) } | ConvertTo-Json -Depth 10 -Compress
            $response = Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/$batch' -Body $payload -ContentType 'application/json' -ErrorAction Stop

            foreach ($item in $response.responses) {
                $requestId = [string]$item.id
                $status = [int]$item.status

                if ($status -eq 429 -or $status -eq 502 -or $status -eq 503 -or $status -eq 504) {
                    $retryAfter = 0
                    try { [void][int]::TryParse([string]$item.headers['Retry-After'], [ref]$retryAfter) } catch { $retryAfter = 0 }
                    $waitSeconds = [Math]::Max($waitSeconds, $retryAfter)
                }
                else {
                    $results[$requestId] = $item
                    $pending.Remove($requestId)
                }
            }
        }
        catch {
            $err = $_
            $code = Get-HttpStatusCode -ErrorRecord $err

            if ($code -eq 400 -or $code -eq 401 -or $code -eq 403) {
                $fatalMessage = "Batch call rejected (HTTP $code): $($err.Exception.Message)"
                Write-Log $fatalMessage -Level ERROR
                break
            }

            Write-Log "Batch call failed (HTTP $code): $($err.Exception.Message)" -Level WARNING
            $waitSeconds = Get-RetryAfterSeconds -ErrorRecord $err
            if ($waitSeconds -le 0) { $waitSeconds = [Math]::Min(60, [Math]::Pow(2, $attempt)) }
        }

        if ($pending.Count -gt 0 -and $attempt -lt $MaxAttempts -and -not $script:CancelRequested) {
            if ($waitSeconds -le 0) { $waitSeconds = [Math]::Min(30, [Math]::Pow(2, $attempt)) }
            Write-Log "$($pending.Count) sub-request(s) throttled or incomplete, retry $attempt of $($MaxAttempts - 1) in $waitSeconds s." -Level WARNING
            Wait-Interruptible -Seconds $waitSeconds
        }
    }

    foreach ($requestId in @($pending.Keys)) {
        $message = if ($fatalMessage) { $fatalMessage } elseif ($cancelled -or $script:CancelRequested) { "Cancelled during retry" } else { "Retries exhausted (throttling or transient errors)" }
        $results[$requestId] = @{ id = $requestId; status = 0; body = @{ error = @{ message = $message } } }
    }

    return $results
}

function New-SoaResult {
    param(
        $Group,
        [string]$Operation,
        [string]$Result,
        [string]$Detail
    )

    $depth = 0
    if ($script:NestingDepth.ContainsKey($Group.Id)) { $depth = $script:NestingDepth[$Group.Id] }

    return [PSCustomObject]@{
        Timestamp   = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Operation   = $Operation
        DisplayName = $Group.DisplayName
        Mail        = $Group.Mail
        GroupType   = $Group.GroupType
        ObjectId    = $Group.Id
        Depth       = $depth
        Result      = $Result
        Detail      = $Detail
        Group       = $Group
    }
}

function Invoke-SoaBatchChange {
    <#
        Converts (CloudManaged = $true) or rolls back (CloudManaged = $false) the given groups using JSON batching.
        Groups are processed by nesting depth level (a level is completed before the next one starts):
        ascending for conversion (children first), descending for rollback (parents first).
        Failed sub-requests are verified with a GET; a group is reported as Success only when confirmed.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [array]$Groups,

        [Parameter(Mandatory=$true)]
        [bool]$CloudManaged
    )

    $operation = if ($CloudManaged) { "Convert" } else { "Rollback" }
    $verb = if ($CloudManaged) { "Converting" } else { "Rolling back" }
    $results = New-Object System.Collections.Generic.List[object]
    $processed = @{}

    Start-Busy
    try {
        $levels = @($Groups | Group-Object -Property { if ($script:NestingDepth.ContainsKey($_.Id)) { [int]$script:NestingDepth[$_.Id] } else { 0 } })

        if ($CloudManaged) {
            $levels = @($levels | Sort-Object { [int]$_.Name })
        } else {
            $levels = @($levels | Sort-Object { [int]$_.Name } -Descending)
        }

        $batches = New-Object System.Collections.Generic.List[object]
        foreach ($level in $levels) {
            $levelGroups = @($level.Group)
            for ($offset = 0; $offset -lt $levelGroups.Count; $offset += 20) {
                $last = [Math]::Min($offset + 20, $levelGroups.Count) - 1
                [void]$batches.Add([PSCustomObject]@{ Groups = @($levelGroups[$offset..$last]) })
            }
        }

        $total = $Groups.Count
        $done = 0
        $watch = [System.Diagnostics.Stopwatch]::StartNew()
        Write-Log "$operation started: $total group(s), $($batches.Count) batch(es), $($levels.Count) depth level(s)."

        foreach ($batch in $batches) {
            if ($script:CancelRequested) {
                Write-Log "$operation cancelled by user. Remaining groups are not processed." -Level WARNING
                break
            }

            $requests = @()
            $map = @{}
            for ($i = 0; $i -lt $batch.Groups.Count; $i++) {
                $requestId = [string]$i
                $map[$requestId] = $batch.Groups[$i]
                $requests += @{
                    id      = $requestId
                    method  = 'PATCH'
                    url     = "/groups/$($batch.Groups[$i].Id)/onPremisesSyncBehavior"
                    headers = @{ 'Content-Type' = 'application/json' }
                    body    = @{ isCloudManaged = $CloudManaged }
                }
            }

            $responses = Invoke-GraphBatch -Requests $requests
            $failed = @()

            foreach ($requestId in $map.Keys) {
                $group = $map[$requestId]
                $response = $responses[$requestId]
                $status = 0
                if ($response) { $status = [int]$response.status }

                if ($status -ge 200 -and $status -lt 300) {
                    $results.Add((New-SoaResult -Group $group -Operation $operation -Result 'Success' -Detail "HTTP $status"))
                }
                else {
                    $message = ''
                    if ($response -and $response.body -and $response.body.error) { $message = [string]$response.body.error.message }
                    $failed += [PSCustomObject]@{ Group = $group; Status = $status; Message = $message }
                }
            }

            if ($failed.Count -gt 0) {
                # Verify the real state of the groups whose sub-request reported an error
                Wait-Interruptible -Seconds 3

                $verifyRequests = @()
                $verifyMap = @{}
                for ($i = 0; $i -lt $failed.Count; $i++) {
                    $requestId = [string]$i
                    $verifyMap[$requestId] = $failed[$i]
                    $verifyRequests += @{
                        id     = $requestId
                        method = 'GET'
                        url    = "/groups/$($failed[$i].Group.Id)/onPremisesSyncBehavior?`$select=isCloudManaged"
                    }
                }

                $verifyResponses = Invoke-GraphBatch -Requests $verifyRequests

                foreach ($requestId in $verifyMap.Keys) {
                    $item = $verifyMap[$requestId]
                    $verify = $verifyResponses[$requestId]
                    $detail = "HTTP $($item.Status) $($item.Message)".Trim()

                    if ($verify -and [int]$verify.status -eq 200 -and $verify.body -and $null -ne $verify.body.isCloudManaged) {
                        if (([bool]$verify.body.isCloudManaged) -eq $CloudManaged) {
                            $results.Add((New-SoaResult -Group $item.Group -Operation $operation -Result 'Success' -Detail "Verified after API error: $detail"))
                        } else {
                            $results.Add((New-SoaResult -Group $item.Group -Operation $operation -Result 'Failed' -Detail $detail))
                        }
                    }
                    else {
                        $results.Add((New-SoaResult -Group $item.Group -Operation $operation -Result 'Unverified' -Detail "$detail (state could not be verified, check the group manually)"))
                    }
                }
            }

            foreach ($group in $batch.Groups) { $processed[$group.Id] = $true }
            $done += $batch.Groups.Count

            $statusLabel.Text = "$verb $done of $total - ETA $(Format-Eta -Watch $watch -Done $done -Total $total)"
            [System.Windows.Forms.Application]::DoEvents()
        }

        foreach ($group in $Groups) {
            if (-not $processed.ContainsKey($group.Id)) {
                $results.Add((New-SoaResult -Group $group -Operation $operation -Result 'Skipped' -Detail 'Cancelled before processing'))
            }
        }

        $resultById = @{}
        foreach ($result in $results) {
            $resultById[$result.ObjectId] = $result

            if ($result.Result -eq 'Success') {
                $result.Group.IsCloudManaged = $CloudManaged
                Write-Log "$operation OK: '$($result.DisplayName)' ($($result.ObjectId)) - $($result.Detail)"
            }
            elseif ($result.Result -eq 'Failed') {
                Write-Log "$operation FAILED: '$($result.DisplayName)' ($($result.ObjectId)) - $($result.Detail)" -Level ERROR
            }
            elseif ($result.Result -eq 'Unverified') {
                Write-Log "$operation UNVERIFIED: '$($result.DisplayName)' ($($result.ObjectId)) - $($result.Detail)" -Level WARNING
            }
        }

        foreach ($row in $dataGridView.Rows) {
            $rowId = $row.Cells["ObjectId"].Value
            if ($rowId -and $resultById.ContainsKey($rowId) -and $resultById[$rowId].Result -eq 'Success') {
                $row.Cells["IsCloudManaged"].Value = if ($CloudManaged) { "True" } else { "False" }
            }
        }

        return $results.ToArray()
    }
    finally {
        Stop-Busy
    }
}

function Export-SoaReport {
    param(
        $Results,
        [string]$Operation
    )

    $reportFile = Join-Path $script:ReportFolder "GroupSOA_${Operation}_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

    try {
        @($Results) | Select-Object Timestamp, Operation, DisplayName, Mail, GroupType, ObjectId, Depth, Result, Detail |
            Export-Csv -Path $reportFile -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        Write-Log "Report written to $reportFile"
    }
    catch {
        Write-Log "Could not write report '$reportFile': $($_.Exception.Message)" -Level WARNING
        $reportFile = "(report not written - see log)"
    }

    return $reportFile
}

# ============================================================
# DISCOVERY (on-premises synced + already converted groups, server-side filters + batching)
# ============================================================

function Get-GroupKind {
    <#
        Returns 'Security', 'MailSecurity' or 'Distribution' for a Graph group item, or $null for groups the tool does not handle
        (Microsoft 365 / Unified groups, or groups that are neither mail- nor security-enabled).
    #>
    param($Item)

    if ($Item.groupTypes -contains 'Unified') { return $null }

    $isMail = ($Item.mailEnabled -eq $true)
    $isSecurity = ($Item.securityEnabled -eq $true)

    if ($isMail -and $isSecurity) { return 'MailSecurity' }
    if ($isMail) { return 'Distribution' }
    if ($isSecurity) { return 'Security' }
    return $null
}

function Get-SearchScope {
    $scope = [PSCustomObject]@{
        Security     = [bool]$checkboxTypeSecurity.Checked
        MailSecurity = [bool]$checkboxTypeMailSecurity.Checked
        Distribution = [bool]$checkboxTypeDistribution.Checked
        Any          = $false
    }
    $scope.Any = ($scope.Security -or $scope.MailSecurity -or $scope.Distribution)
    return $scope
}

function Test-GroupInSearchScope {
    param(
        [string]$Kind,
        $Scope
    )

    if ([string]::IsNullOrEmpty($Kind)) { return $false }
    if (-not $Scope.Any) { return $true }
    return [bool]$Scope.$Kind
}

function New-GroupRecord {
    param(
        $Item,
        $IsCloudManaged
    )

    $groupType = switch (Get-GroupKind -Item $Item) {
        'MailSecurity' { "Mail-Enabled Security Group" }
        'Distribution' { "Distribution Group" }
        default        { "Security Group" }
    }

    return [PSCustomObject]@{
        Id               = $Item.id
        DisplayName      = $Item.displayName
        Mail             = $Item.mail
        GroupType        = $groupType
        IsCloudManaged   = $IsCloudManaged
        SecurityEnabled  = $Item.securityEnabled
        MailEnabled      = $Item.mailEnabled
    }
}

function Get-GraphPagedItems {
    <#
        Follows @odata.nextLink until all pages are read (advanced query: ConsistencyLevel = eventual).
        Returns a List[object], or $null if the user cancelled.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$Uri,

        [string]$Label = "groups"
    )

    $items = New-Object System.Collections.Generic.List[object]
    $nextUri = $Uri

    while ($nextUri) {
        if ($script:CancelRequested) { return $null }

        $page = Invoke-GraphGetWithRetry -Uri $nextUri -Headers @{ ConsistencyLevel = 'eventual' }
        foreach ($item in $page.value) { [void]$items.Add($item) }

        $statusLabel.Text = "Retrieving $Label... $($items.Count)"
        [System.Windows.Forms.Application]::DoEvents()

        $nextUri = $page.'@odata.nextLink'
    }

    return ,$items
}

function Get-SoaStatusMap {
    <#
        Reads isCloudManaged for the given Graph group items through JSON batching (20 per call).
        Returns a hashtable id -> $true / $false / "Unknown", or $null if the user cancelled.
        -RetrySingle retries every failed lookup once with a direct request.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [AllowEmptyCollection()]
        [array]$Items,

        [switch]$RetrySingle
    )

    $map = @{}
    $total = $Items.Count
    $unknownCount = 0
    $batchSize = 20
    $watch = [System.Diagnostics.Stopwatch]::StartNew()

    for ($offset = 0; $offset -lt $total; $offset += $batchSize) {
        if ($script:CancelRequested) { return $null }

        $last = [Math]::Min($offset + $batchSize, $total) - 1
        $chunk = @($Items[$offset..$last])

        $requests = @()
        for ($i = 0; $i -lt $chunk.Count; $i++) {
            $requests += @{
                id     = [string]$i
                method = 'GET'
                url    = "/groups/$($chunk[$i].id)/onPremisesSyncBehavior?`$select=isCloudManaged"
            }
        }

        $responses = Invoke-GraphBatch -Requests $requests

        for ($i = 0; $i -lt $chunk.Count; $i++) {
            $item = $chunk[$i]
            $response = $responses[[string]$i]
            $isCloudManaged = "Unknown"
            $reason = ''

            if ($response -and [int]$response.status -eq 200 -and $response.body -and $null -ne $response.body.isCloudManaged) {
                $isCloudManaged = ($response.body.isCloudManaged -eq $true)
            }
            else {
                if ($response -and $response.body -and $response.body.error) { $reason = [string]$response.body.error.message }
                $reason = "HTTP $(if ($response) { $response.status } else { 'n/a' }) $reason".Trim()

                if ($RetrySingle -and -not $script:CancelRequested) {
                    try {
                        $direct = Invoke-GraphGetWithRetry -Uri "https://graph.microsoft.com/v1.0/groups/$($item.id)/onPremisesSyncBehavior?`$select=isCloudManaged" -MaxAttempts 2
                        if ($null -ne $direct -and $null -ne $direct.isCloudManaged) {
                            $isCloudManaged = ($direct.isCloudManaged -eq $true)
                        }
                    }
                    catch {
                        $reason = "$reason / direct request: $($_.Exception.Message)"
                    }
                }
            }

            if ($isCloudManaged -is [string]) {
                $unknownCount++
                if ($unknownCount -le 100) {
                    Write-Log "Could not retrieve SOA status for group '$($item.displayName)' ($($item.id)): $reason" -Level WARNING
                }
            }

            $map[$item.id] = $isCloudManaged
        }

        $done = $offset + $chunk.Count
        $statusLabel.Text = "Loading SOA status $done of $total - ETA $(Format-Eta -Watch $watch -Done $done -Total $total)"
        [System.Windows.Forms.Application]::DoEvents()
    }

    if ($unknownCount -gt 100) {
        Write-Log "SOA status could not be retrieved for $unknownCount groups in total (only the first 100 are listed above)." -Level WARNING
    }

    return $map
}

function Get-ExchangeGroups {
    $scope = Get-SearchScope

    if ($scope.Any) {
        $scopeNames = @()
        if ($scope.Security) { $scopeNames += "Security Group" }
        if ($scope.MailSecurity) { $scopeNames += "Mail-Enabled Security Group" }
        if ($scope.Distribution) { $scopeNames += "Distribution Group" }
        Write-Log "Search scope: $($scopeNames -join ', ')."
    }
    else {
        Write-Log "Search scope: all synced and converted group types."
    }

    Start-Busy
    try {
        $select = 'id,displayName,mail,mailEnabled,securityEnabled,groupTypes,onPremisesSyncEnabled,onPremisesSecurityIdentifier'

        # Query 1 - on-premises synced groups that are not converted yet (onPremisesSyncEnabled = true)
        Write-Log "Retrieving on-premises synced groups (filter: onPremisesSyncEnabled eq true)..."
        $syncedUri = "https://graph.microsoft.com/v1.0/groups?`$filter=onPremisesSyncEnabled%20eq%20true&`$count=true&`$select=$select&`$top=999"
        $synced = Get-GraphPagedItems -Uri $syncedUri -Label "synced groups"
        if ($null -eq $synced) {
            Write-Log "Search cancelled by user while retrieving groups." -Level WARNING
            return $null
        }
        Write-Log "Retrieved $($synced.Count) on-premises synced groups."

        # Query 2 - groups whose SOA is already converted (onPremisesSyncEnabled becomes null after conversion)
        Write-Log "Retrieving converted groups (filter: onPremisesSyncBehavior/isCloudManaged eq true)..."
        $convertedUri = "https://graph.microsoft.com/v1.0/groups?`$filter=onPremisesSyncBehavior/isCloudManaged%20eq%20true&`$count=true&`$select=$select&`$top=999"
        try {
            $converted = Get-GraphPagedItems -Uri $convertedUri -Label "converted groups"
            if ($null -eq $converted) {
                Write-Log "Search cancelled by user while retrieving groups." -Level WARNING
                return $null
            }
            Write-Log "Retrieved $($converted.Count) converted (cloud-managed) groups."
        }
        catch {
            if ($script:CancelRequested) { throw }
            Write-Log "Could not retrieve converted groups, continuing with synced groups only (converted groups will not be listed): $($_.Exception.Message)" -Level WARNING
            $converted = New-Object System.Collections.Generic.List[object]
        }

        # Merge by id; groups returned by query 2 are known to be cloud-managed
        $knownCloud = @{}
        foreach ($item in $converted) { $knownCloud[$item.id] = $true }

        $merged = [ordered]@{}
        foreach ($item in $synced) { $merged[$item.id] = $item }
        foreach ($item in $converted) { if (-not $merged.Contains($item.id)) { $merged[$item.id] = $item } }

        $lookupCandidates = New-Object System.Collections.Generic.List[object]
        $cloudCandidates = New-Object System.Collections.Generic.List[object]
        $discardedCount = 0
        foreach ($item in $merged.Values) {
            # Safety net: only on-premises synced groups or groups with a converted SOA are processed
            if ($item.onPremisesSyncEnabled -ne $true -and -not $knownCloud.ContainsKey($item.id)) { $discardedCount++; continue }
            if (-not (Test-GroupInSearchScope -Kind (Get-GroupKind -Item $item) -Scope $scope)) { continue }

            if ($knownCloud.ContainsKey($item.id)) {
                [void]$cloudCandidates.Add($item)
            } else {
                [void]$lookupCandidates.Add($item)
            }
        }

        if ($discardedCount -gt 0) {
            Write-Log "Discarded $discardedCount group(s) returned by Graph that are neither on-premises synced nor converted." -Level WARNING
        }
        Write-Log "Kept $($lookupCandidates.Count + $cloudCandidates.Count) of $($merged.Count) group(s) after the search scope filter ($($lookupCandidates.Count) synced, $($cloudCandidates.Count) converted; Microsoft 365 groups excluded)."

        # SOA status through JSON batching, only for groups whose state is not already known
        $soaMap = Get-SoaStatusMap -Items $lookupCandidates.ToArray()
        if ($null -eq $soaMap) {
            Write-Log "Search cancelled by user while loading SOA status." -Level WARNING
            return $null
        }

        $groupResults = New-Object System.Collections.Generic.List[object]
        $unknownCount = 0
        foreach ($item in $lookupCandidates) {
            $state = $soaMap[$item.id]
            if ($state -is [string]) { $unknownCount++ }
            $groupResults.Add((New-GroupRecord -Item $item -IsCloudManaged $state))
        }
        foreach ($item in $cloudCandidates) {
            $groupResults.Add((New-GroupRecord -Item $item -IsCloudManaged $true))
        }

        Write-Log "Finished loading SOA status for $($groupResults.Count) group(s). Unknown: $unknownCount."

        return $groupResults.ToArray()
    }
    catch {
        if ($script:CancelRequested) {
            Write-Log "Search cancelled by user: $($_.Exception.Message)" -Level WARNING
            return $null
        }

        Write-Log "Failed to retrieve groups: $($_.Exception.Message)" -Level ERROR

        [System.Windows.Forms.MessageBox]::Show(
            "Failed to retrieve groups.`n`nError: $($_.Exception.Message)",
            "Retrieval Failed",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )

        return $null
    }
    finally {
        Stop-Busy
    }
}

function Reset-NestingMap {
    param(
        [Parameter(Mandatory=$true)]
        [array]$Groups
    )

    $script:NestingMap = @{}
    $script:NestingDepth = @{}
    $script:NestingAnalyzed = @{}

    foreach ($group in $Groups) {
        $script:NestingDepth[$group.Id] = 0
    }
}

function Reset-LoadedGroups {
    # Clears every loaded group, nesting data and filter so the next Search / CSV import starts from scratch
    $script:AllGroupsUnfiltered = @()
    $script:AllGroups = @()
    $script:GroupById = @{}
    $script:NestingMap = @{}
    $script:NestingDepth = @{}
    $script:NestingAnalyzed = @{}
    $script:CsvFilterIds = $null
    $script:CsvFilterName = $null
    $script:CurrentPage = 1

    if ($labelCsv) { $labelCsv.Text = "CSV: none" }
    if ($checkboxAllFiltered) { $checkboxAllFiltered.Checked = $false }

    Update-GroupGrid -Groups @()
}

function Build-NestingMap {
    <#
        Full mode (no -ParentGroups): analyzes every loaded group.
        Targeted mode (-ParentGroups): analyzes only those groups (not yet analyzed) and merges the result.
        Child lookups use JSON batching and ask Graph for group-type members only.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [array]$Groups,

        [Parameter(Mandatory=$false)]
        [array]$ParentGroups = $null
    )

    if ($null -eq $ParentGroups) {
        Reset-NestingMap -Groups $Groups
        $parents = @($Groups)
    }
    else {
        $parents = @($ParentGroups | Where-Object { -not $script:NestingAnalyzed.ContainsKey($_.Id) })
    }

    Write-Log "Building nested group map for $($parents.Count) group(s)..."

    $groupIds = @{}
    foreach ($group in $Groups) {
        $groupIds[$group.Id] = $true
        if (-not $script:NestingDepth.ContainsKey($group.Id)) { $script:NestingDepth[$group.Id] = 0 }
    }

    Start-Busy
    try {
        $total = $parents.Count
        $batchSize = 20
        $failedCount = 0
        $watch = [System.Diagnostics.Stopwatch]::StartNew()

        for ($offset = 0; $offset -lt $total; $offset += $batchSize) {
            if ($script:CancelRequested) {
                Write-Log "Nesting analysis cancelled by user." -Level WARNING
                break
            }

            $last = [Math]::Min($offset + $batchSize, $total) - 1
            $chunk = @($parents[$offset..$last])

            $requests = @()
            for ($i = 0; $i -lt $chunk.Count; $i++) {
                $requests += @{
                    id     = [string]$i
                    method = 'GET'
                    url    = "/groups/$($chunk[$i].Id)/members/microsoft.graph.group?`$select=id&`$top=999"
                }
            }

            $responses = Invoke-GraphBatch -Requests $requests

            for ($i = 0; $i -lt $chunk.Count; $i++) {
                $parent = $chunk[$i]
                $response = $responses[[string]$i]
                $childIds = @()
                $ok = $false

                if ($response -and [int]$response.status -eq 200 -and $response.body) {
                    $ok = $true
                    foreach ($member in $response.body.value) { $childIds += $member.id }

                    $nextLink = $response.body.'@odata.nextLink'
                    while ($nextLink) {
                        try {
                            $nextPage = Invoke-GraphGetWithRetry -Uri $nextLink
                            foreach ($member in $nextPage.value) { $childIds += $member.id }
                            $nextLink = $nextPage.'@odata.nextLink'
                        }
                        catch {
                            $ok = $false
                            $nextLink = $null
                            Write-Log "Could not page members of group '$($parent.DisplayName)' ($($parent.Id)): $($_.Exception.Message)" -Level WARNING
                        }
                    }
                }
                else {
                    $reason = ''
                    if ($response -and $response.body -and $response.body.error) { $reason = [string]$response.body.error.message }
                    Write-Log "Could not retrieve members for group '$($parent.DisplayName)' ($($parent.Id)): $reason" -Level WARNING
                }

                $script:NestingMap[$parent.Id] = @($childIds | Where-Object { $groupIds.ContainsKey($_) })

                if ($ok) {
                    $script:NestingAnalyzed[$parent.Id] = $true
                } else {
                    $failedCount++
                }
            }

            $done = $offset + $chunk.Count
            $statusLabel.Text = "Analyzing nesting $done of $total - ETA $(Format-Eta -Watch $watch -Done $done -Total $total)"
            [System.Windows.Forms.Application]::DoEvents()
        }

        if ($failedCount -gt 0) {
            Write-Log "Nesting could not be determined for $failedCount group(s); it will be retried on demand when they are selected for an operation." -Level WARNING
        }
    }
    finally {
        Stop-Busy
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

    $nestedCount = @($script:NestingMap.Values | Where-Object { $_.Count -gt 0 }).Count
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
            $childGroup = if ($script:GroupById.ContainsKey($childId)) { $script:GroupById[$childId] } else { $null }
            if ($childGroup -and -not $childGroup.IsCloudManaged) {
                $unconverted += $childGroup
            }
        }
    }

    return $unconverted
}

function Update-GroupFilters {
    $script:GroupTypeFilter = if ($comboGroupType.SelectedItem) { [string]$comboGroupType.SelectedItem } else { "All" }
    $script:SearchText = $textSearch.Text.Trim()
    $script:CurrentPage = 1

    Update-GroupGrid

    $total = @($script:AllGroupsUnfiltered).Count
    if ($total -gt 0) {
        $filterActive = ($script:GroupTypeFilter -ne "All") -or ($script:SearchText.Length -gt 0) -or ($null -ne $script:CsvFilterIds)
        if ($filterActive) {
            $statusLabel.Text = "Filters active - Showing $(@($script:AllGroups).Count) of $total groups"
        } else {
            $statusLabel.Text = "Showing all $total groups"
        }
    }
}

function Get-GroupsFromCsvList {
    param(
        [Parameter(Mandatory=$true)]
        [array]$Identifiers
    )

    Write-Log "Targeted CSV discovery started for $($Identifiers.Count) unique entries..."
    $scope = Get-SearchScope

    Start-Busy
    try {
        $select = 'id,displayName,mail,mailEnabled,securityEnabled,groupTypes,onPremisesSyncEnabled,onPremisesSecurityIdentifier'
        $foundById = @{}
        $pendingById = @{}
        $labelById = @{}
        $notFound = New-Object System.Collections.Generic.List[string]
        $ambiguous = New-Object System.Collections.Generic.List[string]
        $notSynced = New-Object System.Collections.Generic.List[string]
        $outOfScope = New-Object System.Collections.Generic.List[string]

        $batchSize = 20
        $total = $Identifiers.Count
        $watch = [System.Diagnostics.Stopwatch]::StartNew()

        # Phase 1: targeted lookup of the CSV entries in Graph through JSON batching
        for ($offset = 0; $offset -lt $total; $offset += $batchSize) {
            if ($script:CancelRequested) {
                Write-Log "CSV discovery cancelled by user." -Level WARNING
                return $null
            }

            $last = [Math]::Min($offset + $batchSize, $total) - 1
            $chunk = @($Identifiers[$offset..$last])
            $requests = @()

            for ($i = 0; $i -lt $chunk.Count; $i++) {
                $val = $chunk[$i].Trim()
                $escaped = $val.Replace("'", "''")
                $guidRef = [Guid]::Empty

                if ([Guid]::TryParse($val, [ref]$guidRef)) {
                    $url = "/groups/$($guidRef.ToString())?`$select=$select"
                }
                elseif ($val -match '@') {
                    $url = "/groups?`$filter=mail%20eq%20'$([Uri]::EscapeDataString($escaped))'&`$select=$select"
                }
                else {
                    $url = "/groups?`$filter=displayName%20eq%20'$([Uri]::EscapeDataString($escaped))'&`$select=$select"
                }

                $requests += @{
                    id     = [string]$i
                    method = 'GET'
                    url    = $url
                }
            }

            $responses = Invoke-GraphBatch -Requests $requests

            for ($i = 0; $i -lt $chunk.Count; $i++) {
                $val = $chunk[$i]
                $resp = $responses[[string]$i]
                $matchedItems = @()

                if ($resp -and [int]$resp.status -eq 200 -and $resp.body) {
                    if ($null -ne $resp.body.value) {
                        $matchedItems = @($resp.body.value)
                    } elseif ($null -ne $resp.body.id) {
                        $matchedItems = @($resp.body)
                    }
                }

                if ($matchedItems.Count -eq 0) {
                    [void]$notFound.Add($val)
                }
                elseif ($matchedItems.Count -gt 1) {
                    [void]$ambiguous.Add($val)
                }
                else {
                    $g = $matchedItems[0]
                    $kind = Get-GroupKind -Item $g
                    $labelById[$g.id] = "$val ($($g.displayName))"

                    if ($null -eq $kind) {
                        [void]$notSynced.Add($labelById[$g.id])
                    }
                    elseif (-not (Test-GroupInSearchScope -Kind $kind -Scope $scope)) {
                        [void]$outOfScope.Add($labelById[$g.id])
                    }
                    elseif ($g.onPremisesSyncEnabled -eq $true) {
                        $foundById[$g.id] = $g
                    }
                    else {
                        # Not synced: keep it only if its SOA turns out to be converted (onPremisesSyncEnabled is null after conversion)
                        $pendingById[$g.id] = $g
                    }
                }
            }

            $done = $offset + $chunk.Count
            $statusLabel.Text = "Querying CSV groups $done of $total - ETA $(Format-Eta -Watch $watch -Done $done -Total $total)"
            [System.Windows.Forms.Application]::DoEvents()
        }

        # Phase 2: SOA status for the matched groups only
        $candidates = @($foundById.Values) + @($pendingById.Values)
        $soaMap = Get-SoaStatusMap -Items $candidates -RetrySingle
        if ($null -eq $soaMap) {
            Write-Log "CSV SOA lookup cancelled by user." -Level WARNING
            return $null
        }

        $groupResults = New-Object System.Collections.Generic.List[object]
        foreach ($item in $foundById.Values) {
            $groupResults.Add((New-GroupRecord -Item $item -IsCloudManaged $soaMap[$item.id]))
        }
        foreach ($item in $pendingById.Values) {
            if ($soaMap[$item.id] -eq $true -and $soaMap[$item.id] -isnot [string]) {
                $groupResults.Add((New-GroupRecord -Item $item -IsCloudManaged $true))
            } else {
                [void]$notSynced.Add($labelById[$item.id])
            }
        }

        foreach ($entry in $outOfScope) { Write-Log "CSV discovery: out of search scope, skipped: $entry" -Level WARNING }
        foreach ($entry in $notSynced) { Write-Log "CSV discovery: cloud-only or Microsoft 365 group, skipped: $entry" -Level WARNING }

        return [PSCustomObject]@{
            Groups     = $groupResults.ToArray()
            NotFound   = $notFound.ToArray()
            Ambiguous  = $ambiguous.ToArray()
            NotSynced  = $notSynced.ToArray()
            OutOfScope = $outOfScope.ToArray()
        }
    }
    catch {
        if ($script:CancelRequested) {
            Write-Log "CSV discovery cancelled by user: $($_.Exception.Message)" -Level WARNING
            return $null
        }
        throw
    }
    finally {
        Stop-Busy
    }
}

function Import-GroupFilterFromCsv {
    $context = $null
    try { $context = Get-MgContext } catch { $context = $null }

    if (-not $context) {
        [System.Windows.Forms.MessageBox]::Show(
            "You are not connected to Microsoft Graph.`n`nClick 'Connect to Graph' first, then import the CSV.",
            "Not Connected",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = "Select a CSV file with the groups to load/filter"
    $dialog.Filter = "CSV files (*.csv;*.txt)|*.csv;*.txt|All files (*.*)|*.*"
    $dialog.InitialDirectory = $script:ScriptPath

    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        return
    }

    $file = $dialog.FileName
    Write-Log "Importing group list from CSV: $file"

    try {
        $rawLines = @(Get-Content -Path $file -Encoding UTF8 -ErrorAction Stop | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

        if ($rawLines.Count -eq 0) {
            Write-Log "CSV file is empty: $file" -Level WARNING
            [System.Windows.Forms.MessageBox]::Show("The selected file is empty.", "Import CSV", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $delimiter = if ($rawLines[0] -match ';' -and $rawLines[0] -notmatch ',') { ';' } else { ',' }
        $headers = @($rawLines[0] -split [regex]::Escape($delimiter) | ForEach-Object { $_.Trim().Trim('"').Trim() })
        $columnOrder = @('ObjectId','Id','GroupId','Mail','Email','EmailAddress','PrimarySmtpAddress','DisplayName','Name','GroupName','Value')
        $hasHeader = (@($headers | Where-Object { $columnOrder -contains $_ }).Count -gt 0)

        if ($hasHeader) {
            $rows = @(Import-Csv -Path $file -Delimiter $delimiter -Encoding UTF8)
        } else {
            $rows = @($rawLines | ForEach-Object { [PSCustomObject]@{ Value = (($_ -split [regex]::Escape($delimiter))[0]).Trim().Trim('"').Trim() } })
        }

        # Groups already loaded: let the user choose between a fresh targeted discovery and filtering the current list
        $loadFromGraph = $true
        if (@($script:AllGroupsUnfiltered).Count -gt 0) {
            $choice = [System.Windows.Forms.MessageBox]::Show(
                "$(@($script:AllGroupsUnfiltered).Count) groups are already loaded.`n`nYes = load the groups in the CSV fresh from Microsoft Graph (replaces the current list)`nNo = filter the current list to the groups in the CSV`nCancel = do nothing",
                "Import CSV",
                [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )

            if ($choice -eq [System.Windows.Forms.DialogResult]::Cancel) { return }
            $loadFromGraph = ($choice -eq [System.Windows.Forms.DialogResult]::Yes)
        }

        # Case A: targeted discovery in Graph for the groups in the CSV only
        if ($loadFromGraph) {
            $identifiers = New-Object System.Collections.Generic.List[string]
            $seen = @{}

            foreach ($row in $rows) {
                $val = $null
                foreach ($col in $columnOrder) {
                    $prop = $row.PSObject.Properties[$col]
                    if ($null -ne $prop -and -not [string]::IsNullOrWhiteSpace("$($prop.Value)")) {
                        $val = "$($prop.Value)".Trim()
                        break
                    }
                }
                if ($val -and -not $seen.ContainsKey($val.ToLower())) {
                    $seen[$val.ToLower()] = $true
                    [void]$identifiers.Add($val)
                }
            }

            if ($identifiers.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("No valid group identifiers found in the CSV.", "Import CSV", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }

            $discovery = Get-GroupsFromCsvList -Identifiers $identifiers.ToArray()
            if ($null -eq $discovery) { return }

            $loadedGroups = @($discovery.Groups)
            if ($loadedGroups.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show(
                    "None of the entries in the CSV matched an on-premises synced or converted group in the search scope.`n`nThe current list has not been changed.",
                    "Import CSV",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
                return
            }

            Reset-LoadedGroups

            $script:GroupById = @{}
            foreach ($g in $loadedGroups) { $script:GroupById[$g.Id] = $g }

            if ($checkboxSkipNesting.Checked) {
                Reset-NestingMap -Groups $loadedGroups
            } else {
                Build-NestingMap -Groups $loadedGroups
            }

            $script:CsvFilterIds = $null
            $script:CsvFilterName = [System.IO.Path]::GetFileName($file)
            $labelCsv.Text = "CSV: $($script:CsvFilterName) ($($loadedGroups.Count) loaded)"
            $script:CurrentPage = 1
            Update-GroupGrid -Groups $loadedGroups
            $statusLabel.Text = "CSV discovery complete - $($loadedGroups.Count) groups loaded"

            if ($discovery.NotFound.Count -gt 0 -or $discovery.Ambiguous.Count -gt 0 -or $discovery.NotSynced.Count -gt 0 -or $discovery.OutOfScope.Count -gt 0) {
                $summary = "Loaded $($loadedGroups.Count) of $($identifiers.Count) unique entries."
                if ($discovery.NotFound.Count -gt 0) {
                    $summary += "`n`nNot found ($($discovery.NotFound.Count)):`n" + (($discovery.NotFound | Select-Object -First 10) -join "`n")
                }
                if ($discovery.Ambiguous.Count -gt 0) {
                    $summary += "`n`nAmbiguous DisplayName ($($discovery.Ambiguous.Count)):`n" + (($discovery.Ambiguous | Select-Object -First 10) -join "`n")
                }
                if ($discovery.NotSynced.Count -gt 0) {
                    $summary += "`n`nSkipped (cloud-only / M365 group) ($($discovery.NotSynced.Count)):`n" + (($discovery.NotSynced | Select-Object -First 10) -join "`n")
                }
                if ($discovery.OutOfScope.Count -gt 0) {
                    $summary += "`n`nOut of search scope ($($discovery.OutOfScope.Count)):`n" + (($discovery.OutOfScope | Select-Object -First 10) -join "`n")
                }
                $summary += "`n`nFull details are in the log file."
                [System.Windows.Forms.MessageBox]::Show($summary, "CSV Discovery Summary", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            return
        }

        # Case B: groups already loaded in memory -> filter the existing list to the groups in the CSV
        $byId = @{}
        $byMail = @{}
        $byName = @{}
        foreach ($g in $script:AllGroupsUnfiltered) {
            $byId["$($g.Id)".ToLower()] = $g
            if ($g.Mail) { $byMail["$($g.Mail)".ToLower()] = $g }
            if ($g.DisplayName) {
                $nameKey = "$($g.DisplayName)".ToLower()
                if (-not $byName.ContainsKey($nameKey)) { $byName[$nameKey] = @() }
                $byName[$nameKey] += $g
            }
        }

        $matched = @{}
        $notFound = @()
        $ambiguous = @()

        foreach ($row in $rows) {
            $resolved = $null
            $firstValue = $null
            $isAmbiguous = $false

            foreach ($col in $columnOrder) {
                $prop = $row.PSObject.Properties[$col]
                if ($null -eq $prop) { continue }
                $value = "$($prop.Value)".Trim()
                if ([string]::IsNullOrEmpty($value)) { continue }
                if ($null -eq $firstValue) { $firstValue = $value }
                $key = $value.ToLower()

                if ($byId.ContainsKey($key)) { $resolved = $byId[$key]; break }
                if ($byMail.ContainsKey($key)) { $resolved = $byMail[$key]; break }
                if ($byName.ContainsKey($key)) {
                    if ($byName[$key].Count -eq 1) { $resolved = $byName[$key][0]; break }
                    else { $isAmbiguous = $true }
                }
            }

            if ($resolved) {
                $matched[$resolved.Id] = $true
            } elseif ($isAmbiguous) {
                $ambiguous += $firstValue
            } elseif ($null -ne $firstValue) {
                $notFound += $firstValue
            }
        }

        Write-Log "CSV import: $($rows.Count) row(s), $($matched.Count) group(s) matched, $($notFound.Count) not found, $($ambiguous.Count) ambiguous."
        foreach ($item in $notFound) { Write-Log "CSV import: not found in loaded groups: $item" -Level WARNING }
        foreach ($item in $ambiguous) { Write-Log "CSV import: display name matches more than one group, skipped (use ObjectId or Mail): $item" -Level WARNING }

        if ($matched.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "None of the entries in the CSV matched a loaded group.`n`nThe current filter has not been changed.`n`nSupported columns: ObjectId, Mail, DisplayName (or a plain single-column list).",
                "Import CSV",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }

        $script:CsvFilterIds = $matched
        $script:CsvFilterName = [System.IO.Path]::GetFileName($file)
        $labelCsv.Text = "CSV: $($script:CsvFilterName) ($($matched.Count) groups)"

        Update-GroupFilters

        if ($notFound.Count -gt 0 -or $ambiguous.Count -gt 0) {
            $summary = "Matched $($matched.Count) of $($rows.Count) entries."
            if ($notFound.Count -gt 0) {
                $summary += "`n`nNot found ($($notFound.Count)):`n" + (($notFound | Select-Object -First 10) -join "`n")
                if ($notFound.Count -gt 10) { $summary += "`n..." }
            }
            if ($ambiguous.Count -gt 0) {
                $summary += "`n`nAmbiguous display name, skipped ($($ambiguous.Count)):`n" + (($ambiguous | Select-Object -First 10) -join "`n")
                if ($ambiguous.Count -gt 10) { $summary += "`n..." }
            }
            $summary += "`n`nFull details are in the log file."

            [System.Windows.Forms.MessageBox]::Show($summary, "Import CSV", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    }
    catch {
        Write-Log "Failed to import CSV: $($_.Exception.Message)" -Level ERROR
        [System.Windows.Forms.MessageBox]::Show(
        "Failed to import the CSV file.`n`nError: $($_.Exception.Message)",
        "Import Failed",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error)
        }
}

function Export-GroupsToCsv {
    $groups = @($script:AllGroups)

    if ($groups.Count -eq 0) {
        Write-Log "CSV export requested but no groups are loaded." -Level WARNING

        [System.Windows.Forms.MessageBox]::Show(
            "There are no groups to export.`n`nConnect to Microsoft Graph and load the group list first.",
            "Nothing to Export",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )

        return
    }

    $csvFile = Join-Path $script:ScriptPath "GroupSOAExport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

    Write-Log "Exporting $($groups.Count) group(s) to CSV: $csvFile"

    try {
        $exportData = foreach ($group in $groups) {
            $cloudManagedStatus = if ($group.IsCloudManaged -is [string] -and $group.IsCloudManaged -eq "Unknown") {
                "Unknown"
            } elseif ($group.IsCloudManaged -eq $true) {
                "True"
            } else {
                "False"
            }

            [PSCustomObject]@{
                DisplayName     = $group.DisplayName
                Mail            = $group.Mail
                GroupType       = $group.GroupType
                IsCloudManaged  = $cloudManagedStatus
                NestingDepth    = if ($script:NestingDepth.ContainsKey($group.Id)) { $script:NestingDepth[$group.Id] } else { 0 }
                SecurityEnabled = $group.SecurityEnabled
                MailEnabled     = $group.MailEnabled
                ObjectId        = $group.Id
            }
        }

        $exportData | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8 -ErrorAction Stop

        Write-Log "Successfully exported $($groups.Count) group(s) to $csvFile"

        $statusLabel.Text = "Exported $($groups.Count) group(s) to CSV"

        $result = [System.Windows.Forms.MessageBox]::Show(
            "Exported $($groups.Count) group(s) to:`n`n$csvFile`n`nDo you want to open the file now?",
            "Export Complete",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )

        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            try {
                Start-Process -FilePath $csvFile -ErrorAction Stop
            }
            catch {
                Write-Log "Could not open exported CSV file: $($_.Exception.Message)" -Level WARNING
            }
        }
    }
    catch {
        Write-Log "Failed to export groups to CSV: $($_.Exception.Message)" -Level ERROR

        [System.Windows.Forms.MessageBox]::Show(
            "Failed to export the group list to CSV.`n`nError: $($_.Exception.Message)",
            "Export Failed",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

# ============================================================
# MAIN SCRIPT - GUI SETUP
# ============================================================

Write-Log "========================================" -Level INFO
Write-Log "Group SOA Conversion Tool Started" -Level INFO
Write-Log "Target TenantId: $($script:TenantId)" -Level INFO

# Avoid ThreadPool starvation when thousands of sequential Graph calls are made from the GUI thread
try {
    $minWorker = 0
    $minIo = 0
    [System.Threading.ThreadPool]::GetMinThreads([ref]$minWorker, [ref]$minIo)
    if ([System.Threading.ThreadPool]::SetMinThreads([Math]::Max($minWorker, 100), [Math]::Max($minIo, 100))) {
        Write-Log "ThreadPool minimum threads set to $([Math]::Max($minWorker, 100)) (worker) / $([Math]::Max($minIo, 100)) (IO)."
    }
}
catch {
    Write-Log "Could not adjust ThreadPool minimum threads: $($_.Exception.Message)" -Level WARNING
}
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
$labelDescription.Text = "Convert Group Source of Authority for on-premises groups (Security Groups, Mail-Enabled Security Groups & Distribution Groups)"
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

        $statusLabel.Text = "Connected - click 'Search Groups' to load the groups"
    }
    else {
        $buttonConnect.Enabled = $true
    }
})
$form.Controls.Add($buttonConnect)

$buttonRefresh = New-Object System.Windows.Forms.Button
$buttonRefresh.Location = New-Object System.Drawing.Point(205, 105)
$buttonRefresh.Size = New-Object System.Drawing.Size(150, 40)
$buttonRefresh.Text = "Search Groups"
$buttonRefresh.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonRefresh.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#E1E1E1")
$buttonRefresh.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
$buttonRefresh.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$buttonRefresh.FlatAppearance.BorderSize = 0
$buttonRefresh.Cursor = [System.Windows.Forms.Cursors]::Hand
$buttonRefresh.Enabled = $false
$buttonRefresh.Add_Click({
    Start-Busy
    try {
        $groups = Get-ExchangeGroups

        if ($groups) {
            $script:GroupById = @{}
            foreach ($loadedGroup in $groups) { $script:GroupById[$loadedGroup.Id] = $loadedGroup }

            if ($checkboxSkipNesting.Checked) {
                Reset-NestingMap -Groups $groups
                Write-Log "Nesting analysis skipped on Search: it will run on demand for the groups selected for an operation."
            }
            else {
                Build-NestingMap -Groups $groups
            }

            $script:CurrentPage = 1
            Update-GroupGrid -Groups $groups

            $statusLabel.Text = "Search complete - $(@($groups).Count) groups loaded (synced and converted)"
            if ($script:CancelRequested) {
                $statusLabel.Text += " (nesting analysis cancelled, it will be completed on demand)"
            }
        }
        else {
            $script:NestingMap = @{}
            $script:NestingDepth = @{}
            $script:NestingAnalyzed = @{}
            $script:GroupById = @{}
            $script:CurrentPage = 1
            Update-GroupGrid -Groups @()

            if ($script:CancelRequested) {
                $statusLabel.Text = "Search cancelled"
            } else {
                $statusLabel.Text = "Search complete - no groups found (check the log if an error was reported)"
            }
        }
    }
    finally {
        Stop-Busy
    }
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

        Reset-LoadedGroups

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


# --- Export Button ---
$buttonExportCsv = New-Object System.Windows.Forms.Button
$buttonExportCsv.Location = New-Object System.Drawing.Point(565, 105)
$buttonExportCsv.Size = New-Object System.Drawing.Size(170, 40)
$buttonExportCsv.Text = "Export List to CSV"
$buttonExportCsv.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonExportCsv.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#E1E1E1")
$buttonExportCsv.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
$buttonExportCsv.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$buttonExportCsv.FlatAppearance.BorderSize = 0
$buttonExportCsv.Cursor = [System.Windows.Forms.Cursors]::Hand
$buttonExportCsv.Add_Click({
    Export-GroupsToCsv
})
$form.Controls.Add($buttonExportCsv)

# --- Exit Button ---
$buttonExit = New-Object System.Windows.Forms.Button
$buttonExit.Location = New-Object System.Drawing.Point(975, 105)
$buttonExit.Size = New-Object System.Drawing.Size(105, 40)
$buttonExit.Text = "Exit"
$buttonExit.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonExit.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#E1E1E1")
$buttonExit.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
$buttonExit.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$buttonExit.FlatAppearance.BorderSize = 0
$buttonExit.Cursor = [System.Windows.Forms.Cursors]::Hand
$buttonExit.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$buttonExit.Add_Click({
    # Closing the form triggers the FormClosing handler (Graph disconnect + logging)
    $form.Close()
})
$form.Controls.Add($buttonExit)

# --- Search scope (applied when Search is run): none ticked = all synced groups ---
function New-TypeCheckBox {
    param([string]$Text, [int]$X, [int]$Width)

    $box = New-Object System.Windows.Forms.CheckBox
    $box.Location = New-Object System.Drawing.Point($X, 150)
    $box.Size = New-Object System.Drawing.Size($Width, 26)
    $box.Text = $Text
    $box.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $box.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
    $box.Cursor = [System.Windows.Forms.Cursors]::Hand
    $box.Checked = $false
    return $box
}

$checkboxTypeSecurity = New-TypeCheckBox -Text "Security groups" -X 20 -Width 125
$checkboxTypeMailSecurity = New-TypeCheckBox -Text "Mail-enabled security" -X 150 -Width 160
$checkboxTypeDistribution = New-TypeCheckBox -Text "Distribution lists" -X 315 -Width 135
$form.Controls.Add($checkboxTypeSecurity)
$form.Controls.Add($checkboxTypeMailSecurity)
$form.Controls.Add($checkboxTypeDistribution)

$toolTip = New-Object System.Windows.Forms.ToolTip
foreach ($typeBox in @($checkboxTypeSecurity, $checkboxTypeMailSecurity, $checkboxTypeDistribution)) {
    $toolTip.SetToolTip($typeBox, "Search scope, applied when you click Search Groups. None ticked = all synced groups; tick one or two types to narrow the search.")
}

$checkboxSkipNesting = New-Object System.Windows.Forms.CheckBox
$checkboxSkipNesting.Location = New-Object System.Drawing.Point(470, 150)
$checkboxSkipNesting.Size = New-Object System.Drawing.Size(495, 26)
$checkboxSkipNesting.Text = "Skip nesting analysis on Search (done on demand when converting)"
$checkboxSkipNesting.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$checkboxSkipNesting.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
$checkboxSkipNesting.Cursor = [System.Windows.Forms.Cursors]::Hand
$checkboxSkipNesting.Checked = $false
$form.Controls.Add($checkboxSkipNesting)

$buttonCancel = New-Object System.Windows.Forms.Button
$buttonCancel.Location = New-Object System.Drawing.Point(975, 148)
$buttonCancel.Size = New-Object System.Drawing.Size(105, 28)
$buttonCancel.Text = "Cancel job"
$buttonCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonCancel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#E1E1E1")
$buttonCancel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#A4262C")
$buttonCancel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$buttonCancel.FlatAppearance.BorderSize = 0
$buttonCancel.Cursor = [System.Windows.Forms.Cursors]::Hand
$buttonCancel.Enabled = $false
$buttonCancel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$buttonCancel.Add_Click({
    $script:CancelRequested = $true
    $statusLabel.Text = "Cancelling - the current batch will finish first..."
    Write-Log "Cancel requested by user." -Level WARNING
})
$form.Controls.Add($buttonCancel)

# --- Filter Row ---
$labelGroupType = New-Object System.Windows.Forms.Label
$labelGroupType.Location = New-Object System.Drawing.Point(20, 183)
$labelGroupType.Size = New-Object System.Drawing.Size(75, 24)
$labelGroupType.Text = "Group type:"
$labelGroupType.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$labelGroupType.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
$labelGroupType.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$form.Controls.Add($labelGroupType)

$comboGroupType = New-Object System.Windows.Forms.ComboBox
$comboGroupType.Location = New-Object System.Drawing.Point(95, 183)
$comboGroupType.Size = New-Object System.Drawing.Size(210, 24)
$comboGroupType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$comboGroupType.Font = New-Object System.Drawing.Font("Segoe UI", 9)
foreach ($typeName in @("All", "Security Group", "Mail-Enabled Security Group", "Distribution Group")) {
    [void]$comboGroupType.Items.Add($typeName)
}
$comboGroupType.SelectedIndex = 0
$comboGroupType.Add_SelectedIndexChanged({
    Update-GroupFilters
})
$form.Controls.Add($comboGroupType)

$labelSearch = New-Object System.Windows.Forms.Label
$labelSearch.Location = New-Object System.Drawing.Point(320, 183)
$labelSearch.Size = New-Object System.Drawing.Size(55, 24)
$labelSearch.Text = "Search:"
$labelSearch.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$labelSearch.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
$labelSearch.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$form.Controls.Add($labelSearch)

$textSearch = New-Object System.Windows.Forms.TextBox
$textSearch.Location = New-Object System.Drawing.Point(375, 183)
$textSearch.Size = New-Object System.Drawing.Size(200, 24)
$textSearch.Font = New-Object System.Drawing.Font("Segoe UI", 9)
# Debounce: apply the search 400 ms after typing stops (or immediately on Enter) instead of on every keystroke
$searchTimer = New-Object System.Windows.Forms.Timer
$searchTimer.Interval = 400
$searchTimer.Add_Tick({
    $searchTimer.Stop()
    Update-GroupFilters
})
$textSearch.Add_TextChanged({
    $searchTimer.Stop()
    $searchTimer.Start()
})
$textSearch.Add_KeyDown({
    param($eventSender, $e)

    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $e.SuppressKeyPress = $true
        $searchTimer.Stop()
        Update-GroupFilters
    }
})
$form.Controls.Add($textSearch)

$buttonImportCsv = New-Object System.Windows.Forms.Button
$buttonImportCsv.Location = New-Object System.Drawing.Point(590, 180)
$buttonImportCsv.Size = New-Object System.Drawing.Size(130, 30)
$buttonImportCsv.Text = "Import CSV..."
$buttonImportCsv.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonImportCsv.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#E1E1E1")
$buttonImportCsv.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
$buttonImportCsv.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$buttonImportCsv.FlatAppearance.BorderSize = 0
$buttonImportCsv.Cursor = [System.Windows.Forms.Cursors]::Hand
$buttonImportCsv.Add_Click({
    Import-GroupFilterFromCsv
})
$form.Controls.Add($buttonImportCsv)

$buttonClearFilters = New-Object System.Windows.Forms.Button
$buttonClearFilters.Location = New-Object System.Drawing.Point(730, 180)
$buttonClearFilters.Size = New-Object System.Drawing.Size(130, 30)
$buttonClearFilters.Text = "Clear Filters"
$buttonClearFilters.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonClearFilters.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#E1E1E1")
$buttonClearFilters.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
$buttonClearFilters.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$buttonClearFilters.FlatAppearance.BorderSize = 0
$buttonClearFilters.Cursor = [System.Windows.Forms.Cursors]::Hand
$buttonClearFilters.Add_Click({
    $script:CsvFilterIds = $null
    $script:CsvFilterName = $null
    $labelCsv.Text = "CSV: none"
    $textSearch.Text = ""
    $comboGroupType.SelectedIndex = 0
    Update-GroupFilters
})
$form.Controls.Add($buttonClearFilters)

$labelCsv = New-Object System.Windows.Forms.Label
$labelCsv.Location = New-Object System.Drawing.Point(870, 183)
$labelCsv.Size = New-Object System.Drawing.Size(210, 24)
$labelCsv.Text = "CSV: none"
$labelCsv.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$labelCsv.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#605E5C")
$labelCsv.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$labelCsv.AutoEllipsis = $true
$labelCsv.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($labelCsv)

# --- DataGridView ---
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(20, 220)
$dataGridView.Size = New-Object System.Drawing.Size(1060, 305)
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

$dataGridView.Add_ColumnHeaderMouseClick({
    param($eventSender, $e)

    $columnName = $eventSender.Columns[$e.ColumnIndex].Name

    # Don't sort on ObjectId column
    if ($columnName -eq "ObjectId") {
        return
    }

    # Toggle sort direction if clicking the same column
    if ($script:SortColumn -eq $columnName) {
        $script:SortAscending = -not $script:SortAscending
    } else {
        $script:SortColumn = $columnName
        $script:SortAscending = $true
    }

    # Reset to first page when sorting
    $script:CurrentPage = 1

    # Update the grid with sorted data
    Update-GroupGrid
})

# A manual row click never silently means "all pages": untick "Select all" first
$dataGridView.Add_CellMouseDown({
    param($eventSender, $e)

    if ($e.RowIndex -ge 0 -and $checkboxAllFiltered.Checked) {
        $checkboxAllFiltered.Checked = $false
        $statusLabel.Text = "'Select all' was cleared because a row was clicked"
    }
})

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
$pageInfo.Size = New-Object System.Drawing.Size(380, 30)
$pageInfo.Text = "No groups to display"
$pageInfo.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$pageInfo.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#605E5C")
$pageInfo.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$pageInfo.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($pageInfo)

$checkboxAllFiltered = New-Object System.Windows.Forms.CheckBox
$checkboxAllFiltered.Location = New-Object System.Drawing.Point(525, 570)
$checkboxAllFiltered.Size = New-Object System.Drawing.Size(260, 30)
$checkboxAllFiltered.Text = "Select all groups in the list (all pages)"
$checkboxAllFiltered.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$checkboxAllFiltered.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#A4262C")
$checkboxAllFiltered.Cursor = [System.Windows.Forms.Cursors]::Hand
$checkboxAllFiltered.Checked = $false
$checkboxAllFiltered.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$checkboxAllFiltered.Add_CheckedChanged({
    if ($checkboxAllFiltered.Checked) {
        $dataGridView.SelectAll()
    } else {
        $dataGridView.ClearSelection()
    }
})
$form.Controls.Add($checkboxAllFiltered)

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
            "The required permission 'Group-OnPremisesSyncBehavior.ReadWrite.All' is not consented.`n`nPlease reconnect to Graph to grant the required permission.",
            "Permission Required",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    if ($dataGridView.SelectedRows.Count -eq 0 -and -not $checkboxAllFiltered.Checked) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select at least one group from the list.",
            "No Group Selected",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $selectedGroups = @()
    if ($checkboxAllFiltered.Checked) {
        # "Select all": every group in the list (all pages) that is in the right SOA state
        $selectedGroups = @($script:AllGroups | Where-Object { $_.IsCloudManaged -eq $false })
    }
    else {
        foreach ($selectedRow in $dataGridView.SelectedRows) {
            $groupId = $selectedRow.Cells["ObjectId"].Value
            $group = $script:GroupById[$groupId]
            if ($group) {
                $selectedGroups += $group
            }
        }
    }

    if ($selectedGroups.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "There are no eligible groups to process.",
            "No Eligible Groups",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return
    }

    # Make sure the nesting of the groups to process is known (needed when nesting analysis was skipped or incomplete)
    $notAnalyzed = @($selectedGroups | Where-Object { -not $script:NestingAnalyzed.ContainsKey($_.Id) })
    if ($notAnalyzed.Count -gt 0) {
        Build-NestingMap -Groups $script:AllGroupsUnfiltered -ParentGroups $notAnalyzed

        if ($script:CancelRequested) {
            $statusLabel.Text = "Nesting analysis cancelled - operation aborted"
            return
        }

        $stillMissing = @($selectedGroups | Where-Object { -not $script:NestingAnalyzed.ContainsKey($_.Id) })
        if ($stillMissing.Count -gt 0) {
            $continueAnyway = [System.Windows.Forms.MessageBox]::Show(
                "Nesting could not be determined for $($stillMissing.Count) selected group(s) (see log).`n`nThe processing order may not be correct for these groups.`n`nDo you want to continue anyway?",
                "Nesting Incomplete",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )

            if ($continueAnyway -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        }
    }

    $selectedIdSet = @{}
    foreach ($selectedGroup in $selectedGroups) { $selectedIdSet[$selectedGroup.Id] = $true }

    $allWarnings = @()
    foreach ($group in $selectedGroups) {
        $unconvertedChildren = Get-UnconvertedChildren -GroupId $group.Id
        $childrenNotInSelection = @($unconvertedChildren | Where-Object { -not $selectedIdSet.ContainsKey($_.Id) })

        if ($childrenNotInSelection.Count -gt 0) {
            $childNames = ($childrenNotInSelection | ForEach-Object { $_.DisplayName }) -join ", "
            $allWarnings += "Group '$($group.DisplayName)' has unconverted nested children not in your selection: $childNames"
        }
    }

    if ($allWarnings.Count -gt 0) {
        foreach ($warningLine in $allWarnings) { Write-Log $warningLine -Level WARNING }

        $warningText = "WARNING: Nested group ordering issue detected! ($($allWarnings.Count) group(s) affected)`n`n"
        $warningText += (@($allWarnings | Select-Object -First 10) -join "`n`n")
        if ($allWarnings.Count -gt 10) { $warningText += "`n`n...and $($allWarnings.Count - 10) more (see log)." }
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

    $countText = if ($checkboxAllFiltered.Checked) { "ALL $selectedCount eligible groups in the filtered list (all pages)" } else { "$selectedCount groups" }

    if ($selectedCount -eq 1 -and -not $checkboxAllFiltered.Checked) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Are you sure you want to convert group '$($sortedGroups[0].DisplayName)' to Cloud Managed?",
            "Confirm Conversion",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
    } else {
        $result = Show-ConfirmationDialog -Title "Confirm Conversion" `
            -Message "Are you sure you want to convert $countText to Cloud Managed?`n`nGroups will be converted in bottom-up order (deepest nested first)." `
            -GroupList $groupList `
            -Icon "Question"
    }

    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $runResults = @(Invoke-SoaBatchChange -Groups $sortedGroups -CloudManaged $true)
        $successCount = @($runResults | Where-Object { $_.Result -eq 'Success' }).Count
        $failCount = $runResults.Count - $successCount
        $reportFile = Export-SoaReport -Results $runResults -Operation 'Convert'

        $summaryMessage = "Batch conversion completed.`n`nSuccessful: $successCount`nFailed / unverified / skipped: $failCount`n`nReport: $reportFile"

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
            "The required permission 'Group-OnPremisesSyncBehavior.ReadWrite.All' is not consented.`n`nPlease reconnect to Graph to grant the required permission.",
            "Permission Required",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    if ($dataGridView.SelectedRows.Count -eq 0 -and -not $checkboxAllFiltered.Checked) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select at least one group from the list.",
            "No Group Selected",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $selectedGroups = @()
    if ($checkboxAllFiltered.Checked) {
        # "Select all": every group in the list (all pages) that is in the right SOA state
        $selectedGroups = @($script:AllGroups | Where-Object { $_.IsCloudManaged -eq $true })
    }
    else {
        foreach ($selectedRow in $dataGridView.SelectedRows) {
            $groupId = $selectedRow.Cells["ObjectId"].Value
            $group = $script:GroupById[$groupId]
            if ($group) {
                $selectedGroups += $group
            }
        }
    }

    if ($selectedGroups.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "There are no eligible groups to process.",
            "No Eligible Groups",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return
    }

    # Make sure the nesting of the groups to process is known (needed when nesting analysis was skipped or incomplete)
    $notAnalyzed = @($selectedGroups | Where-Object { -not $script:NestingAnalyzed.ContainsKey($_.Id) })
    if ($notAnalyzed.Count -gt 0) {
        Build-NestingMap -Groups $script:AllGroupsUnfiltered -ParentGroups $notAnalyzed

        if ($script:CancelRequested) {
            $statusLabel.Text = "Nesting analysis cancelled - operation aborted"
            return
        }

        $stillMissing = @($selectedGroups | Where-Object { -not $script:NestingAnalyzed.ContainsKey($_.Id) })
        if ($stillMissing.Count -gt 0) {
            $continueAnyway = [System.Windows.Forms.MessageBox]::Show(
                "Nesting could not be determined for $($stillMissing.Count) selected group(s) (see log).`n`nThe processing order may not be correct for these groups.`n`nDo you want to continue anyway?",
                "Nesting Incomplete",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )

            if ($continueAnyway -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        }
    }

    $sortedGroups = $selectedGroups | Sort-Object { 
        if ($script:NestingDepth.ContainsKey($_.Id)) { $script:NestingDepth[$_.Id] } else { 0 }
    } -Descending

    $selectedCount = $sortedGroups.Count
    $groupList = ($sortedGroups | ForEach-Object { "$($_.DisplayName) (Depth: $(if ($script:NestingDepth.ContainsKey($_.Id)) { $script:NestingDepth[$_.Id] } else { 0 }))" }) -join "`n"

    $countText = if ($checkboxAllFiltered.Checked) { "ALL $selectedCount eligible groups in the filtered list (all pages)" } else { "$selectedCount groups" }

    if ($selectedCount -eq 1 -and -not $checkboxAllFiltered.Checked) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Are you sure you want to roll back group '$($sortedGroups[0].DisplayName)' to On-Premises Managed?`n`nIMPORTANT: Make sure to remove cloud users from the group and remove the group from access packages before rolling back.",
            "Confirm Rollback",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
    } else {
        $result = Show-ConfirmationDialog -Title "Confirm Rollback" `
            -Message "Are you sure you want to roll back $countText to On-Premises Managed?`n`nGroups will be rolled back in top-down order (parents first).`n`nIMPORTANT: Make sure to remove cloud users from the groups and remove the groups from access packages before rolling back." `
            -GroupList $groupList `
            -Icon "Warning"
    }

    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $runResults = @(Invoke-SoaBatchChange -Groups $sortedGroups -CloudManaged $false)
        $successCount = @($runResults | Where-Object { $_.Result -eq 'Success' }).Count
        $failCount = $runResults.Count - $successCount
        $reportFile = Export-SoaReport -Results $runResults -Operation 'Rollback'

        $summaryMessage = "Batch rollback completed.`n`nSuccessful: $successCount`nFailed / unverified / skipped: $failCount`n`nReport: $reportFile`n`nNote: The rollback is only complete after the next scheduled or forced run of Connect Sync."

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
    $pageInfo.Size = New-Object System.Drawing.Size(($form.ClientSize.Width - 570), 30)
    $buttonNextPage.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 120), $paginationY)
    $checkboxAllFiltered.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 395), $paginationY)

    $gridHeight = $paginationY - 230
    $dataGridView.Size = New-Object System.Drawing.Size(($form.ClientSize.Width - 40), $gridHeight)
})

# --- Form Closing Handler ---
$form.Add_FormClosing({
    param($eventSender, $e)

    if ($script:BusyDepth -gt 0) {
        $script:CancelRequested = $true
        $e.Cancel = $true
        [System.Windows.Forms.MessageBox]::Show(
            "A job is still running. A cancel has been requested: wait for the current batch to finish, then close the window again.",
            "Job Running",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

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
