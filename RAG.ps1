# [REMOTE AUTOMATION GUI - PART 1/5]
# Requires: Windows PowerShell 5+ (or PowerShell 7+ with Windows forms), SQLServer module (for Invoke-Sqlcmd)
# Run: powershell.exe -ExecutionPolicy Bypass -File .\RemoteAutomationGUI.ps1

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 2. DPI / visual style setup — must come *before* any form or control is created
try {
    if ([enum]::GetNames([System.Windows.Forms.HighDpiMode])) {
        [System.Windows.Forms.Application]::SetHighDpiMode(
            [System.Windows.Forms.HighDpiMode]::SystemAware
        )
    }
} catch { }

[System.Windows.Forms.Application]::EnableVisualStyles()
try { [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false) } catch { }


# === CONFIG ===
try {
    if ($MyInvocation.MyCommand.Path) {
        # Running as script
        $basePath = Split-Path -Parent $MyInvocation.MyCommand.Path
    } elseif ([System.Reflection.Assembly]::GetEntryAssembly()) {
        # Running as compiled EXE
        $basePath = Split-Path -Parent ([System.Reflection.Assembly]::GetEntryAssembly().Location)
    } else {
        throw "Unable to detect base path"
    }

    $basePath = (Resolve-Path $basePath).Path
}
catch {
    [System.Windows.Forms.MessageBox]::Show(
        "Failed to resolve base path: $($_.Exception.Message)",
        "Startup Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit
}

$configFile   = Join-Path $env:APPDATA "RemoteAutomationGUI.config"
$secureFolder = Join-Path $basePath "Secure"
$fixedCredPath= Join-Path $secureFolder "adminCred.xml"
$scriptFolder = Join-Path $basePath "Scripts"
$listFolder   = Join-Path $basePath "Lists"
$logFolder    = Join-Path $basePath "Logs"
$iconPath     = Join-Path $basePath "RAG.ico"


# Ensure base folders exist (best-effort)
foreach ($p in @($basePath,$secureFolder,$scriptFolder,$listFolder,$logFolder)) {
    try { if (-not (Test-Path $p)) { New-Item -Path $p -ItemType Directory -Force | Out-Null } } catch {}
}

# Global transfer queue state
$script:ActiveCopyJobs = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:CopyTimer = $null

$script:ActiveScriptJobs = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:ScriptTimer = $null

$script:ActiveRunJobs = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:RunTimer      = $null

$script:ActiveSqlJobs = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:SqlTimer      = $null

# === RUNTIME STATE ===
$script:StopRequested     = $false
$script:SqlFullOutput     = @{}
$script:LastCheckedIndex  = $null
$script:IsBulkChecking    = $false

# === DEFAULTS ===
$lastFile         = ""
$lastServerList   = ""
$lastRemoteFolder = "C$\Kit"
$lastScriptFolder = $scriptFolder
$selectedScriptPath = ""
$scriptPathMap       = @{}  # name -> full path
$serverListPathMap   = @{}  # file name -> full path

# === LOAD CONFIG (if present) ===
if (Test-Path $configFile) {
    try {
        $cfg = Get-Content $configFile -Raw | ConvertFrom-Json
        if ($cfg.LastFile)         { $lastFile = $cfg.LastFile }
        if ($cfg.LastServerList)   { $lastServerList = $cfg.LastServerList }
        if ($cfg.LastRemoteFolder) { $lastRemoteFolder = $cfg.LastRemoteFolder }
        if ($cfg.LastScriptFolder) { $lastScriptFolder = $cfg.LastScriptFolder }
    } catch {}
}

# === FORM ===
$form = New-Object Windows.Forms.Form
$form.Text          = "Remote Automation GUI"
$form.Size          = New-Object Drawing.Size(1200,750)
$form.StartPosition = "CenterScreen"
$form.Icon = New-Object System.Drawing.Icon($iconPath)
if ($Host.Name -notlike "*ISE*") {
    $form.Add_FormClosed({ [System.Windows.Forms.Application]::Exit() })
} else {
    $form.Add_FormClosed({ $form.Dispose() })
}


# === TOP BAR: File selection and credentials ===
$lblFile               = New-Object Windows.Forms.Label
$lblFile.Text          = "Selected File:"
$lblFile.Location      = New-Object Drawing.Point(20,20)
$lblFile.Size          = New-Object Drawing.Size(100,20)

$txtFile               = New-Object Windows.Forms.TextBox
$txtFile.Location      = New-Object Drawing.Point(120,20)
$txtFile.Size          = New-Object Drawing.Size(430,20)
$txtFile.Text          = $lastFile
$txtFile.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

$btnBrowse             = New-Object Windows.Forms.Button
$btnBrowse.Text        = "Browse file"
$btnBrowse.Location    = New-Object Drawing.Point(560,18)
$btnBrowse.Size        = New-Object Drawing.Size(100,25)

$btnCreateCred         = New-Object Windows.Forms.Button
$btnCreateCred.Text    = "Create Admin Cred..."
$btnCreateCred.Location= New-Object Drawing.Point(850,18)
$btnCreateCred.Size    = New-Object Drawing.Size(140,25)

$btnCreateSqlCred = New-Object System.Windows.Forms.Button
$btnCreateSqlCred.Text = "Create SQL Credential"
$btnCreateSqlCred.Location = '1000, 18'  # adjust as needed
$btnCreateSqlCred.Size = '140, 25'

# === Encrypt VNC Passwords Button ===
if ($null -ne $btnEncryptVnc) {
    # Remove any existing event handlers to avoid double prompt
    $btnEncryptVnc.Remove_Click($btnEncryptVnc_Click)
    $form.Controls.Remove($btnEncryptVnc)
}

$btnEncryptVnc = New-Object Windows.Forms.Button
$btnEncryptVnc.Text = "Create VNC Pass..."
$btnEncryptVnc.Location = New-Object Drawing.Point(700, 18)
$btnEncryptVnc.Size = New-Object Drawing.Size(140,25)


# === Remote path / folder field ===
$lblRemoteFolder       = New-Object Windows.Forms.Label
$lblRemoteFolder.Text  = "Remote Folder:"
$lblRemoteFolder.Location = New-Object Drawing.Point(20,52)
$lblRemoteFolder.Size  = New-Object Drawing.Size(120,20)

$txtRemoteFolder       = New-Object Windows.Forms.TextBox
$txtRemoteFolder.Location = New-Object Drawing.Point(150,52)
$txtRemoteFolder.Size  = New-Object Drawing.Size(400,20)
$txtRemoteFolder.Text  = $lastRemoteFolder
$txtRemoteFolder.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

$chkCreateFolder = New-Object System.Windows.Forms.CheckBox
$chkCreateFolder.Text = "Create folder if missing"
$chkCreateFolder.Checked = $false  # ✅ Default OFF
$chkCreateFolder.AutoSize = $true
$chkCreateFolder.Location = New-Object System.Drawing.Point(22, 75)
# Match font to rest of GUI
$chkCreateFolder.Font = New-Object System.Drawing.Font("Segoe UI", 6)


# === Server list dropdown + browse ===
$lblServers            = New-Object Windows.Forms.Label
$lblServers.Text       = "Server List:"
$lblServers.Location   = New-Object Drawing.Point(20,90)
$lblServers.Size       = New-Object Drawing.Size(100,20)

$cmbServers            = New-Object Windows.Forms.ComboBox
$cmbServers.Location   = New-Object Drawing.Point(120,90)
$cmbServers.Size       = New-Object Drawing.Size(300,18)
$cmbServers.DropDownStyle = "DropDownList"


$btnBrowseServers          = New-Object Windows.Forms.Button
$btnBrowseServers.Text     = "Browse List..."
$btnBrowseServers.Location = New-Object Drawing.Point(450,90)
$btnBrowseServers.Size     = New-Object Drawing.Size(100,25)

# === Filter + list view with checkboxes ===
$lblFilter            = New-Object Windows.Forms.Label
$lblFilter.Text       = "Filter:"
$lblFilter.Location   = New-Object Drawing.Point(20,122)
$lblFilter.Size       = New-Object Drawing.Size(50,20)

$txtFilter            = New-Object Windows.Forms.TextBox
$txtFilter.Location   = New-Object Drawing.Point(80,122)
$txtFilter.Size       = New-Object Drawing.Size(160,20)
$txtfilter.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

$lblServersbar = New-Object System.Windows.Forms.Label
$lblServersbar.Text = "Servers"
$lblServersbar.Location = New-Object Drawing.Point(20,150)
$lblServersbar.AutoSize = $true
$lblServersbar.Font = New-Object Drawing.Font("Segoe UI",10)


$listViewServers                  = New-Object System.Windows.Forms.ListView
$listViewServers.Location         = New-Object Drawing.Point(20,170)
$listViewServers.Size             = New-Object Drawing.Size(400,370)
$listViewServers.View             = 'Details'
$listViewServers.CheckBoxes       = $true
$listViewServers.FullRowSelect    = $true
$listViewServers.MultiSelect      = $true
[void]$listViewServers.Columns.Add("Server", 380)
$listViewServers.HeaderStyle = 'None'
$listViewServers.BorderStyle = 'FixedSingle'


# Shift+Click range support is in Part 2

# === Selection helpers ===
$btnSaveSelection = New-Object System.Windows.Forms.Button
$btnSaveSelection.Text = "Save Selection"
$btnSaveSelection.Size = New-Object Drawing.Size(100, 25)
$btnSaveSelection.Location = New-Object Drawing.Point(320, 545)

$btnResetFilter          = New-Object System.Windows.Forms.Button
$btnResetFilter.Text     = "Clear"
$btnResetFilter.Location = New-Object Drawing.Point(246, 120)
$btnResetFilter.Size     = New-Object Drawing.Size(50, 25)
$btnResetFilter.Add_Click({
    $script:IsBulkChecking = $true
    try {
        if ($script:CheckedServers) {
            $script:CheckedServers.Clear()
        }
        foreach ($item in $listViewServers.Items) {
            $item.Checked = $false
        }
        $script:LastCheckedIndex = $null
    }
    finally {
        $script:IsBulkChecking = $false
    }

    if ($txtFilter.Text -ne "") {
        $txtFilter.Text = ""
    }
    $txtFilter.Focus()
})

$btnCheckAll = New-Object System.Windows.Forms.Button
$btnCheckAll.Text = "Check All"
$btnCheckAll.Size = New-Object Drawing.Size(60, 25)
$btnCheckAll.Location = New-Object Drawing.Point(296, 120)

$btnUncheckAll = New-Object System.Windows.Forms.Button
$btnUncheckAll.Text = "Uncheck All"
$btnUncheckAll.Size = New-Object Drawing.Size(65, 25)
$btnUncheckAll.Location = New-Object Drawing.Point(356, 120)

# === Right side actions ===
$btnCopy                  = New-Object Windows.Forms.Button
$btnCopy.Text             = "Copy"
$btnCopy.Location         = New-Object Drawing.Point(560,50)
$btnCopy.Size             = New-Object Drawing.Size(100,30)

$btnRunFile               = New-Object Windows.Forms.Button
$btnRunFile.Text          = "Run file"
$btnRunFile.Location      = New-Object Drawing.Point(670,50)
$btnRunFile.Size          = New-Object Drawing.Size(120,30)

$btnExecSQL               = New-Object Windows.Forms.Button
$btnExecSQL.Text          = "Execute SQL"
$btnExecSQL.Location      = New-Object Drawing.Point(800,50)
$btnExecSQL.Size          = New-Object Drawing.Size(120,30)

$btnStop          = New-Object Windows.Forms.Button
$btnStop.Text     = "STOP"
$btnStop.Location = New-Object Drawing.Point(1072, 85)
$btnStop.Size     = New-Object Drawing.Size(70, 30)
$btnStop.BackColor = [System.Drawing.Color]::Crimson
$btnStop.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$btnStop.ForeColor = [System.Drawing.Color]::Black

$btnPing                  = New-Object Windows.Forms.Button
$btnPing.Text             = "Ping"
$btnPing.Location         = New-Object Drawing.Point(450, 645)
$btnPing.Size             = New-Object Drawing.Size(70, 25)

$btnConnectVNC            = New-Object Windows.Forms.Button
$btnConnectVNC.Text       = "Connect VNC"
$btnConnectVNC.Location   = New-Object Drawing.Point(530,645)
$btnConnectVNC.Size       = New-Object Drawing.Size(100,25)

$btnRDP                   = New-Object Windows.Forms.Button
$btnRDP.Text              = "Connect RDP"
$btnRDP.Location          = New-Object Drawing.Point(640,645)
$btnRDP.Size              = New-Object Drawing.Size(100,25)

$btnOpenC                 = New-Object Windows.Forms.Button
$btnOpenC.Text            = "Open C$"
$btnOpenC.Location        = New-Object Drawing.Point(750,645)
$btnOpenC.Size            = New-Object Drawing.Size(70,25)

$btnOpenD                 = New-Object Windows.Forms.Button
$btnOpenD.Text            = "Open D$"
$btnOpenD.Location        = New-Object Drawing.Point(830,645)
$btnOpenD.Size            = New-Object Drawing.Size(70,25)

# === Progress + grid ===
$progress                        = New-Object Windows.Forms.ProgressBar
$progress.Location               = New-Object Drawing.Point(450,123)
$progress.Size                   = New-Object Drawing.Size(690,20)
$progress.Minimum                = 0

$grid                            = New-Object Windows.Forms.DataGridView
$grid.Location                   = New-Object Drawing.Point(450,150)
$grid.Size                       = New-Object Drawing.Size(690,455)
$grid.ColumnCount                = 3
$grid.Columns[0].Name            = "Server"
$grid.Columns[1].Name            = "Status"
$grid.Columns[2].Name            = "Output"
$grid.AutoSizeColumnsMode        = "Fill"
$grid.Columns[0].FillWeight      = 20
$grid.Columns[1].FillWeight      = 20
$grid.Columns[2].FillWeight      = 60
$grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 8)

# === Scripts dropdown with autocomplete ===
$comboScripts                    = New-Object Windows.Forms.ComboBox
$comboScripts.Location           = New-Object Drawing.Point(20,580)
$comboScripts.Size               = New-Object Drawing.Size(400,20)
$comboScripts.DropDownStyle      = "DropDown"
$comboScripts.AutoCompleteMode   = "SuggestAppend"
$comboScripts.AutoCompleteSource = "ListItems"

$btnBrowseScript                 = New-Object Windows.Forms.Button
$btnBrowseScript.Text            = "Browse Scripts..."
$btnBrowseScript.Location        = New-Object Drawing.Point(350,610)
$btnBrowseScript.Size            = New-Object Drawing.Size(70,25)

$btnRunScripts          = New-Object System.Windows.Forms.Button
$btnRunScripts.Text     = "Run Script"
$btnRunScripts.Location = New-Object System.Drawing.Point(20, 645)
$btnRunScripts.Size     = New-Object System.Drawing.Size(100, 30)
$btnRunScripts.ForeColor = [System.Drawing.Color]::White
$btnRunScripts.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup

$btnFolderCommands = New-Object System.Windows.Forms.Button
$btnFolderCommands.Text = "Commands"
$btnFolderCommands.Size = New-Object System.Drawing.Size(78, 24)
$btnFolderCommands.Location = New-Object System.Drawing.Point(20, 610)

$btnFolderDiagnostic = New-Object System.Windows.Forms.Button
$btnFolderDiagnostic.Text = "Diagnostic"
$btnFolderDiagnostic.Size = New-Object System.Drawing.Size(78, 24)
$btnFolderDiagnostic.Location = New-Object System.Drawing.Point(100, 610)

$btnFolderSql = New-Object System.Windows.Forms.Button
$btnFolderSql.Text = "SQL"
$btnFolderSql.Size = New-Object System.Drawing.Size(40, 24)
$btnFolderSql.Location = New-Object System.Drawing.Point(180, 610)


# 1. Create the status label first
$statusLabel = New-Object Windows.Forms.Label
$statusLabel.ForeColor = [System.Drawing.Color]::LimeGreen
$statusLabel.Text = "Ready"
$statusLabel.Location = New-Object Drawing.Point(730, 95)
$statusLabel.Size = New-Object Drawing.Size(300, 30)
$statusLabel.Font = New-Object Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)


# 2. Create the checkbox
$chkDryRun = New-Object System.Windows.Forms.CheckBox
$chkDryRun.Text = "Dry Run (simulate only)"
$chkDryRun.Checked = $false
$chkDryRun.AutoSize = $true
$chkDryRun.Location = New-Object System.Drawing.Point(22, 685)
$chkDryRun.Font = New-Object System.Drawing.Font("Segoe UI", 6)

function Update-DryRunVisuals {
    if ($chkDryRun.Checked) {
        $statusLabel.Text         = "🧪 Dry Run Mode Active"
        $statusLabel.ForeColor    = [System.Drawing.Color]::DarkOrange
        $btnRunScripts.BackColor  = [System.Drawing.Color]::DarkOrange
        $btnRunScripts.ForeColor  = [System.Drawing.Color]::Black
    } else {
        $statusLabel.Text         = "Ready"
        $statusLabel.ForeColor    = [System.Drawing.Color]::LimeGreen
        $btnRunScripts.BackColor  = [System.Drawing.Color]::LimeGreen
        $btnRunScripts.ForeColor  = [System.Drawing.Color]::Black
    }
}

$chkDryRun.Add_CheckedChanged({
    Update-DryRunVisuals
})

# Set initial state when the GUI loads
Update-DryRunVisuals

# 4. Add tooltip
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.SetToolTip($chkDryRun, "Simulates actions without executing them. Safe for testing.")

# 5. Set dry run ON by default
$chkDryRun.Checked = $true  # This will now trigger the CheckedChanged event and update the label10, [System.Drawing.FontStyle]::Bold)

$btnSaveLog = New-Object Windows.Forms.Button
$btnSaveLog.Text     = "Save Log"
$btnSaveLog.Location = New-Object Drawing.Point(1040,610)
$btnSaveLog.Size     = New-Object Drawing.Size(100,25)

# Main button
$btnCreateCredentials = New-Object System.Windows.Forms.Button
$btnCreateCredentials.Text = "Create Credentials ▼"
$btnCreateCredentials.Size = New-Object Drawing.Size(140,30)
$btnCreateCredentials.Location = New-Object Drawing.Point(1002,18)

# Context menu
$credMenu = New-Object System.Windows.Forms.ContextMenuStrip

$adminItem = $credMenu.Items.Add("Admin Credentials")
$sqlItem   = $credMenu.Items.Add("SQL Credentials")
$vncItem   = $credMenu.Items.Add("VNC Passwords")


# Show menu when button clicked
$btnCreateCredentials.Add_Click({
    $credMenu.Show(
        $btnCreateCredentials,
        0,
        $btnCreateCredentials.Height
    )
})


# === Global button font ===
$buttonFont = New-Object System.Drawing.Font("Segoe UI", 10)
@(
    $btnCopy, $btnRunFile, $btnRunScripts, $btnExecSQL, $btnStop,
    $btnConnectVNC, $btnPing, $btnRDP, $btnOpenC, $btnOpenD,
    $btnSaveSelection, $btnCheckAll, $btnUncheckAll,
    $btnBrowse, $btnBrowseServers, $btnBrowseScript, $btnCreateCred, $btnCreateSqlCred,
    $btnEncryptVnc, $btnResetFilter, $btnCreateCredentials, $credMenu, $btnSaveLog,
    $btnFolderCommands, $btnFolderDiagnostic, $btnFolderSql, $btnThemeToggle
) | Where-Object { $null -ne $_ } | ForEach-Object { $_.Font = $buttonFont }

# === Unicode font for controls with special characters ===
$unicodeFont = New-Object System.Drawing.Font("Segoe UI", 11)

$listViewServers.Font     = $unicodeFont
$cmbServers.Font          = $unicodeFont
$txtFilter.Font           = $unicodeFont
$txtFile.Font             = $unicodeFont
$txtRemoteFolder.Font     = $unicodeFont
$comboScripts.Font = New-Object System.Drawing.Font("Segoe UI", 11)


# [REMOTE AUTOMATION GUI - PART 2/5]
# === HELPERS & FUNCTIONS ===

function Save-Config {
    $cfg = @{
        LastFile         = $txtFile.Text
        LastServerList   = $cmbServers.SelectedItem
        LastRemoteFolder = $txtRemoteFolder.Text
        LastScriptFolder = $lastScriptFolder
    }
    try {
        $cfg | ConvertTo-Json | Set-Content -Path $configFile -Encoding UTF8
    } catch {}
}

# === DPAPI encryption/decryption helpers ===
Add-Type -AssemblyName System.Security


function Protect-String {
    param(
        [Parameter(Mandatory)][string]$PlainText,
        [ValidateSet("CurrentUser","LocalMachine")][string]$Scope = "CurrentUser"
    )
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($PlainText)
    $scopeEnum = if ($Scope -eq "CurrentUser") {
        [System.Security.Cryptography.DataProtectionScope]::CurrentUser
    } else {
        [System.Security.Cryptography.DataProtectionScope]::LocalMachine
    }
    $enc = [System.Security.Cryptography.ProtectedData]::Protect($bytes, $null, $scopeEnum)
    return [Convert]::ToBase64String($enc)
}

function Unprotect-String {
    param(
        [Parameter(Mandatory)][string]$CipherText,
        [ValidateSet("CurrentUser","LocalMachine")][string]$Scope = "CurrentUser"
    )
    $bytes = [Convert]::FromBase64String($CipherText)
    $scopeEnum = if ($Scope -eq "CurrentUser") {
        [System.Security.Cryptography.DataProtectionScope]::CurrentUser
    } else {
        [System.Security.Cryptography.DataProtectionScope]::LocalMachine
    }
    $plainBytes = [System.Security.Cryptography.ProtectedData]::Unprotect($bytes, $null, $scopeEnum)
    return [System.Text.Encoding]::UTF8.GetString($plainBytes)
}



# Create credential: export with DPAPI and tighten ACL best-effort
function Save-FixedCredential {
    param ([System.Management.Automation.PSCredential]$Cred)
    try {
        if (-not (Test-Path $secureFolder)) {
            New-Item -Path $secureFolder -ItemType Directory -Force | Out-Null
        }
        $Cred | Export-Clixml -Path $fixedCredPath -Force

        try {
            $file = Get-Item -LiteralPath $fixedCredPath
            $acl  = $file.GetAccessControl('Access')
            $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

            foreach ($existingRule in $acl.Access) {
                try { $acl.RemoveAccessRule($existingRule) | Out-Null } catch {}
            }

            $inherit = [System.Security.AccessControl.InheritanceFlags]::None
            $propag  = [System.Security.AccessControl.PropagationFlags]::None
            $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                $currentUser,
                [System.Security.AccessControl.FileSystemRights]::FullControl,
                $inherit, $propag,
                [System.Security.AccessControl.AccessControlType]::Allow
            )
            $acl.AddAccessRule($rule)
            $file.SetAccessControl($acl)
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Saved credential but failed to set strict ACLs.`n`nReason: $($_.Exception.Message)",
                "Warning",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
        }
        return $true
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to save credentials:`n$($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return $false
    }
}

# Update a grid row by server name
function Update-GridStatus {
    param (
        [string]$Server,
        [string]$Status,
        [System.Drawing.Color]$Color,
        [string]$Output = ""
    )

    foreach ($row in $grid.Rows) {

        if ($row.Cells[0].Value -eq $Server) {

            $row.Cells[1].Value = $Status
            $row.Cells[2].Value = $Output

            # Color only Status column
            $row.Cells[1].Style.ForeColor = $Color

            # Output always theme color
            $row.Cells[2].Style.ForeColor = $script:ThemeFgPrimary

            break
        }
    }
}

# Get checked servers (normalized, unique)
function Get-SelectedServers {
    return $listViewServers.Items |
        Where-Object { $_.Checked } |
        ForEach-Object {
            $ipMatch = [regex]::Match($_.Text, '\b\d{1,3}(\.\d{1,3}){3}\b')
            if ($ipMatch.Success) { $ipMatch.Value.ToUpper() }
        } |
        Sort-Object -Unique
}

# Get checked display lines (full annotated text)
function Get-SelectedDisplayLines {
    return $listViewServers.Items |
        Where-Object { $_.Checked } |
        ForEach-Object { $_.Text.Trim() } |
        Sort-Object -Unique
}


if (-not $script:CheckedServers) {
    $script:CheckedServers = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
}

function Load-ServerList {
    param(
        [Parameter(Mandatory = $false)]
        $ServerListName
    )

    if (-not $ServerListName) { return }

    $listName = [string]$ServerListName
    if ([string]::IsNullOrWhiteSpace($listName)) { return }

    if (-not ($script:CheckedServers -is [System.Collections.Generic.HashSet[string]])) {
        $script:CheckedServers = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
    }

    # Reset tracked selections and last index when switching server lists
    $script:CheckedServers.Clear()
    $script:LastCheckedIndex = $null

    $appRoot = if ($PSScriptRoot) {
        $PSScriptRoot
    } elseif ($MyInvocation.MyCommand.Path) {
        Split-Path -Parent $MyInvocation.MyCommand.Path
    } else {
        (Get-Location).Path
    }

    $listsDir = Join-Path $appRoot "Lists"
    $servers  = @()

    if (Test-Path -LiteralPath $listsDir) {
        $candidateFile = Join-Path $listsDir $listName
        if (-not (Test-Path -LiteralPath $candidateFile) -and (Test-Path -LiteralPath "$candidateFile.txt")) {
            $candidateFile = "$candidateFile.txt"
        }
        if (Test-Path -LiteralPath $candidateFile) {
            $servers = @(
                Get-Content -LiteralPath $candidateFile |
                    ForEach-Object { $_.Trim() } |
                    Where-Object { $_ -and -not $_.StartsWith('#') }
            )
        }
    }

    $script:CachedServers = $servers

    $script:IsBulkChecking = $true
    $script:IsFilteringServers = $true
    $listViewServers.BeginUpdate()
    try {
        $listViewServers.Items.Clear()

        $filter = if ($txtFilter) { $txtFilter.Text.Trim() } else { "" }
        $filtered = if ([string]::IsNullOrWhiteSpace($filter)) {
            $script:CachedServers
        } else {
            $script:CachedServers | Where-Object { $_ -like "*$filter*" }
        }

        $newItems = foreach ($srv in $filtered) {
            New-Object System.Windows.Forms.ListViewItem($srv)
        }
        if ($newItems) {
            $listViewServers.Items.AddRange($newItems)
        }
    }
    finally {
        $listViewServers.EndUpdate()
        $script:IsFilteringServers = $false
        $script:IsBulkChecking = $false
    }
}


$listViewServers.Add_ItemCheck({
    param($sender, $e)
    if ($script:IsBulkChecking -or $script:IsFilteringServers) { return }

    if (-not ($script:CheckedServers -is [System.Collections.Generic.HashSet[string]])) {
        $script:CheckedServers = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
    }

    $shiftPressed = ([System.Windows.Forms.Control]::ModifierKeys -band [System.Windows.Forms.Keys]::Shift) -ne 0
    $ctrlPressed  = ([System.Windows.Forms.Control]::ModifierKeys -band [System.Windows.Forms.Keys]::Control) -ne 0
    $nowChecked   = ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked)
    $targetText   = $listViewServers.Items[$e.Index].Text

    $script:IsBulkChecking = $true
    try {
        if ($shiftPressed -and $null -ne $script:LastCheckedIndex) {
            $start = [Math]::Min($script:LastCheckedIndex, $e.Index)
            $end   = [Math]::Max($script:LastCheckedIndex, $e.Index)
            for ($i = $start; $i -le $end; $i++) {
                if ($i -ne $e.Index) {
                    $listViewServers.Items[$i].Checked = $nowChecked
                    $rowText = $listViewServers.Items[$i].Text
                    if ($nowChecked) { [void]$script:CheckedServers.Add($rowText) }
                    else { [void]$script:CheckedServers.Remove($rowText) }
                }
            }
        }
        elseif ($nowChecked -and -not $ctrlPressed) {
            $script:CheckedServers.Clear()
            for ($i = 0; $i -lt $listViewServers.Items.Count; $i++) {
                if ($i -ne $e.Index -and $listViewServers.Items[$i].Checked) {
                    $listViewServers.Items[$i].Checked = $false
                }
            }
        }

        if ($nowChecked) { [void]$script:CheckedServers.Add($targetText) }
        else { [void]$script:CheckedServers.Remove($targetText) }
    }
    finally {
        $script:IsBulkChecking = $false
    }

    $script:LastCheckedIndex = $e.Index

    if ($nowChecked -and -not $shiftPressed -and -not $ctrlPressed -and -not $script:IsFilteringServers -and
        $comboScripts -and $comboScripts.Visible -and $comboScripts.Enabled) {
        $comboScripts.Focus()
        $comboScripts.SelectAll()
    }
})



function Copy-ToServers($fileToCopy, $remoteFolderLive, $servers, $cred, $createFolderIfMissing) {
    # Stop any existing background timer if running
    if ($script:CopyTimer) {
        $script:CopyTimer.Stop()
        $script:CopyTimer.Dispose()
        $script:CopyTimer = $null
    }

    $script:ActiveCopyJobs.Clear()
    $grid.Rows.Clear()

    $fileName   = Split-Path $fileToCopy -Leaf
    $fileInfo   = Get-Item $fileToCopy
    $fileSizeMB = $fileInfo.Length / 1MB
    $threshold  = 10 # Megabyte limit for WinRM serialization

    # Pre-read file bytes if small to avoid reading repeatedly in jobs
    $fileBytes = $null
    if ($fileSizeMB -lt $threshold) {
        $fileBytes = [System.IO.File]::ReadAllBytes($fileToCopy)
    }

    # Populate grid with initial "Pending" rows
    foreach ($s in $servers) {
        [void]$grid.Rows.Add($s, "Pending", "")
    }

    $progress.Maximum = $servers.Count
    $progress.Value   = 0

    # Launch a background job for EACH server simultaneously (or sequentially in background)
    foreach ($srv in $servers) {
        $job = Start-Job -ScriptBlock {
            param($srv, $fileToCopy, $remoteFolderLive, $cred, $createFolderIfMissing, $fileSizeMB, $threshold, $fileBytes, $fileName)
            
            $result = [PSCustomObject]@{
                Server  = $srv
                Success = $false
                Message = ""
            }

            try {
                # Convert path format
                if ($remoteFolderLive -match '^([a-zA-Z])\$\\(.+)$') {
                    $drive = $matches[1]
                    $sub   = $matches[2]
                    $folderPath   = "$drive`:\$sub"
                    $uncSharePath = "\\$srv\${drive}`$\$sub"
                } elseif ($remoteFolderLive -match '^([a-zA-Z]):\\(.*)$') {
                    $drive = $matches[1]
                    $sub   = $matches[2]
                    $folderPath   = $remoteFolderLive
                    $uncSharePath = "\\$srv\${drive}`$\$sub"
                } else {
                    $folderPath = $remoteFolderLive
                    $uncSharePath = $null
                }

                # Create folder if missing via WinRM
                if ($createFolderIfMissing) {
                    Invoke-Command -ComputerName $srv -Credential $cred -ErrorAction Stop -ScriptBlock {
                        param($folderPath)
                        if (-not (Test-Path $folderPath)) {
                            New-Item -Path $folderPath -ItemType Directory -Force | Out-Null
                        }
                    } -ArgumentList $folderPath
                }

                # Transfer Logic
                if ($fileSizeMB -lt $threshold) {
                    Invoke-Command -ComputerName $srv -Credential $cred -ErrorAction Stop -ScriptBlock {
                        param($folderPath, $fileName, $fileBytes)
                        if (-not (Test-Path $folderPath)) {
                            throw "Destination folder '$folderPath' does not exist"
                        }
                        $destPath = Join-Path $folderPath $fileName
                        [System.IO.File]::WriteAllBytes($destPath, $fileBytes)
                    } -ArgumentList $folderPath, $fileName, $fileBytes

                    $result.Success = $true
                    $result.Message = "Copied $fileName via WinRM to $folderPath"
                } 
                else {
                    if (-not $uncSharePath) {
                        throw "Could not resolve UNC admin share path for: $remoteFolderLive"
                    }
                    if (-not (Test-Path $uncSharePath)) {
                        New-Item -Path $uncSharePath -ItemType Directory -Force | Out-Null
                    }
                    $destFilePath = Join-Path $uncSharePath $fileName
                    Copy-Item -Path $fileToCopy -Destination $destFilePath -Force

                    $result.Success = $true
                    $result.Message = "Copied $fileName asynchronously via SMB Share to $folderPath"
                }
            }
            catch {
                $result.Success = $false
                $result.Message = $_.Exception.Message
            }

            return $result
        } -ArgumentList $srv, $fileToCopy, $remoteFolderLive, $cred, $createFolderIfMissing, $fileSizeMB, $threshold, $fileBytes, $fileName

        # Store job reference alongside its target server name
        $script:ActiveCopyJobs.Add([PSCustomObject]@{ Server = $srv; Job = $job })
    }

    # Create a UI Timer to poll job progress asynchronously (ticks every 250ms)
    $script:CopyTimer = New-Object System.Windows.Forms.Timer
    $script:CopyTimer.Interval = 250

    $script:CopyTimer.Add_Tick({
        $allDone = $true

        foreach ($item in $script:ActiveCopyJobs) {
            $job = $item.Job
            $srv = $item.Server

            if ($null -eq $job) {
                continue
            }

            if ($job.State -eq 'Completed' -or $job.State -eq 'Failed' -or $job.State -eq 'Stopped') {
                # Retrieve result if not already processed
                $data = Receive-Job $job -ErrorAction SilentlyContinue
                Remove-Job $job -Force -ErrorAction SilentlyContinue

                # Mark as processed by clearing job reference
                $item.Job = $null

                # Update grid UI safely on main thread
                foreach ($row in $grid.Rows) {
                    if ($row.Cells[0].Value -eq $srv) {
                        if ($data -and $data.Success) {
                            $row.Cells[1].Value = "✅ Copied"
                            $row.Cells[1].Style.ForeColor = [System.Drawing.Color]::LimeGreen
                            $row.Cells[2].Value = $data.Message
                        } else {
                            $errMsg = if ($data -and $data.Message) { $data.Message } else { "Copy job ended with state '$($job.State)' and no result object." }
                            $row.Cells[1].Value = "❌ Copy failed"
                            $row.Cells[1].Style.ForeColor = [System.Drawing.Color]::Red
                            $row.Cells[2].Value = $errMsg
                        }
                        break
                    }
                }

                $progress.Value = [Math]::Min($progress.Value + 1, $progress.Maximum)
            }
            else {
                # Still running
                $allDone = $false
            }
        }

 # If all jobs finished, stop and cleanup timer
        if ($allDone) {
            $script:CopyTimer.Stop()
            $script:CopyTimer.Dispose()
            $script:CopyTimer = $null
            $script:ActiveCopyJobs.Clear()
            $statusLabel.Text = if ($script:StopRequested) { "Stopped by user" } else { "Copy Completed" }
        }
    })

    $script:CopyTimer.Start()
}

# Invoke a remote file on servers, route execution by extension
function Start-RunOnServersAsync {
    param(
        [string]$RemotePath,
        [string[]]$Servers,
        [System.Management.Automation.PSCredential]$Credential
    )

    if ($script:RunTimer) {
        $script:RunTimer.Stop()
        $script:RunTimer.Dispose()
        $script:RunTimer = $null
    }

    $script:ActiveRunJobs.Clear()
    $grid.Rows.Clear()
    foreach ($s in $Servers) {
        [void]$grid.Rows.Add($s, "Pending", "")
    }

    $progress.Maximum = $Servers.Count
    $progress.Value   = 0
    $statusLabel.Text = "Running file..."

    $extension = [System.IO.Path]::GetExtension($RemotePath).ToLower()

    foreach ($srv in $Servers) {
        $job = Start-Job -ScriptBlock {
            param($srv, $RemotePath, $extension, $Credential)

            $result = [PSCustomObject]@{
                Server  = $srv
                Success = $false
                Status  = "❌ Run failed"
                Details = ""
            }

            try {
                $out = Invoke-Command -ComputerName $srv -Credential $Credential -ErrorAction Stop -ScriptBlock {
                    param($batFile, $ext)
                    if (Test-Path $batFile) {
                        try {
                            switch ($ext) {
                                ".bat" { & cmd.exe /c "`"$batFile`"" 2>&1 }
                                ".cmd" { & cmd.exe /c "`"$batFile`"" 2>&1 }
                                ".ps1" { & powershell.exe -ExecutionPolicy Bypass -File "`"$batFile`"" 2>&1 }
                                ".vbs" { & cscript.exe //nologo "`"$batFile`"" 2>&1 }
                                ".exe" {
                                    try {
                                        $args = "/VERYSILENT /NORESTART /SUPPRESSMSGBOXES"
                                        Start-Process -FilePath $batFile -ArgumentList $args -Wait
                                        "✅ EXE launched successfully via Start-Process"
                                    } catch { "❌ EXE Start-Process failed: $($_.Exception.Message)" }
                                }
                                ".msi" {
                                    try {
                                        $args = "/i `"$batFile`" /quiet /norestart"
                                        Start-Process -FilePath "msiexec.exe" -ArgumentList $args -Wait
                                        "✅ MSI launched successfully via Start-Process"
                                    } catch { "❌ MSI Start-Process failed: $($_.Exception.Message)" }
                                }
                                default { "❌ Unsupported file type: $ext" }
                            }
                        } catch { "❌ Execution failed: $($_.Exception.Message)" }
                    } else {
                        "❌ File not found: $batFile"
                    }
                } -ArgumentList $RemotePath, $extension

                $outputText = ($out | Out-String).Trim()
                if ($outputText -match "❌|^File not found") {
                    $result.Success = $false
                    $result.Status  = "❌ Failed"
                    $result.Details = $outputText
                } else {
                    $result.Success = $true
                    $result.Status  = "✅ Success"
                    $result.Details = $outputText
                }
            }
            catch {
                $result.Success = $false
                $result.Status  = "❌ Run failed"
                $result.Details = $_.Exception.Message
            }

            return $result
        } -ArgumentList $srv, $RemotePath, $extension, $Credential

        $script:ActiveRunJobs.Add([PSCustomObject]@{ Server = $srv; Job = $job })
    }

    $script:RunTimer = New-Object System.Windows.Forms.Timer
    $script:RunTimer.Interval = 200

    $script:RunTimer.Add_Tick({
        $allDone = $true

        foreach ($item in $script:ActiveRunJobs) {
            $job = $item.Job
            $srv = $item.Server

            if ($null -eq $job) {
                continue
            }

            if ($job.State -eq 'Completed' -or $job.State -eq 'Failed' -or $job.State -eq 'Stopped') {
                $data = Receive-Job $job -ErrorAction SilentlyContinue
                Remove-Job $job -Force -ErrorAction SilentlyContinue
                $item.Job = $null

                $statusText = if ($data -and $data.Status) { $data.Status } else { "❌ Run failed" }
                $detailsMsg = if ($data -and $data.Details) { $data.Details } else { "No output returned or job stopped before returning result." }
                $color      = if ($statusText -like "✅*") { [System.Drawing.Color]::LimeGreen } else { [System.Drawing.Color]::Red }

                Update-GridStatus $srv $statusText $color $detailsMsg
                $progress.Value = [Math]::Min($progress.Value + 1, $progress.Maximum)
            }
            else {
                $allDone = $false
            }
        }

        if ($allDone) {
            $script:RunTimer.Stop()
            $script:RunTimer.Dispose()
            $script:RunTimer = $null
            $script:ActiveRunJobs.Clear()
            $btnRunFile.Enabled = $true
            $statusLabel.Text = if ($script:StopRequested) { "Stopped by user" } else { "Run Completed" }
        }
    })

    $script:RunTimer.Start()
}

function Run-ScriptOnServers {
    param (
        [string[]]$Servers,
        [string]$ScriptPath,
        [string]$ScriptText,
        [string]$FilePath,
        [bool]$UsesSqlAuth = $false,
        [string]$SqlUser = "",
        [string]$SqlPass = "",
        [bool]$IsDryRun = $false
    )

    if ($IsDryRun) {
        $simulated = Simulate-DryRun -Action "Would run script" `
            -DetailsTemplate "Simulated script:`n{target}`n--- on {{server}}" `
            -Servers $Servers -Target $ScriptText

        Show-DryRunSummary $simulated
        $statusLabel.Text = "Dry Run Completed"
        return
    }

    # Stop any existing script timer if already running
    if ($script:ScriptTimer) {
        $script:ScriptTimer.Stop()
        $script:ScriptTimer.Dispose()
        $script:ScriptTimer = $null
    }

    $script:ActiveScriptJobs.Clear()
    $grid.Rows.Clear()
    foreach ($s in $Servers) {
        [void]$grid.Rows.Add($s, "Pending", "")
    }

    $progress.Value = 0
    $progress.Maximum = $Servers.Count
    $statusLabel.Text = "Running Script..."

    # Launch background jobs for each server simultaneously
    foreach ($srv in $Servers) {
        $job = Start-Job -ScriptBlock {
            param($srv, $ScriptPath, $FilePath, $UsesSqlAuth, $SqlUser, $SqlPass)

            $resultObject = [PSCustomObject]@{
                Server  = $srv
                Status  = ""
                Output  = ""
                IsError = $false
            }

            try {
                if ($UsesSqlAuth) {
                    $res = & $ScriptPath -FilePath $FilePath -Server $srv -SqlUser $SqlUser -SqlPass $SqlPass -UseSqlAuth
                } else {
                    $res = & $ScriptPath -FilePath $FilePath -Server $srv
                }

                if ($res -and $res.Status) {

    $resultObject.Status = $res.Status
    $resultObject.Output = $res.Output

}
elseif ($res) {

    $resultObject.Status = "✅ Completed"
    $resultObject.Output = ($res | Out-String).Trim()

}
else {

    $resultObject.Status = "✅ Completed"
    $resultObject.Output = "Script completed successfully."
}
            } catch {
                $resultObject.Status = "❌ Script Error"
                $resultObject.Output = $_.Exception.Message
                $resultObject.IsError = $true
            }

            return $resultObject
        } -ArgumentList $srv, $ScriptPath, $FilePath, $UsesSqlAuth, $SqlUser, $SqlPass

        $script:ActiveScriptJobs.Add([PSCustomObject]@{ Server = $srv; Job = $job })
    }

    # Create a UI Timer to poll job completion safely in the background (ticks every 250ms)
    $script:ScriptTimer = New-Object System.Windows.Forms.Timer
    $script:ScriptTimer.Interval = 250

    $script:ScriptTimer.Add_Tick({
        $allDone = $true

        foreach ($item in $script:ActiveScriptJobs) {
            $job = $item.Job
            $srv = $item.Server

            if ($job -and ($job.State -eq 'Completed' -or $job.State -eq 'Failed' -or $job.State -eq 'Stopped')) {
                $data = Receive-Job $job -ErrorAction SilentlyContinue
                Remove-Job $job -Force -ErrorAction SilentlyContinue
                $item.Job = $null

                # Update grid UI safely on the main thread
                foreach ($row in $grid.Rows) {
                    if ($row.Cells[0].Value -eq $srv) {
                        $statusText = if ($data -and $data.Status) { $data.Status } else { "❌ Script Error" }
                        $outputMsg  = if ($data -and $data.Output) { $data.Output } else { "No output returned." }

                        $color = if ($statusText -like "✅*") { [System.Drawing.Color]::LimeGreen } else { [System.Drawing.Color]::Red }

                        $row.Cells[1].Value = $statusText
                        $row.Cells[1].Style.ForeColor = $color
                        $row.Cells[2].Value = $outputMsg
                        break
                    }
                }

                $progress.Value = [Math]::Min($progress.Value + 1, $progress.Maximum)
            }
            elseif ($null -ne $job) {
                $allDone = $false
            }
        }

        # Once all jobs finish, wrap up UI states and stop the timer
        if ($allDone) {
            $script:ScriptTimer.Stop()
            $script:ScriptTimer.Dispose()
            $script:ScriptTimer = $null
            $script:ActiveScriptJobs.Clear()
            $statusLabel.Text = if ($script:StopRequested) { "Stopped by user" } else { "Script Execution Completed" }
        }
    })

    $script:ScriptTimer.Start()
}

function Start-ExecuteSqlOnServersAsync {
    param (
        [string[]]$Servers,
        [string]$SqlFile,
        [string]$SqlText,
        [string]$SqlUser,
        [string]$SqlPass,
        [string]$LogFolder
    )

    $progress.Value   = 0
    $progress.Maximum = $Servers.Count
    $statusLabel.Text = "Executing SQL via .NET SqlClient..."

    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ApartmentState = 'STA'
    $runspace.ThreadOptions  = 'ReuseThread'
    $runspace.Open()

    $queue = [System.Collections.Concurrent.ConcurrentQueue[hashtable]]::new()
    $state = [hashtable]::Synchronized(@{
        Done = $false
    })

    $runspace.SessionStateProxy.SetVariable('Servers', $Servers)
    $runspace.SessionStateProxy.SetVariable('SqlText', $SqlText)
    $runspace.SessionStateProxy.SetVariable('SqlUser', $SqlUser)
    $runspace.SessionStateProxy.SetVariable('SqlPass', $SqlPass)
    $runspace.SessionStateProxy.SetVariable('LogFolder', $LogFolder)
    $runspace.SessionStateProxy.SetVariable('Queue', $queue)
    $runspace.SessionStateProxy.SetVariable('State', $state)
    $runspace.SessionStateProxy.SetVariable('StopFlagRef', ([ref]$script:StopRequested))

    $ps = [powershell]::Create()
    $ps.Runspace = $runspace

    [void]$ps.AddScript({
        Add-Type -AssemblyName System.Data
        $sqlBatches = $SqlText -split "(?m)^\s*GO\s*\r?$"

        foreach ($srv in $Servers) {
            if ($StopFlagRef.Value) {
                $Queue.Enqueue(@{
                    Server  = $srv
                    Status  = "⏹️ Stopped"
                    Color   = "Gray"
                    Details = "User interrupted"
                    Step    = $true
                })
                break
            }

            $conn = $null
            $cmd  = $null

            try {
                $connString = "Server=$srv;Database=master;User ID=$SqlUser;Password=$SqlPass;TrustServerCertificate=True;Connection Timeout=15"
                $conn = New-Object System.Data.SqlClient.SqlConnection $connString
                $conn.Open()
                $cmd = $conn.CreateCommand()
                $cmd.CommandTimeout = 30

                $results = New-Object System.Collections.Generic.List[string]
          foreach ($batch in $sqlBatches) {
                        if ($batch.Trim()) {
                 $cmd.CommandText = $batch
                try {
            $reader = $cmd.ExecuteReader()
            try {
                if ($reader.FieldCount -gt 0) {
                    while ($reader.Read()) {
                        $row = for ($i = 0; $i -lt $reader.FieldCount; $i++) {
                            $reader.GetValue($i)
                        }
                        $results.Add(($row -join ", "))
                    }
                } else {
                    $affected = if ($reader.RecordsAffected -ge 0) { "$($reader.RecordsAffected) row(s) affected" } else { "Non-query batch executed successfully." }
                    $results.Add($affected)
                }
            } finally {
                if ($reader) { $reader.Close(); $reader.Dispose() }
            }
            } catch {
                $results.Add("❌ Batch failed: $($_.Exception.Message)")
            }
        }
    }

                $ts = Get-Date -Format "yyyyMMdd_HHmmss"
                if (-not (Test-Path $LogFolder)) {
                    New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null
                }
                $logFile = Join-Path $LogFolder "SqlResult_${srv}_$ts.txt"
                $results | Set-Content -Path $logFile -Encoding UTF8

                $outputText = if ($results.Count -gt 0) {
                    $results -join "`n"
                } else {
                    "Executed successfully with no result set."
                }

                $hasBatchError = @($results | Where-Object { $_ -like "❌ Batch failed:*" }).Count -gt 0
                $statusText = if ($hasBatchError) { "⚠️ Completed with batch errors" } else { "✅ Success" }
                $colorName  = if ($hasBatchError) { "DarkOrange" } else { "LimeGreen" }

                $Queue.Enqueue(@{
                    Server  = $srv
                    Status  = $statusText
                    Color   = $colorName
                    Details = $outputText
                    Step    = $true
                })
            } catch {
                $Queue.Enqueue(@{
                    Server  = $srv
                    Status  = "❌ Error"
                    Color   = "Red"
                    Details = $_.Exception.Message
                    Step    = $true
                })
            } finally {
                if ($cmd)  { $cmd.Dispose() }
                if ($conn) { $conn.Close(); $conn.Dispose() }
            }
        }

        $State.Done = $true
    })

    $jobEntry = [pscustomobject]@{
        PowerShell  = $ps
        Runspace    = $runspace
        AsyncHandle = $ps.BeginInvoke()
        Queue       = $queue
        State       = $state
    }
    $script:ActiveSqlJobs.Add($jobEntry)

    if (-not $script:SqlTimer) {
        $script:SqlTimer = New-Object System.Windows.Forms.Timer
        $script:SqlTimer.Interval = 150
        $script:SqlTimer.Add_Tick({
            $finishedJobs = [System.Collections.Generic.List[PSCustomObject]]::new()

            foreach ($job in $script:ActiveSqlJobs) {
                $msg = $null
                while ($job.Queue.TryDequeue([ref]$msg)) {
                    $drawColor = [System.Drawing.Color]::FromName($msg.Color)
                    Update-GridStatus $msg.Server $msg.Status $drawColor $msg.Details
                    if ($msg.Step -and $progress.Value -lt $progress.Maximum) {
                        $progress.Value += 1
                    }
                }

                if ($job.State.Done -and $job.Queue.IsEmpty) {
                    try { $job.PowerShell.EndInvoke($job.AsyncHandle) | Out-Null } catch {}
                    $job.PowerShell.Dispose()
                    $job.Runspace.Close()
                    $job.Runspace.Dispose()
                    $finishedJobs.Add($job)
                }
            }

            foreach ($doneJob in $finishedJobs) {
                [void]$script:ActiveSqlJobs.Remove($doneJob)
            }

            if ($script:ActiveSqlJobs.Count -eq 0) {
                $script:SqlTimer.Stop()
                $btnExecSQL.Enabled = $true
                $statusLabel.Text = if (-not $script:StopRequested) { "SQL Execution Completed" } else { "Stopped by user" }
            }
        })
    }

    $script:SqlTimer.Start()
}

# Scripts dropdown loader
# Load scripts from folder, optionally recurse into subfolders
# Load scripts from folder, optionally recurse into subfolders, compatible with older .NET
function Load-ScriptsFromFolder {
    param(
        [string]$folder,
        [bool]$Recurse = $true
    )

    $comboScripts.Items.Clear()
    $scriptPathMap.Clear()
    $comboScripts.SelectedIndex = -1
    $comboScripts.Text = ""

    if (-not (Test-Path $folder)) { return }

    # Gather files
    if ($Recurse) {
        $scriptFiles = Get-ChildItem -Path $folder -Filter '*.ps1' -Recurse -File -ErrorAction SilentlyContinue
    } else {
        $scriptFiles = Get-ChildItem -Path $folder -Filter '*.ps1' -File -ErrorAction SilentlyContinue
    }

    $root = $folder.TrimEnd('\','/')

    foreach ($f in $scriptFiles) {
        $full = $f.FullName
        if ($full.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase)) {
            $rel = $full.Substring($root.Length).TrimStart('\','/')
        } else {
            $rel = $f.Name
        }

        $display = $rel -replace '/','\'

        # Ensure unique display keys
        $key = $display
        $i = 1
        while ($scriptPathMap.ContainsKey($key)) {
            $key = "{0} ({1})" -f $display, $i
            $i++
        }

        $scriptPathMap[$key] = $full

        # Add only the object (not also the string)
        $item = [pscustomobject]@{
            Display = $f.BaseName
            Path    = $full
            Rel     = $display
        }

        $comboScripts.Items.Add($item) | Out-Null
    }

    # Tell ComboBox which property to show
    $comboScripts.DisplayMember = 'Display'

    if ($comboScripts.Items.Count -gt 0) {
        $comboScripts.SelectedIndex = 0
    } else {
        $comboScripts.SelectedIndex = -1
        $comboScripts.Text = ""
    }
}



# Simple details dialog to inspect large cell contents
function Show-DetailsDialog {
    param([string]$Title, [string]$FullText)

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = $Title
    $dlg.Size = New-Object System.Drawing.Size(900,600)
    $dlg.StartPosition = "CenterParent"

    # Use current GUI theme
    $dlg.BackColor = $script:ThemeBgForm
    $dlg.ForeColor = $script:ThemeFgPrimary

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Multiline = $true
    $txt.ReadOnly = $true
    $txt.ScrollBars = "Both"
    $txt.WordWrap = $false
    $txt.Dock = "Fill"
    $txt.Font = New-Object System.Drawing.Font("Consolas",9)
    $txt.Text = $FullText

    # Theme colors
    $txt.BackColor = $script:ThemeBgControl
    $txt.ForeColor = $script:ThemeFgPrimary
    $txt.BorderStyle = "FixedSingle"

    # --- Fix: clear selection so text is not highlighted ---
    $txt.SelectionStart = 0
    $txt.SelectionLength = 0

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = "Save As..."
    $btnSave.Dock = "Bottom"

    # Theme colors
    $btnSave.BackColor = $script:ThemeBgControl
    $btnSave.ForeColor = $script:ThemeFgPrimary
    $btnSave.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup

    $btnSave.Add_Click({
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
        $sfd.FileName = ($Title -replace '[\\/:]', '_') + ".log"

        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txt.Text | Set-Content -Path $sfd.FileName -Encoding UTF8
        }
    })

    $dlg.Controls.Add($txt)
    $dlg.Controls.Add($btnSave)

    [void]$dlg.ShowDialog()
}

function Add-DryRunRow {
    param (
        [string]$Server,
        [string]$Action,
        [string]$Details
    )

    $rowIndex = $grid.Rows.Add($Server, "🧪 $Action", $Details)

    $grid.Rows[$rowIndex].Cells[1].Style.ForeColor = [System.Drawing.Color]::DarkOrange
    $grid.Rows[$rowIndex].Cells[2].Style.ForeColor = [System.Drawing.Color]::DarkOrange
    $grid.Rows[$rowIndex].Cells[1].Style.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
}

function Simulate-DryRun {
    param (
        [string]$Action,              # e.g. "Would run script"
        [string]$DetailsTemplate,     # e.g. "Simulated execution of {target} on {server}"
        [string[]]$Servers,
        [string]$Target               # e.g. script path, file name, SQL query
    )

    $simulatedLines = @()

    foreach ($srv in $Servers) {
        $details = $DetailsTemplate -replace '{server}', $srv -replace '{target}', $Target
        Add-DryRunRow -Server $srv -Action $Action -Details $details
        $simulatedLines += "🧪 $details"
    }

    return $simulatedLines
}

function Show-DryRunSummary {
    param([string[]]$Lines)

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Dry Run Summary"
    $dlg.Size = New-Object System.Drawing.Size(800,500)
    $dlg.StartPosition = "CenterParent"

    # Use current GUI theme
    $dlg.BackColor = $script:ThemeBgForm
    $dlg.ForeColor = $script:ThemeFgPrimary

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Multiline = $true
    $txt.ReadOnly = $true
    $txt.ScrollBars = "Vertical"
    $txt.Dock = "Fill"
    $txt.Font = New-Object System.Drawing.Font("Consolas",9)
    $txt.Text = ($Lines -join "`r`n")

    # Theme colors
    $txt.BackColor = $script:ThemeBgControl
    $txt.ForeColor = $script:ThemeFgPrimary
    $txt.BorderStyle = "FixedSingle"

    # Fix: clear selection so text is not highlighted
    $txt.SelectionStart = 0
    $txt.SelectionLength = 0

    $dlg.Controls.Add($txt)

    [void]$dlg.ShowDialog()
}

function Get-AppScriptsRoot {
    $appRoot = if ($PSScriptRoot) {
        $PSScriptRoot
    } else {
        (Get-Location).Path
    }

    return (Join-Path $appRoot "Scripts")
}

function Open-ScriptShortcutFolder {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SubFolderName
    )

    $appScriptsRoot = Get-AppScriptsRoot
    $targetPath = Join-Path $appScriptsRoot $SubFolderName

    if (Test-Path -LiteralPath $targetPath) {
        $script:lastScriptFolder = $targetPath
        Load-ScriptsFromFolder -folder $targetPath -Recurse $true
    }
    else {
        $dialog = New-Object Windows.Forms.FolderBrowserDialog
        $dialog.Description = "Select folder for $SubFolderName"
        $dialog.SelectedPath = if (Test-Path -LiteralPath $appScriptsRoot) { $appScriptsRoot } else { "" }

        if ($dialog.ShowDialog() -eq [Windows.Forms.DialogResult]::OK) {
            $script:lastScriptFolder = $dialog.SelectedPath
            Load-ScriptsFromFolder -folder $dialog.SelectedPath -Recurse $true
        }
    }

    if ($comboScripts -and $comboScripts.Visible -and $comboScripts.Enabled) {
        $comboScripts.Focus()
        $comboScripts.SelectAll()
    }
}


# [REMOTE AUTOMATION GUI - PART 3/5]
# === INITIAL DATA LOAD ===

# Populate server list dropdown from Lists folder
try {
    Get-ChildItem -Path $listFolder -Filter *.txt -ErrorAction Stop | ForEach-Object {
        [void]$cmbServers.Items.Add($_.Name)
        $serverListPathMap[$_.Name] = $_.FullName
    }
} catch {}

# Auto-select last used list or default to first
if ($lastServerList -and $cmbServers.Items.Contains($lastServerList)) {
    $cmbServers.SelectedItem = $lastServerList
    Load-ServerList $lastServerList
} elseif ($cmbServers.Items.Count -gt 0) {
    $cmbServers.SelectedIndex = 0
    Load-ServerList $cmbServers.SelectedItem
}

# Load scripts from last script folder, fallback to default Scripts\
$defaultScriptsFolder = if (Get-Command Get-AppScriptsRoot -ErrorAction SilentlyContinue) {
    Get-AppScriptsRoot
} else {
    $scriptFolder
}

if ($script:lastScriptFolder -and (Test-Path -LiteralPath $script:lastScriptFolder)) {
    Load-ScriptsFromFolder -folder $script:lastScriptFolder -Recurse $true
} elseif ($defaultScriptsFolder -and (Test-Path -LiteralPath $defaultScriptsFolder)) {
    $script:lastScriptFolder = $defaultScriptsFolder
    Load-ScriptsFromFolder -folder $defaultScriptsFolder -Recurse $true
}

# === EVENTS: grid double-click details ===

$grid.Add_CellDoubleClick({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }  # ignore header clicks

    $row    = $grid.Rows[$e.RowIndex]
    $server = $row.Cells[0].Value
    $status = $row.Cells[1].Value
    $output = $row.Cells[2].Value

    if ($server) {
        # Build the same structured object as Save Log
        $entry = [PSCustomObject]@{
            Server = $server
            Status = $status
            Output = if ($output) { $output.ToString() } else { "" }
        }

        # Format it like Save Log (Out-String for nice text block)
        $fullText = ($entry | Format-List | Out-String).Trim()

        Show-DetailsDialog -Title "Details: $server" -FullText $fullText
    }
})

# === EVENTS: top-level browsing ===
$btnBrowse.Add_Click({
    $dialog = New-Object Windows.Forms.OpenFileDialog
    $dialog.Filter = "All files (*.*)|*.*"
    if ($dialog.ShowDialog() -eq "OK") { $txtFile.Text = $dialog.FileName }
})

$btnBrowseServers.Add_Click({
    $dialog = New-Object Windows.Forms.OpenFileDialog
    $dialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    if ($dialog.ShowDialog() -eq "OK") {
        $fileName = [System.IO.Path]::GetFileName($dialog.FileName)
        if (-not $cmbServers.Items.Contains($fileName)) { [void]$cmbServers.Items.Add($fileName) }
        $serverListPathMap[$fileName] = $dialog.FileName
        $cmbServers.SelectedItem = $fileName
        Load-ServerList $fileName
        $listViewServers.Focus()
    }
})

$btnFolderCommands.Add_Click({
    Open-ScriptShortcutFolder -SubFolderName "commands"
})

$btnFolderDiagnostic.Add_Click({
    Open-ScriptShortcutFolder -SubFolderName "Diagnostic Scripts"
})

$btnFolderSql.Add_Click({
    Open-ScriptShortcutFolder -SubFolderName "SQL scripts"
})

# Browse script folder and reload (keeps recursion enabled)
$btnBrowseScript.Add_Click({
    $dialog = New-Object Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Select a folder containing PowerShell scripts"
    $dialog.SelectedPath = $script:lastScriptFolder
    if ($dialog.ShowDialog() -eq [Windows.Forms.DialogResult]::OK) {
        $script:lastScriptFolder = $dialog.SelectedPath
        Load-ScriptsFromFolder -folder $script:lastScriptFolder -Recurse $true
        Save-Config
    }
})

# Server list changed -> reload
$cmbServers.Add_SelectedIndexChanged({
    Load-ServerList $cmbServers.SelectedItem
    if (Get-Command Set-GuiTheme -ErrorAction SilentlyContinue) {
        Set-GuiTheme -RootControl $cmbServers -Dark $script:isDarkMode
    }
    $listViewServers.Focus()
})

# Scripts dropdown changed -> focus Run Scripts
$comboScripts.Add_SelectedIndexChanged({
    $scriptName    = $comboScripts.SelectedItem
    $possiblePath  = if ($scriptName) { Join-Path $lastScriptFolder $scriptName } else { $null }
    # If manually typed full path, remember it (optional)
    if (-not $selectedScriptPath -and (Test-Path $comboScripts.Text)) { $selectedScriptPath = $comboScripts.Text }

    #if ($form.IsHandleCreated -and $btnRunScripts.Visible -and $btnRunScripts.Enabled) { $btnRunScripts.Focus() }

})

# Filter typing -> refilter while preserving checked items


$txtFilter.Add_TextChanged({
    if (-not $script:CachedServers) { return }

    if (-not ($script:CheckedServers -is [System.Collections.Generic.HashSet[string]])) {
        $script:CheckedServers = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
    }

    $script:IsFilteringServers = $true
    $listViewServers.BeginUpdate()
    try {
        $filter = $txtFilter.Text.Trim()
        $listViewServers.Items.Clear()

        $filtered = if ([string]::IsNullOrWhiteSpace($filter)) {
            $script:CachedServers
        } else {
            $script:CachedServers | Where-Object { $_ -like "*$filter*" }
        }

        $newListItems = foreach ($srv in $filtered) {
            $item = New-Object System.Windows.Forms.ListViewItem($srv)
            if ($script:CheckedServers.Contains($srv)) {
                $item.Checked = $true
            }
            $item
        }

        if ($newListItems) {
            $listViewServers.Items.AddRange($newListItems)
        }
    }
    finally {
        $listViewServers.EndUpdate()
        $script:IsFilteringServers = $false
    }
})

# Enable form-wide key interception
$form.KeyPreview = $true

$form.Add_KeyDown({
    param($sender, $e)

    # Ctrl+F -> Focus $comboScripts and highlight existing text without opening the full dropdown first
    if ($e.Control -and -not $e.Shift -and -not $e.Alt -and $e.KeyCode -eq [System.Windows.Forms.Keys]::F) {
        $e.SuppressKeyPress = $true
        $e.Handled = $true
        $comboScripts.Focus()
        $comboScripts.SelectAll()
    }
    # Ctrl+S -> Focus server filter box ($txtFilter) and highlight existing text
    elseif ($e.Control -and -not $e.Shift -and -not $e.Alt -and $e.KeyCode -eq [System.Windows.Forms.Keys]::S) {
        $e.SuppressKeyPress = $true
        $e.Handled = $true
        if ($null -ne $txtFilter) {
            $txtFilter.Focus()
            $txtFilter.SelectAll()
        }
    }
    # Esc -> Clear active search/filter box
    elseif ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
        # 1. If focused on $comboScripts and it has text -> clear script search first
        if ($comboScripts.Focused -and -not [string]::IsNullOrEmpty($comboScripts.Text)) {
            $e.SuppressKeyPress = $true
            $e.Handled = $true
            $comboScripts.Text = ""
            $comboScripts.DroppedDown = $false

            if ($script:allScripts) {
                $comboScripts.BeginUpdate()
                $comboScripts.Items.Clear()
                [void]$comboScripts.Items.AddRange(@($script:allScripts))
                $comboScripts.EndUpdate()
            }
        }
        # 2. Else if focused on $txtFilter and it has text -> clear server filter first
        elseif ($txtFilter.Focused -and -not [string]::IsNullOrEmpty($txtFilter.Text)) {
            $e.SuppressKeyPress = $true
            $e.Handled = $true
            $txtFilter.Clear()
        }
        # 3. Fallback (when focus is on the grid, list, or a button):
        # Clear whichever box currently has text (script search first, then server filter)
        elseif (-not [string]::IsNullOrEmpty($comboScripts.Text)) {
            $e.SuppressKeyPress = $true
            $e.Handled =true
            $comboScripts.Text = ""
            $comboScripts.DroppedDown = $false

            if ($script:allScripts) {
                $comboScripts.BeginUpdate()
                $comboScripts.Items.Clear()
                [void]$comboScripts.Items.AddRange(@($script:allScripts))
                $comboScripts.EndUpdate()
            }
        }
        elseif (-not [string]::IsNullOrEmpty($txtFilter.Text)) {
            $e.SuppressKeyPress = $true
            $e.Handled = $true
            $txtFilter.Clear()
        }
    }
})



# [REMOTE AUTOMATION GUI - PART 4/5]
# === ACTION BUTTONS ===

$btnSaveSelection.Add_Click({
    $lines = Get-SelectedDisplayLines
    if ($lines.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No servers selected to save.")
        return
    }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter    = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    $sfd.Title     = "Save Selected Servers"
    $sfd.FileName  = "SelectedServers.txt"
    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $lines | Set-Content -Path $sfd.FileName -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("Selection saved to:`n$($sfd.FileName)")
    }
})

$btnCheckAll.Add_Click({
    if (-not $listViewServers -or $listViewServers.Items.Count -eq 0) { return }

    if (-not $script:CheckedServers) {
        $script:CheckedServers = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
    }

    $script:IsBulkChecking = $true
    $listViewServers.BeginUpdate()
    try {
        foreach ($item in $listViewServers.Items) {
            $item.Checked = $true
            [void]$script:CheckedServers.Add($item.Text)
        }
        if ($listViewServers.Items.Count -gt 0) {
            $script:LastCheckedIndex = $listViewServers.Items.Count - 1
        }
    }
    finally {
        $listViewServers.EndUpdate()
        $script:IsBulkChecking = $false
    }

    if ($chkDryRun) { $chkDryRun.Checked = $true }
    if ($statusLabel) {
        $statusLabel.Text = "🧪 Dry Run Enabled (via Check All)"
        $statusLabel.ForeColor = [System.Drawing.Color]::DarkOrange
    }
})

$btnUncheckAll.Add_Click({
    if (-not $listViewServers) { return }

    $script:IsBulkChecking = $true
    $listViewServers.BeginUpdate()
    try {
        if ($script:CheckedServers) {
            $script:CheckedServers.Clear()
        }
        foreach ($item in $listViewServers.Items) {
            $item.Checked = $false
        }
        $script:LastCheckedIndex = $null
    }
    finally {
        $listViewServers.EndUpdate()
        $script:IsBulkChecking = $false
    }
})

$btnCopy.Add_Click({
    if ($script:ActiveCopyJobs.Count -gt 0) {
        [System.Windows.Forms.MessageBox]::Show("A copy operation is already running.")
        return
    }

    $script:StopRequested = $false

    Save-Config
    $fileToCopy = $txtFile.Text.Trim()
    if (-not (Test-Path $fileToCopy)) {
        [System.Windows.Forms.MessageBox]::Show("Select a valid file first.")
        return
    }

    $rawRemoteFolder = $txtRemoteFolder.Text.Trim()
    if (-not $rawRemoteFolder) {
        [System.Windows.Forms.MessageBox]::Show("Please enter the remote folder path, e.g. C:\Kit or C$\Kit")
        return
    }

    # Convert C:\... to C$\... for UNC copy
    if ($rawRemoteFolder -match "^[a-zA-Z]:\\") {
        $drive = $rawRemoteFolder.Substring(0,1)
        $sub   = $rawRemoteFolder.Substring(3)
        $rawRemoteFolder = "$drive`$\$sub"
    }

    $remoteFolderLive = $rawRemoteFolder.TrimEnd('\')

    $servers = Get-SelectedServers
    if ($servers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one server.")
        return
    }

    if ($chkDryRun.Checked) {
        $fileName = Split-Path $txtFile.Text -Leaf
        $simulated = Simulate-DryRun -Action "Would copy file" `
            -DetailsTemplate "Simulated copy of {target} to {server}" `
            -Servers $servers -Target $fileName

        Show-DryRunSummary $simulated
        $statusLabel.Text = "Dry Run Completed"
        return
    }

    try {
        $cred = Import-Clixml -Path $fixedCredPath
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to load credentials from:`n$fixedCredPath")
        return
    }

    $statusLabel.Text = "Copying in background..."

    # Kick off the asynchronous copy function (manages grid and timer internally)
    Copy-ToServers $fileToCopy $remoteFolderLive $servers $cred $chkCreateFolder.Checked
})

$btnRunScripts.Add_Click({
    if ($script:ActiveScriptJobs.Count -gt 0) {
        [System.Windows.Forms.MessageBox]::Show("A script execution is already running.")
        return
    }

    $script:SqlFullOutput = @{}
    $script:StopRequested = $false

    Save-Config
    $selectedServers = Get-SelectedServers
    if ($selectedServers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one server.")
        return
    }

    $selected = $comboScripts.SelectedItem
    if (-not $selected) {
        [System.Windows.Forms.MessageBox]::Show("Please select a script from the dropdown.")
        return
    }

    $scriptPath = $selected.Path
    $filePath   = $txtFile.Text

    if (-not (Test-Path $scriptPath)) {
        [System.Windows.Forms.MessageBox]::Show("Script not found:`n$scriptPath")
        return
    }

    $scriptText = Get-Content -Path $scriptPath -Raw -Encoding UTF8
    $scriptContent = $scriptText
    $usesSqlAuth = $scriptContent -match 'Invoke-Sqlcmd' -or $scriptContent -match '\$UseSqlAuth'

    # Initialize empty defaults so they always exist in scope
    $sqlUser = ""
    $sqlPass = ""

    if ($usesSqlAuth) {
        $userFile = Join-Path $secureFolder "sqluser.txt"
        $passFile = Join-Path $secureFolder "sqlpass.txt"

        if (-not (Test-Path $userFile) -or -not (Test-Path $passFile)) {
            [System.Windows.Forms.MessageBox]::Show("Missing sqluser.txt or sqlpass.txt in:`n$secureFolder")
            return
        }

        $sqlUser = (Get-Content -Path $userFile -Raw -ErrorAction Stop).Trim()
        $rawPass = (Get-Content -Path $passFile -Raw -ErrorAction Stop).Trim()

        try {
            $securePass = ConvertTo-SecureString -String $rawPass -ErrorAction Stop
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Stored SQL password could not be decrypted. Re-save sqlpass.txt for the current user.")
            return
        }

        $cred = New-Object System.Management.Automation.PSCredential($sqlUser, $securePass)
        $sqlPass = $cred.GetNetworkCredential().Password
    }

    Run-ScriptOnServers -Servers $selectedServers -ScriptPath $scriptPath -ScriptText $scriptText -FilePath $filePath `
        -UsesSqlAuth:$usesSqlAuth -SqlUser $sqlUser -SqlPass $sqlPass -IsDryRun:$chkDryRun.Checked
})

$btnRunFile.Add_Click({
    if ($script:ActiveRunJobs.Count -gt 0) {
        [System.Windows.Forms.MessageBox]::Show("A file execution is already running.")
        return
    }

    $script:SqlFullOutput = @{}
    $script:StopRequested = $false

    Save-Config

    $remotePath = $txtRemoteFolder.Text.Trim()
    if (-not $remotePath) {
        [System.Windows.Forms.MessageBox]::Show("Please enter the full remote file path, e.g. C:\Kit\EnableWinRM.bat")
        return
    }

    # Convert C$\... to C:\... when used inside remote session path resolution
    if ($remotePath -match "^[a-zA-Z]\$\\") {
        $drive = $remotePath.Substring(0,1)
        $sub   = $remotePath.Substring(3)
        $remotePath = "$drive`:\$sub"
    }

    $servers = Get-SelectedServers
    if ($servers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one server.")
        return
    }

    if ($chkDryRun.Checked) {
        $targetName = Split-Path $remotePath -Leaf
        $simulated = Simulate-DryRun -Action "Would run file" `
            -DetailsTemplate "Simulated execution of {target} on {server}" `
            -Servers $servers -Target $targetName

        Show-DryRunSummary $simulated
        $statusLabel.Text = "Dry Run Completed"
        return
    }

    try {
        $cred = Import-Clixml -Path $fixedCredPath
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to load credentials from:`n$fixedCredPath")
        return
    }

    $btnRunFile.Enabled = $false
    Start-RunOnServersAsync -RemotePath $remotePath -Servers $servers -Credential $cred
})

$btnExecSQL.Add_Click({
    if ($script:ActiveSqlJobs.Count -gt 0) {
        [System.Windows.Forms.MessageBox]::Show("A SQL execution is already running.")
        return
    }

    $script:StopRequested = $false
    $servers = Get-SelectedServers
    if ($servers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one server.")
        return
    }

    $sqlFile = $txtFile.Text
    if (-not (Test-Path $sqlFile)) {
        [System.Windows.Forms.MessageBox]::Show("Please select a valid SQL file in Selected File.")
        return
    }

    Save-Config

    $sqlText = Get-Content -Path $sqlFile -Raw -Encoding UTF8

    if ($chkDryRun.Checked) {
        $simulated = Simulate-DryRun -Action "Would execute SQL" `
            -DetailsTemplate "Simulated SQL query '{target}' on {server}" `
            -Servers $servers -Target $sqlText

        Show-DryRunSummary $simulated
        $statusLabel.Text = "Dry Run Completed"
        return
    }

    $sqlUser = (Get-Content "$secureFolder\sqluser.txt" -Raw -ErrorAction Stop).Trim()
    $rawPass = (Get-Content "$secureFolder\sqlpass.txt" -Raw -ErrorAction Stop).Trim()

    if ([string]::IsNullOrWhiteSpace($sqlUser) -or [string]::IsNullOrWhiteSpace($rawPass)) {
        [System.Windows.Forms.MessageBox]::Show("SQL username or password file is empty.")
        return
    }

    try {
        $securePass = ConvertTo-SecureString -String $rawPass -ErrorAction Stop
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Stored SQL password could not be decrypted. Re-save sqlpass.txt for the current user.")
        return
    }

    $cred = New-Object System.Management.Automation.PSCredential($sqlUser, $securePass)
    $sqlPass = $cred.GetNetworkCredential().Password

    $grid.Rows.Clear()
    foreach ($s in $servers) {
        [void]$grid.Rows.Add($s, "Pending", "")
    }

    $btnExecSQL.Enabled = $false
    Start-ExecuteSqlOnServersAsync `
        -Servers $servers `
        -SqlFile $sqlFile `
        -SqlText $sqlText `
        -SqlUser $sqlUser `
        -SqlPass $sqlPass `
        -LogFolder $logFolder
})

$btnStop.Add_Click({
    $script:StopRequested = $true

    # Stop all timers immediately so no tick fires mid-cleanup
    foreach ($timerName in @('CopyTimer', 'ScriptTimer', 'RunTimer', 'SqlTimer', 'PingTimer')) {
        $timerVar = Get-Variable -Name $timerName -Scope Script -ErrorAction SilentlyContinue
        if ($timerVar -and $timerVar.Value) {
            try {
                $timerVar.Value.Stop()
                $timerVar.Value.Dispose()
            } catch {}
            Set-Variable -Name $timerName -Scope Script -Value $null
        }
    }

    # Re-enable action buttons right away
    if ($btnRunFile) { $btnRunFile.Enabled = $true }
    if ($btnExecSQL) { $btnExecSQL.Enabled = $true }

    # Update grid and status immediately
    foreach ($row in $grid.Rows) {
        if ($row.Cells[1].Value -eq "Pending") {
            $row.Cells[1].Value = "⏹️ Stopped"
            $row.Cells[1].Style.ForeColor = [System.Drawing.Color]::Gray
            $row.Cells[2].Value = "Operation cancelled by user"
        }
    }
    $statusLabel.Text = "Stopped (cleaning up background jobs...)"

    # Snapshot and clear active job/runspace lists on the UI thread
    $jobsToStop = @()
    foreach ($listName in @('ActiveCopyJobs', 'ActiveScriptJobs', 'ActiveRunJobs', 'ActivePingJobs')) {
        $listVar = Get-Variable -Name $listName -Scope Script -ErrorAction SilentlyContinue
        if ($listVar -and $listVar.Value) {
            foreach ($item in @($listVar.Value)) {
                if ($item -and $item.Job) {
                    $jobsToStop += $item.Job
                }
            }
            $listVar.Value.Clear()
        }
    }

    $sqlToStop = @()
    if ($script:ActiveSqlJobs) {
        $sqlToStop = @($script:ActiveSqlJobs)
        $script:ActiveSqlJobs.Clear()
    }

    # Offload blocking Stop-Job / Remove-Job teardown to a background runspace
    if ($jobsToStop.Count -gt 0 -or $sqlToStop.Count -gt 0) {
        $cleanupRs = [runspacefactory]::CreateRunspace()
        $cleanupRs.Open()
        $cleanupRs.SessionStateProxy.SetVariable('JobsToStop', $jobsToStop)
        $cleanupRs.SessionStateProxy.SetVariable('SqlToStop', $sqlToStop)

        $cleanupPs = [powershell]::Create()
        $cleanupPs.Runspace = $cleanupRs
        [void]$cleanupPs.AddScript({
            foreach ($job in $JobsToStop) {
                try {
                    if ($job.State -eq 'Running') {
                        Stop-Job -Job $job -ErrorAction SilentlyContinue
                    }
                    Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
                } catch {}
            }

            foreach ($sqlJob in $sqlToStop) {
                try {
                    if ($sqlJob.PowerShell) {
                        $sqlJob.PowerShell.Stop()
                        $sqlJob.PowerShell.Dispose()
                    }
                } catch {}
                try {
                    if ($sqlJob.Runspace) {
                        $sqlJob.Runspace.Close()
                        $sqlJob.Runspace.Dispose()
                    }
                } catch {}
            }
        })

        [void]$cleanupPs.BeginInvoke()
    }

    $statusLabel.Text = "Stopped by user"
})

$btnConnectVNC.Add_Click({
    $servers = Get-SelectedServers
    if ($servers.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Please select at least one server."); return }

    $vncViewer = Join-Path $basePath 'Tools\UltraVNC\vncviewer.exe'
    $xmlPath   = Join-Path $secureFolder 'VncPasswords.xml'

    if (-not (Test-Path $vncViewer)) { [System.Windows.Forms.MessageBox]::Show("UltraVNC viewer not found at $vncViewer"); return }
    if (-not (Test-Path $xmlPath))   { [System.Windows.Forms.MessageBox]::Show("Password file not found at $xmlPath"); return }

    try { $xml = [xml](Get-Content $xmlPath -Encoding UTF8) } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed loading XML: $($_.Exception.Message)")
        return
    }

    $credMap = @{}
    foreach ($entry in $xml.VncCredentials.Server) {
        $ip = $entry.ip.Trim(); $pw = $entry.password.Trim()
        if ($ip -and $pw) { $credMap[$ip] = $pw }
    }

    foreach ($srv in $servers) {
        $srv = $srv.Trim(); if (-not $srv) { continue }
        $args = @($srv)
        $status = "Launching"
        $output = "Attempting connection to $srv"

        if ($credMap.ContainsKey($srv)) {
            $pw = $credMap[$srv]
            if ($pw -like "dpapi:*") {
                $b64 = $pw.Substring(6)
                try { $pw = Unprotect-String -CipherText $b64 -Scope CurrentUser }
                catch { $pw = ""; $output = "Failed to decrypt DPAPI password." }
            }
            $args += '/password'; $args += $pw
            $output = "Using stored password for $srv"
        }

        try {
            & $vncViewer @args
            $status = "✅ Connected"
        } catch {
            $status = "❌ Failed"; $output = $_.Exception.Message
        }

        $rowIndex = $grid.Rows.Add($srv, $status, $output)
        $color    = if ($status -like "✅*") { [System.Drawing.Color]::LimeGreen } else { [System.Drawing.Color]::Red }
        $grid.Rows[$rowIndex].Cells[1].Style.ForeColor = $color
    }
    $statusLabel.Text = "Ready"
})

if (-not $script:ActivePingJobs) {
    $script:ActivePingJobs = [System.Collections.Generic.List[PSCustomObject]]::new()
}

$btnPing.Add_Click({
    $script:StopRequested = $false
    $servers = Get-SelectedServers
    if ($servers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one server.")
        return
    }

    if ($script:PingTimer) {
        $script:PingTimer.Stop()
        $script:PingTimer.Dispose()
        $script:PingTimer = $null
    }
    $script:ActivePingJobs.Clear()

    $script:SqlFullOutput = @{}
    $grid.Rows.Clear()
    foreach ($s in $servers) {
        [void]$grid.Rows.Add($s, "Pending", "")
    }

    $progress.Value   = 0
    $progress.Maximum = $servers.Count
    $statusLabel.Text = "Pinging servers in parallel..."

    foreach ($srv in $servers) {
        $job = Start-Job -ScriptBlock {
            param($TargetServer)

            $res = [PSCustomObject]@{
                Server  = $TargetServer
                Status  = "❌ No Response"
                Color   = "Red"
                Message = "Ping to ${TargetServer}: No Response"
            }

            try {
                $ping = Test-Connection -ComputerName $TargetServer -Count 1 -Quiet -ErrorAction Stop
                if ($ping) {
                    $res.Status  = "✅ Reachable"
                    $res.Color   = "LimeGreen"
                    $res.Message = "Ping to ${TargetServer}: Reachable"
                } else {
                    $res.Status  = "❌ No Response"
                    $res.Color   = "Red"
                    $res.Message = "Ping to ${TargetServer}: No Response"
                }
            } catch {
                $res.Status  = "❌ Error"
                $res.Color   = "Red"
                $res.Message = $_.Exception.Message
            }

            return $res
        } -ArgumentList $srv

        $script:ActivePingJobs.Add([PSCustomObject]@{
            Server = $srv
            Job    = $job
        })
    }

    $script:PingTimer = New-Object System.Windows.Forms.Timer
    $script:PingTimer.Interval = 200

    $script:PingTimer.Add_Tick({
        if ($script:StopRequested) {
            if ($script:PingTimer) {
                $script:PingTimer.Stop()
                $script:PingTimer.Dispose()
                $script:PingTimer = $null
            }
            return
        }

        $completed = [System.Collections.Generic.List[PSCustomObject]]::new()

        foreach ($item in $script:ActivePingJobs) {
            $job = $item.Job
            if ($null -eq $job) { continue }

            if ($job.State -eq 'Completed' -or $job.State -eq 'Failed' -or $job.State -eq 'Stopped') {
                $data = Receive-Job $job -ErrorAction SilentlyContinue
                Remove-Job $job -Force -ErrorAction SilentlyContinue

                $statusText = if ($data -and $data.Status)  { $data.Status }  else { "❌ Failed" }
                $detailsMsg = if ($data -and $data.Message) { $data.Message } else { "No response data returned." }
                $color = switch ($data.Color) {
                    "LimeGreen" { [System.Drawing.Color]::LimeGreen }
                    "Red"   { [System.Drawing.Color]::Red }
                    default { if ($statusText -like "✅*") { [System.Drawing.Color]::LimeGreen } else { [System.Drawing.Color]::Red } }
                }

                Update-GridStatus $item.Server $statusText $color $detailsMsg
                $progress.Value = [Math]::Min($progress.Value + 1, $progress.Maximum)
                $completed.Add($item)
            }
        }

        foreach ($done in $completed) {
            [void]$script:ActivePingJobs.Remove($done)
        }

        if ($script:ActivePingJobs.Count -eq 0) {
            $script:PingTimer.Stop()
            $script:PingTimer.Dispose()
            $script:PingTimer = $null
            $statusLabel.Text = "Ping Completed"
        }
    })

    $script:PingTimer.Start()
})
$btnRDP.Add_Click({
    $servers = Get-SelectedServers
    if ($servers.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Please select at least one server."); return }

    foreach ($srv in $servers) {
        $srv = $srv.Trim(); if (-not $srv) { continue }
        try {
            Start-Process -FilePath "mstsc.exe" -ArgumentList "/v:$srv"
            $rowIndex = $grid.Rows.Add($srv, "✅ Launched", "RDP session launched")
            $grid.Rows[$rowIndex].Cells[1].Style.ForeColor = [System.Drawing.Color]::LimeGreen
        } catch {
            $rowIndex = $grid.Rows.Add($srv, "❌ Failed", $_.Exception.Message)
            $grid.Rows[$rowIndex].Cells[1].Style.ForeColor = [System.Drawing.Color]::Red
        }
    }
    $statusLabel.Text = "RDP launch complete"
})

$btnOpenC.Add_Click({
    $servers = Get-SelectedServers
    if ($servers.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Please select at least one server."); return }
    foreach ($srv in $servers) {
        $srv = $srv.Trim(); if (-not $srv) { continue }
        try {
            $unc = "\\$srv\C$"
            Start-Process -FilePath "explorer.exe" -ArgumentList $unc
            $rowIndex = $grid.Rows.Add($srv, "✅ Opened C$", $unc)
            $grid.Rows[$rowIndex].Cells[1].Style.ForeColor = [System.Drawing.Color]::LimeGreen
        } catch {
            $rowIndex = $grid.Rows.Add($srv, "❌ Open failed", $_.Exception.Message)
            $grid.Rows[$rowIndex].Cells[1].Style.ForeColor = [System.Drawing.Color]::Red
        }
    }
})

$btnOpenD.Add_Click({
    $servers = Get-SelectedServers
    if ($servers.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Please select at least one server."); return }
    foreach ($srv in $servers) {
        $srv = $srv.Trim(); if (-not $srv) { continue }
        try {
            $unc = "\\$srv\D$"
            Start-Process -FilePath "explorer.exe" -ArgumentList $unc
            $rowIndex = $grid.Rows.Add($srv, "✅ Opened D$", $unc)
            $grid.Rows[$rowIndex].Cells[1].Style.ForeColor = [System.Drawing.Color]::LimeGreen
        } catch {
            $rowIndex = $grid.Rows.Add($srv, "❌ Open failed", $_.Exception.Message)
            $grid.Rows[$rowIndex].Cells[1].Style.ForeColor = [System.Drawing.Color]::Red
        }
    }
})

$btnCreateCred_Click = {

    try {
        $prompt = "Enter the administrator credential to store at $fixedCredPath"
        $cred = $Host.UI.PromptForCredential("Create Credentials", $prompt, "Administrator", "")

        if (-not $cred) {
            [System.Windows.Forms.MessageBox]::Show(
                "No credentials entered. Operation cancelled.",
                "Cancelled",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            return
        }

        $ok = Save-FixedCredential -Cred $cred

        if ($ok) {
            [System.Windows.Forms.MessageBox]::Show(
                "Credentials saved to:`n$fixedCredPath",
                "Saved",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }

        $userName      = $cred.UserName
        $plainPassword = $cred.GetNetworkCredential().Password
        $servers       = Get-SelectedServers
        $successCount  = 0
        $failCount     = 0

        foreach ($srv in $servers) {

            try {
                cmdkey /delete:$srv | Out-Null
                cmdkey /add:$srv /user:$userName /pass:$plainPassword

                Update-GridStatus $srv "✅ Credential Added" ([System.Drawing.Color]::LimeGreen) "Stored Windows credentials"
                $successCount++
            }
            catch {
                Update-GridStatus $srv "❌ Failed" ([System.Drawing.Color]::Red) $_.Exception.Message
                $failCount++
            }
        }

        [System.Windows.Forms.MessageBox]::Show(
            "Windows credentials created for $successCount servers.`nFailures: $failCount",
            "Summary",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Unexpected error: $($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

$btnCreateCred.Add_Click($btnCreateCred_Click)

# Define the click handler once and store it in a variable
$btnEncryptVnc_Click = {
    try {
        $xmlPath = Join-Path $secureFolder 'VncPasswords.xml'
        if (-not (Test-Path $secureFolder)) {
            New-Item -Path $secureFolder -ItemType Directory -Force | Out-Null
        }

        # --- Ask once for new password ---
        $prompt = New-Object Windows.Forms.Form
        $prompt.Text = "Encrypt Selected Servers"
        $prompt.Size = New-Object Drawing.Size(360,160)
        $prompt.StartPosition = "CenterParent"

        $lbl = New-Object Windows.Forms.Label
        $lbl.Text = "Enter new VNC password to apply to selected servers:"
        $lbl.Location = '15,15'; $lbl.Size = '330,20'

        $txtPass = New-Object Windows.Forms.TextBox
        $txtPass.Location = '15,40'; $txtPass.Size = '320,25'
        $txtPass.UseSystemPasswordChar = $true

        $btnOK = New-Object Windows.Forms.Button
        $btnOK.Text = "OK"; $btnOK.Location = '70,80'; $btnOK.Size = '100,25'
        $btnCancel = New-Object Windows.Forms.Button
        $btnCancel.Text = "Cancel"; $btnCancel.Location = '190,80'; $btnCancel.Size = '100,25'

        $prompt.Controls.AddRange(@($lbl,$txtPass,$btnOK,$btnCancel))

        $password = $null
        $btnOK.Add_Click({
            $password = $txtPass.Text
            if (-not $password) {
                [System.Windows.Forms.MessageBox]::Show("Please enter a password.") ; return
            }
            $prompt.Tag = $password
            $prompt.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $prompt.Close()
        })
        $btnCancel.Add_Click({ $prompt.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $prompt.Close() })

        $result = $prompt.ShowDialog()
        if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $password = $prompt.Tag
        if (-not $password) { return }

        # --- Load or create XML ---
        [xml]$xml = $null
        if (Test-Path $xmlPath) {
            try { $xml = [xml](Get-Content $xmlPath -Raw -Encoding UTF8) } catch { $xml = $null }
        }
        if (-not $xml -or -not $xml.VncCredentials) {
            $xml = New-Object xml
            $root = $xml.CreateElement("VncCredentials")
            $xml.AppendChild($root) | Out-Null
        }

        # --- Get selected IPs from your checklist ---
        $selectedServers = Get-SelectedServers
        if ($selectedServers.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one server from the list.") ; return
        }

        $existingIPs = @{}
        foreach ($srv in $xml.VncCredentials.Server) {
            $existingIPs[$srv.ip] = $srv
        }

        # --- Update or create entries for selected servers ---
        foreach ($srv in $selectedServers) {
            if ($existingIPs.ContainsKey($srv)) {
                $node = $existingIPs[$srv]
            } else {
                $node = $xml.CreateElement("Server")
                $node.SetAttribute("ip", $srv)
                $xml.VncCredentials.AppendChild($node) | Out-Null
            }
            $enc = Protect-String -PlainText $password -Scope CurrentUser
            $node.SetAttribute("password", "dpapi:$enc")
        }

        $xml.Save($xmlPath)
        [System.Windows.Forms.MessageBox]::Show("Passwords updated for selected servers:`n`n$($selectedServers -join ', ')`n`nSaved in:`n$xmlPath`n`nOnly this user can decrypt them.",
            "Done", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Encryption failed:`n$($_.Exception.Message)","Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
    }
}


# Attach the handler
$btnEncryptVnc.Add_Click($btnEncryptVnc_Click)

$btnCreateSqlCred_Click = {

    try {
        $prompt = "Enter the SQL credential to store at $secureFolder"
        $cred = $Host.UI.PromptForCredential("Create SQL Credential", $prompt, "", "")
        if (-not $cred) {
            [System.Windows.Forms.MessageBox]::Show("No credentials entered. Operation cancelled.","Cancelled",
                [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }

        if (-not (Test-Path $secureFolder)) {
            New-Item -Path $secureFolder -ItemType Directory | Out-Null
        }

        $cred.UserName | Set-Content "$secureFolder\sqluser.txt"
        $cred.Password | ConvertFrom-SecureString | Set-Content "$secureFolder\sqlpass.txt"

        [System.Windows.Forms.MessageBox]::Show(
            "SQL credentials saved to:`n$secureFolder`n`nOnly this Windows user on this machine can decrypt them.",
            "Saved",
            [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Unexpected error: $($_.Exception.Message)","Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
    }

}

$btnCreateSqlCred.Add_Click($btnCreateSqlCred_Click)



$adminItem.Add_Click($btnCreateCred_Click)

$sqlItem.Add_Click($btnCreateSqlCred_Click)

$vncItem.Add_Click($btnEncryptVnc_Click)


$btnCreateCred.Visible = $false
$btnEncryptVnc.Visible = $false
$btnCreateSqlCred.Visible = $false

$btnSaveLog.Add_Click({
    $entries = @()

    foreach ($row in $grid.Rows) {
        if ($row.IsNewRow) { continue }

        $server = $row.Cells[0].Value
        $status = $row.Cells[1].Value
        $output = $row.Cells[2].Value

        if ($server -or $status -or $output) {
            $entries += [PSCustomObject]@{
                Server = if ($server) { $server.ToString() } else { "" }
                Status = if ($status) { $status.ToString() } else { "" }
                Output = if ($output) { $output.ToString() } else { "" }
            }
        }
    }

    if ($entries.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "There are no grid results to save.",
            "Save Log",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }

    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Title = "Save Output Log"
    $dialog.Filter = "Log Files (*.log)|*.log|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    $dialog.FileName = "ScriptOutput_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date)

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $formattedBlocks = foreach ($entry in $entries) {
            @(
                "Server : $($entry.Server)"
                "Status : $($entry.Status)"
                "Output : $($entry.Output)"
            ) -join "`r`n"
        }

        $logContent = $formattedBlocks -join ("`r`n" + ("-" * 50) + "`r`n")
        Set-Content -LiteralPath $dialog.FileName -Value $logContent -Encoding UTF8
    }
})

# [REMOTE AUTOMATION GUI - PART 5/5]
# === ADD CONTROLS & SHOW FORM ===

$form.Controls.Add($lblFile)
$form.Controls.Add($txtFile)
$form.Controls.Add($btnBrowse)

$form.Controls.Add($lblRemoteFolder)
$form.Controls.Add($txtRemoteFolder)

$form.Controls.Add($lblServers)
$form.Controls.Add($cmbServers)
$form.Controls.Add($btnBrowseServers)

$form.Controls.Add($lblFilter)
$form.Controls.Add($txtFilter)
$form.Controls.Add($listViewServers)

$form.Controls.Add($btnSaveSelection)
$form.Controls.Add($btnResetFilter)
$form.Controls.Add($btnCheckAll)
$form.Controls.Add($btnUncheckAll)
$form.Controls.Add($btnCreateCred)

$form.Controls.Add($btnCopy)
$form.Controls.Add($btnRunFile)
$form.Controls.Add($btnExecSQL)
$form.Controls.Add($btnStop)

$form.Controls.Add($btnPing)
$form.Controls.Add($btnConnectVNC)
$form.Controls.Add($btnRDP)
$form.Controls.Add($btnOpenC)
$form.Controls.Add($btnOpenD)

$form.Controls.Add($progress)
$form.Controls.Add($grid)
$form.Controls.Add($btnRunScripts)
$form.Controls.Add($btnBrowseScript)
$form.Controls.Add($comboScripts)
$form.Controls.Add($btnSaveLog)
$form.Controls.Add($statusLabel)
$form.Controls.Add($btnCreateSqlCred)
$form.Controls.Add($btnEncryptVnc)
$form.Controls.Add($chkCreateFolder)
$form.Controls.Add($chkDryRun)

$form.Controls.Add($btnCreateCredentials)
$form.Controls.AddRange(@($btnFolderCommands, $btnFolderDiagnostic, $btnFolderSql))
$form.Controls.Add($lblServersbar)

# === Minimal Dark Grey Title Bar Helper ===
if (-not ("Win32.DarkTitleBar" -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

namespace Win32 {
    public class DarkTitleBar {
        [DllImport("dwmapi.dll", PreserveSig = true)]
        public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
    }
}
"@
}

function Set-WindowTitleBarTheme {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.Form]$Form,
        [Parameter(Mandatory = $true)]
        [bool]$Dark
    )
    if (-not $Form.IsHandleCreated) { [void]$Form.Handle }
    [int]$useDark = if ($Dark) { 1 } else { 0 }
    [void][Win32.DarkTitleBar]::DwmSetWindowAttribute($Form.Handle, 20, [ref]$useDark, 4)
}

# === Minimal Grey / Light Theme Switcher ===
$script:isDarkMode = $true

function Set-GuiTheme {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.Control]$RootControl,

        [Parameter(Mandatory = $true)]
        [bool]$Dark
    )

    if ($Dark) {

        # Softer Grey Theme

        $bgForm      = [System.Drawing.Color]::FromArgb(72,72,72)
        $bgControl   = [System.Drawing.Color]::FromArgb(82,82,82)
        $bgGridAlt   = [System.Drawing.Color]::FromArgb(80,80,80)
        $bgHeader    = [System.Drawing.Color]::FromArgb(102,102,102)

        $fgPrimary   = [System.Drawing.Color]::White

        $borderColor = [System.Drawing.Color]::FromArgb(135,135,135)
    }
    else {

        $bgForm      = [System.Drawing.SystemColors]::Control
        $bgControl   = [System.Drawing.Color]::White
        $bgGridAlt   = [System.Drawing.Color]::FromArgb(245,247,250)
        $bgHeader    = [System.Drawing.Color]::FromArgb(232,235,240)

        $fgPrimary   = [System.Drawing.SystemColors]::ControlText

        $borderColor = [System.Drawing.SystemColors]::ControlDark
    }

    # Save current theme colors globally
    $script:ThemeBgForm      = $bgForm
    $script:ThemeBgControl   = $bgControl
    $script:ThemeBgHeader    = $bgHeader
    $script:ThemeFgPrimary   = $fgPrimary
    $script:ThemeBorderColor = $borderColor

    $applyStyle = {

        param([System.Windows.Forms.Control]$ctrl)

        if ($ctrl -eq $form) {

            $ctrl.BackColor = $bgForm
            $ctrl.ForeColor = $fgPrimary
        }

        elseif ($ctrl -is [System.Windows.Forms.DataGridView]) {

            $ctrl.BackgroundColor = $bgControl
            $ctrl.GridColor = $borderColor

            $ctrl.EnableHeadersVisualStyles = $false

            $ctrl.ColumnHeadersDefaultCellStyle.BackColor = $bgHeader
            $ctrl.ColumnHeadersDefaultCellStyle.ForeColor = $fgPrimary

            $ctrl.DefaultCellStyle.BackColor = $bgControl
            $ctrl.DefaultCellStyle.ForeColor = $fgPrimary

            $ctrl.AlternatingRowsDefaultCellStyle.BackColor = $bgControl
            $ctrl.AlternatingRowsDefaultCellStyle.ForeColor = $fgPrimary

            $ctrl.DefaultCellStyle.SelectionBackColor =
                [System.Drawing.Color]::FromArgb(110,110,110)

            $ctrl.DefaultCellStyle.SelectionForeColor =
                [System.Drawing.Color]::White

            $ctrl.RowHeadersVisible = $false
        }

        elseif ($ctrl -is [System.Windows.Forms.ListView]) {

            $ctrl.OwnerDraw = $false
            $ctrl.BackColor = $bgControl
            $ctrl.ForeColor = $fgPrimary
            $ctrl.HideSelection = $Dark

            $ctrl.Refresh()
        }

        elseif ($ctrl -is [System.Windows.Forms.ComboBox]) {

            $ctrl.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $ctrl.BackColor = $bgControl
            $ctrl.ForeColor = $fgPrimary

            $ctrl.Invalidate()
        }

        elseif ($ctrl -is [System.Windows.Forms.TextBox]) {

            $ctrl.BackColor = $bgControl
            $ctrl.ForeColor = $fgPrimary
            $ctrl.BorderStyle = 'FixedSingle'
        }

        elseif ($ctrl -is [System.Windows.Forms.ListBox]) {

            $ctrl.BackColor = $bgControl
            $ctrl.ForeColor = $fgPrimary
        }

        elseif ($ctrl -is [System.Windows.Forms.Button]) {

            if ($ctrl -notin @($btnRunScripts, $btnStop)) {

                $ctrl.BackColor = $bgControl
                $ctrl.ForeColor = $fgPrimary

                $ctrl.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
            }
        }

        else {

            if ($ctrl -ne $statusLabel) {

                $ctrl.BackColor = $bgForm
                $ctrl.ForeColor = $fgPrimary
            }
        }

        foreach ($child in $ctrl.Controls) {
            & $applyStyle $child
        }
    }

    & $applyStyle $RootControl

    if ($RootControl -eq $form) {
        Set-WindowTitleBarTheme -Form $form -Dark $Dark
    }

    if ($btnThemeToggle) {
        $btnThemeToggle.Text = if ($Dark) { "Light Mode" } else { "Dark Mode" }
    }

    if (Get-Command Update-DryRunVisuals -ErrorAction SilentlyContinue) {
        Update-DryRunVisuals
    }
}



$btnThemeToggle          = New-Object System.Windows.Forms.Button
$btnThemeToggle.Text     = "Light Mode"
$btnThemeToggle.Location = New-Object Drawing.Point(1042, 645)
$btnThemeToggle.Size     = New-Object Drawing.Size(100, 25)
$btnThemeToggle.Add_Click({
    $script:isDarkMode = -not $script:isDarkMode
    Set-GuiTheme -RootControl $form -Dark $script:isDarkMode
})
$form.Controls.Add($btnThemeToggle)

$form.Font = New-Object Drawing.Font("Segoe UI", 10)

$scaleFactor = 1.1
foreach ($ctrl in $form.Controls) {
    $ctrl.Location = New-Object Drawing.Point(
        [int]($ctrl.Location.X * $scaleFactor),
        [int]($ctrl.Location.Y * $scaleFactor)
    )
    $ctrl.Size = New-Object Drawing.Size(
        [int]($ctrl.Width * $scaleFactor),
        [int]($ctrl.Height * $scaleFactor)
    )
}
$form.Size = New-Object Drawing.Size(
    [int]($form.Width * $scaleFactor),
    [int]($form.Height * $scaleFactor)
)

Set-GuiTheme -RootControl $form -Dark $script:isDarkMode

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public static class DpiHelper {
    [DllImport("user32.dll")]
    private static extern bool SetProcessDpiAwarenessContext(IntPtr dpiContext);

    [DllImport("user32.dll")]
    private static extern bool SetProcessDPIAware();

    public static void EnableDpiAwareness() {
        try {
            // -4 = DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2
            if (!SetProcessDpiAwarenessContext(new IntPtr(-4))) {
                SetProcessDPIAware();
            }
        } catch {
            try { SetProcessDPIAware(); } catch {}
        }
    }
}
"@
[DpiHelper]::EnableDpiAwareness()

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

function Set-SafeAnchor {
    param(
        [object]$Control,
        [System.Windows.Forms.AnchorStyles]$Anchor
    )
    if ($null -ne $Control -and ($Control.PSObject.Properties.Name -contains 'Anchor')) {
        $Control.Anchor = $Anchor
    }
}

$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$form.MaximizeBox     = $true
$form.MinimumSize     = $form.Size

# Output grid expands with window
Set-SafeAnchor -Control $grid -Anchor ([System.Windows.Forms.AnchorStyles]::Top -bor
                                       [System.Windows.Forms.AnchorStyles]::Bottom -bor
                                       [System.Windows.Forms.AnchorStyles]::Left -bor
                                       [System.Windows.Forms.AnchorStyles]::Right)

# Bottom action buttons move down with the bottom edge
foreach ($btn in @($btnPing, $btnConnectVNC, $btnRDP, $btnOpenC, $btnOpenD)) {
    Set-SafeAnchor -Control $btn -Anchor ([System.Windows.Forms.AnchorStyles]::Bottom -bor
                                          [System.Windows.Forms.AnchorStyles]::Left)
}

# Save Log and Dark/Light toggle stay pinned to the bottom-right
foreach ($rBtn in @($btnSaveLog, $btnThemeToggle)) {
    Set-SafeAnchor -Control $rBtn -Anchor ([System.Windows.Forms.AnchorStyles]::Bottom -bor
                                           [System.Windows.Forms.AnchorStyles]::Right)


# Progress bar stretches horizontally above $grid to match $grid's width
Set-SafeAnchor -Control $progress -Anchor ([System.Windows.Forms.AnchorStyles]::Top -bor
                                          [System.Windows.Forms.AnchorStyles]::Left -bor
                                          [System.Windows.Forms.AnchorStyles]::Right)

# STOP button stays pinned to the top-right above $progress / $grid
Set-SafeAnchor -Control $btnStop -Anchor ([System.Windows.Forms.AnchorStyles]::Top -bor
                                         [System.Windows.Forms.AnchorStyles]::Right)

# Run Script button moves down with the bottom edge alongside the bottom row
Set-SafeAnchor -Control $btnRunScripts -Anchor ([System.Windows.Forms.AnchorStyles]::Bottom -bor
                                                [System.Windows.Forms.AnchorStyles]::Left)

# Server list stretches vertically while keeping its left margin and width
Set-SafeAnchor -Control $listViewServers -Anchor ([System.Windows.Forms.AnchorStyles]::Top -bor
                                                  [System.Windows.Forms.AnchorStyles]::Bottom -bor
                                                  [System.Windows.Forms.AnchorStyles]::Left)

# Save Selection sits below $listViewServers and tracks the bottom edge
Set-SafeAnchor -Control $btnSaveSelection -Anchor ([System.Windows.Forms.AnchorStyles]::Bottom -bor
                                                     [System.Windows.Forms.AnchorStyles]::Left)

# Keep Create Credentials dropdown button pinned to the top-right corner
Set-SafeAnchor -Control $btnCreateCredentials -Anchor ([System.Windows.Forms.AnchorStyles]::Top -bor
                                                        [System.Windows.Forms.AnchorStyles]::Right)


# Bottom-left script controls move down together with $btnRunScripts
foreach ($scriptCtrl in @(
    $chkDryRun,
    $comboScripts,
    $btnFolderCommands,
    $btnFolderDiagnostic,
    $btnFolderSql,
    $btnBrowseScript
)) {
    Set-SafeAnchor -Control $scriptCtrl -Anchor ([System.Windows.Forms.AnchorStyles]::Bottom -bor
                                                   [System.Windows.Forms.AnchorStyles]::Left)
}

# Keep status label anchored Top, Left above the grid
Set-SafeAnchor -Control $statusLabel -Anchor ([System.Windows.Forms.AnchorStyles]::Top -bor
                                              [System.Windows.Forms.AnchorStyles]::Left)
}




[void]$form.ShowDialog()
$form.Dispose()
exit