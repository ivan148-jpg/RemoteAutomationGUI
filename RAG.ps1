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
$iconPath     = Join-Path $basePath "Assets\RAG.ico"


# Ensure base folders exist (best-effort)
foreach ($p in @($basePath,$secureFolder,$scriptFolder,$listFolder,$logFolder)) {
    try { if (-not (Test-Path $p)) { New-Item -Path $p -ItemType Directory -Force | Out-Null } } catch {}
}

# Global transfer queue state
$script:ActiveCopyJobs = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:CopyTimer = $null

$script:ActiveScriptJobs = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:ScriptTimer = $null

# === RUNTIME STATE ===
$script:StopRequested     = $false
$script:SqlFullOutput     = @{}
$script:LastCheckedIndex  = $null
$script:IsBulkChecking    = $false

# === DEFAULTS ===
$lastFile         = ""
$lastServerList   = ""
$lastRemoteFolder = "C$"
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
$form.Size          = New-Object Drawing.Size(1200,700)
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
$txtFile.Size          = New-Object Drawing.Size(400,20)
$txtFile.Text          = $lastFile

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
$txtRemoteFolder.Size  = New-Object Drawing.Size(370,20)
$txtRemoteFolder.Text  = $lastRemoteFolder

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

$listViewServers                  = New-Object System.Windows.Forms.ListView
$listViewServers.Location         = New-Object Drawing.Point(20,150)
$listViewServers.Size             = New-Object Drawing.Size(400,390)
$listViewServers.View             = 'Details'
$listViewServers.CheckBoxes       = $true
$listViewServers.FullRowSelect    = $true
$listViewServers.MultiSelect      = $true
[void]$listViewServers.Columns.Add("Server", 380)

# Shift+Click range support is in Part 2

# === Selection helpers ===
$btnSaveSelection = New-Object System.Windows.Forms.Button
$btnSaveSelection.Text = "Save Selection"
$btnSaveSelection.Size = New-Object Drawing.Size(110, 25)
$btnSaveSelection.Location = New-Object Drawing.Point(310, 540)

$btnResetFilter       = New-Object System.Windows.Forms.Button
$btnResetFilter.Text  = "Clear"
$btnResetFilter.Location = New-Object Drawing.Point(248, 120)
$btnResetFilter.Size  = New-Object Drawing.Size(50, 25)
$btnResetFilter.Add_Click({
    $txtFilter.Text = ""
})

$btnCheckAll = New-Object System.Windows.Forms.Button
$btnCheckAll.Text = "Check All"
$btnCheckAll.Size = New-Object Drawing.Size(60, 25)
$btnCheckAll.Location = New-Object Drawing.Point(298, 120)

$btnUncheckAll = New-Object System.Windows.Forms.Button
$btnUncheckAll.Text = "Uncheck All"
$btnUncheckAll.Size = New-Object Drawing.Size(65, 25)
$btnUncheckAll.Location = New-Object Drawing.Point(357, 120)

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

$btnStop                  = New-Object Windows.Forms.Button
$btnStop.Text             = "Stop"
$btnStop.Location         = New-Object Drawing.Point(1080,50)
$btnStop.Size             = New-Object Drawing.Size(70,30)

$btnPing                  = New-Object Windows.Forms.Button
$btnPing.Text             = "Ping"
$btnPing.Location         = New-Object Drawing.Point(560, 90)
$btnPing.Size             = New-Object Drawing.Size(100, 25)

$btnConnectVNC            = New-Object Windows.Forms.Button
$btnConnectVNC.Text       = "Connect VNC"
$btnConnectVNC.Location   = New-Object Drawing.Point(670,90)
$btnConnectVNC.Size       = New-Object Drawing.Size(100,25)

$btnRDP                   = New-Object Windows.Forms.Button
$btnRDP.Text              = "Connect RDP"
$btnRDP.Location          = New-Object Drawing.Point(780,90)
$btnRDP.Size              = New-Object Drawing.Size(100,25)

$btnOpenC                 = New-Object Windows.Forms.Button
$btnOpenC.Text            = "Open C$"
$btnOpenC.Location        = New-Object Drawing.Point(890,90)
$btnOpenC.Size            = New-Object Drawing.Size(70,25)

$btnOpenD                 = New-Object Windows.Forms.Button
$btnOpenD.Text            = "Open D$"
$btnOpenD.Location        = New-Object Drawing.Point(970,90)
$btnOpenD.Size            = New-Object Drawing.Size(70,25)

# === Progress + grid ===
$progress                        = New-Object Windows.Forms.ProgressBar
$progress.Location               = New-Object Drawing.Point(450,123)
$progress.Size                   = New-Object Drawing.Size(670,20)
$progress.Minimum                = 0

$grid                            = New-Object Windows.Forms.DataGridView
$grid.Location                   = New-Object Drawing.Point(450,150)
$grid.Size                       = New-Object Drawing.Size(670,390)
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
$comboScripts.Size               = New-Object Drawing.Size(250,20)
$comboScripts.DropDownStyle      = "DropDown"
$comboScripts.AutoCompleteMode   = "SuggestAppend"
$comboScripts.AutoCompleteSource = "ListItems"

$comboScripts.Add_TextChanged({
    $match = $comboScripts.Items | Where-Object { $_.Display -eq $comboScripts.Text }
    if ($match) { $comboScripts.SelectedItem = $match }
})



$btnBrowseScript                 = New-Object Windows.Forms.Button
$btnBrowseScript.Text            = "Browse Scripts..."
$btnBrowseScript.Location        = New-Object Drawing.Point(280,578)
$btnBrowseScript.Size            = New-Object Drawing.Size(120,25)

$btnRunScripts                   = New-Object Windows.Forms.Button
$btnRunScripts.Text              = "Run Script"
$btnRunScripts.Location          = New-Object Drawing.Point(20,550)
$btnRunScripts.Size              = New-Object Drawing.Size(100,30)

$btnRestartMachine            = New-Object Windows.Forms.Button
$btnRestartMachine.Text       = " Restart Machine"
$btnRestartMachine.Location   = New-Object Drawing.Point(990, 610)
$btnRestartMachine.Size       = New-Object Drawing.Size(135, 32)

# 1. Create the status label first
$statusLabel = New-Object Windows.Forms.Label
$statusLabel.ForeColor = [System.Drawing.Color]::Green
$statusLabel.Text = "Ready"
$statusLabel.Location = New-Object Drawing.Point(450, 550)
$statusLabel.Size = New-Object Drawing.Size(300, 30)
$statusLabel.Font = New-Object Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)


# 2. Create the checkbox
$chkDryRun = New-Object System.Windows.Forms.CheckBox
$chkDryRun.Text = "Dry Run (simulate only)"
$chkDryRun.Checked = $false
$chkDryRun.AutoSize = $true
$chkDryRun.Location = New-Object System.Drawing.Point(22, 605)
$chkDryRun.Font = New-Object System.Drawing.Font("Segoe UI", 6)

# 3. Hook up the event
$chkDryRun.Add_CheckedChanged({
    if ($chkDryRun.Checked) {
        $statusLabel.Text = "🧪 Dry Run Mode Active"
        $statusLabel.ForeColor = [System.Drawing.Color]::DarkOrange
    } else {
        $statusLabel.Text = "Ready"
        $statusLabel.ForeColor = [System.Drawing.Color]::Green
    }
})

# 4. Add tooltip
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.SetToolTip($chkDryRun, "Simulates actions without executing them. Safe for testing.")

# 5. Set dry run ON by default
$chkDryRun.Checked = $true  # This will now trigger the CheckedChanged event and update the label10, [System.Drawing.FontStyle]::Bold)

$btnSaveLog = New-Object Windows.Forms.Button
$btnSaveLog.Text     = "Save Log"
$btnSaveLog.Location = New-Object Drawing.Point(750,570)
$btnSaveLog.Size     = New-Object Drawing.Size(120,30)

# === Global button font ===
$buttonFont = New-Object System.Drawing.Font("Segoe UI", 8)
@(
    $btnCopy, $btnRunFile, $btnRunScripts, $btnExecSQL, $btnStop,
    $btnConnectVNC, $btnPing, $btnRDP, $btnOpenC, $btnOpenD,
    $btnSaveSelection, $btnCheckAll, $btnUncheckAll, $btnSaveSelection,
    $btnBrowse, $btnBrowseServers, $btnBrowseScript, $btnCreateCred, $btnCreateSqlCred, $btnEncryptVnc,$btnResetFilter, $btnRestartMachine  
) | ForEach-Object { $_.Font = $buttonFont }

# === Unicode font for controls with special characters ===
$unicodeFont = New-Object System.Drawing.Font("Segoe UI", 9)

$listViewServers.Font     = $unicodeFont
$cmbServers.Font          = $unicodeFont
$txtFilter.Font           = $unicodeFont
$txtFile.Font             = $unicodeFont
$txtRemoteFolder.Font     = $unicodeFont
$comboScripts.Font        = $unicodeFont


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
            $row.Cells[1].Style.ForeColor = $Color
            $row.Cells[2].Style.ForeColor = $Color
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

# ListView range-check support (Shift+Click)
function Get-ListViewItemIndexAtCursor {
    $pt  = $listViewServers.PointToClient([System.Windows.Forms.Cursor]::Position)
    $hit = $listViewServers.HitTest($pt)
    if ($hit -and $hit.Item) { return $hit.Item.Index } else { return $null }
}

$listViewServers.Add_ItemCheck({
    param($sender, $e)
    if ($script:IsBulkChecking) { return }
    
    $shiftPressed = ([System.Windows.Forms.Control]::ModifierKeys -band [System.Windows.Forms.Keys]::Shift) -ne 0
    if ($shiftPressed -and $script:LastCheckedIndex -ne $null) {
        $script:IsBulkChecking = $true
        try {
            $start    = [Math]::Min($script:LastCheckedIndex, $e.Index)
            $end      = [Math]::Max($script:LastCheckedIndex, $e.Index)
            $newState = ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked)
            for ($i = $start; $i -le $end; $i++) { 
                $listViewServers.Items[$i].Checked = $newState 
            }
        } finally { 
            $script:IsBulkChecking = $false 
        }
    }
    $script:LastCheckedIndex = $e.Index
})

$listViewServers.Add_MouseDown({ $script:LastCheckedIndex = Get-ListViewItemIndexAtCursor })


# Load servers from a list file name found in path map
function Load-ServerList {
    param([string]$listName)
    if (-not $listName) { return }
    if (-not $serverListPathMap.ContainsKey($listName)) { return }
    $selectedListPath = $serverListPathMap[$listName]
    if (-not (Test-Path $selectedListPath)) { return }

    $listViewServers.BeginUpdate()
    try {
        $listViewServers.Items.Clear()
        Get-Content $selectedListPath | ForEach-Object {
            if (-not [string]::IsNullOrWhiteSpace($_)) {
                $parts = $_ -split '\|'
                $ip    = $parts[0].Trim()
                $label = if ($parts.Count -gt 1) { $parts[1].Trim() } else { "" }
                $display = if ($label) { "{0} | {1}" -f $ip, $label } else { $ip }
                $item = New-Object System.Windows.Forms.ListViewItem($display)
                [void]$listViewServers.Items.Add($item)
            }
        }
    } finally { $listViewServers.EndUpdate() }
}



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

            if ($job.State -eq 'Completed' -or $job.State -eq 'Failed' -or $job.State -eq 'Stopped') {
                # Retrieve result if not already processed
                $data = Receive-Job $job -ErrorAction SilentlyContinue
                Remove-Job $job -Force -ErrorAction SilentlyContinue

                # Mark as processed by clearing job reference
                $item.Job = $null

                # Update grid UI safely on main thread
                foreach ( $row in $grid.Rows ) {
                    if ($row.Cells[0].Value -eq $srv) {
                        if ($data.Success) {
                            $row.Cells[1].Value = "✅ Copied"
                            $row.Cells[1].Style.ForeColor = [System.Drawing.Color]::Green
                            $row.Cells[2].Value = $data.Message
                        } else {
                            $row.Cells[1].Value = "❌ Copy failed"
                            $row.Cells[1].Style.ForeColor = [System.Drawing.Color]::Red
                            $row.Cells[2].Value = $data.Message
                        }
                        break
                    }
                }

                $progress.Value = [Math]::Min($progress.Value + 1, $progress.Maximum)
            }
            elseif ($null -ne $job) {
                # Still running
                $allDone = $false
            }
        }

        # If all jobs finished, stop and cleanup timer
        if ($allDone) {
            $script:CopyTimer.Stop()
            $script:CopyTimer.Dispose()
            $script:CopyTimer = $null
        }
    })

    $script:CopyTimer.Start()
}

# Invoke a remote file on servers, route execution by extension
function Run-OnServers($remotePath, $servers, $cred) {
    $progress.Maximum = $servers.Count
    $progress.Value   = 0
    $extension = [System.IO.Path]::GetExtension($remotePath).ToLower()

    foreach ($srv in $servers) {
        if ($script:StopRequested) {
            Update-GridStatus $srv "⏹️ Stopped" ([System.Drawing.Color]::Gray) "User interrupted"
            break
        }

        $progress.Value += 1
        $form.Refresh(); [System.Windows.Forms.Application]::DoEvents()

        try {
            $out = Invoke-Command -ComputerName $srv -Credential $cred -ErrorAction Stop -ScriptBlock {
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
            } -ArgumentList $remotePath, $extension

            $outputText = ($out | Out-String).Trim()
            $status = if ($outputText -match "❌|error|not found") { "❌ Failed" } else { "✅ Success" }
            $color  = if ($status -like "✅*") { [System.Drawing.Color]::Green } else { [System.Drawing.Color]::Red }
            Update-GridStatus $srv $status $color $outputText
        }
        catch {
            Update-GridStatus $srv "❌ Run failed" ([System.Drawing.Color]::Red) $_.Exception.Message
        }

    }
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
                Server = $srv
                Status = ""
                Output = ""
                IsError = $false
            }

            try {
                if ($UsesSqlAuth) {
                    $res = & $ScriptPath -FilePath $FilePath -Server $srv -SqlUser $SqlUser -SqlPass $SqlPass -UseSqlAuth
                } else {
                    $res = & $ScriptPath -FilePath $FilePath -Server $srv
                }

                $resultObject.Status = $res.Status
                $resultObject.Output = $res.Output
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
                        $statusText = if ($data.Status) { $data.Status } else { "❌ Script Error" }
                        $outputMsg  = if ($data.Output) { $data.Output } else { "No output returned" }
                        
                        $color = if ($statusText -like "✅*") { [System.Drawing.Color]::Green } else { [System.Drawing.Color]::Red }
                        
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
            $statusLabel.Text = "Script Execution Completed"
        }
    })

    $script:ScriptTimer.Start()
}

function Execute-SqlOnServers {
    param (
        [string[]]$Servers,
        [string]$SqlFile,
        [string]$SqlText,
        [string]$SqlUser,
        [string]$SqlPass,
        [bool]$IsDryRun = $false
    )

         if ($IsDryRun) {
        $simulated = Simulate-DryRun -Action "Would execute SQL" `
            -DetailsTemplate "Simulated SQL query '{target}' on {server}" `
            -Servers $Servers -Target $SqlText

        Show-DryRunSummary $simulated
        $statusLabel.Text = "Dry Run Completed"
        return
    }

    $grid.Rows.Clear()
    foreach ($s in $Servers) {
        [void]$grid.Rows.Add($s, "Pending", "")
    }

    $sqlBatches = $SqlText -split "(?m)^\s*GO\s*$"
    $progress.Value = 0
    $progress.Maximum = $Servers.Count
    $statusLabel.Text = "Executing SQL via .NET SqlClient..."
    $form.Refresh(); [System.Windows.Forms.Application]::DoEvents()

    foreach ($srv in $Servers) {
        if ($script:StopRequested) {
            Update-GridStatus $srv "⏹️ Stopped" ([System.Drawing.Color]::Gray) "User interrupted"
            break
        }

        $progress.Value += 1
        $form.Refresh(); [System.Windows.Forms.Application]::DoEvents()

        try {
            $connString = "Server=$srv;Database=master;User ID=$SqlUser;Password=$SqlPass;TrustServerCertificate=True;Connection Timeout=15"
            $conn = New-Object System.Data.SqlClient.SqlConnection $connString
            $conn.Open()
            $cmd = $conn.CreateCommand()
            $cmd.CommandTimeout = 30

            $results = @()
            foreach ($batch in $sqlBatches) {
                if ($batch.Trim()) {
                    $cmd.CommandText = $batch
                    try {
                        $reader = $cmd.ExecuteReader()
                        while ($reader.Read()) {
                            $row = @()
                            for ($i = 0; $i -lt $reader.FieldCount; $i++) {
                                $row += $reader.GetValue($i)
                            }
                            $results += ($row -join ", ")
                        }
                        $reader.Close()
                    } catch {
                        try {
                            $cmd.ExecuteNonQuery() | Out-Null
                            $results += "Non-query batch executed successfully."
                        } catch {
                            $results += "❌ Batch failed: $($_.Exception.Message)"
                        }
                    }
                }
            }

            $conn.Close()

            $ts = Get-Date -Format "yyyyMMdd_HHmmss"
            if (-not (Test-Path $logFolder)) {
                New-Item -Path $logFolder -ItemType Directory | Out-Null
            }
            $logFile = Join-Path $logFolder "SqlResult_${srv}_$ts.txt"
            $results | Set-Content -Path $logFile -Encoding UTF8

            $outputText = if ($results.Count -gt 0) { $results -join "`n" } else { "Executed successfully with no result set." }
            Update-GridStatus $srv "✅ Success" ([System.Drawing.Color]::Green) $outputText
        } catch {
            Update-GridStatus $srv "❌ Error" ([System.Drawing.Color]::Red) $_.Exception.Message
        }
    }

    $statusLabel.Text = if (-not $script:StopRequested) { "SQL Execution Completed" } else { "Stopped by user" }
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
            Display = $f.BaseName   # what user sees in dropdown
            Path    = $full         # full path for execution
            Rel     = $display      # optional: relative path
        }
        $comboScripts.Items.Add($item) | Out-Null
    }

    if ($comboScripts.Items.Count -gt 0) { $comboScripts.SelectedIndex = 0 }

    # Tell ComboBox which property to show
    $comboScripts.DisplayMember = 'Display'
}



# Simple details dialog to inspect large cell contents
function Show-DetailsDialog {
    param([string]$Title, [string]$FullText)

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = $Title
    $dlg.Size = New-Object System.Drawing.Size(900,600)
    $dlg.StartPosition = "CenterParent"

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Multiline = $true
    $txt.ReadOnly = $true
    $txt.ScrollBars = "Both"
    $txt.WordWrap = $false
    $txt.Dock = "Fill"
    $txt.Font = New-Object System.Drawing.Font("Consolas",9)
    $txt.Text = $FullText

    # --- Fix: clear selection so text is not highlighted ---
    $txt.SelectionStart = 0
    $txt.SelectionLength = 0

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = "Save As..."
    $btnSave.Dock = "Bottom"
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

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Multiline = $true
    $txt.ReadOnly = $true
    $txt.ScrollBars = "Vertical"
    $txt.Dock = "Fill"
    $txt.Font = New-Object System.Drawing.Font("Consolas",9)
    $txt.Text = ($Lines -join "`r`n")

    # Fix: clear selection so text is not highlighted
    $txt.SelectionStart = 0
    $txt.SelectionLength = 0

    $dlg.Controls.Add($txt)
    [void]$dlg.ShowDialog()
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
if ($lastScriptFolder -and (Test-Path $lastScriptFolder) -and 
    (Get-ChildItem -Path $lastScriptFolder -Filter '*.ps1' -ErrorAction SilentlyContinue)) {
    Load-ScriptsFromFolder $lastScriptFolder
} else {
    Load-ScriptsFromFolder $scriptFolder
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

# Browse script folder and reload (keeps recursion enabled)
$btnBrowseScript.Add_Click({
    $dialog = New-Object Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Select a folder containing PowerShell scripts"
    $dialog.SelectedPath = $lastScriptFolder
    if ($dialog.ShowDialog() -eq [Windows.Forms.DialogResult]::OK) {
        $lastScriptFolder = $dialog.SelectedPath
        Load-ScriptsFromFolder -folder $lastScriptFolder -Recurse $true
        Save-Config
    }
})

# Server list changed -> reload
$cmbServers.Add_SelectedIndexChanged({
    Load-ServerList $cmbServers.SelectedItem
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


$script:CachedServers = @()

$txtFilter.Add_TextChanged({
    $listViewServers.BeginUpdate()
    try {
        $filter = $txtFilter.Text.Trim()

        # Cache master server list once
        if (-not $script:CachedServers) {
            $selected = $cmbServers.SelectedItem
            if (-not $selected) { return }
            $selectedListPath = $serverListPathMap[$selected]
            if (-not (Test-Path $selectedListPath)) { return }
            $script:CachedServers = Get-Content $selectedListPath
        }

        # Remember which servers were checked in the CURRENT view before filtering
        $checkedBefore = @()
        foreach ($item in $listViewServers.Items) {
            if ($item.Checked) { $checkedBefore += $item.Text }
        }

        $listViewServers.Items.Clear()

        # Filter the cached array
        $filtered = if ([string]::IsNullOrWhiteSpace($filter)) { 
            $script:CachedServers 
        } else { 
            $script:CachedServers | Where-Object { $_ -like "*$filter*" } 
        }

        # Rebuild items; if filter is empty, everything starts unchecked (fresh reset)
        $newListItems = foreach ($srv in $filtered) {
            $i = New-Object System.Windows.Forms.ListViewItem($srv)
            if (-not [string]::IsNullOrWhiteSpace($filter) -and ($checkedBefore -contains $srv)) {
                $i.Checked = $true
            }
            $i
        }

        if ($newListItems) {
            $listViewServers.Items.AddRange($newListItems)
        }
    } 
    finally { 
        $listViewServers.EndUpdate() 
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
    foreach ($item in $listViewServers.Items) {
        $item.Checked = $true
    }

    # Automatically enable dry run
    $chkDryRun.Checked = $true
    $statusLabel.Text = "🧪 Dry Run Enabled (via Check All)"
    $statusLabel.ForeColor = [System.Drawing.Color]::DarkOrange
})

$btnUncheckAll.Add_Click({
    foreach ($item in $listViewServers.Items) { $item.Checked = $false }
})

$btnCopy.Add_Click({
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
    $scriptContent = Get-Content $scriptPath -Raw
    $usesSqlAuth = $scriptContent -match 'Invoke-Sqlcmd' -or $scriptContent -match '\$UseSqlAuth'

    # Initialize empty defaults so they always exist in scope
    $sqlUser = ""
    $sqlPass = ""

    if ($usesSqlAuth) {
        $sqlUser = Get-Content "$secureFolder\sqluser.txt"
        $securePass = Get-Content "$secureFolder\sqlpass.txt" | ConvertTo-SecureString
        $cred = New-Object System.Management.Automation.PSCredential($sqlUser, $securePass)
        $sqlPass = $cred.GetNetworkCredential().Password
    }

    Run-ScriptOnServers -Servers $selectedServers -ScriptPath $scriptPath -ScriptText $scriptText -FilePath $filePath `
    -UsesSqlAuth:$usesSqlAuth -SqlUser $sqlUser -SqlPass $sqlPass -IsDryRun:$chkDryRun.Checked
})

$btnRunFile.Add_Click({
    $script:SqlFullOutput = @{}
    $script:StopRequested = $false
    $statusLabel.Text     = "Running file..."
    $form.Refresh(); [System.Windows.Forms.Application]::DoEvents()

    Save-Config

    $remotePath = $txtRemoteFolder.Text.Trim()
    if (-not $remotePath) {
        [System.Windows.Forms.MessageBox]::Show("Please enter the full destination path, e.g. C:\Deploy\Script.bat", "Input Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    # Convert C$\... to C:\... (when used inside remote session path resolution)
    if ($remotePath -match "^[a-zA-Z]\$\\") {
        $drive = $remotePath.Substring(0,1)
        $sub   = $remotePath.Substring(3)
        $remotePath = "$drive`:\$sub"
    }

    $servers = Get-SelectedServers
    if ($servers.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Please select at least one server."); return }

    try { $cred = Import-Clixml -Path $fixedCredPath } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to load credentials from:`n$fixedCredPath")
        return
    }

    if ($chkDryRun.Checked) {
    $fileName = Split-Path $txtFile.Text -Leaf
    $simulated = Simulate-DryRun -Action "Would run file" `
    -DetailsTemplate "Simulated execution of {target} on {server}" `
    -Servers $servers -Target $fileName

    Show-DryRunSummary $simulated
    $statusLabel.Text = "Dry Run Completed"
    return
}

    $grid.Rows.Clear()
    foreach ($s in $servers) { [void]$grid.Rows.Add($s, "Pending", "") }

    Run-OnServers $remotePath $servers $cred
    if (-not $script:StopRequested) { $statusLabel.Text = "Run Completed" }
    $form.Refresh(); [System.Windows.Forms.Application]::DoEvents()
})

$btnExecSQL.Add_Click({
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
    $sqlUser = Get-Content "$secureFolder\sqluser.txt"
    $securePass = Get-Content "$secureFolder\sqlpass.txt" | ConvertTo-SecureString
    $cred = New-Object System.Management.Automation.PSCredential($sqlUser, $securePass)
    $sqlPass = $cred.GetNetworkCredential().Password

    Execute-SqlOnServers -Servers $servers -SqlFile $sqlFile -SqlText $sqlText -SqlUser $sqlUser -SqlPass $sqlPass -IsDryRun:$chkDryRun.Checked
})

$btnStop.Add_Click({
    $script:StopRequested = $true
    $statusLabel.Text = "Stopping active operations..."
    $form.Refresh(); [System.Windows.Forms.Application]::DoEvents()

    # Stop and remove any active copy jobs
    if ($script:ActiveCopyJobs) {
        foreach ($item in $script:ActiveCopyJobs) {
            if ($item.Job) {
                if ($item.Job.State -eq 'Running') {
                    Stop-Job $item.Job -ErrorAction SilentlyContinue
                }
                Remove-Job $item.Job -Force -ErrorAction SilentlyContinue
            }
        }
        $script:ActiveCopyJobs.Clear()
    }

    # Stop and remove any active script jobs
    if ($script:ActiveScriptJobs) {
        foreach ($item in $script:ActiveScriptJobs) {
            if ($item.Job) {
                if ($item.Job.State -eq 'Running') {
                    Stop-Job $item.Job -ErrorAction SilentlyContinue
                }
                Remove-Job $item.Job -Force -ErrorAction SilentlyContinue
            }
        }
        $script:ActiveScriptJobs.Clear()
    }

    # Stop timers if active
    if ($script:CopyTimer) {
        $script:CopyTimer.Stop()
        $script:CopyTimer.Dispose()
        $script:CopyTimer = $null
    }
    if ($script:ScriptTimer) {
        $script:ScriptTimer.Stop()
        $script:ScriptTimer.Dispose()
        $script:ScriptTimer = $null
    }

    # Immediately mark any unfinished "Pending" rows as Stopped in the grid
    foreach ($row in $grid.Rows) {
        if ($row.Cells[1].Value -eq "Pending") {
            $row.Cells[1].Value = "⏹️ Stopped"
            $row.Cells[1].Style.ForeColor = [System.Drawing.Color]::Gray
            $row.Cells[2].Value = "Operation cancelled by user"
        }
    }

    $statusLabel.Text = "Stopped by user"
})

$btnConnectVNC.Add_Click({
    $servers = Get-SelectedServers
    if ($servers.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Please select at least one server."); return }

    if ($chkDryRun.Checked) {
    $simulated = Simulate-DryRun -Action "Would connect VNC" `
        -DetailsTemplate "Simulated VNC connection to {server}" `
        -Servers $servers -Target ""

    Show-DryRunSummary $simulated
    $statusLabel.Text = "Dry Run Completed"
    return
}

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
        $color    = if ($status -like "✅*") { [System.Drawing.Color]::Green } else { [System.Drawing.Color]::Red }
        $grid.Rows[$rowIndex].Cells[1].Style.ForeColor = $color
    }
    $statusLabel.Text = "Ready"
})

$btnPing.Add_Click({
    $script:StopRequested = $false
    $servers = Get-SelectedServers
    if ($servers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one server.")
        return
    }

    $script:SqlFullOutput = @{}
    $grid.Rows.Clear()
    foreach ($s in $servers) {
        [void]$grid.Rows.Add($s, "Pending", "")
    }

    $progress.Value   = 0
    $progress.Maximum = $servers.Count
    $statusLabel.Text = "Pinging servers..."
    $form.Refresh(); [System.Windows.Forms.Application]::DoEvents()

    foreach ($srv in $servers) {
        if ($script:StopRequested) {
            Update-GridStatus $srv "⏹️ Stopped" ([System.Drawing.Color]::Gray) "User interrupted"
            $statusLabel.Text = "Stopped by user"
            break
        }

        $progress.Value += 1
        $form.Refresh(); [System.Windows.Forms.Application]::DoEvents()

        try {
            $ping = Test-Connection -ComputerName $srv -Count 1 -Quiet -ErrorAction Stop
            $status = if ($ping) { "✅ Reachable" } else { "❌ No Response" }
            $color  = if ($ping) { [System.Drawing.Color]::Green } else { [System.Drawing.Color]::Red }
            $message = "Ping to ${srv}: $status"
            Update-GridStatus $srv $status $color $message
        } catch {
            Update-GridStatus $srv "❌ Error" ([System.Drawing.Color]::Red) $_.Exception.Message
        }

        Start-Sleep -Milliseconds 200
    }

    if (-not $script:StopRequested) {
        $statusLabel.Text = "Ping Completed"
    }
})

$btnRDP.Add_Click({
    $servers = Get-SelectedServers
    if ($servers.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Please select at least one server."); return }
    
    if ($chkDryRun.Checked) {
    $simulated = Simulate-DryRun -Action "Would connect RDP" `
        -DetailsTemplate "Simulated RDP connection to {server}" `
        -Servers $servers -Target ""

    Show-DryRunSummary $simulated
    $statusLabel.Text = "Dry Run Completed"
    return
}
    
    foreach ($srv in $servers) {
        $srv = $srv.Trim(); if (-not $srv) { continue }
        try {
            Start-Process -FilePath "mstsc.exe" -ArgumentList "/v:$srv"
            $rowIndex = $grid.Rows.Add($srv, "✅ Connected", "RDP session launched")
            $grid.Rows[$rowIndex].Cells[1].Style.ForeColor = [System.Drawing.Color]::Green
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
            $grid.Rows[$rowIndex].Cells[1].Style.ForeColor = [System.Drawing.Color]::Green
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
            $grid.Rows[$rowIndex].Cells[1].Style.ForeColor = [System.Drawing.Color]::Green
        } catch {
            $rowIndex = $grid.Rows.Add($srv, "❌ Open failed", $_.Exception.Message)
            $grid.Rows[$rowIndex].Cells[1].Style.ForeColor = [System.Drawing.Color]::Red
        }
    }
})

$btnCreateCred.Add_Click({
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

        # === 1. Save to adminCred.xml ===
        $ok = Save-FixedCredential -Cred $cred
        if ($ok) {
            [System.Windows.Forms.MessageBox]::Show(
                "Credentials saved to:`n$fixedCredPath",
                "Saved",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }

        # === 2. Prepare values for cmdkey ===
        $userName      = $cred.UserName
        $plainPassword = $cred.GetNetworkCredential().Password
        $servers       = Get-SelectedServers
        $successCount  = 0
        $failCount     = 0

        foreach ($srv in $servers) {
            try {
                # Delete any old entry first
                cmdkey /delete:$srv | Out-Null

                # Add fresh entry
                cmdkey /add:$srv /user:$userName /pass:$plainPassword

                Update-GridStatus $srv "✅ Credential Added" ([System.Drawing.Color]::Green) "Stored Windows credentials"
                $successCount++
            }
            catch {
                Update-GridStatus $srv "❌ Failed" ([System.Drawing.Color]::Red) $_.Exception.Message
                $failCount++
            }
        }

        # === 3. Show summary popup ===
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
})

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

$btnCreateSqlCred.Add_Click({
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
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Unexpected error: $($_.Exception.Message)","Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
    }
})



$restartIconPath = Join-Path $basePath "Assets\Restart_icon.png" # or Assets\Restart_icon.png depending on your exact tree

if (Test-Path $restartIconPath) {
    try {
        $imgRest = [System.Drawing.Image]::FromFile($restartIconPath)
        $iconSize = 24
        $scaledImgRest = New-Object System.Drawing.Bitmap($iconSize, $iconSize)
        $gRest = [System.Drawing.Graphics]::FromImage($scaledImgRest)
        $gRest.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $gRest.DrawImage($imgRest, 0, 0, $iconSize, $iconSize)
        $gRest.Dispose()
        $imgRest.Dispose()

        $btnRestartMachine.Image = $scaledImgRest
        $btnRestartMachine.ImageAlign = [System.Drawing.ContentAlignment]::MiddleLeft
        $btnRestartMachine.TextImageRelation = [System.Windows.Forms.TextImageRelation]::ImageBeforeText
    } catch { }
}

$btnRestartMachine.Add_Click({
    $script:StopRequested = $false
    Save-Config

    $selectedServers = Get-SelectedServers
    if ($selectedServers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one server.")
        return
    }

    if ($chkDryRun.Checked) {
        $simulated = Simulate-DryRun -Action "Would restart remote machine(s)" `
            -DetailsTemplate "Simulated Restart-Computer -Force on {server}" `
            -Servers $selectedServers -Target "OS Reboot"

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

    $grid.Rows.Clear()
    foreach ($s in $selectedServers) {
        [void]$grid.Rows.Add($s, "Pending", "")
    }

    $progress.Value = 0
    $progress.Maximum = $selectedServers.Count
    $statusLabel.Text = "Sending Restart Commands..."
    $form.Refresh(); [System.Windows.Forms.Application]::DoEvents()

    foreach ($srv in $selectedServers) {
        if ($script:StopRequested) {
            Update-GridStatus $srv "⏹️ Stopped" ([System.Drawing.Color]::Gray) "User interrupted"
            break
        }

        $progress.Value += 1
        $form.Refresh(); [System.Windows.Forms.Application]::DoEvents()

        try {
            Invoke-Command -ComputerName $srv -Credential $cred -ErrorAction Stop -ScriptBlock {
                Restart-Computer -Force
            }

            Update-GridStatus $srv "✅ Restarting" ([System.Drawing.Color]::Green) "Restart command sent successfully"
        }
        catch {
            $err = $_.Exception.Message
            Update-GridStatus $srv "❌ Failed" ([System.Drawing.Color]::Red) $err
        }
    }

    $statusLabel.Text = if (-not $script:StopRequested) { "Restart Machine Completed" } else { "Stopped by user" }
    $form.Refresh(); [System.Windows.Forms.Application]::DoEvents()
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

$form.Controls.Add($btnRestartMachine)

# === Scale entire GUI by 1.1 (10% larger) ===
$scaleFactor = 1.1

# Make sure all controls are attached before scaling
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

# Optionally bump the base font slightly
$form.Font = New-Object Drawing.Font("Segoe UI", 9.5)



[void]$form.ShowDialog()
$form.Dispose()
exit