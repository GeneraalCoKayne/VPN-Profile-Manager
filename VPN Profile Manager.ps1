param(
    [switch]$Debug
)

# Clear console and show logo
Clear-Host

# Load required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Data

# --- DPI Awareness & Visual Styles ---
Add-Type -Namespace WinAPI -Name Dpi -MemberDefinition @"
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern bool SetProcessDPIAware();
    [System.Runtime.InteropServices.DllImport("Shcore.dll")]
    public static extern int SetProcessDpiAwareness(int value);
"@
try {
    # 2 = PROCESS_PER_MONITOR_DPI_AWARE, falls back gracefully on older Windows
    [void][WinAPI.Dpi]::SetProcessDpiAwareness(2)
} catch {
    try { [void][WinAPI.Dpi]::SetProcessDPIAware() } catch {}
}
[System.Windows.Forms.Application]::EnableVisualStyles()
# --- End DPI Awareness ---

function Set-StandardControlStyles([System.Windows.Forms.Control] $root) {
    foreach ($c in $root.Controls) {
        try {
            $c.AutoSize = $false
        } catch {}
        if ($c -is [System.Windows.Forms.Button] -or
            $c -is [System.Windows.Forms.TextBox] -or
            $c -is [System.Windows.Forms.ComboBox]) {
            $c.Height = 27
            try { $c.UseVisualStyleBackColor = $false } catch {}
        }
        if ($c -is [System.Windows.Forms.DataGridView]) {
            $c.RowHeadersVisible = $false
            $c.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        }
        Set-StandardControlStyles $c
    }
}


Write-Host "VPN Profile Manager v0.0.1" -ForegroundColor Green
Write-Host "Starting application..." -ForegroundColor Gray

if ($Debug) {
    Write-Host "DEBUG MODE ENABLED" -ForegroundColor Yellow
}

Write-Host ""

# Global debug variable
$global:DebugMode = $Debug

function Write-DebugLog {
    param([string]$Message)
    if ($global:DebugMode) {
        $timestamp = Get-Date -Format "HH:mm:ss"
        Write-Host "[$timestamp] DEBUG: $Message" -ForegroundColor Cyan
    }
}

$form = New-Object System.Windows.Forms.Form
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point)
$form.SuspendLayout()
$form.Text = "VPN Profile Manager v0.0.1"
$form.ClientSize = New-Object System.Drawing.Size(685, 1025)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.BackColor = [System.Drawing.Color]::WhiteSmoke
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false

function Add-FormField {
    param(
        [string]$label,
        [int]$yPos,
        [string]$defaultValue = "",
        [int]$width = 300
    )
    
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(30, $yPos)
    $lbl.Size = New-Object System.Drawing.Size(200, 25)
    $lbl.Text = $label
    $form.Controls.Add($lbl)
    
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(250, $yPos)
    $textBox.Size = New-Object System.Drawing.Size($width, 25)
    $textBox.Text = $defaultValue
    $form.Controls.Add($textBox)
    
    return $textBox
}

function Add-FormDropDown {
    param(
        [string]$label,
        [int]$yPos,
        [string[]]$options,
        [int]$width = 300
    )
    
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(30, $yPos)
    $lbl.Size = New-Object System.Drawing.Size(200, 25)
    $lbl.Text = $label
    $form.Controls.Add($lbl)
    
    $dropDown = New-Object System.Windows.Forms.ComboBox
    $dropDown.Location = New-Object System.Drawing.Point(250, $yPos)
    $dropDown.Size = New-Object System.Drawing.Size($width, 25)
    $dropDown.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    foreach ($option in $options) {
        [void]$dropDown.Items.Add($option)
    }
    if ($options.Length -gt 0) {
        $dropDown.SelectedIndex = 0
    }
    $form.Controls.Add($dropDown)
    
    return $dropDown
}

function Add-FormCheckBox {
    param(
        [string]$label,
        [int]$yPos,
        [bool]$checked = $false
    )
    
    $checkBox = New-Object System.Windows.Forms.CheckBox
    $checkBox.Location = New-Object System.Drawing.Point(250, $yPos)
    $checkBox.Size = New-Object System.Drawing.Size(400, 25)
    $checkBox.Text = $label
    $checkBox.Checked = $checked
    $form.Controls.Add($checkBox)
    
    return $checkBox
}

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Location = New-Object System.Drawing.Point(30, 20)
$titleLabel.Size = New-Object System.Drawing.Size(600, 30)
$titleLabel.Text = "VPN Profile Manager"
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($titleLabel)

$descLabel = New-Object System.Windows.Forms.Label
$descLabel.Location = New-Object System.Drawing.Point(30, 60)
$descLabel.Size = New-Object System.Drawing.Size(600, 40)
$descLabel.Text = "Add a new VPN profile to Windows" + $(if ($global:DebugMode) { " (DEBUG MODE)" } else { "" })
$descLabel.ForeColor = $(if ($global:DebugMode) { [System.Drawing.Color]::Red } else { [System.Drawing.Color]::Black })
$form.Controls.Add($descLabel)

$separatorLine = New-Object System.Windows.Forms.Panel
$separatorLine.Location = New-Object System.Drawing.Point(30, 100)
$separatorLine.Size = New-Object System.Drawing.Size(620, 2)
$separatorLine.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($separatorLine)

$nameTextBox = Add-FormField -label "Name:" -yPos 130
$serverTextBox = Add-FormField -label "Server Address:" -yPos 170

$tunnelTypeDropDown = Add-FormDropDown -label "Tunnel Type:" -yPos 210 -options @("Pptp", "L2tp", "Sstp", "Ikev2", "Automatic")
$encryptionDropDown = Add-FormDropDown -label "Encryption Level:" -yPos 250 -options @("NoEncryption", "Optional", "Required", "Maximum", "Custom")
$authMethodDropDown = Add-FormDropDown -label "Authentication Method:" -yPos 290 -options @("Pap", "Chap", "MSChapv2", "Eap", "MachineCertificate")
$l2tpPskTextBox = Add-FormField -label "L2TP Pre-shared key:" -yPos 330
$dnsSuffixTextBox = Add-FormField -label "Dns Suffix:" -yPos 370
$idleDisconnectTextBox = Add-FormField -label "Idle Disconnect Seconds:" -yPos 410
$idleDisconnectTextBox.Text = "0"

$splitTunnelingCheckBox = Add-FormCheckBox -label "Split Tunneling" -yPos 450
$allUserConnectionCheckBox = Add-FormCheckBox -label "All User Connection" -yPos 490
$rememberCredentialCheckBox = Add-FormCheckBox -label "Remember Credentials" -yPos 530
$useWinlogonCredentialCheckBox = Add-FormCheckBox -label "Use Winlogon Credentials" -yPos 570

$routeLabel = New-Object System.Windows.Forms.Label
$routeLabel.Location = New-Object System.Drawing.Point(30, 610)
$routeLabel.Size = New-Object System.Drawing.Size(600, 25)
$routeLabel.Text = "VPN Profile Routes:"
$routeLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($routeLabel)

$routesDataGridView = New-Object System.Windows.Forms.DataGridView
$routesDataGridView.Location = New-Object System.Drawing.Point(30, 640)
$routesDataGridView.Size = New-Object System.Drawing.Size(620, 120)
$routesDataGridView.AllowUserToAddRows = $false
$routesDataGridView.AllowUserToDeleteRows = $false
$routesDataGridView.AllowUserToResizeRows = $false
$routesDataGridView.MultiSelect = $false
$routesDataGridView.SelectionMode = "FullRowSelect"
$routesDataGridView.ReadOnly = $true
$routesDataGridView.ColumnHeadersHeightSizeMode = "AutoSize"
$routesDataGridView.BackgroundColor = [System.Drawing.Color]::White
$routesDataGridView.AutoSizeColumnsMode = "Fill"
$form.Controls.Add($routesDataGridView)

$routesTable = New-Object System.Data.DataTable
[void]$routesTable.Columns.Add("DestinationPrefix", [string])
[void]$routesTable.Columns.Add("RouteMetric", [string])
[void]$routesTable.Columns.Add("PassThru", [bool])

$routesDataGridView.DataSource = $routesTable

$routeGroupBox = New-Object System.Windows.Forms.GroupBox
$routeGroupBox.Location = New-Object System.Drawing.Point(30, 770)
$routeGroupBox.Size = New-Object System.Drawing.Size(620, 130)
$routeGroupBox.Text = "Add route to table"
$form.Controls.Add($routeGroupBox)

$destinationPrefixLabel = New-Object System.Windows.Forms.Label
$destinationPrefixLabel.Location = New-Object System.Drawing.Point(10, 25)
$destinationPrefixLabel.Size = New-Object System.Drawing.Size(120, 25)
$destinationPrefixLabel.Text = "IP Subnet:"
$routeGroupBox.Controls.Add($destinationPrefixLabel)

$destinationPrefixTextBox = New-Object System.Windows.Forms.TextBox
$destinationPrefixTextBox.Location = New-Object System.Drawing.Point(130, 25)
$destinationPrefixTextBox.Size = New-Object System.Drawing.Size(220, 25)
$routeGroupBox.Controls.Add($destinationPrefixTextBox)

$routeMetricLabel = New-Object System.Windows.Forms.Label
$routeMetricLabel.Location = New-Object System.Drawing.Point(10, 60)
$routeMetricLabel.Size = New-Object System.Drawing.Size(110, 25)
$routeMetricLabel.Text = "Route Metric:"
$routeGroupBox.Controls.Add($routeMetricLabel)

$routeMetricTextBox = New-Object System.Windows.Forms.TextBox
$routeMetricTextBox.Location = New-Object System.Drawing.Point(130, 60)
$routeMetricTextBox.Size = New-Object System.Drawing.Size(100, 25)
$routeMetricTextBox.Text = "1"
$routeGroupBox.Controls.Add($routeMetricTextBox)

$routePassThruCheckLabel = New-Object System.Windows.Forms.Label
$routePassThruCheckLabel.Location = New-Object System.Drawing.Point(10, 95)
$routePassThruCheckLabel.Size = New-Object System.Drawing.Size(110, 25)
$routePassThruCheckLabel.Text = "Pass Thru:"
$routeGroupBox.Controls.Add($routePassThruCheckLabel)

$routePassThruCheckBox = New-Object System.Windows.Forms.CheckBox
$routePassThruCheckBox.Location = New-Object System.Drawing.Point(130, 95)
$routePassThruCheckBox.Size = New-Object System.Drawing.Size(100, 25)
$routePassThruCheckBox.Checked = $true
$routeGroupBox.Controls.Add($routePassThruCheckBox)

$addRouteButton = New-Object System.Windows.Forms.Button
$addRouteButton.Location = New-Object System.Drawing.Point(485, 40)
$addRouteButton.Size = New-Object System.Drawing.Size(120, 30)
$addRouteButton.Text = "Add Route"
$addRouteButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$addRouteButton.ForeColor = [System.Drawing.Color]::White
$addRouteButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$routeGroupBox.Controls.Add($addRouteButton)

$removeRouteButton = New-Object System.Windows.Forms.Button
$removeRouteButton.Location = New-Object System.Drawing.Point(485, 80)
$removeRouteButton.Size = New-Object System.Drawing.Size(120, 30)
$removeRouteButton.Text = "Delete Route"
$removeRouteButton.BackColor = [System.Drawing.Color]::DarkGray
$removeRouteButton.ForeColor = [System.Drawing.Color]::White
$removeRouteButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$removeRouteButton.Enabled = $false
$routeGroupBox.Controls.Add($removeRouteButton)

$addRouteButton.Add_Click({
    $destinationPrefix = $destinationPrefixTextBox.Text.Trim()
    $routeMetric = $routeMetricTextBox.Text.Trim()
    
    Write-DebugLog "Adding route: $destinationPrefix (Metric: $routeMetric, PassThru: $($routePassThruCheckBox.Checked))"
    
    if ([string]::IsNullOrWhiteSpace($destinationPrefix)) {
        Write-DebugLog "Route validation failed: Empty destination prefix"
        [System.Windows.MessageBox]::Show("Enter a valid IP subnet (e.g., 192.168.1.0/24).", "Missing information", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    if (![string]::IsNullOrWhiteSpace($routeMetric) -and ![uint32]::TryParse($routeMetric, [ref]$null)) {
        Write-DebugLog "Route validation failed: Invalid route metric"
        [System.Windows.MessageBox]::Show("Route Metric must be a positive number.", "Invalid value", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    $newRow = $routesTable.NewRow()
    $newRow["DestinationPrefix"] = $destinationPrefix
    $newRow["RouteMetric"] = $routeMetric
    $newRow["PassThru"] = $routePassThruCheckBox.Checked
    $routesTable.Rows.Add($newRow)
    
    Write-DebugLog "Route added successfully to table"

    $destinationPrefixTextBox.Text = ""
    $routeMetricTextBox.Text = "1"
})

$routesDataGridView.Add_SelectionChanged({
    $removeRouteButton.Enabled = $routesDataGridView.SelectedRows.Count -gt 0
})

$removeRouteButton.Add_Click({
    if ($routesDataGridView.SelectedRows.Count -gt 0) {
        $selectedIndex = $routesDataGridView.SelectedRows[0].Index
        Write-DebugLog "Removing route at index: $selectedIndex"
        $routesTable.Rows.RemoveAt($selectedIndex)
    }
})

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.AutoSize = $false
$statusLabel.Dock = [System.Windows.Forms.DockStyle]::Bottom
$statusLabel.Height = 25
$statusLabel.Padding = New-Object System.Windows.Forms.Padding(5,0,0,0)
$statusLabel.ForeColor = [System.Drawing.Color]::DarkBlue
$form.Controls.Add($statusLabel)

$addButton = New-Object System.Windows.Forms.Button
$addButton.Location = New-Object System.Drawing.Point(430, 950)
$addButton.Size = New-Object System.Drawing.Size(225, 35)
$addButton.Text = "Add VPN Profile"
$addButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$addButton.ForeColor = [System.Drawing.Color]::White
$addButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($addButton)

$form.Controls.Add($addButton)
$form.Controls.SetChildIndex($addButton, 0)   # Zorg dat hij boven $statusLabel staat


$ikev2GroupBox = New-Object System.Windows.Forms.GroupBox
$ikev2GroupBox.Location = New-Object System.Drawing.Point(30, 450) 
$ikev2GroupBox.Size = New-Object System.Drawing.Size(620, 250)
$ikev2GroupBox.Text = "IKEv2 IPsec Configuration"
$ikev2GroupBox.Visible = $false
$form.Controls.Add($ikev2GroupBox)

function Update-FormPositions {
    param([bool]$showIPsec = $false)
    
    $offset = if ($showIPsec) { 265 } else { 0 }
    
    $yPosSplit = 450 + $offset
    $splitTunnelingCheckBox.Location = New-Object System.Drawing.Point(250, $yPosSplit)
    
    $yPosAllUser = 490 + $offset
    $allUserConnectionCheckBox.Location = New-Object System.Drawing.Point(250, $yPosAllUser)
    
    $yPosRemember = 530 + $offset
    $rememberCredentialCheckBox.Location = New-Object System.Drawing.Point(250, $yPosRemember)
    
    $yPosWinlogon = 570 + $offset
    $useWinlogonCredentialCheckBox.Location = New-Object System.Drawing.Point(250, $yPosWinlogon)
    
    $yPosRouteLabel = 610 + $offset
    $routeLabel.Location = New-Object System.Drawing.Point(30, $yPosRouteLabel)
    
    $yPosRoutesDataGridView = 640 + $offset
    $routesDataGridView.Location = New-Object System.Drawing.Point(30, $yPosRoutesDataGridView)
    
    $yPosRouteGroupBox = 770 + $offset
    $routeGroupBox.Location = New-Object System.Drawing.Point(30, $yPosRouteGroupBox)
    
    $yPosStatus = 890 + $offset
    $statusLabel.Location = New-Object System.Drawing.Point(30, $yPosStatus)
    
    $yPosButtons = 920 + $offset
    $addButton.Location = New-Object System.Drawing.Point(430, $yPosButtons)
    
    $newHeight = 1025 + $offset
    $form.ClientSize = New-Object System.Drawing.Size(685, $newHeight)
}

function Add-IKEv2DropDown {
    param(
        [string]$label,
        [int]$yPos,
        [string[]]$options,
        [int]$width = 300
    )
    
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(20, $yPos)
    $lbl.Size = New-Object System.Drawing.Size(200, 25)
    $lbl.Text = $label
    $ikev2GroupBox.Controls.Add($lbl)
    
    $dropDown = New-Object System.Windows.Forms.ComboBox
    $dropDown.Location = New-Object System.Drawing.Point(240, $yPos)
    $dropDown.Size = New-Object System.Drawing.Size($width, 25)
    $dropDown.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    foreach ($option in $options) {
        [void]$dropDown.Items.Add($option)
    }
    if ($options.Length -gt 0) {
        $dropDown.SelectedIndex = 0
    }
    $ikev2GroupBox.Controls.Add($dropDown)
    
    return $dropDown
}

$authTransformDropDown = Add-IKEv2DropDown -label "Authentication Transform:" -yPos 30 -options @("MD596", "SHA196", "SHA256128", "GCMAES128", "GCMAES192", "GCMAES256", "None")
$cipherTransformDropDown = Add-IKEv2DropDown -label "Cipher Transform:" -yPos 65 -options @("DES", "DES3", "AES128", "AES192", "AES256", "GCMAES128", "GCMAES192", "GCMAES256", "None")
$dhGroupDropDown = Add-IKEv2DropDown -label "DH Group:" -yPos 100 -options @("None", "Group1", "Group2", "Group14", "ECP256", "ECP384", "Group24")
$encryptionMethodDropDown = Add-IKEv2DropDown -label "Encryption Method:" -yPos 135 -options @("DES", "DES3", "AES128", "AES192", "AES256", "GCMAES128", "GCMAES256")
$integrityCheckDropDown = Add-IKEv2DropDown -label "Integrity Check Method:" -yPos 170 -options @("MD5", "SHA1", "SHA256", "SHA384")
$pfsGroupDropDown = Add-IKEv2DropDown -label "PFS Group:" -yPos 205 -options @("None", "PFS1", "PFS2", "PFS2048", "ECP256", "ECP384", "PFSMM", "PFS24")

$tunnelTypeDropDown.Add_SelectedIndexChanged({
    Write-DebugLog "Tunnel type changed to: $($tunnelTypeDropDown.SelectedItem)"
    
    # L2TP opties
    if ($tunnelTypeDropDown.SelectedItem -eq "L2tp") {
        $l2tpPskTextBox.Enabled = $true
        Write-DebugLog "L2TP PSK field enabled"
    } else {
        $l2tpPskTextBox.Enabled = $false
        $l2tpPskTextBox.Text = ""
        Write-DebugLog "L2TP PSK field disabled"
    }
    
    # IKEv2 opties
    if ($tunnelTypeDropDown.SelectedItem -eq "Ikev2") {
        $ikev2GroupBox.Visible = $true
        Update-FormPositions -showIPsec $true
        Write-DebugLog "IKEv2 configuration panel shown"
    } else {
        $ikev2GroupBox.Visible = $false
        Update-FormPositions -showIPsec $false
        Write-DebugLog "IKEv2 configuration panel hidden"
    }
})

$addButton.Add_Click({
    Write-Host ""
    Write-Host "============================================" -ForegroundColor Yellow
    Write-Host "STARTING VPN PROFILE CREATION" -ForegroundColor Yellow
    Write-Host "============================================" -ForegroundColor Yellow
    
    if ([string]::IsNullOrWhiteSpace($nameTextBox.Text)) {
        Write-Host "ERROR: VPN name is required" -ForegroundColor Red
        Write-DebugLog "Validation failed: Empty VPN name"
        [System.Windows.MessageBox]::Show("Enter a name for the VPN Profile.", "Missing information", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    if ([string]::IsNullOrWhiteSpace($serverTextBox.Text)) {
        Write-Host "ERROR: Server address is required" -ForegroundColor Red
        Write-DebugLog "Validation failed: Empty server address"
        [System.Windows.MessageBox]::Show("Enter a server address.", "Missing information", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    Write-Host "VPN Configuration:" -ForegroundColor Green
    Write-Host "  Name: $($nameTextBox.Text)" -ForegroundColor White
    Write-Host "  Server: $($serverTextBox.Text)" -ForegroundColor White
    Write-Host "  Tunnel Type: $($tunnelTypeDropDown.SelectedItem)" -ForegroundColor White
    Write-Host "  Encryption: $($encryptionDropDown.SelectedItem)" -ForegroundColor White
    Write-Host "  Authentication: $($authMethodDropDown.SelectedItem)" -ForegroundColor White
    
    Write-DebugLog "Building VPN command with parameters:"
    Write-DebugLog "  Name: $($nameTextBox.Text)"
    Write-DebugLog "  Server: $($serverTextBox.Text)"
    Write-DebugLog "  Tunnel Type: $($tunnelTypeDropDown.SelectedItem)"
    Write-DebugLog "  Encryption: $($encryptionDropDown.SelectedItem)"
    Write-DebugLog "  Auth Method: $($authMethodDropDown.SelectedItem)"
    
    $vpnCommand = "Add-VpnConnection -Name '$($nameTextBox.Text)' -ServerAddress '$($serverTextBox.Text)'"
    
    if ($tunnelTypeDropDown.SelectedItem) {
        $vpnCommand += " -TunnelType $($tunnelTypeDropDown.SelectedItem)"
    }
    
    if ($encryptionDropDown.SelectedItem) {
        $vpnCommand += " -EncryptionLevel $($encryptionDropDown.SelectedItem)"
    }
    
    if ($authMethodDropDown.SelectedItem) {
        $vpnCommand += " -AuthenticationMethod $($authMethodDropDown.SelectedItem)"
    }
    
    # Additional options
    if ($splitTunnelingCheckBox.Checked) {
        $vpnCommand += " -SplitTunneling"
        Write-Host "  Split Tunneling: Enabled" -ForegroundColor White
        Write-DebugLog "  Split Tunneling: Enabled"
    }
    
    if ($allUserConnectionCheckBox.Checked) {
        $vpnCommand += " -AllUserConnection"
        Write-Host "  All User Connection: Enabled" -ForegroundColor White
        Write-DebugLog "  All User Connection: Enabled"
    }
    
    if ($l2tpPskTextBox.Text -and $tunnelTypeDropDown.SelectedItem -eq "L2tp") {
        $vpnCommand += " -L2tpPsk '$($l2tpPskTextBox.Text)'"
        Write-Host "  L2TP PSK: Configured" -ForegroundColor White
        Write-DebugLog "  L2TP PSK: Configured"
    }
    
    if ($rememberCredentialCheckBox.Checked) {
        $vpnCommand += " -RememberCredential"
        Write-Host "  Remember Credentials: Enabled" -ForegroundColor White
        Write-DebugLog "  Remember Credentials: Enabled"
    }
    
    if ($useWinlogonCredentialCheckBox.Checked) {
        $vpnCommand += " -UseWinlogonCredential"
        Write-Host "  Use Winlogon Credentials: Enabled" -ForegroundColor White
        Write-DebugLog "  Use Winlogon Credentials: Enabled"
    }
    
    if ($dnsSuffixTextBox.Text) {
        $vpnCommand += " -DnsSuffix '$($dnsSuffixTextBox.Text)'"
        Write-Host "  DNS Suffix: $($dnsSuffixTextBox.Text)" -ForegroundColor White
        Write-DebugLog "  DNS Suffix: $($dnsSuffixTextBox.Text)"
    }
    
    if ($idleDisconnectTextBox.Text -and [int]::TryParse($idleDisconnectTextBox.Text, [ref]$null)) {
        $vpnCommand += " -IdleDisconnectSeconds $($idleDisconnectTextBox.Text)"
        Write-Host "  Idle Disconnect: $($idleDisconnectTextBox.Text) seconds" -ForegroundColor White
        Write-DebugLog "  Idle Disconnect: $($idleDisconnectTextBox.Text) seconds"
    }
    
    $vpnCommand += " -Force"
    
    Write-DebugLog "Final VPN command: $vpnCommand"
    
    try {
        $statusLabel.Text = "Adding the VPN Profile..."
        $statusLabel.ForeColor = [System.Drawing.Color]::Blue
        $form.Refresh()
        
        Write-Host ""
        Write-Host "Creating VPN Profile..." -ForegroundColor Yellow
        Write-DebugLog "Executing VPN creation command"
        
        if ($allUserConnectionCheckBox.Checked) {
            Write-Host "Running with elevated privileges..." -ForegroundColor Yellow
            Write-DebugLog "Running with elevated privileges (All User Connection)"
            $scriptBlock = [Scriptblock]::Create($vpnCommand)
            Start-Process powershell -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", "$scriptBlock" -Verb RunAs -Wait
        } else {
            Write-DebugLog "Running with current user privileges"
            Invoke-Expression $vpnCommand
        }
        
        Write-Host "✓ VPN Profile created successfully" -ForegroundColor Green
        Write-DebugLog "VPN Profile created successfully"
        
        # IKEv2 IPsec configuration
        if ($tunnelTypeDropDown.SelectedItem -eq "Ikev2") {
            $statusLabel.Text = "VPN Profile added, configuring IPsec..."
            $form.Refresh()
            
            Write-Host ""
            Write-Host "Configuring IKEv2 IPsec settings..." -ForegroundColor Yellow
            Write-Host "  Auth Transform: $($authTransformDropDown.SelectedItem)" -ForegroundColor White
            Write-Host "  Cipher Transform: $($cipherTransformDropDown.SelectedItem)" -ForegroundColor White
            Write-Host "  DH Group: $($dhGroupDropDown.SelectedItem)" -ForegroundColor White
            Write-Host "  Encryption Method: $($encryptionMethodDropDown.SelectedItem)" -ForegroundColor White
            Write-Host "  Integrity Check: $($integrityCheckDropDown.SelectedItem)" -ForegroundColor White
            Write-Host "  PFS Group: $($pfsGroupDropDown.SelectedItem)" -ForegroundColor White
            
            Write-DebugLog "Configuring IKEv2 IPsec settings:"
            Write-DebugLog "  Auth Transform: $($authTransformDropDown.SelectedItem)"
            Write-DebugLog "  Cipher Transform: $($cipherTransformDropDown.SelectedItem)"
            Write-DebugLog "  DH Group: $($dhGroupDropDown.SelectedItem)"
            Write-DebugLog "  Encryption Method: $($encryptionMethodDropDown.SelectedItem)"
            Write-DebugLog "  Integrity Check: $($integrityCheckDropDown.SelectedItem)"
            Write-DebugLog "  PFS Group: $($pfsGroupDropDown.SelectedItem)"
            
            $ipsecCommand = "Set-VpnConnectionIPsecConfiguration -ConnectionName '$($nameTextBox.Text)'"
            $ipsecCommand += " -AuthenticationTransformConstants $($authTransformDropDown.SelectedItem)"
            $ipsecCommand += " -CipherTransformConstants $($cipherTransformDropDown.SelectedItem)"
            $ipsecCommand += " -DHGroup $($dhGroupDropDown.SelectedItem)"
            $ipsecCommand += " -EncryptionMethod $($encryptionMethodDropDown.SelectedItem)"
            $ipsecCommand += " -IntegrityCheckMethod $($integrityCheckDropDown.SelectedItem)"
            $ipsecCommand += " -PfsGroup $($pfsGroupDropDown.SelectedItem)"
            
            if ($allUserConnectionCheckBox.Checked) {
                $ipsecCommand += " -AllUserConnection"
            }
            
            $ipsecCommand += " -Force"
            
            Write-DebugLog "IPsec command: $ipsecCommand"
            
            if ($allUserConnectionCheckBox.Checked) {
                Write-DebugLog "Running IPsec configuration with elevated privileges"
                $scriptBlockIPsec = [Scriptblock]::Create($ipsecCommand)
                Start-Process powershell -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", "$scriptBlockIPsec" -Verb RunAs -Wait
            } else {
                Write-DebugLog "Running IPsec configuration with current user privileges"
                Invoke-Expression $ipsecCommand
            }
            
            Write-Host "✓ IPsec configuration completed" -ForegroundColor Green
            Write-DebugLog "IPsec configuration completed"
        }
        
        # Add routes if configured
        if ($routesTable.Rows.Count -gt 0) {
            $statusLabel.Text = "VPN Profile added, setting up VPN routes..."
            $form.Refresh()
            
            Write-Host ""
            Write-Host "Adding VPN routes ($($routesTable.Rows.Count) routes)..." -ForegroundColor Yellow
            Write-DebugLog "Adding $($routesTable.Rows.Count) routes to VPN Profile"
            
            foreach ($route in $routesTable.Rows) {
                Write-Host "  Adding route: $($route.DestinationPrefix)" -ForegroundColor White
                Write-DebugLog "Adding route: $($route.DestinationPrefix) (Metric: $($route.RouteMetric), PassThru: $($route.PassThru))"
                
                $routeCommand = "Add-VpnConnectionRoute -ConnectionName '$($nameTextBox.Text)' -DestinationPrefix '$($route.DestinationPrefix)'"
                
                if (![string]::IsNullOrWhiteSpace($route.RouteMetric)) {
                    $routeCommand += " -RouteMetric $($route.RouteMetric)"
                }
                
                if ($route.PassThru) {
                    $routeCommand += " -PassThru"
                }
                
                if ($allUserConnectionCheckBox.Checked) {
                    $routeCommand += " -AllUserConnection"
                }
                
                Write-DebugLog "Route command: $routeCommand"
                
                try {
                    if ($allUserConnectionCheckBox.Checked) {
                        $scriptBlockRoute = [Scriptblock]::Create($routeCommand)
                        Start-Process powershell -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", "$scriptBlockRoute" -Verb RunAs -Wait
                    } else {
                        Invoke-Expression $routeCommand
                    }
                    Write-Host "    ✓ Route added successfully" -ForegroundColor Green
                    Write-DebugLog "Route added successfully: $($route.DestinationPrefix)"
                } catch {
                    Write-Host "    ✗ Failed to add route: $($_.Exception.Message)" -ForegroundColor Red
                    Write-DebugLog "Failed to add route $($route.DestinationPrefix): $($_.Exception.Message)"
                }
            }
            
            Write-DebugLog "All routes processed"
        }
        
        $statusLabel.Text = "VPN Profile '$($nameTextBox.Text)' has been successfully added!"
        $statusLabel.ForeColor = [System.Drawing.Color]::Green
        
        Write-Host ""
        Write-Host "============================================" -ForegroundColor Green
        Write-Host "VPN PROFILE CREATED SUCCESSFULLY!" -ForegroundColor Green
        Write-Host "Connection Name: $($nameTextBox.Text)" -ForegroundColor Green
        Write-Host "============================================" -ForegroundColor Green
        Write-Host ""
        
        Write-DebugLog "VPN Profile setup completed successfully"

        [System.Windows.MessageBox]::Show("The VPN Profile '$($nameTextBox.Text)' has been successfully added to Windows.", "VPN Profile Added", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
    }
    catch {
        $statusLabel.Text = "Error adding the VPN Profile."
        $statusLabel.ForeColor = [System.Drawing.Color]::Red
        
        Write-Host ""
        Write-Host "============================================" -ForegroundColor Red
        Write-Host "ERROR CREATING VPN PROFILE" -ForegroundColor Red
        Write-Host "============================================" -ForegroundColor Red
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host ""
        
        Write-DebugLog "ERROR: Failed to create VPN Profile: $($_.Exception.Message)"
        Write-DebugLog "ERROR: Stack trace: $($_.ScriptStackTrace)"

        [System.Windows.MessageBox]::Show("An error occurred while adding the VPN Profile:`n$($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
})

$l2tpPskTextBox.Enabled = $false

Write-Host "Application loaded successfully!" -ForegroundColor Green
Write-Host ""

Set-StandardControlStyles $form
$form.ResumeLayout($true)
[void]$form.ShowDialog()

Write-Host "Application closed." -ForegroundColor Gray