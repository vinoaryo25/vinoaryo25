Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 1. Gather System Information
$computerSystem = Get-CimInstance Win32_ComputerSystem
$osInfo         = Get-CimInstance Win32_OperatingSystem
$cpuInfo        = Get-CimInstance Win32_Processor | Select-Object -First 1
$biosInfo       = Get-CimInstance Win32_Bios
$videoInfo      = Get-CimInstance Win32_VideoController | Select-Object -First 1

# Get ALL local hard drives/partitions & calculate individual + total sizes
$disks = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3"
$diskList = @()
$totalStorageGB = 0

foreach ($disk in $disks) {
    $sizeGB = [math]::Round($disk.Size / 1GB)
    $diskList += "$($disk.DeviceID) $sizeGB GB"
    $totalStorageGB += $sizeGB  # Sum up the total capacity
}

# Combine into one clean string with the grand total at the end
$allStorage = ($diskList -join " | ") + " (Total: $totalStorageGB GB)"

$ramGB = [math]::Round($computerSystem.TotalPhysicalMemory / 1GB)
$resolution = if ($videoInfo.CurrentHorizontalResolution) {
    "$($videoInfo.CurrentHorizontalResolution)x$($videoInfo.CurrentVerticalResolution)"
} else {
    "Detected via OS"
}

# Map data exactly to your Excel headers
$data = [ordered]@{
    "Computer Name" = $env:COMPUTERNAME
    "Merk / Tipe"   = "$($computerSystem.Manufacturer) - $($computerSystem.Model)"
    "OS"            = $osInfo.Caption
    "Processor"     = $cpuInfo.Name.Trim()
    "RAM"           = "$ramGB GB"
    "Storage"       = $allStorage
    "Layar"         = "$resolution ($($videoInfo.Name))"
    "Serial Number" = $biosInfo.SerialNumber
}

# 2. Build the GUI Window
$form = New-Object System.Windows.Forms.Form
$form.Text = "System Info - Stock Opname Tool"
$form.Size = [System.Drawing.Size]::new(520, 440)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::White
$form.ControlBox = $false

# Styling Fonts
$fontLabel = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$fontValue = New-Object System.Drawing.Font("Segoe UI", 9)

$y = 20

# Generate Labels and Selectable Textboxes dynamically
foreach ($key in $data.Keys) {
    # Label
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $key
    $lbl.Location = [System.Drawing.Point]::new(20, $y)
    $lbl.Size = [System.Drawing.Size]::new(120, 23)
    $lbl.Font = $fontLabel
    $lbl.ForeColor = [System.Drawing.Color]::DarkSlateGray
    $form.Controls.Add($lbl)

    # TextBox (ReadOnly so users can still highlight & copy text if needed)
    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Text = $data[$key]
    $txt.Location = [System.Drawing.Point]::new(150, $y - 3)
    $txt.Size = [System.Drawing.Size]::new(330, 23)
    $txt.Font = $fontValue
    $txt.ReadOnly = $true
    $txt.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
    $form.Controls.Add($txt)

    $y += 38
}

# Close Button
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"
$btnClose.Location = [System.Drawing.Point]::new(330, $y + 10)
$btnClose.Size = [System.Drawing.Size]::new(150, 35)
$btnClose.Font = $fontLabel
$btnClose.BackColor = [System.Drawing.Color]::LightGray
$btnClose.FlatStyle = "Flat"
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

# Display the window
$form.ShowDialog() | Out-Null
