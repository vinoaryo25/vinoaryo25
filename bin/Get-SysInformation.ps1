# 1. Gather System Information
$computerSystem = Get-CimInstance Win32_ComputerSystem
$osInfo         = Get-CimInstance Win32_OperatingSystem
$cpuInfo        = Get-CimInstance Win32_Processor | Select-Object -First 1
$biosInfo       = Get-CimInstance Win32_Bios
$videoInfo      = Get-CimInstance Win32_VideoController | Select-Object -First 1

# Extract the Physical Screen Inch Size via WMI Display Parameters
$screenInches = try {
    $monitorParams = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorBasicDisplayParams | Select-Object -First 1
    if ($monitorParams -and $monitorParams.MaxHorizontalImageSize -gt 0) {
        $widthInches = $monitorParams.MaxHorizontalImageSize / 2.54
        $heightInches = $monitorParams.MaxVerticalImageSize / 2.54
        $diagonal = [Math]::Sqrt(($widthInches * $widthInches) + ($heightInches * $heightInches))
        "$([Math]::Round($diagonal, 1))`""
    } else {
        "Unknown Size"
    }
} catch {
    "Unknown Size"
}

# Get ALL local hard drives/partitions & calculate individual + total sizes
$disks = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3"
$diskList = @()
$totalStorageGB = 0

foreach ($disk in $disks) {
    $sizeGB = [math]::Round($disk.Size / 1GB)
    $diskList += "$($disk.DeviceID) $sizeGB GB"
    $totalStorageGB += $sizeGB  
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
    "Layar"         = "$resolution ($screenInches - $($videoInfo.Name))"
    "Serial Number" = $biosInfo.SerialNumber
}

# 2. Build the Clean GUI Window
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Net.Http

$form = New-Object System.Windows.Forms.Form
$form.Text = "System Info - Stock Opname Tool"
$form.Size = [System.Drawing.Size]::new(520, 480)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::White

# --- APPLICATION ICON ENGINE ---
# Converts the logo image seamlessly into an active native window application icon
try {
    $imageUrl = "https://github.com/vinoaryo25/vinoaryo25/blob/main/bin/logo-wk.png?raw=true"
    $httpClient = New-Object System.Net.Http.HttpClient
    $imageBytes = $httpClient.GetByteArrayAsync($imageUrl).GetAwaiter().GetResult()
    $ms = New-Object System.IO.MemoryStream($imageBytes, 0, $imageBytes.Length)
    $bitmap = New-Object System.Drawing.Bitmap($ms)
    $hIcon = $bitmap.GetHicon()
    $form.Icon = [System.Drawing.Icon]::FromHandle($hIcon)
    $bitmap.Dispose()
    $httpClient.Dispose()
} catch {
    # Fallback to standard OS styling if offline
}

# Styling Fonts
$fontLabel = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$fontValue = New-Object System.Drawing.Font("Segoe UI", 9)

$y = 20

# --- USER INPUT FIELD FOR NTFY TITLE ---
$lblUser = New-Object System.Windows.Forms.Label
$lblUser.Text = "User Name"
$lblUser.Location = [System.Drawing.Point]::new(20, $y)
$lblUser.Size = [System.Drawing.Size]::new(120, 23)
$lblUser.Font = $fontLabel
$lblUser.ForeColor = [System.Drawing.Color]::Navy
$lblUser.BackColor = [System.Drawing.Color]::Transparent
$form.Controls.Add($lblUser)

$txtUser = New-Object System.Windows.Forms.TextBox
$txtUser.Text = "" # Left blank for manual inventory user assignment
$txtUser.Location = [System.Drawing.Point]::new(150, $y - 3)
$txtUser.Size = [System.Drawing.Size]::new(330, 23)
$txtUser.Font = $fontValue
$txtUser.BackColor = [System.Drawing.Color]::LightYellow  # Distinct visual accent color
$form.Controls.Add($txtUser)

$y += 38

# Generate Labels and Read-only Textboxes dynamically for specs
foreach ($key in $data.Keys) {
    # Label
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $key
    $lbl.Location = [System.Drawing.Point]::new(20, $y)
    $lbl.Size = [System.Drawing.Size]::new(120, 23)
    $lbl.Font = $fontLabel
    $lbl.ForeColor = [System.Drawing.Color]::DarkSlateGray
    $lbl.BackColor = [System.Drawing.Color]::Transparent
    $form.Controls.Add($lbl)

    # TextBox (Full width now since the interior graphic is removed)
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

# Submit & Close Button
$btnSubmit = New-Object System.Windows.Forms.Button
$btnSubmit.Text = "Submit & Close"
$btnSubmit.Location = [System.Drawing.Point]::new(330, $y + 10)
$btnSubmit.Size = [System.Drawing.Size]::new(150, 35)
$btnSubmit.Font = $fontLabel
$btnSubmit.BackColor = [System.Drawing.Color]::LightGreen
$btnSubmit.FlatStyle = "Flat"

# Action Trigger: Bundles and pushes string row upon confirmation click
$btnSubmit.Add_Click({ 
    $finalUser = if ([string]::IsNullOrWhiteSpace($txtUser.Text)) { "Unknown User" } else { $txtUser.Text.Trim() }
    
    # Pack semicolon matrix payload
    $semicolonRow = ($data.Values -join ";")
    $ntfyUrl = "https://ntfy.sh/sop"
    $headers = @{
        "Title"    = "Stock Opname: $finalUser ($($data.'Computer Name'))"
        "Priority" = "default"
        "Tags"     = "computer,clipboard"
    }

    try {
        Invoke-RestMethod -Method Post -Uri $ntfyUrl -Headers $headers -Body $semicolonRow -ContentType "text/plain; charset=utf-8" | Out-Null
    } catch {
        # Silent exception catch to close cleanly even if network drops
    }

    $form.Close() 
})
$form.Controls.Add($btnSubmit)

# --- KEYBOARD ENTER SUBMIT FIX ---
# This line tells the form: "If the user hits Enter anywhere, click the submit button"
$form.AcceptButton = $btnSubmit

# Display the window
$form.ShowDialog() | Out-Null
