# 1. Gather System Information
$computerSystem = Get-CimInstance Win32_ComputerSystem
$osInfo         = Get-CimInstance Win32_OperatingSystem
$cpuInfo        = Get-CimInstance Win32_Processor | Select-Object -First 1
$biosInfo       = Get-CimInstance Win32_Bios
$videoInfo      = Get-CimInstance Win32_VideoController | Select-Object -First 1

# Extract Active Network Adapter Information (IP, Gateway, MAC)
$netAdapter = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' -and $_.HardwareInterface -eq $true } | Select-Object -First 1
$ipInfo     = if ($netAdapter) { Get-NetIPAddress -InterfaceAlias $netAdapter.Name -AddressFamily IPv4 | Select-Object -First 1 } else { $null }
$ipAddress  = if ($ipInfo) { $ipInfo.IPAddress } else { "Not Connected" }
$macAddress = if ($netAdapter) { $netAdapter.MacAddress } else { "Unknown" }

$gateway = try {
    $route = Get-NetRoute -DestinationPrefix "0.0.0.0/0" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($route) { $route.NextHop } else { "None" }
} catch {
    "None"
}

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

# 2. Build the Clean GUI Window
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Net.Http

$form = New-Object System.Windows.Forms.Form
$form.Text = "System Info - Stock Opname Tool"
$form.Size = [System.Drawing.Size]::new(520, 710)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::White

# --- APPLICATION ICON ENGINE ---
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

# --- USER INPUT FIELDS FOR INVENTORY ---

# 1. User Name Input
$lblUser = New-Object System.Windows.Forms.Label
$lblUser.Text = "User Name"
$lblUser.Location = [System.Drawing.Point]::new(20, $y)
$lblUser.Size = [System.Drawing.Size]::new(120, 23)
$lblUser.Font = $fontLabel
$lblUser.ForeColor = [System.Drawing.Color]::Navy
$lblUser.BackColor = [System.Drawing.Color]::Transparent
$form.Controls.Add($lblUser)

$txtUser = New-Object System.Windows.Forms.TextBox
$txtUser.Text = ""
$txtUser.Location = [System.Drawing.Point]::new(150, $y - 3)
$txtUser.Size = [System.Drawing.Size]::new(330, 23)
$txtUser.Font = $fontValue
$txtUser.BackColor = [System.Drawing.Color]::LightYellow
$form.Controls.Add($txtUser)

$y += 38

# 2. Email Address Input
$lblEmail = New-Object System.Windows.Forms.Label
$lblEmail.Text = "Email Address"
$lblEmail.Location = [System.Drawing.Point]::new(20, $y)
$lblEmail.Size = [System.Drawing.Size]::new(120, 23)
$lblEmail.Font = $fontLabel
$lblEmail.ForeColor = [System.Drawing.Color]::Navy
$lblEmail.BackColor = [System.Drawing.Color]::Transparent
$form.Controls.Add($lblEmail)

$txtEmail = New-Object System.Windows.Forms.TextBox
$txtEmail.Text = ""
$txtEmail.Location = [System.Drawing.Point]::new(150, $y - 3)
$txtEmail.Size = [System.Drawing.Size]::new(330, 23)
$txtEmail.Font = $fontValue
$txtEmail.BackColor = [System.Drawing.Color]::LightYellow
$form.Controls.Add($txtEmail)

$y += 38

# 3. Perangkat Kerja Input
$lblPerangkat = New-Object System.Windows.Forms.Label
$lblPerangkat.Text = "Perangkat Kerja"
$lblPerangkat.Location = [System.Drawing.Point]::new(20, $y)
$lblPerangkat.Size = [System.Drawing.Size]::new(120, 23)
$lblPerangkat.Font = $fontLabel
$lblPerangkat.ForeColor = [System.Drawing.Color]::Navy
$lblPerangkat.BackColor = [System.Drawing.Color]::Transparent
$form.Controls.Add($lblPerangkat)

$txtPerangkat = New-Object System.Windows.Forms.TextBox
$txtPerangkat.Text = ""
$txtPerangkat.Location = [System.Drawing.Point]::new(150, $y - 3)
$txtPerangkat.Size = [System.Drawing.Size]::new(330, 23)
$txtPerangkat.Font = $fontValue
$txtPerangkat.BackColor = [System.Drawing.Color]::LightYellow
$form.Controls.Add($txtPerangkat)

$y += 38

# 4. Kepemilikan Input
$lblKepemilikan = New-Object System.Windows.Forms.Label
$lblKepemilikan.Text = "Kepemilikan"
$lblKepemilikan.Location = [System.Drawing.Point]::new(20, $y)
$lblKepemilikan.Size = [System.Drawing.Size]::new(120, 23)
$lblKepemilikan.Font = $fontLabel
$lblKepemilikan.ForeColor = [System.Drawing.Color]::Navy
$lblKepemilikan.BackColor = [System.Drawing.Color]::Transparent
$form.Controls.Add($lblKepemilikan)

$txtKepemilikan = New-Object System.Windows.Forms.TextBox
$txtKepemilikan.Text = ""
$txtKepemilikan.Location = [System.Drawing.Point]::new(150, $y - 3)
$txtKepemilikan.Size = [System.Drawing.Size]::new(330, 23)
$txtKepemilikan.Font = $fontValue
$txtKepemilikan.BackColor = [System.Drawing.Color]::LightYellow
$form.Controls.Add($txtKepemilikan)

$y += 38

# System Specs Display Hash Table (System Info + Networking)
$systemSpecs = [ordered]@{
    "Computer Name"   = $env:COMPUTERNAME
    "Merk / Tipe"     = "$($computerSystem.Manufacturer) - $($computerSystem.Model)"
    "OS"              = $osInfo.Caption
    "Processor"       = $cpuInfo.Name.Trim()
    "RAM"             = "$ramGB GB"
    "Storage"         = $allStorage
    "Layar"           = "$resolution ($screenInches - $($videoInfo.Name))"
    "Serial Number"   = $biosInfo.SerialNumber
    "IP Address"      = $ipAddress
    "Default Gateway" = $gateway
    "MAC Address"     = $macAddress
}

# Generate Labels and Read-only Textboxes dynamically for specs
foreach ($key in $systemSpecs.Keys) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $key
    $lbl.Location = [System.Drawing.Point]::new(20, $y)
    $lbl.Size = [System.Drawing.Size]::new(120, 23)
    $lbl.Font = $fontLabel
    $lbl.ForeColor = [System.Drawing.Color]::DarkSlateGray
    $lbl.BackColor = [System.Drawing.Color]::Transparent
    $form.Controls.Add($lbl)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Text = $systemSpecs[$key]
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
    $finalUser       = if ([string]::IsNullOrWhiteSpace($txtUser.Text))       { "Unknown User" } else { $txtUser.Text.Trim() }
    $finalEmail      = if ([string]::IsNullOrWhiteSpace($txtEmail.Text))      { "No Email" }      else { $txtEmail.Text.Trim() }
    $finalPerangkat  = if ([string]::IsNullOrWhiteSpace($txtPerangkat.Text))  { "-" }             else { $txtPerangkat.Text.Trim() }
    $finalKepemilikan = if ([string]::IsNullOrWhiteSpace($txtKepemilikan.Text)) { "-" }             else { $txtKepemilikan.Text.Trim() }

    # Map all data dynamically into $data ordered hash table upon submit
    $data = [ordered]@{
        "Perangkat Kerja" = $finalPerangkat
        "Kepemilikan"     = $finalKepemilikan
        "Computer Name"   = $systemSpecs["Computer Name"]
        "Merk / Tipe"     = $systemSpecs["Merk / Tipe"]
        "OS"              = $systemSpecs["OS"]
        "Processor"       = $systemSpecs["Processor"]
        "RAM"             = $systemSpecs["RAM"]
        "Storage"         = $systemSpecs["Storage"]
        "Layar"           = $systemSpecs["Layar"]
        "Serial Number"   = $systemSpecs["Serial Number"]
        "IP Address"      = $systemSpecs["IP Address"]
        "Default Gateway" = $systemSpecs["Default Gateway"]
        "MAC Address"     = $systemSpecs["MAC Address"]
    }

    # Pack semicolon matrix payload
    $semicolonRow = ($data.Values -join ";")
    $ntfyUrl = "https://ntfy.sh/sop"
    
    # Title Format: Stock Opname: Mahuraja (DESKTOP-11JKUEJ, mahuraja@waskita.co.id)
    $headers = @{
        "Title"    = "Stock Opname: $finalUser ($($data.'Computer Name'), $finalEmail)"
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

# Key ENTER triggers submit
$form.AcceptButton = $btnSubmit

# Display the window
$form.ShowDialog() | Out-Null
