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
    "Layar"         = "$screenInches"
    "Serial Number" = $biosInfo.SerialNumber
}

# 2. Build the Raw TEXTSPLIT Formula String
$excelFormula = '=TEXTSPLIT("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}", "|")' -f $data."Computer Name", $data."Merk / Tipe", $data."OS", $data."Processor", $data."RAM", $data."Storage", $data."Layar", $data."Serial Number"

# 3. Build the Mixed GUI (Specs on Top, Formula on Bottom)
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Net.Http

$form = New-Object System.Windows.Forms.Form
$form.Text = "System Info & Excel Formula Generator"
$form.Size = [System.Drawing.Size]::new(850, 680)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "Sizable"
$form.MaximizeBox = $true
$form.MinimumSize = [System.Drawing.Size]::new(650, 550)
$form.BackColor = [System.Drawing.Color]::White

# Track current text size scalers
$script:currentSpecSize = 10
$script:currentFormulaSize = 12

# --- IMAGE LOGO ENGINE ---
$logoBox = New-Object System.Windows.Forms.PictureBox
try {
    $imageUrl = "https://github.com/vinoaryo25/vinoaryo25/blob/main/bin/logo-wk.png?raw=true"
    $httpClient = New-Object System.Net.Http.HttpClient
    $imageBytes = $httpClient.GetByteArrayAsync($imageUrl).GetAwaiter().GetResult()
    $ms = New-Object System.IO.MemoryStream($imageBytes, 0, $imageBytes.Length)
    $logoBox.Image = [System.Drawing.Image]::FromStream($ms)
    $httpClient.Dispose()
} catch {
    # Fail silently if offline
}
$logoBox.Size = [System.Drawing.Size]::new(120, 120)
$logoBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$logoBox.Anchor = [System.Windows.Forms.AnchorStyles]"Top, Right"
$form.Controls.Add($logoBox)

# --- TOP SECTION: LIVE SPECIFICATIONS ---
$y = 20
$specControls = @()

foreach ($key in $data.Keys) {
    # Label
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $key
    $lbl.Location = [System.Drawing.Point]::new(20, $y)
    $lbl.Size = [System.Drawing.Size]::new(150, 28)
    $lbl.Font = New-Object System.Drawing.Font("Segoe UI", $script:currentSpecSize, [System.Drawing.FontStyle]::Bold)
    $lbl.BackColor = [System.Drawing.Color]::Transparent
    $form.Controls.Add($lbl)
    $specControls += $lbl

    # Value TextBox
    $txtSpec = New-Object System.Windows.Forms.TextBox
    $txtSpec.Text = $data[$key]
    $txtSpec.Location = [System.Drawing.Point]::new(180, $y - 3)
    $txtSpec.Font = New-Object System.Drawing.Font("Segoe UI", $script:currentSpecSize)
    $txtSpec.ReadOnly = $true
    $txtSpec.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 250)
    $txtSpec.Size = [System.Drawing.Size]::new(($form.ClientSize.Width - 330), 28)
    $txtSpec.Anchor = [System.Windows.Forms.AnchorStyles]"Top, Left, Right"
    $form.Controls.Add($txtSpec)
    $specControls += $txtSpec

    $y += 42
}

# Divider Label
$lblDivider = New-Object System.Windows.Forms.Label
$lblDivider.Text = "Excel Formula (Auto-spreads across Columns A-H)"
$lblDivider.Location = [System.Drawing.Point]::new(20, $y + 15)
$lblDivider.Size = [System.Drawing.Size]::new(500, 25)
$lblDivider.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$lblDivider.ForeColor = [System.Drawing.Color]::DarkGreen
$lblDivider.BackColor = [System.Drawing.Color]::Transparent
$form.Controls.Add($lblDivider)

# --- BOTTOM SECTION: EXCEL TEXTSPLIT FORMULA ---
$txtFormula = New-Object System.Windows.Forms.TextBox
$txtFormula.Text = $excelFormula
$txtFormula.Location = [System.Drawing.Point]::new(20, $y + 45)
$txtFormula.Size = [System.Drawing.Size]::new(($form.ClientSize.Width - 50), 120) 
$txtFormula.Font = New-Object System.Drawing.Font("Consolas", $script:currentFormulaSize, [System.Drawing.FontStyle]::Bold)
$txtFormula.Multiline = $true
$txtFormula.WordWrap = $true
$txtFormula.ReadOnly = $true
$txtFormula.ScrollBars = "Vertical"
$txtFormula.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 250)
$txtFormula.Anchor = [System.Windows.Forms.AnchorStyles]"Top, Bottom, Left, Right"
$form.Controls.Add($txtFormula)

# Size Functions
function Update-FormulaFontSize ($delta) {
    $script:currentFormulaSize += $delta
    if ($script:currentFormulaSize -lt 6) { $script:currentFormulaSize = 6 }
    if ($script:currentFormulaSize -gt 48) { $script:currentFormulaSize = 48 }
    $txtFormula.Font = New-Object System.Drawing.Font("Consolas", $script:currentFormulaSize, [System.Drawing.FontStyle]::Bold)
}

function Update-SpecFontSize ($delta) {
    $script:currentSpecSize += $delta
    if ($script:currentSpecSize -lt 6) { $script:currentSpecSize = 6 }
    foreach ($control in $specControls) {
        $style = if ($control -is [System.Windows.Forms.Label]) { [System.Drawing.FontStyle]::Bold } else { [System.Drawing.FontStyle]::Regular }
        $control.Font = New-Object System.Drawing.Font("Segoe UI", $script:currentSpecSize, $style)
    }
}

# Scrollwheel Engine (Ctrl + Mouse Wheel over Formula Field)
$txtFormula.Add_MouseWheel({
    param($sender, $e)
    if ([System.Windows.Forms.Control]::ModifierKeys -eq [System.Windows.Forms.Keys]::Control) {
        if ($e.Delta -gt 0) { Update-FormulaFontSize(1.5) } else { Update-FormulaFontSize(-1.5) }
    }
})

# --- MANUAL ACCESSIBILITY BUTTONS ---
$btnSizeDown = New-Object System.Windows.Forms.Button
$btnSizeDown.Text = "A- Smaller Text"
$btnSizeDown.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnSizeDown.Size = [System.Drawing.Size]::new(120, 35)
$btnSizeDown.Anchor = [System.Windows.Forms.AnchorStyles]"Bottom, Left"
$btnSizeDown.Add_Click({ Update-FormulaFontSize(-1.5); Update-SpecFontSize(-1.5) })
$form.Controls.Add($btnSizeDown)

$btnSizeUp = New-Object System.Windows.Forms.Button
$btnSizeUp.Text = "A+ Larger Text"
$btnSizeUp.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnSizeUp.Size = [System.Drawing.Size]::new(120, 35)
$btnSizeUp.Anchor = [System.Windows.Forms.AnchorStyles]"Bottom, Left"
$btnSizeUp.Add_Click({ Update-FormulaFontSize(1.5); Update-SpecFontSize(1.5) })
$form.Controls.Add($btnSizeUp)

$btnCopy = New-Object System.Windows.Forms.Button
$btnCopy.Text = "Copy Formula"
$btnCopy.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnCopy.Size = [System.Drawing.Size]::new(130, 35)
$btnCopy.Anchor = [System.Windows.Forms.AnchorStyles]"Bottom, Right"
$btnCopy.Add_Click({ 
    [System.Windows.Forms.Clipboard]::SetText($txtFormula.Text)
    $btnCopy.Text = "Copied!"
    Start-Sleep -Seconds 1
    $btnCopy.Text = "Copy Formula"
})
$form.Controls.Add($btnCopy)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Close"
$btnCancel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnCancel.Size = [System.Drawing.Size]::new(100, 35)
$btnCancel.Anchor = [System.Windows.Forms.AnchorStyles]"Bottom, Right"
$btnCancel.Add_Click({ $form.Close() })
$form.Controls.Add($btnCancel)

# Window Resize Coordinator
$form.Add_Resize({
    $btnCancel.Location = [System.Drawing.Point]::new(($form.ClientSize.Width - 120), ($form.ClientSize.Height - 50))
    $btnCopy.Location = [System.Drawing.Point]::new(($form.ClientSize.Width - 260), ($form.ClientSize.Height - 50))
    $btnSizeDown.Location = [System.Drawing.Point]::new(20, ($form.ClientSize.Height - 50))
    $btnSizeUp.Location = [System.Drawing.Point]::new(150, ($form.ClientSize.Height - 50))
    $logoBox.Location = [System.Drawing.Point]::new(($form.ClientSize.Width - 140), 20)
})

# Launch fully maximized
$form.WindowState = [System.Windows.Forms.FormWindowState]::Maximized

# Highlight formula instantly
$form.Add_Shown({ $txtFormula.SelectAll(); $txtFormula.Focus() })

$form.ShowDialog() | Out-Null
