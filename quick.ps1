Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Quick Email Sender
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$storePath = Join-Path $scriptDir "QuickEmailSender_history.json"
$sapDir = Join-Path $scriptDir "SapOrderReader"
$sapResultPath = Join-Path $sapDir "SapOrderReader_result.json"
$sapDllPath = Join-Path $sapDir "bin\Debug\net8.0-windows10.0.22621.0\SapOrderReader.dll"
if (-not (Test-Path $sapDllPath)) {
    $sapDllPath = Join-Path $sapDir "bin\net8.0-windows10.0.22621.0\SapOrderReader.dll"
}
$recipient = "konfig.scandinavia@also.com"

$script:Suppress = $false
$script:SuppressSuggestions = $false

# Default history. Older JSON files with Tenants values as strings remain supported.
$script:History = [PSCustomObject]@{
    Last = [PSCustomObject]@{
        Tenant = ""
        GroupTag = ""
        Qty = "1"
        OrderType = "Autopilot upload"
    }
    Tenants = [PSCustomObject]@{}
}

if (Test-Path $storePath) {
    try {
        $loaded = Get-Content $storePath -Raw | ConvertFrom-Json
        if ($loaded) { $script:History = $loaded }
    } catch {}
}

if (-not $script:History.Last) {
    $script:History | Add-Member NoteProperty Last ([PSCustomObject]@{}) -Force
}
if (-not $script:History.Tenants) {
    $script:History | Add-Member NoteProperty Tenants ([PSCustomObject]@{}) -Force
}
foreach ($item in @(
    @{Name="Tenant"; Value=""}
    @{Name="GroupTag"; Value=""}
    @{Name="Qty"; Value="1"}
    @{Name="OrderType"; Value="Autopilot upload"}
)) {
    if (-not ($script:History.Last.PSObject.Properties.Name -contains $item.Name)) {
        $script:History.Last | Add-Member NoteProperty $item.Name $item.Value -Force
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Quick Email Sender"
$form.Size = New-Object System.Drawing.Size(560, 340)
$form.StartPosition = "CenterScreen"
$form.TopMost = $true
$form.FormBorderStyle = "Sizable"

function New-Label($text, $x, $y, $w, $h=20) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $text
    $label.Location = New-Object System.Drawing.Point($x,$y)
    $label.Size = New-Object System.Drawing.Size($w,$h)
    return $label
}

$lblType = New-Label "Order Type" 10 20 105
$lblOrder = New-Label "Order Number" 10 55 105
$lblDelivery = New-Label "Delivery Number" 10 90 105
$lblQty = New-Label "Qty" 10 125 105
$lblTenant = New-Label "Tenant Name" 10 160 105
$lblGroup = New-Label "GroupTag" 10 195 105

$cbType = New-Object System.Windows.Forms.ComboBox
$cbType.Location = New-Object System.Drawing.Point(120,20)
$cbType.Size = New-Object System.Drawing.Size(410,21)
$cbType.DropDownStyle = "DropDownList"
[void]$cbType.Items.Add("Autopilot upload")

$txtOrder = New-Object System.Windows.Forms.TextBox
$txtOrder.Location = New-Object System.Drawing.Point(120,55)
$txtOrder.Size = New-Object System.Drawing.Size(410,20)

$txtDelivery = New-Object System.Windows.Forms.TextBox
$txtDelivery.Location = New-Object System.Drawing.Point(120,90)
$txtDelivery.Size = New-Object System.Drawing.Size(410,20)

$txtQty = New-Object System.Windows.Forms.TextBox
$txtQty.Location = New-Object System.Drawing.Point(120,125)
$txtQty.Size = New-Object System.Drawing.Size(70,20)

$cbTenant = New-Object System.Windows.Forms.ComboBox
$cbTenant.Location = New-Object System.Drawing.Point(120,160)
$cbTenant.Size = New-Object System.Drawing.Size(410,21)
$cbTenant.DropDownStyle = "DropDown"
$cbTenant.AutoCompleteMode = "None"
$cbTenant.AutoCompleteSource = "None"

$txtGroup = New-Object System.Windows.Forms.TextBox
$txtGroup.Location = New-Object System.Drawing.Point(120,195)
$txtGroup.Size = New-Object System.Drawing.Size(300,20)

# Small, regular-font button for cycling GroupTag history.
$btnFetch = New-Object System.Windows.Forms.Button
$btnFetch.Text = "Fetch"
$btnFetch.Location = New-Object System.Drawing.Point(425,192)
$btnFetch.Size = New-Object System.Drawing.Size(55,25)
$btnFetch.Font = New-Object System.Drawing.Font("Segoe UI",8,[System.Drawing.FontStyle]::Regular)
$btnFetch.UseVisualStyleBackColor = $true

$lblNew = New-Label "" 490 195 55
$lblNew.ForeColor = [System.Drawing.Color]::Green
$lblStatus = New-Label "" 120 285 410 30
$lblStatus.ForeColor = [System.Drawing.Color]::Green

$script:AllTenantNames = @(
    $script:History.Tenants.PSObject.Properties |
    ForEach-Object { [string]$_.Name } |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
)
foreach ($name in $script:AllTenantNames) {
    [void]$cbTenant.Items.Add($name)
}

function Clear-Status {
    if (-not $script:Suppress) { $lblStatus.Text = "" }
}

function Normalize-Text([string]$text) {
    if ($null -eq $text) { return "" }
    return (($text.ToLowerInvariant()) -replace "[^a-z0-9]", "")
}

function Tenant-GroupValues([string]$tenantLine) {
    $p = $script:History.Tenants.PSObject.Properties[$tenantLine]
    if (-not $p) { return @() }
    $v = $p.Value
    if ($v -is [string]) {
        if ([string]::IsNullOrWhiteSpace($v)) { return @() }
        return @([string]$v)
    }
    if ($v.PSObject.Properties.Name -contains "GroupTags") {
        return @($v.GroupTags | ForEach-Object { [string]$_ } | Where-Object { $_ })
    }
    if ($v.PSObject.Properties.Name -contains "LastGroupTag") {
        if ($v.LastGroupTag) { return @([string]$v.LastGroupTag) }
    }
    return @()
}

function Add-TenantGroup([string]$tenantLine, [string]$group) {
    if ([string]::IsNullOrWhiteSpace($tenantLine) -or
        [string]::IsNullOrWhiteSpace($group)) { return }

    $p = $script:History.Tenants.PSObject.Properties[$tenantLine]
    $values = @(Tenant-GroupValues $tenantLine)
    if ($values -notcontains $group) { $values += $group }

    if ($p -and $p.Value -isnot [string]) {
        $p.Value.LastGroupTag = $group
        $p.Value.GroupTags = @($values)
    } else {
        $obj = [PSCustomObject]@{
            LastGroupTag = $group
            GroupTags = @($values)
        }
        if ($p) { $p.Value = $obj }
        else { $script:History.Tenants | Add-Member NoteProperty $tenantLine $obj -Force }
    }
}

function Resolve-FullTenantLine([string]$value) {
    if ([string]::IsNullOrWhiteSpace($value)) { return $null }
    $needle = Normalize-Text $value
    if (-not $needle) { return $null }

    foreach ($p in @($script:History.Tenants.PSObject.Properties)) {
        $line = [string]$p.Name
        $whole = Normalize-Text $line
        if ($whole.Contains($needle) -or $needle.Contains($whole)) { return $line }
        foreach ($part in ($line -split "/")) {
            $partNorm = Normalize-Text $part
            if ($partNorm -and ($partNorm.Contains($needle) -or $needle.Contains($partNorm))) {
                return $line
            }
        }
    }
    return $null
}

function Update-NewIndicator {
    $tenant = $cbTenant.Text.Trim()
    $group = $txtGroup.Text.Trim()
    $lblNew.Text = ""
    $p = $script:History.Tenants.PSObject.Properties[$tenant]
    if ($p) {
        $old = @(Tenant-GroupValues $tenant)
        if ($group -and $old -notcontains $group) { $lblNew.Text = "NEW" }
    }
}

function Update-TenantSuggestions {
    if ($script:SuppressSuggestions) { return }
    $search = $cbTenant.Text
    $pos = $cbTenant.SelectionStart
    $matches = @($script:AllTenantNames | Where-Object {
        [string]::IsNullOrWhiteSpace($search) -or
        $_.IndexOf($search,[StringComparison]::OrdinalIgnoreCase) -ge 0
    })
    $script:SuppressSuggestions = $true
    try {
        $cbTenant.BeginUpdate()
        $cbTenant.Items.Clear()
        foreach ($m in $matches) { [void]$cbTenant.Items.Add($m) }
        $cbTenant.Text = $search
        $cbTenant.SelectionStart = [Math]::Min($pos,$cbTenant.Text.Length)
        $cbTenant.SelectionLength = 0
    } finally {
        $cbTenant.EndUpdate()
        $script:SuppressSuggestions = $false
    }
    if ($search -and $matches.Count -gt 0) { $cbTenant.DroppedDown = $true }
}

function Save-History {
    $tenant = $cbTenant.Text.Trim()
    $group = $txtGroup.Text.Trim()
    $script:History.Last.Tenant = $tenant
    $script:History.Last.GroupTag = $group
    $script:History.Last.Qty = $txtQty.Text.Trim()
    $script:History.Last.OrderType = $cbType.Text
    if ($tenant -and $group) { Add-TenantGroup $tenant $group }
    $script:History | ConvertTo-Json -Depth 10 | Set-Content $storePath -Encoding UTF8
}

function Import-SapResult {
    if (-not (Test-Path $sapResultPath)) { throw "SAP result file was not found:`r`n$sapResultPath" }
    $r = Get-Content $sapResultPath -Raw | ConvertFrom-Json
    $order = [string]$r.OrderNumber
    $ocrTenant = [string]$r.Tenant
    $sapGroup = [string]$r.GroupTag
    $qty = [string]$r.Qty
    if (-not $order) { throw "SAP result did not contain OrderNumber." }
    if (-not $ocrTenant) { throw "SAP result did not contain Tenant." }

    $full = Resolve-FullTenantLine $ocrTenant
    $tenant = if ($full) { $full } else { $ocrTenant }
    $group = $sapGroup
    if ([string]::IsNullOrWhiteSpace($group) -and $full) {
        $saved = @(Tenant-GroupValues $full)
        if ($saved.Count -gt 0) { $group = [string]$saved[-1] }
    }

    $script:SuppressSuggestions = $true
    try {
        $txtOrder.Text = $order
        $txtQty.Text = $qty
        $cbTenant.Text = $tenant
        $txtGroup.Text = $group
    } finally { $script:SuppressSuggestions = $false }

    if ($full) {
        $lblStatus.ForeColor = [Drawing.Color]::Green
        $lblStatus.Text = if ($sapGroup) { "SAP values imported" } else { "SAP imported; saved GroupTag used" }
    } else {
        $lblStatus.ForeColor = [Drawing.Color]::DarkOrange
        $lblStatus.Text = "SAP imported; new tenant"
    }
    Update-NewIndicator
}

function Start-SapReader {
    if (-not (Test-Path $sapDllPath)) { throw "SAP helper DLL was not found:`r`n$sapDllPath" }
    if (Test-Path $sapResultPath) { Remove-Item $sapResultPath -Force }
    $dotnet = Join-Path ${env:ProgramFiles} "dotnet\dotnet.exe"
    if (-not (Test-Path $dotnet)) { $dotnet = "dotnet.exe" }
    $p = Start-Process -FilePath $dotnet -ArgumentList "`"$sapDllPath`"" `
        -WorkingDirectory (Split-Path $sapDllPath -Parent) -WindowStyle Hidden -PassThru
    $deadline = (Get-Date).AddMinutes(2)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 500
        if (Test-Path $sapResultPath) {
            try {
                $r = Get-Content $sapResultPath -Raw | ConvertFrom-Json
                if ($r.OrderNumber -and $r.Tenant) { break }
            } catch {}
        }
        if ($p.HasExited -and -not (Test-Path $sapResultPath)) { throw "SAP reader exited without a result." }
    }
    if (-not (Test-Path $sapResultPath)) { throw "SAP reader did not create a result within 2 minutes." }
    try {
        $p.WaitForExit(5000)
        if (-not $p.HasExited) { $p.Kill() }
    } catch {}
}

$cbTenant.Add_TextChanged({
    if (-not $script:SuppressSuggestions) {
        Update-TenantSuggestions
        $p = $script:History.Tenants.PSObject.Properties[$cbTenant.Text.Trim()]
        if ($p) {
            $values = @(Tenant-GroupValues $cbTenant.Text.Trim())
            if ($values.Count -gt 0) { $txtGroup.Text = [string]$values[-1] }
        }
    }
    Update-NewIndicator
    Clear-Status
})
$txtGroup.Add_TextChanged({ Update-NewIndicator; Clear-Status })
$txtOrder.Add_TextChanged({ Clear-Status })
$txtDelivery.Add_TextChanged({ Clear-Status })
$txtQty.Add_TextChanged({ Clear-Status })
$cbType.Add_SelectedIndexChanged({ Clear-Status })

$script:FetchIndex = 0
$btnFetch.Add_Click({
    $tenant = $cbTenant.Text.Trim()
    if (-not $tenant) { $lblStatus.Text = "Select a tenant first."; return }
    $values = @(Tenant-GroupValues $tenant)
    if ($values.Count -eq 0) { $lblStatus.Text = "No saved GroupTags for this tenant."; return }
    $script:FetchIndex = ($script:FetchIndex + 1) % $values.Count
    $txtGroup.Text = [string]$values[$script:FetchIndex]
    $lblStatus.ForeColor = [Drawing.Color]::DarkBlue
    $lblStatus.Text = "Saved GroupTag selected"
})

$btnSend = New-Object System.Windows.Forms.Button
$btnSend.Text = "Send Email"
$btnSend.Location = New-Object System.Drawing.Point(120,225)
$btnSend.Size = New-Object System.Drawing.Size(100,28)

$btnReadSap = New-Object System.Windows.Forms.Button
$btnReadSap.Text = "Read SAP values"
$btnReadSap.Location = New-Object System.Drawing.Point(230,225)
$btnReadSap.Size = New-Object System.Drawing.Size(125,28)

$btnReadSap.Add_Click({
    try {
        $btnReadSap.Enabled = $false
        $btnSend.Enabled = $false
        $lblStatus.Text = "Running SAP reader..."
        $form.Hide()
        Start-SapReader
        $form.Show(); $form.Activate()
        Import-SapResult
        [void]$txtDelivery.Focus()
    } catch {
        $form.Show(); $form.Activate()
        $lblStatus.ForeColor = [Drawing.Color]::Red
        $lblStatus.Text = "SAP import failed"
        [Windows.Forms.MessageBox]::Show($_.Exception.Message,"SAP import failed",[Windows.Forms.MessageBoxButtons]::OK,[Windows.Forms.MessageBoxIcon]::Error)
    } finally {
        $btnReadSap.Enabled = $true
        $btnSend.Enabled = $true
    }
})

$btnSend.Add_Click({
    $order=$txtOrder.Text.Trim(); $delivery=$txtDelivery.Text.Trim()
    $qty=$txtQty.Text.Trim(); $type=$cbType.Text
    $tenant=$cbTenant.Text.Trim(); $group=$txtGroup.Text.Trim()
    if (-not $order) {$lblStatus.Text="Order Number is required.";return}
    if (-not $delivery) {$lblStatus.Text="Delivery Number is required.";return}
    if (-not $qty) {$lblStatus.Text="Qty is required.";return}
    if (-not $type) {$lblStatus.Text="Order Type is required.";return}
    if (-not $tenant) {$lblStatus.Text="Tenant Name is required.";return}
    if (-not $group) {$lblStatus.Text="GroupTag is required.";return}
    try {
        $mail=(New-Object -ComObject Outlook.Application).CreateItem(0)
        $mail.To=$recipient
        $mail.Subject="$order, $delivery - ${qty}stk $type"
        $mail.Body="Tenant Name: $tenant`r`nGroupTag: $group"
        $mail.Send()
        Save-History
        $lblStatus.ForeColor=[Drawing.Color]::Green
        $lblStatus.Text="Email sent"
        $script:Suppress=$true
        $txtOrder.Text=""; $txtDelivery.Text=""
        [void]$txtOrder.Focus()
        $script:Suppress=$false
    } catch {
        $script:Suppress=$false
        $lblStatus.ForeColor=[Drawing.Color]::Red
        $lblStatus.Text="Failed: "+$_.Exception.Message
    }
})

$cbTenant.Text = [string]$script:History.Last.Tenant
$txtGroup.Text = [string]$script:History.Last.GroupTag
$txtQty.Text = [string]$script:History.Last.Qty
$lastType = [string]$script:History.Last.OrderType
if ($cbType.Items.Contains($lastType)) { $cbType.SelectedItem=$lastType } else { $cbType.SelectedIndex=0 }

$form.Controls.AddRange(@(
    $lblType,$lblOrder,$lblDelivery,$lblQty,$lblTenant,$lblGroup,
    $cbType,$txtOrder,$txtDelivery,$txtQty,$cbTenant,$txtGroup,
    $lblNew,$btnFetch,$btnSend,$btnReadSap,$lblStatus
))
Update-NewIndicator
[void]$form.ShowDialog()
