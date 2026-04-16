<#
MIT License

Copyright (c) 2026 akaminski11

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

-------------------------------------------------------------------------------
Responsible Use Disclaimer

By using this script, you confirm that you have the appropriate authorization
and permissions to access and retrieve data from Microsoft Graph and any
associated systems. You agree to use this tool in a lawful, ethical, and
responsible manner, and acknowledge that misuse of administrative tools or
tenant data may violate organizational policy, contractual obligations, or law.
-------------------------------------------------------------------------------
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing


Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# -----------------------------
# Utility helpers
# -----------------------------
function Encode-Html {
    param($s)
    $text = [string]$s
    try { return [System.Web.HttpUtility]::HtmlEncode($text) }
    catch {
        try { return [System.Net.WebUtility]::HtmlEncode($text) }
        catch {
            $text = $text -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;' -replace "'",'&#39;'
            return $text
        }
    }
}

function Get-ObjValue {
    param(
        [Parameter(Mandatory=$true)] $Obj,
        [Parameter(Mandatory=$true)][string] $Name
    )
    if ($null -eq $Obj) { return $null }

    if ($Obj -is [hashtable]) {
        if ($Obj.ContainsKey($Name)) { return $Obj[$Name] }
        return $null
    } else {
        $prop = $Obj.PSObject.Properties | Where-Object { $_.Name -eq $Name }
        if ($prop) { return $prop.Value }
        return $null
    }
}

function Normalize-TypeString {
    param($raw)
    if (-not $raw) { return "" }
    $s = $raw.ToString().Trim()
    if ($s.StartsWith("#")) { $s = $s.Substring(1) }
    return $s.ToLower()
}

# Convert an image file to a resized PNG data URI (base64) for embedding in HTML
function Convert-ImageToDataUriPng {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [int]$MaxWidth = 260,
        [int]$MaxHeight = 80
    )

    if (-not (Test-Path $Path)) { return $null }

    $img = $null
    $bmp = $null
    $ms  = $null
    try {
        $img = [System.Drawing.Image]::FromFile($Path)

        $scaleW = $MaxWidth  / [double]$img.Width
        $scaleH = $MaxHeight / [double]$img.Height
        $scale  = [Math]::Min([Math]::Min($scaleW, $scaleH), 1.0)

        $newW = [int][Math]::Round($img.Width  * $scale)
        $newH = [int][Math]::Round($img.Height * $scale)

        $bmp = New-Object System.Drawing.Bitmap($newW, $newH)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $g.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
        $g.CompositingQuality= [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
        $g.Clear([System.Drawing.Color]::Transparent)
        $g.DrawImage($img, 0, 0, $newW, $newH)
        $g.Dispose()

        $ms = New-Object System.IO.MemoryStream
        $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
        $bytes = $ms.ToArray()
        $b64 = [Convert]::ToBase64String($bytes)
        return "data:image/png;base64,$b64"
    } catch {
        return $null
    } finally {
        if ($ms) { $ms.Dispose() }
        if ($bmp) { $bmp.Dispose() }
        if ($img) { $img.Dispose() }
    }
}

# Graph paging helper for Invoke-MgGraphRequest
function Invoke-GraphPagedGet {
    param(
        [Parameter(Mandatory=$true)][string]$Uri
    )
    $all = @()
    $next = $Uri
    do {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $next
        if ($resp -and $resp.value) { $all += $resp.value }
        $next = $null
        if ($resp -and ($resp.PSObject.Properties.Name -contains '@odata.nextLink')) {
            $next = $resp.'@odata.nextLink'
        }
    } while ($next)
    return $all
}

# Group name cache
$script:GroupNameCache = @{}
function Get-GroupNameFromId {
    param([string]$groupId)

    if ([string]::IsNullOrWhiteSpace($groupId)) { return "Unknown Group" }

    if ($script:GroupNameCache.ContainsKey($groupId)) {
        return $script:GroupNameCache[$groupId]
    }

    try {
        $g = Get-MgGroup -GroupId $groupId -Property "displayName"
        $name = $g.DisplayName
        if ([string]::IsNullOrWhiteSpace($name)) { $name = "Unknown Group ($groupId)" }
    } catch {
        $name = "Unknown Group ($groupId)"
    }

    $script:GroupNameCache[$groupId] = $name
    return $name
}

# App classification
function Get-IntuneAppClassification {
    param([Parameter(Mandatory=$true)] $App)

    $tRaw = Get-ObjValue -Obj $App -Name '@odata.type'
    $t    = Normalize-TypeString $tRaw

    $cls = [PSCustomObject]@{ Platform='Other'; Type='Other'; Source='' }

    switch -Regex ($t) {
        'microsoft\.graph\.win32lobapp' { $cls.Platform='Windows'; $cls.Type='Windows app (Win32)'; $cls.Source='Line of Business'; return $cls }
        'microsoft\.graph\.wingetapp'   { $cls.Platform='Windows'; $cls.Type='Windows app (Win32)'; $cls.Source='Microsoft Store (WinGet)'; return $cls }
        '(microsoft\.graph\.windowsstoreapp|microsoft\.graph\.windowsmicrosoftstoreforbusinessapp)' {
            $cls.Platform='Windows'; $cls.Type='Windows app (Win32)'; $cls.Source='Microsoft Store (Legacy)'; return $cls
        }
        'microsoft\.graph\.androidstoreapp' { $cls.Platform='Android'; $cls.Type='Android (Public Store)'; return $cls }
        'microsoft\.graph\.managedgoogleplay(app|storeapp)' { $cls.Platform='Android'; $cls.Type='Android Enterprise (Managed Play)'; return $cls }
        'microsoft\.graph\.iosstoreapp' { $cls.Platform='iOS/iPadOS'; $cls.Type='iOS/iPadOS (App Store)'; return $cls }
        'microsoft\.graph\.macoslobapp' {
            $cls.Platform='macOS'
            $pkgType = Get-ObjValue -Obj $App -Name 'packageType'
            if ($pkgType) {
                $pt = $pkgType.ToString().ToUpper()
                if ($pt -eq 'PKG') { $cls.Type='macOS (LOB – PKG)'; return $cls }
                if ($pt -eq 'DMG') { $cls.Type='macOS (LOB – DMG)'; return $cls }
            }
            $cls.Type='macOS (LOB)'; return $cls
        }
        'microsoft\.graph\.windowsmicrosoftedgeapp' { $cls.Platform='Windows'; $cls.Type='Microsoft Edge (77+)'; return $cls }
        '(microsoft\.graph\.officesuiteapp|microsoft\.graph\.windowsmicrosoftoffice365(app)?)' { $cls.Platform='Windows'; $cls.Type='Microsoft 365 (Office Suite)'; return $cls }
        'microsoft\.graph\.webapp$' { $cls.Platform='Other'; $cls.Type='Web App (Browser-based)'; return $cls }
        'microsoft\.graph\.ioswebclip' { $cls.Platform='iOS/iPadOS'; $cls.Type='Platform Link Shortcut'; return $cls }
        'microsoft\.graph\.windowswebapp' { $cls.Platform='Windows'; $cls.Type='Platform Link Shortcut'; return $cls }
        'microsoft\.graph\.macoswebclip' { $cls.Platform='macOS'; $cls.Type='Platform Link Shortcut'; return $cls }
        'microsoft\.graph\.androidenterprisesystemapp' { $cls.Platform='Android'; $cls.Type='Android Enterprise (System App)'; return $cls }
    }

    if ($cls.Platform -eq 'Other') {
        if     ($t -match 'microsoft\.graph\..*windows') { $cls.Platform = 'Windows' }
        elseif ($t -match 'microsoft\.graph\..*macos')   { $cls.Platform = 'macOS' }
        elseif ($t -match 'microsoft\.graph\..*android') { $cls.Platform = 'Android' }
        elseif ($t -match 'microsoft\.graph\..*ios')     { $cls.Platform = 'iOS/iPadOS' }
    }

    return $cls
}

# -----------------------------
# UI State
# -----------------------------
$script:ConnectedTenantName = ""
$script:ConnectedTenantId   = ""
$script:LogoPath            = ""
$script:LogoDataUri         = $null

# -----------------------------
# UI (WinForms)
# -----------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Intune Assessment Tool"
$form.StartPosition = "CenterScreen"
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.ClientSize = New-Object System.Drawing.Size(720, 330)

# Button bar using FlowLayoutPanel (prevents cut-off / DPI issues)
$btnPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$btnPanel.Location = New-Object System.Drawing.Point(20, 15)
$btnPanel.Size = New-Object System.Drawing.Size(680, 40)
$btnPanel.WrapContents = $false
$btnPanel.AutoSize = $true
$btnPanel.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$btnPanel.Padding = New-Object System.Windows.Forms.Padding(0)
$btnPanel.Margin = New-Object System.Windows.Forms.Padding(0)

$connectBtn = New-Object System.Windows.Forms.Button
$connectBtn.Text = "Connect"
$connectBtn.AutoSize = $true
$connectBtn.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$connectBtn.Padding = New-Object System.Windows.Forms.Padding(12,6,12,6)

$disconnectBtn = New-Object System.Windows.Forms.Button
$disconnectBtn.Text = "Disconnect"
$disconnectBtn.AutoSize = $true
$disconnectBtn.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$disconnectBtn.Padding = New-Object System.Windows.Forms.Padding(12,6,12,6)

$selectLogoBtn = New-Object System.Windows.Forms.Button
$selectLogoBtn.Text = "Select Logo"
$selectLogoBtn.AutoSize = $true
$selectLogoBtn.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$selectLogoBtn.Padding = New-Object System.Windows.Forms.Padding(12,6,12,6)

$clearLogoBtn = New-Object System.Windows.Forms.Button
$clearLogoBtn.Text = "Clear Logo"
$clearLogoBtn.AutoSize = $true
$clearLogoBtn.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$clearLogoBtn.Padding = New-Object System.Windows.Forms.Padding(12,6,12,6)

$generateBtn = New-Object System.Windows.Forms.Button
$generateBtn.Text = "Generate Report"
$generateBtn.AutoSize = $true
$generateBtn.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$generateBtn.Padding = New-Object System.Windows.Forms.Padding(12,6,12,6)

$btnPanel.Controls.AddRange(@($connectBtn,$disconnectBtn,$selectLogoBtn,$clearLogoBtn,$generateBtn))
$form.Controls.Add($btnPanel)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Status: Not Connected"
$statusLabel.Location = New-Object System.Drawing.Point(20, 65)
$statusLabel.Size = New-Object System.Drawing.Size(680, 20)
$form.Controls.Add($statusLabel)

$tenantLabel = New-Object System.Windows.Forms.Label
$tenantLabel.Text = "Tenant: (not connected)"
$tenantLabel.Location = New-Object System.Drawing.Point(20, 88)
$tenantLabel.Size = New-Object System.Drawing.Size(680, 20)
$tenantLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($tenantLabel)

$logoLabel = New-Object System.Windows.Forms.Label
$logoLabel.Text = "Logo: (none selected)"
$logoLabel.Location = New-Object System.Drawing.Point(20, 112)
$logoLabel.Size = New-Object System.Drawing.Size(520, 20)
$form.Controls.Add($logoLabel)

$logoPreview = New-Object System.Windows.Forms.PictureBox
$logoPreview.Location = New-Object System.Drawing.Point(20, 140)
$logoPreview.Size = New-Object System.Drawing.Size(320, 90)
$logoPreview.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$logoPreview.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$form.Controls.Add($logoPreview)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20, 245)
$progressBar.Size = New-Object System.Drawing.Size(680, 20)
$form.Controls.Add($progressBar)

$hintLabel = New-Object System.Windows.Forms.Label
$hintLabel.Text = "Tip: The selected logo will be embedded into the report (portable HTML)."
$hintLabel.Location = New-Object System.Drawing.Point(20, 270)
$hintLabel.Size = New-Object System.Drawing.Size(680, 20)
$hintLabel.ForeColor = [System.Drawing.Color]::FromArgb(90,90,90)
$form.Controls.Add($hintLabel)

# -----------------------------
# UI Actions
# -----------------------------
$connectBtn.Add_Click({
    try {
        Connect-MgGraph -Scopes "DeviceManagementApps.Read.All DeviceManagementConfiguration.Read.All DeviceManagementManagedDevices.Read.All Group.Read.All Organization.Read.All"
        $org = (Get-MgOrganization -All | Select-Object -First 1)
        $script:ConnectedTenantName = $org.DisplayName
        $script:ConnectedTenantId   = $org.Id
        $statusLabel.Text = "Status: Connected"
        $tenantLabel.Text = "Tenant: $($script:ConnectedTenantName)"
        $form.Text = "Intune Assessment Tool — $($script:ConnectedTenantName)"
    } catch {
        $statusLabel.Text = "Status: Connection failed"
        $tenantLabel.Text = "Tenant: (not connected)"
        $form.Text = "Intune Assessment Tool"
    }
})

$disconnectBtn.Add_Click({
    Disconnect-MgGraph
    $script:ConnectedTenantName = ""
    $script:ConnectedTenantId = ""
    $statusLabel.Text = "Status: Disconnected"
    $tenantLabel.Text = "Tenant: (not connected)"
    $form.Text = "Intune Assessment Tool"
})

$selectLogoBtn.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Title = "Select a logo image"
    $ofd.Filter = "Image Files (*.png;*.jpg;*.jpeg;*.gif;*.bmp)|*.png;*.jpg;*.jpeg;*.gif;*.bmp|All Files (*.*)|*.*"
    $ofd.Multiselect = $false
    $ofd.RestoreDirectory = $true

    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:LogoPath = $ofd.FileName
        $logoLabel.Text = "Logo: $([System.IO.Path]::GetFileName($script:LogoPath))"

        try {
            if ($logoPreview.Image) { $logoPreview.Image.Dispose(); $logoPreview.Image = $null }
            $logoPreview.Image = [System.Drawing.Image]::FromFile($script:LogoPath)
        } catch {
            $logoLabel.Text = "Logo: (unable to preview image)"
        }

        $script:LogoDataUri = Convert-ImageToDataUriPng -Path $script:LogoPath -MaxWidth 260 -MaxHeight 80
    }
})

$clearLogoBtn.Add_Click({
    $script:LogoPath = ""
    $script:LogoDataUri = $null
    $logoLabel.Text = "Logo: (none selected)"
    if ($logoPreview.Image) { $logoPreview.Image.Dispose(); $logoPreview.Image = $null }
})

$generateBtn.Add_Click({
    $progressBar.Value = 5
    $statusLabel.Text = "Collecting data..."

    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $collectionErrors = New-Object System.Collections.Generic.List[string]

    # Tenant details
    try {
        $orgInfo = (Get-MgOrganization -All | Select-Object -First 1)
        $tenantName = $orgInfo.DisplayName
        $tenantId = $orgInfo.Id
        $verifiedDomains = ($orgInfo.VerifiedDomains | ForEach-Object { $_.Name }) -join ", "
    } catch {
        $tenantName = "Unknown"
        $tenantId = "Unknown"
        $verifiedDomains = "Error retrieving domains"
        $collectionErrors.Add("Tenant Info: $($_.Exception.Message)")
    }

    # Counters
    $devicesCount = 0
    $nonCompliantCount = 0
    $configPoliciesCount = 0
    $configAssignmentFailures = 0
    $compliancePoliciesCount = 0
    $appsCount = 0

    # Devices
    $devices = @()
    try {
        $devices = Get-MgDeviceManagementManagedDevice -All -Property "deviceName,operatingSystem,osVersion,complianceState,isEncrypted"
        $devicesCount = $devices.Count
        $nonCompliantCount = ($devices | Where-Object { $_.ComplianceState -ne "compliant" }).Count
    } catch {
        $devices = @()
        $devicesCount = "ERROR"
        $nonCompliantCount = "ERROR"
        $collectionErrors.Add("Devices: $($_.Exception.Message)")
    }

    $progressBar.Value = 20

    # Config Policies (paged)
    $configPolicies = @()
    try {
        $configPolicies = Invoke-GraphPagedGet -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies"
        $configPoliciesCount = $configPolicies.Count
    } catch {
        $configPolicies = @()
        $configPoliciesCount = "ERROR"
        $collectionErrors.Add("Configuration Policies: $($_.Exception.Message)")
    }

    $progressBar.Value = 35

    # Compliance Policies (paged)
    $compliancePolicies = @()
    try {
        $compliancePolicies = Invoke-GraphPagedGet -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies"
        $compliancePoliciesCount = $compliancePolicies.Count
    } catch {
        $compliancePolicies = @()
        $compliancePoliciesCount = "ERROR"
        $collectionErrors.Add("Compliance Policies: $($_.Exception.Message)")
    }

    $progressBar.Value = 50

    # Apps (paged)
    $apps = @()
    try {
        $apps = Invoke-GraphPagedGet -Uri "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps"
        $appsCount = $apps.Count
    } catch {
        $apps = @()
        $appsCount = "ERROR"
        $collectionErrors.Add("Applications: $($_.Exception.Message)")
    }

    $progressBar.Value = 65

    # -----------------------------
    # Client-ready HTML (with optional logo)
    # -----------------------------

    # CHANGE #1: If no logo selected, leave blank (no placeholder)
    $logoHtml = ""
    if ($script:LogoDataUri) {
        $logoHtml = "<img class='logo' src='$($script:LogoDataUri)' alt='Logo'/>"
    }

    # CHANGE #2: Replace subtitle with disclaimer text
    $disclaimerText = "Disclaimer: This report was generated using an automated script after obtaining the appropriate permissions to access Microsoft Graph data. Results may vary by tenant configuration, licensing, role-based access, and data availability, and may not reflect a complete assessment."

    $html = @"
<html>
<head>
<title>Intune Assessment Report</title>
<style>
  body {font-family: Arial; background-color: #f9f9f9; color: #333; margin: 16px;}
  .header {display:flex; align-items:center; justify-content:space-between; background:#ffffff; border:1px solid #e5e5e5; border-radius:8px; padding:14px 16px; margin-bottom:16px;}
  .header-left {display:flex; align-items:center; gap:14px;}
  .logo {max-height:80px; max-width:260px;}
  .title-block h1 {margin:0; color:#2c3e50; font-size:22px;}
  .title-block .sub {margin-top:6px; color:#555; font-size:12px; line-height:1.35; max-width:560px;}
  .meta {text-align:right; color:#555; font-size:12px;}
  h2 {color:#34495e; border-bottom: 2px solid #ccc; padding-bottom: 4px; margin-top: 22px;}
  h3 {color:#34495e; margin:0 0 6px 0;}
  table {border-collapse: collapse; width: 100%; margin-bottom: 20px; background:#fff;}
  th {background-color:#2980b9; color:#fff; padding:8px; text-align:left;}
  td {border:1px solid #ccc; padding:8px; vertical-align:top;}
  .summary-box {background-color:#ecf0f1; padding:12px; border-radius:6px; margin-bottom:20px;}
  .recommendations {background-color:#fdf2e9; padding:12px; border-radius:6px;}
  .warnings {background-color:#fdecea; padding:12px; border-radius:6px; border:1px solid #f5c6cb; margin-bottom:20px;}
  .muted {color:#666;}
  .small {font-size:12px;}
  .badge {display:inline-block; padding:2px 10px; border-radius:999px; background:#eef2ff; font-size:12px; font-weight:bold; margin-left:6px;}
  .badge-warn {background:#fdecea;}
  .section-note {
    margin-top: 6px;
    margin-bottom: 14px;
    color: #555;
    font-size: 13px;
    line-height: 1.4;
  }
</style>
</head>
<body>

<div class="header">
  <div class="header-left">
    $logoHtml
    <div class="title-block">
      <h1>Intune Assessment Report</h1>
      <div class="sub">$(Encode-Html $disclaimerText)</div>
    </div>
  </div>
  <div class="meta">
    <div><strong>Tenant:</strong> $(Encode-Html $tenantName)</div>
    <div><strong>Generated:</strong> $(Encode-Html $timestamp)</div>
  </div>
</div>

<div class='summary-box'>
  <h3>Executive Summary</h3>
  <p class="small">
    This report provides a high-level assessment of the current Microsoft Intune configuration for
    <strong>$(Encode-Html $tenantName)</strong>. It highlights device compliance posture, policy coverage,
    and application deployment state, and identifies priority areas for optimization and risk reduction.
  </p>
</div>

<h2>Tenant Information</h2>
<table>
<tr><th>Tenant Name</th><th>Tenant ID</th><th>Verified Domains</th></tr>
<tr>
  <td>$(Encode-Html $tenantName)</td>
  <td>$(Encode-Html $tenantId)</td>
  <td>$(Encode-Html $verifiedDomains)</td>
</tr>
</table>

<div class='summary-box'>
  <h2>High-Level Summary</h2>
  <p>
    <strong>Total Devices:</strong> $(Encode-Html $devicesCount)<br/>
    <strong>Noncompliant Devices:</strong> $(Encode-Html $nonCompliantCount) <span class="badge badge-warn">Needs review</span><br/>
    <strong>Total Configuration Policies:</strong> $(Encode-Html $configPoliciesCount)<br/>
    <strong>Assignment Failures:</strong> $(Encode-Html $configAssignmentFailures)<br/>
    <strong>Total Compliance Policies:</strong> $(Encode-Html $compliancePoliciesCount)<br/>
    <strong>Total Applications:</strong> $(Encode-Html $appsCount)
  </p>
  <p class='muted small'>Note: Some datasets use Microsoft Graph <em>/beta</em> endpoints.</p>
</div>
"@

    if ($collectionErrors.Count -gt 0) {
        $html += "<div class='warnings'><h2>Data Collection Warnings</h2><ul>"
        foreach ($e in $collectionErrors) { $html += "<li>$(Encode-Html $e)</li>" }
        $html += "</ul></div>"
    }

    # Device inventory
    $html += "<h2>Device Inventory</h2>"
    $html += "<p class='section-note'>This section provides a list of devices currently managed by Microsoft Intune. It highlights each device’s operating system, compliance status, and whether disk encryption is enabled.</p>"

    if ($devices -and $devices.Count -gt 0) {
        $html += "<table><tr><th>Name</th><th>Platform</th><th>OS Version</th><th>Compliance</th><th>Encryption</th></tr>"
        foreach ($device in $devices) {
            $enc = "Unknown"
            if ($null -ne $device.IsEncrypted) { $enc = if ($device.IsEncrypted) { "Encrypted" } else { "Not Encrypted" } }

            $html += "<tr>
              <td>$(Encode-Html $device.DeviceName)</td>
              <td>$(Encode-Html $device.OperatingSystem)</td>
              <td>$(Encode-Html $device.OsVersion)</td>
              <td>$(Encode-Html $device.ComplianceState)</td>
              <td>$(Encode-Html $enc)</td>
            </tr>"
        }
        $html += "</table>"
    } else {
        $html += "<p class='muted'>No device data returned.</p>"
    }

    $progressBar.Value = 75

    # Configuration policies
    $html += "<h2>Configuration Policies</h2>"
    $html += "<p class='section-note'>Configuration policies define security controls, system settings, and user experience standards applied to devices and users. This section shows existing policies and the groups or device scopes they are assigned to.</p>"

    if ($configPolicies -and $configPolicies.Count -gt 0) {
        $html += "<table><tr><th>Name</th><th>Description</th><th>Last Modified</th><th>Assignments</th></tr>"

        foreach ($policy in $configPolicies) {
            $name = if ($policy.name) { $policy.name } else { "Unnamed Policy" }
            $description = if ($policy.description) { $policy.description } else { "No description" }
            $lastModified = if ($policy.lastModifiedDateTime) { $policy.lastModifiedDateTime } else { "Unknown" }

            $assignmentsHtml = "None"
            try {
                $assignmentUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/$($policy.id)/assignments"
                $assignmentItems = Invoke-GraphPagedGet -Uri $assignmentUri

                $assignmentNames = @()
                foreach ($assignment in $assignmentItems) {
                    if ($assignment.target.groupId) {
                        $groupName = Get-GroupNameFromId $assignment.target.groupId
                        $assignmentNames += (Encode-Html $groupName)
                    } else {
                        $tt = $assignment.target.'@odata.type'
                        $assignmentNames += (Encode-Html ("Target Type: $tt"))
                    }
                }

                if ($assignmentNames.Count -gt 0) {
                    $assignmentsHtml = ($assignmentNames -join "<br/>")
                }
            } catch {
                $assignmentsHtml = "<span class='muted'>Error retrieving assignments</span>"
                $configAssignmentFailures++
            }

            $html += "<tr>
              <td>$(Encode-Html $name)</td>
              <td>$(Encode-Html $description)</td>
              <td>$(Encode-Html $lastModified)</td>
              <td><div class='small muted'>$assignmentsHtml</div></td>
            </tr>"
        }

        $html += "</table>"
    } else {
        $html += "<p class='muted'>No configuration policy data returned.</p>"
    }

    $progressBar.Value = 85

    # Compliance policies
    $html += "<h2>Compliance Policies</h2>"
    $html += "<p class='section-note'>Compliance policies set minimum security and health requirements that devices must meet. Devices that fail these checks may be marked noncompliant and restricted from accessing corporate resources.</p>"

    if ($compliancePolicies -and $compliancePolicies.Count -gt 0) {
        $html += "<table><tr><th>Name</th><th>Description</th><th>Last Modified</th></tr>"
        foreach ($cp in $compliancePolicies) {
            $html += "<tr>
              <td>$(Encode-Html $cp.displayName)</td>
              <td>$(Encode-Html $cp.description)</td>
              <td>$(Encode-Html $cp.lastModifiedDateTime)</td>
            </tr>"
        }
        $html += "</table>"
    } else {
        $html += "<p class='muted'>No compliance policy data returned.</p>"
    }

    # Applications
    $html += "<h2>Applications</h2>"
    $html += "<p class='section-note'>This section summarizes applications deployed through Intune, including how they are sourced and maintained. It helps identify standardization opportunities and gaps in application management.</p>"

    if ($apps -and $apps.Count -gt 0) {
        $byKey = @{}
        foreach ($app in $apps) {
            $cls = Get-IntuneAppClassification -App $app
            $key = $cls.Type
            if ($cls.Platform -eq 'Windows' -and (-not [string]::IsNullOrWhiteSpace($cls.Source))) {
                $key = "$($cls.Type) — $($cls.Source)"
            }
            if (-not $byKey.ContainsKey($key)) { $byKey[$key] = 0 }
            $byKey[$key]++
        }

        $html += "<div class='summary-box'><h3>Application Types Summary</h3>"
        $html += "<p class='section-note'>Application counts are grouped by platform and deployment method to provide a high-level overview of the application strategy.</p>"
        $html += "<p class='small'>"
        foreach ($k in ($byKey.Keys | Sort-Object)) {
            $html += "<strong>$(Encode-Html $k):</strong> $($byKey[$k]) &nbsp;&nbsp; "
        }
        $html += "</p></div>"

        $html += "<table><tr>
          <th>Name</th><th>Description</th><th>Version/Metadata</th><th>Publisher</th>
          <th>Last Modified</th><th>Platform</th><th>Type</th><th>Windows Source</th>
        </tr>"

        foreach ($app in $apps) {
            $cls = Get-IntuneAppClassification -App $app
            $odataNorm = Normalize-TypeString (Get-ObjValue -Obj $app -Name '@odata.type')

            $displayName = Get-ObjValue -Obj $app -Name 'displayName'
            $description = Get-ObjValue -Obj $app -Name 'description'
            $publisher   = Get-ObjValue -Obj $app -Name 'publisher'
            $lastMod     = Get-ObjValue -Obj $app -Name 'lastModifiedDateTime'

            $metaParts = @()
            $versionProp = Get-ObjValue -Obj $app -Name 'version'
            if ($versionProp) { $metaParts += "Version: $versionProp" }

            $publishingState = Get-ObjValue -Obj $app -Name 'publishingState'
            if ($publishingState) { $metaParts += "State: $publishingState" }

            $fileName = Get-ObjValue -Obj $app -Name 'fileName'
            if ($fileName) { $metaParts += "File: $fileName" }

            $committed = Get-ObjValue -Obj $app -Name 'committedContentVersion'
            if ($committed) { $metaParts += "Content Version: $committed" }

            if ($metaParts.Count -eq 0) {
                if ($odataNorm -match 'storeapp|wingetapp') { $metaParts += "Store-managed (version not provided here)" }
                else { $metaParts += "N/A" }
            }

            $metaText = $metaParts -join " | "

            $html += "<tr>
              <td>$(Encode-Html $displayName)</td>
              <td>$(Encode-Html $description)</td>
              <td>$(Encode-Html $metaText)</td>
              <td>$(Encode-Html $publisher)</td>
              <td>$(Encode-Html $lastMod)</td>
              <td>$(Encode-Html $cls.Platform)</td>
              <td>$(Encode-Html $cls.Type)</td>
              <td>$(Encode-Html $cls.Source)</td>
            </tr>"
        }
        $html += "</table>"
    } else {
        $html += "<p class='muted'>No application data returned.</p>"
    }

    $html += @"
<div class='recommendations'>
  <h2>Actionable Recommendations</h2>
  <ul>
    <li>Review and remediate <strong>$(Encode-Html $nonCompliantCount) devices</strong> that are not meeting compliance requirements.</li>
    <li>Investigate any configuration policies where assignment details could not be retrieved to ensure proper targeting.</li>
    <li>Validate policy scope to ensure all required user/device populations are covered.</li>
    <li>Review application deployment strategy and confirm required apps are targeted appropriately.</li>
  </ul>
</div>

</body></html>
"@

    $progressBar.Value = 95

    try {
        $downloads = [System.IO.Path]::Combine($env:USERPROFILE, 'Downloads')
        $reportPath = Join-Path $downloads "IntuneAssessmentReport.html"
        $html | Out-File -FilePath $reportPath -Encoding UTF8
        $statusLabel.Text = "Report saved to Downloads: IntuneAssessmentReport.html"
        $progressBar.Value = 100
    } catch {
        $statusLabel.Text = "Error: Could not save report"
    }
})

$form.Topmost = $true
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()
