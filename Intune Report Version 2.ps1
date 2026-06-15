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
        $Obj,
        [string] $Name
    )

    if ($null -eq $Obj) { return $null }
    if (-not $Name) { return $null }

    if ($Obj -is [hashtable]) {
        if ($Obj.ContainsKey($Name)) { return $Obj[$Name] }
        return $null
    }
    else {
        $prop = $Obj.PSObject.Properties | Where-Object { $_.Name -eq $Name }
        if ($prop) { return $prop.Value }
        return $null
    }
}


function Normalize-TypeString {
    param($raw)
    if (-not $raw) { return "" }
    $s = [string]$raw
    $s = $s.Trim()
    if ($s.StartsWith("#")) { $s = $s.Substring(1) }
    return $s.ToLower()
}

function Normalize-Text {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return "" }
    return $Value.ToLower().Trim()
}

function Get-Percent {
    param(
        [int]$Part,
        [int]$Total
    )
    if ($Total -le 0) { return 0 }
    return [int][Math]::Round(($Part / [double]$Total) * 100, 0)
}

function Get-RiskLevel {
    param([int]$Score)

    if ($Score -ge 90) {
        return [PSCustomObject]@{ Label = "Low Risk"; Css = "risk-good" }
    }
    elseif ($Score -ge 75) {
        return [PSCustomObject]@{ Label = "Moderate Risk"; Css = "risk-warn" }
    }
    else {
        return [PSCustomObject]@{ Label = "High Risk"; Css = "risk-bad" }
    }
}

function Get-NormalizedNameKey {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return "" }
    return (($Name.ToLower() -replace '[^a-z0-9]','').Trim())
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
        $g.InterpolationMode  = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $g.SmoothingMode      = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
        $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
        $g.Clear([System.Drawing.Color]::Transparent)
        $g.DrawImage($img, 0, 0, $newW, $newH)
        $g.Dispose()

        $ms = New-Object System.IO.MemoryStream
        $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
        $bytes = $ms.ToArray()
        $b64 = [Convert]::ToBase64String($bytes)
        return "data:image/png;base64,$b64"
    }
    catch {
        return $null
    }
    finally {
        if ($ms)  { $ms.Dispose() }
        if ($bmp) { $bmp.Dispose() }
        if ($img) { $img.Dispose() }
    }
}

# Graph paging helper for Invoke-MgGraphRequest
function Invoke-GraphPagedGet {
    param(
        [Parameter(Mandatory=$true)][string]$Uri,
        [int]$MaxRetries = 3
    )

    $all = @()
    $next = $Uri

    do {
        $resp = $null
        $attempt = 0
        $success = $false

        do {
            try {
                $resp = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType PSObject
                $success = $true
            }
            catch {
                $attempt++
                if ($attempt -ge $MaxRetries) { throw }
                Start-Sleep -Seconds ([Math]::Min(2 * $attempt, 6))
            }
        } while (-not $success)

        if ($resp) {
            if ($resp.PSObject.Properties.Name -contains 'value') {
                $all += @($resp.value)
            }
            else {
                $all += @($resp)
            }
        }

        $next = $null
        if ($resp -and ($resp.PSObject.Properties.Name -contains '@odata.nextLink')) {
            $next = $resp.'@odata.nextLink'
        }
    } while ($next)

    return @($all)
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
    }
    catch {
        $name = "Unknown Group ($groupId)"
    }

    $script:GroupNameCache[$groupId] = $name
    return $name
}

function Get-AssignmentTargetDisplay {
    param($Assignment)

    if ($null -eq $Assignment) { return "Unknown Target" }

    $target = Get-ObjValue -Obj $Assignment -Name 'target'
    if ($null -eq $target) { return "Unknown Target" }

    $groupId = Get-ObjValue -Obj $target -Name 'groupId'
    if (-not [string]::IsNullOrWhiteSpace($groupId)) {
        return (Get-GroupNameFromId $groupId)
    }

    $odataType = Normalize-TypeString (Get-ObjValue -Obj $target -Name '@odata.type')
    if (-not [string]::IsNullOrWhiteSpace($odataType)) {
        switch -Regex ($odataType) {
            'alllicensedusersassignmenttarget' { return "All Users" }
            'alldevicesassignmenttarget'       { return "All Devices" }
            'exclusiongroupassignmenttarget'   {
                $gid = Get-ObjValue -Obj $target -Name 'groupId'
                if ($gid) { return "Exclude: $(Get-GroupNameFromId $gid)" }
                return "Exclude Group"
            }
            default { return "Target Type: $odataType" }
        }
    }

    return "Unknown Target"
}

function Convert-PlatformValue {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }

    $v = $Value.Trim().ToLower()

    # Remove platform noise such as "mdm", "configmgr", etc.
    if ($v -match '^(mdm|configmgr|unknown|all|none)$') { return $null }

    switch -Regex ($v) {
        'windows'            { return 'Windows' }
        'mac|macos'          { return 'macOS' }
        'ios|ipad'           { return 'iOS/iPadOS' }
        'android'            { return 'Android' }
        'linux'              { return 'Linux' }
        'chrome'             { return 'ChromeOS' }
        default              { return $null }
    }
}

function Get-PolicyPlatformLabel {
    param([Parameter(Mandatory=$true)]$Policy)

    $platforms = New-Object System.Collections.Generic.List[string]

    foreach ($propName in @('platforms','platform','platformType')) {
        $propVal = Get-ObjValue -Obj $Policy -Name $propName
        if ($null -eq $propVal) { continue }

        if ($propVal -is [System.Array]) {
            foreach ($item in $propVal) {
                $friendly = Convert-PlatformValue ([string]$item)
                if (-not [string]::IsNullOrWhiteSpace($friendly)) {
                    [void]$platforms.Add($friendly)
                }
            }
        }
        else {
            $friendly = Convert-PlatformValue ([string]$propVal)
            if (-not [string]::IsNullOrWhiteSpace($friendly)) {
                [void]$platforms.Add($friendly)
            }
        }
    }

    $odataType = Normalize-TypeString (Get-ObjValue -Obj $Policy -Name '@odata.type')
    switch -Regex ($odataType) {
        'windows|wufb|featureupdate|qualityupdate' { [void]$platforms.Add('Windows') }
        'macos|mac'                                { [void]$platforms.Add('macOS') }
        'ios|ipad'                                 { [void]$platforms.Add('iOS/iPadOS') }
        'android'                                  { [void]$platforms.Add('Android') }
        'linux'                                    { [void]$platforms.Add('Linux') }
        'chrome'                                   { [void]$platforms.Add('ChromeOS') }
    }

    $templateRef = Get-ObjValue -Obj $Policy -Name 'templateReference'
    $nameBlob = @(
        (Get-ObjValue -Obj $Policy -Name 'displayName')
        (Get-ObjValue -Obj $Policy -Name 'name')
        (Get-ObjValue -Obj $Policy -Name 'description')
        (Get-ObjValue -Obj $Policy -Name 'templateDisplayName')
        (Get-ObjValue -Obj $templateRef -Name 'displayName')
        (Get-ObjValue -Obj $templateRef -Name 'templateFamily')
    ) -join ' | '

    $nameBlob = Normalize-Text $nameBlob

    if ($nameBlob -match '\bwindows\b')     { [void]$platforms.Add('Windows') }
    if ($nameBlob -match '\bmac\b|\bmacos\b') { [void]$platforms.Add('macOS') }
    if ($nameBlob -match '\bios\b|\bipad\b')  { [void]$platforms.Add('iOS/iPadOS') }
    if ($nameBlob -match '\bandroid\b')     { [void]$platforms.Add('Android') }
    if ($nameBlob -match '\blinux\b')       { [void]$platforms.Add('Linux') }
    if ($nameBlob -match '\bchrome\b')      { [void]$platforms.Add('ChromeOS') }

    $distinct = @($platforms | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)

    if ($distinct.Count -eq 0) {
        return "Unknown / Not Exposed by Graph"
    }

    return ($distinct -join ', ')
}

function Get-CompliancePolicyPlatformLabel {
    param([Parameter(Mandatory=$true)]$Policy)

    $odataType = Normalize-TypeString (Get-ObjValue -Obj $Policy -Name '@odata.type')
    switch -Regex ($odataType) {
        'windows' { return 'Windows' }
        'macos|mac' { return 'macOS' }
        'ios|ipad' { return 'iOS/iPadOS' }
        'android' { return 'Android' }
        'linux' { return 'Linux' }
        'chrome' { return 'ChromeOS' }
    }

    $displayName = Normalize-Text (Get-ObjValue -Obj $Policy -Name 'displayName')
    if ($displayName -match '\bwindows\b')   { return 'Windows' }
    if ($displayName -match '\bmac\b|\bmacos\b') { return 'macOS' }
    if ($displayName -match '\bios\b|\bipad\b')  { return 'iOS/iPadOS' }
    if ($displayName -match '\bandroid\b')   { return 'Android' }
    if ($displayName -match '\blinux\b')     { return 'Linux' }
    if ($displayName -match '\bchrome\b')    { return 'ChromeOS' }

    return "Unknown / Not Exposed by Graph"
}

function Get-ClassicConfigPolicyCategory {
    param([Parameter(Mandatory=$true)]$Policy)

    $odataType = Normalize-TypeString (Get-ObjValue -Obj $Policy -Name '@odata.type')

    if ($odataType -eq 'microsoft.graph.windowsupdateforbusinessconfiguration') {
        return 'Windows Update'
    }

    return 'Classic Device Configuration'
}

function Get-AssignmentSummary {
    param([Parameter(Mandatory=$true)][string]$Uri)

    $assignmentNames = New-Object System.Collections.Generic.List[string]
    $assignmentHtml = "<span class='muted'>None assigned</span>"
    $assignmentCount = 0

    $assignmentItems = Invoke-GraphPagedGet -Uri $Uri

    foreach ($assignment in $assignmentItems) {
        $targetName = Get-AssignmentTargetDisplay -Assignment $assignment
        if (-not [string]::IsNullOrWhiteSpace($targetName)) {
            [void]$assignmentNames.Add($targetName)
        }
    }

    $assignmentNames = @($assignmentNames | Sort-Object -Unique)
    $assignmentCount = @($assignmentNames).Count

    if ($assignmentCount -gt 0) {
        $assignmentHtml = ((@($assignmentNames) | ForEach-Object { Encode-Html $_ }) -join "<br/>")
    }

    return [PSCustomObject]@{
        Html  = $assignmentHtml
        Names = @($assignmentNames)
        Count = $assignmentCount
    }
}

function New-PlatformCounter {
    return [ordered]@{
        'Windows' = 0
        'macOS' = 0
        'iOS/iPadOS' = 0
        'Android' = 0
        'Linux' = 0
        'ChromeOS' = 0
        'Unknown / Not Exposed by Graph' = 0
        'Other' = 0
    }
}

function Add-PlatformCount {
    param(
        [Parameter(Mandatory=$true)]$Map,
        [string]$PlatformLabel
    )

    if ([string]::IsNullOrWhiteSpace($PlatformLabel)) {
        $Map['Unknown / Not Exposed by Graph']++
        return
    }

    $parts = $PlatformLabel -split '\s*,\s*'
    foreach ($part in $parts) {
        if ($Map.Contains($part)) {
            $Map[$part]++
        }
        else {
            $Map['Other']++
        }
    }
}

function Get-EndpointSecurityCategoryFromIntent {
    param([Parameter(Mandatory=$true)] $Policy)

    $templateId          = Normalize-Text (Get-ObjValue -Obj $Policy -Name 'templateId')
    $templateDisplayName = Normalize-Text (Get-ObjValue -Obj $Policy -Name 'templateDisplayName')
    $displayName         = Normalize-Text (Get-ObjValue -Obj $Policy -Name 'displayName')
    $description         = Normalize-Text (Get-ObjValue -Obj $Policy -Name 'description')

    $composite = @(
        $templateId
        $templateDisplayName
        $displayName
        $description
    ) -join " | "

    if ($composite -match 'antivirus|defender antivirus|defender update controls|windows security experience') {
        return "Antivirus"
    }
    if ($composite -match 'bitlocker|disk encryption') {
        return "Disk Encryption"
    }
    if ($composite -match 'firewall') {
        return "Firewall"
    }
    if ($composite -match 'endpoint detection and response|edr|sense|defender for endpoint') {
        return "Endpoint Detection and Response"
    }
    if ($composite -match 'attack surface reduction|asr|web protection') {
        return "Attack Surface Reduction"
    }
    if ($composite -match 'account protection|laps|hello|credential') {
        return "Account Protection"
    }

    return "Other Endpoint Security"
}

function Get-EndpointSecurityCategoryFromConfigPolicy {
    param([Parameter(Mandatory=$true)] $Policy)

    $name         = Normalize-Text (Get-ObjValue -Obj $Policy -Name 'name')
    $description  = Normalize-Text (Get-ObjValue -Obj $Policy -Name 'description')
    $templateRef  = Get-ObjValue -Obj $Policy -Name 'templateReference'
    $templateFam  = ""
    $templateName = ""

    if ($templateRef) {
        $templateFam  = Normalize-Text (Get-ObjValue -Obj $templateRef -Name 'templateFamily')
        $templateName = Normalize-Text (Get-ObjValue -Obj $templateRef -Name 'displayName')
    }

    $composite = @(
        $name
        $description
        $templateFam
        $templateName
    ) -join " | "

    if ($composite -match 'antivirus|defender antivirus|defender update controls|windows security experience') {
        return "Antivirus"
    }
    if ($composite -match 'bitlocker|disk encryption') {
        return "Disk Encryption"
    }
    if ($composite -match 'firewall') {
        return "Firewall"
    }
    if ($composite -match 'endpoint detection and response|edr|sense|defender for endpoint') {
        return "Endpoint Detection and Response"
    }
    if ($composite -match 'attack surface reduction|asr|web protection|app control|exploit') {
        return "Attack Surface Reduction"
    }
    if ($composite -match 'account protection|laps|hello|credential|local admin') {
        return "Account Protection"
    }

    return $null
}

function Get-WindowsUpdatePolicyCategory {
    param([Parameter(Mandatory=$true)] $Policy)

    $odataType    = Normalize-Text (Get-ObjValue -Obj $Policy -Name '@odata.type')
    $displayName  = Normalize-Text (Get-ObjValue -Obj $Policy -Name 'displayName')
    $name         = Normalize-Text (Get-ObjValue -Obj $Policy -Name 'name')
    $description  = Normalize-Text (Get-ObjValue -Obj $Policy -Name 'description')
    $templateRef  = Get-ObjValue -Obj $Policy -Name 'templateReference'
    $templateFam  = ""
    $templateName = ""

    if ($templateRef) {
        $templateFam  = Normalize-Text (Get-ObjValue -Obj $templateRef -Name 'templateFamily')
        $templateName = Normalize-Text (Get-ObjValue -Obj $templateRef -Name 'displayName')
    }

    $composite = @(
        $odataType
        $displayName
        $name
        $description
        $templateFam
        $templateName
    ) -join " | "

    if ($odataType -eq 'microsoft.graph.windowsupdateforbusinessconfiguration') {
        return "Update Ring"
    }
    if ($composite -match 'feature update') {
        return "Feature Update Profile"
    }
    if ($composite -match 'quality update') {
        return "Quality Update Profile"
    }
    if ($composite -match 'windows update|update ring|expedite update|quality updates|feature updates|windows updates for business|wufb') {
        return "Other Windows Update Policy"
    }

    return $null
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
        'microsoft\.graph\.androidstoreapp'                  { $cls.Platform='Android'; $cls.Type='Android (Public Store)'; return $cls }
        'microsoft\.graph\.managedgoogleplay(app|storeapp)' { $cls.Platform='Android'; $cls.Type='Android Enterprise (Managed Play)'; return $cls }
        'microsoft\.graph\.iosstoreapp'                     { $cls.Platform='iOS/iPadOS'; $cls.Type='iOS/iPadOS (App Store)'; return $cls }
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
        'microsoft\.graph\.webapp$'                 { $cls.Platform='Other'; $cls.Type='Web App (Browser-based)'; return $cls }
        'microsoft\.graph\.ioswebclip'              { $cls.Platform='iOS/iPadOS'; $cls.Type='Platform Link Shortcut'; return $cls }
        'microsoft\.graph\.windowswebapp'           { $cls.Platform='Windows'; $cls.Type='Platform Link Shortcut'; return $cls }
        'microsoft\.graph\.macoswebclip'            { $cls.Platform='macOS'; $cls.Type='Platform Link Shortcut'; return $cls }
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

    # -----------------------------
    # Tenant details
    # -----------------------------
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

    # -----------------------------
    # Counters, data, and state
    # -----------------------------
    $devicesCount = 0
    $nonCompliantCount = 0
    $compliantCount = 0
    $encryptedCount = 0
    $staleDeviceCount = 0

    $modernConfigPoliciesCount = 0
    $classicDeviceConfigurationsCount = 0
    $configPoliciesCount = 0
    $configAssignmentFailures = 0
    $unassignedConfigPolicyCount = 0

    $compliancePoliciesCount = 0
    $appsCount = 0

    $staleDeviceThresholdDays = 30
    $staleDevices = @()
    $unassignedConfigPolicies = @()
    $configPolicyAssignmentMap = @{}

    $modernConfigPolicies = @()
    $classicDeviceConfigurations = @()
    $allAssessmentConfigPolicies = @()

    $modernConfigPoliciesForSection = @()
    $classicConfigPoliciesForSection = @()
    $generalConfigPolicies = @()

    $endpointSecurityPolicies = @()
    $endpointSecurityCategoryCounts = [ordered]@{
        "Antivirus" = 0
        "Disk Encryption" = 0
        "Firewall" = 0
        "Endpoint Detection and Response" = 0
        "Attack Surface Reduction" = 0
        "Account Protection" = 0
        "Other Endpoint Security" = 0
    }
    $endpointSecurityAssignmentMap = @{}
    $endpointSecurityDisplayRows = @()
    $endpointSecurityConfigPolicyIds = @{}

    $updateRingPolicies = @()
    $featureUpdateProfiles = @()
    $qualityUpdateProfiles = @()
    $otherWindowsUpdatePolicies = @()
    $windowsUpdateAssignmentMap = @{}
    $updateRingAssignmentMap = @{}

    $conflictFindings = New-Object System.Collections.Generic.List[object]

    # -----------------------------
    # Devices
    # -----------------------------
    $devices = @()
    try {
        $devices = Get-MgDeviceManagementManagedDevice -All -Property "deviceName,operatingSystem,osVersion,complianceState,isEncrypted,lastSyncDateTime"
        $devicesCount = $devices.Count
        $nonCompliantCount = ($devices | Where-Object { $_.ComplianceState -ne "compliant" }).Count
        $compliantCount = ($devices | Where-Object { $_.ComplianceState -eq "compliant" }).Count
        $encryptedCount = ($devices | Where-Object { $_.IsEncrypted -eq $true }).Count

        $staleCutoff = (Get-Date).AddDays(-$staleDeviceThresholdDays)
        $staleDevices = $devices | Where-Object {
            $_.LastSyncDateTime -and ([datetime]$_.LastSyncDateTime -lt $staleCutoff)
        }
        $staleDeviceCount = $staleDevices.Count
    } catch {
        $devices = @()
        $devicesCount = 0
        $nonCompliantCount = 0
        $compliantCount = 0
        $encryptedCount = 0
        $staleDeviceCount = 0
        $collectionErrors.Add("Devices: $($_.Exception.Message)")
    }

    $progressBar.Value = 16

    # -----------------------------
    # Modern Configuration Policies
    # -----------------------------
    try {
        $modernRaw = Invoke-GraphPagedGet -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?`$top=100"
        foreach ($policy in $modernRaw) {
            $policyName = if ($policy.name) { $policy.name } else { "Unnamed Policy" }

            $modernConfigPolicies += [PSCustomObject]@{
                Id            = $policy.id
                Name          = $policyName
                Description   = $policy.description
                LastModified  = $policy.lastModifiedDateTime
                Platform      = Get-PolicyPlatformLabel -Policy $policy
                PolicyStore   = "Modern Configuration Policy"
                Category      = "Modern Configuration Policy"
                AssignmentUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/$($policy.id)/assignments"
                Raw           = $policy
            }
        }
        $modernConfigPoliciesCount = @($modernConfigPolicies).Count
    } catch {
        $modernConfigPolicies = @()
        $modernConfigPoliciesCount = 0
        $collectionErrors.Add("Modern Configuration Policies: $($_.Exception.Message)")
    }

    # -----------------------------
    # Classic Device Configurations
    # -----------------------------
    try {
        $classicRaw = Invoke-GraphPagedGet -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations?`$top=100"
        foreach ($policy in $classicRaw) {
            $category = Get-ClassicConfigPolicyCategory -Policy $policy
            $policyName = if ($policy.displayName) { $policy.displayName } else { "Unnamed Classic Policy" }

            $classicDeviceConfigurations += [PSCustomObject]@{
                Id            = $policy.id
                Name          = $policyName
                Description   = $policy.description
                LastModified  = $policy.lastModifiedDateTime
                Platform      = Get-PolicyPlatformLabel -Policy $policy
                PolicyStore   = "Classic Device Configuration"
                Category      = $category
                AssignmentUri = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations/$($policy.id)/assignments"
                Raw           = $policy
            }
        }
        $classicDeviceConfigurationsCount = @($classicDeviceConfigurations).Count
    } catch {
        $classicDeviceConfigurations = @()
        $classicDeviceConfigurationsCount = 0
        $collectionErrors.Add("Classic Device Configurations: $($_.Exception.Message)")
    }

    $allAssessmentConfigPolicies = @($modernConfigPolicies + $classicDeviceConfigurations)
    $configPoliciesCount = @($allAssessmentConfigPolicies).Count

    if ($allAssessmentConfigPolicies.Count -gt 0) {
        foreach ($policy in $allAssessmentConfigPolicies) {
            try {
                $assignmentSummary = Get-AssignmentSummary -Uri $policy.AssignmentUri
                $configPolicyAssignmentMap[$policy.Id] = $assignmentSummary

                if ($assignmentSummary.Count -eq 0) {
                    $unassignedConfigPolicies += $policy
                }
            } catch {
                $configAssignmentFailures++
                $configPolicyAssignmentMap[$policy.Id] = [PSCustomObject]@{
                    Html  = "<span class='muted'>Error retrieving assignments</span>"
                    Names = @()
                    Count = 0
                }
            }
        }
        $unassignedConfigPolicyCount = @($unassignedConfigPolicies).Count
    }

    $progressBar.Value = 28

    # -----------------------------
    # Compliance Policies
    # -----------------------------
    $compliancePolicies = @()
    try {
        $compliancePolicies = Invoke-GraphPagedGet -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies?`$top=100"
        $compliancePoliciesCount = @($compliancePolicies).Count
    } catch {
        $compliancePolicies = @()
        $compliancePoliciesCount = 0
        $collectionErrors.Add("Compliance Policies: $($_.Exception.Message)")
    }

    $progressBar.Value = 40

    # -----------------------------
    # Applications
    # -----------------------------
    $apps = @()
    try {
        $apps = Invoke-GraphPagedGet -Uri "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?`$top=100"
        $appsCount = $apps.Count
    } catch {
        $apps = @()
        $appsCount = 0
        $collectionErrors.Add("Applications: $($_.Exception.Message)")
    }

    $progressBar.Value = 52

    # -----------------------------
    # Endpoint Security policies
    # Combines:
    # 1) Legacy Endpoint Security intents
    # 2) Endpoint Security-backed configuration policies
    # -----------------------------
    try {
        $endpointSecurityPolicies = @()
        $endpointSecurityDisplayRows = @()
        $endpointSecurityConfigPolicyIds = @{}

        $endpointSecurityCategoryCounts = [ordered]@{
            "Antivirus" = 0
            "Disk Encryption" = 0
            "Firewall" = 0
            "Endpoint Detection and Response" = 0
            "Attack Surface Reduction" = 0
            "Account Protection" = 0
            "Other Endpoint Security" = 0
        }

        # -------------------------
        # Legacy intents
        # -------------------------
        $legacyIntentPolicies = @()
        try {
            $legacyIntentPolicies = Invoke-GraphPagedGet -Uri "https://graph.microsoft.com/beta/deviceManagement/intents?`$top=100"
            $legacyIntentPolicies = $legacyIntentPolicies |
                Group-Object id |
                ForEach-Object { $_.Group | Select-Object -First 1 }
        } catch {
            $legacyIntentPolicies = @()
            $collectionErrors.Add("Endpoint Security Legacy Intents: $($_.Exception.Message)")
        }

        foreach ($policy in $legacyIntentPolicies) {
            $category = Get-EndpointSecurityCategoryFromIntent -Policy $policy

            $assignmentNames = New-Object System.Collections.Generic.List[string]
            $assignmentHtml  = "<span class='muted'>None assigned</span>"
            $assignmentCount = 0
            $isAssigned = $false

            try {
                $intentAssignments = Invoke-GraphPagedGet -Uri "https://graph.microsoft.com/beta/deviceManagement/intents/$($policy.id)/assignments?`$top=100"

                foreach ($assignment in $intentAssignments) {
                    $targetName = Get-AssignmentTargetDisplay -Assignment $assignment
                    if (-not [string]::IsNullOrWhiteSpace($targetName)) {
                        [void]$assignmentNames.Add($targetName)
                    }
                }

                $assignmentNames = $assignmentNames | Sort-Object -Unique
                $assignmentCount = @($assignmentNames).Count

                if ($assignmentCount -gt 0) {
                    $assignmentHtml = ((@($assignmentNames) | ForEach-Object { Encode-Html $_ }) -join "<br/>")
                    $isAssigned = $true
                }
            } catch {
                $assignmentHtml = "<span class='muted'>Error retrieving assignments</span>"
                $collectionErrors.Add("Endpoint Security Assignments [$($policy.displayName)]: $($_.Exception.Message)")
            }

            if (-not $isAssigned -and $policy.isAssigned -eq $true) {
                $isAssigned = $true
            }

            $policyName = $policy.displayName
            if ([string]::IsNullOrWhiteSpace($policyName)) {
                $policyName = $policy.templateDisplayName
            }
            if ([string]::IsNullOrWhiteSpace($policyName)) {
                $policyName = "Unnamed Endpoint Security Policy ($($policy.id))"
            }

            $platform = Get-PolicyPlatformLabel -Policy $policy

            $endpointSecurityAssignmentMap[$policy.id] = [PSCustomObject]@{
                Html       = $assignmentHtml
                Names      = @($assignmentNames)
                Count      = $assignmentCount
                Category   = $category
                IsAssigned = $isAssigned
                Name       = $policyName
                Source     = "Legacy Intent"
                Platform   = $platform
            }

            $endpointSecurityDisplayRows += [PSCustomObject]@{
                Id           = $policy.id
                Name         = $policyName
                Category     = $category
                Platform     = $platform
                Assigned     = $isAssigned
                Assignments  = $assignmentHtml
                LastModified = $policy.lastModifiedDateTime
                SortName     = $policyName
                Source       = "Legacy Intent"
            }

            if ($endpointSecurityCategoryCounts.Contains($category)) {
                $endpointSecurityCategoryCounts[$category]++
            } else {
                $endpointSecurityCategoryCounts["Other Endpoint Security"]++
            }
        }

        # -------------------------
        # Endpoint Security-backed modern configuration policies
        # -------------------------
        foreach ($wrapper in $modernConfigPolicies) {
            $policy = $wrapper.Raw
            $category = Get-EndpointSecurityCategoryFromConfigPolicy -Policy $policy

            if ($null -ne $category) {
                $endpointSecurityConfigPolicyIds[$wrapper.Id] = $true

                $assignmentHtml = "<span class='muted'>None assigned</span>"
                $assignmentNames = @()
                $assignmentCount = 0
                $isAssigned = $false

                if ($configPolicyAssignmentMap.ContainsKey($wrapper.Id)) {
                    $assignmentHtml = $configPolicyAssignmentMap[$wrapper.Id].Html
                    $assignmentNames = $configPolicyAssignmentMap[$wrapper.Id].Names
                    $assignmentCount = $configPolicyAssignmentMap[$wrapper.Id].Count
                    if ($assignmentCount -gt 0) { $isAssigned = $true }
                }

                $endpointSecurityAssignmentMap[$wrapper.Id] = [PSCustomObject]@{
                    Html       = $assignmentHtml
                    Names      = @($assignmentNames)
                    Count      = $assignmentCount
                    Category   = $category
                    IsAssigned = $isAssigned
                    Name       = $wrapper.Name
                    Source     = "Endpoint Security (Config Policy)"
                    Platform   = $wrapper.Platform
                }

                $endpointSecurityDisplayRows += [PSCustomObject]@{
                    Id           = $wrapper.Id
                    Name         = $wrapper.Name
                    Category     = $category
                    Platform     = $wrapper.Platform
                    Assigned     = $isAssigned
                    Assignments  = $assignmentHtml
                    LastModified = $wrapper.LastModified
                    SortName     = $wrapper.Name
                    Source       = "Endpoint Security (Config Policy)"
                }

                if ($endpointSecurityCategoryCounts.Contains($category)) {
                    $endpointSecurityCategoryCounts[$category]++
                } else {
                    $endpointSecurityCategoryCounts["Other Endpoint Security"]++
                }
            }
        }

        $endpointSecurityDisplayRows = $endpointSecurityDisplayRows | Sort-Object Category, SortName
        $endpointSecurityPolicies = $endpointSecurityDisplayRows
    }
    catch {
        $endpointSecurityPolicies = @()
        $endpointSecurityDisplayRows = @()
        $collectionErrors.Add("Endpoint Security Policies: $($_.Exception.Message)")
    }

    # -----------------------------
    # Windows Update policies
    # -----------------------------
    try {
        $updateRingPolicies = @($classicDeviceConfigurations | Where-Object {
            $_.Category -eq 'Windows Update'
        })

        foreach ($ring in $updateRingPolicies) {
            try {
                $summary = Get-AssignmentSummary -Uri $ring.AssignmentUri
                $updateRingAssignmentMap[$ring.Id] = $summary
            } catch {
                $updateRingAssignmentMap[$ring.Id] = [PSCustomObject]@{
                    Html  = "<span class='muted'>Error retrieving assignments</span>"
                    Names = @()
                    Count = 0
                }
                $collectionErrors.Add("Classic Update Ring Assignments [$($ring.Name)]: $($_.Exception.Message)")
            }
        }
    } catch {
        $updateRingPolicies = @()
        $collectionErrors.Add("Classic Update Rings: $($_.Exception.Message)")
    }

    try {
        $featureUpdateProfiles = Invoke-GraphPagedGet -Uri "https://graph.microsoft.com/beta/deviceManagement/windowsFeatureUpdateProfiles?`$top=100"
    } catch {
        $featureUpdateProfiles = @()
        $collectionErrors.Add("Feature Update Profiles: $($_.Exception.Message)")
    }

    try {
        $qualityUpdateProfiles = Invoke-GraphPagedGet -Uri "https://graph.microsoft.com/beta/deviceManagement/windowsQualityUpdateProfiles?`$top=100"
    } catch {
        $qualityUpdateProfiles = @()
        $collectionErrors.Add("Quality Update Profiles: $($_.Exception.Message)")
    }

    try {
        $otherWindowsUpdatePolicies = @(
            $modernConfigPolicies | Where-Object {
                (Get-WindowsUpdatePolicyCategory -Policy $_.Raw) -eq 'Other Windows Update Policy'
            }
        )

        foreach ($policy in $otherWindowsUpdatePolicies) {
            if ($configPolicyAssignmentMap.ContainsKey($policy.Id)) {
                $windowsUpdateAssignmentMap[$policy.Id] = [PSCustomObject]@{
                    Html     = $configPolicyAssignmentMap[$policy.Id].Html
                    Names    = @($configPolicyAssignmentMap[$policy.Id].Names)
                    Count    = $configPolicyAssignmentMap[$policy.Id].Count
                    Category = "Other Windows Update Policy"
                }
            }
            else {
                $windowsUpdateAssignmentMap[$policy.Id] = [PSCustomObject]@{
                    Html     = "<span class='muted'>None assigned</span>"
                    Names    = @()
                    Count    = 0
                    Category = "Other Windows Update Policy"
                }
            }
        }
    } catch {
        $otherWindowsUpdatePolicies = @()
        $collectionErrors.Add("Other Windows Update Policies: $($_.Exception.Message)")
    }

    # -----------------------------
    # Build general config sections AFTER classification
    # -----------------------------
    $modernConfigPoliciesForSection = @(
        $modernConfigPolicies | Where-Object {
            -not $endpointSecurityConfigPolicyIds.ContainsKey($_.Id) -and
            (Get-WindowsUpdatePolicyCategory -Policy $_.Raw) -eq $null
        }
    )

    $classicConfigPoliciesForSection = @(
        $classicDeviceConfigurations | Where-Object {
            $_.Category -ne 'Windows Update'
        }
    )

    $generalConfigPolicies = @($modernConfigPoliciesForSection + $classicConfigPoliciesForSection)

    $progressBar.Value = 74

    # -----------------------------
    # Scoring / top risks / dynamic recommendations
    # -----------------------------
    $complianceScore = Get-Percent -Part $compliantCount -Total $devicesCount
    $encryptionScore = Get-Percent -Part $encryptedCount -Total $devicesCount

    $assignedPolicyCount = [Math]::Max(($configPoliciesCount - $unassignedConfigPolicyCount), 0)
    $configHygieneScore = Get-Percent -Part $assignedPolicyCount -Total $configPoliciesCount

    $endpointSecurityImplementedCategories = ($endpointSecurityCategoryCounts.GetEnumerator() | Where-Object {
        $_.Key -ne "Other Endpoint Security" -and $_.Value -gt 0
    }).Count
    $endpointSecurityScore = Get-Percent -Part $endpointSecurityImplementedCategories -Total 6

    $overallScore = [int][Math]::Round((($complianceScore + $encryptionScore + $configHygieneScore + $endpointSecurityScore) / 4), 0)

    $overallRisk    = Get-RiskLevel -Score $overallScore
    $complianceRisk = Get-RiskLevel -Score $complianceScore
    $encryptionRisk = Get-RiskLevel -Score $encryptionScore
    $configRisk     = Get-RiskLevel -Score $configHygieneScore
    $endpointRisk   = Get-RiskLevel -Score $endpointSecurityScore

    $topRisks = New-Object System.Collections.Generic.List[string]
    $dynamicRecommendations = New-Object System.Collections.Generic.List[string]

    if ($nonCompliantCount -gt 0) {
        [void]$topRisks.Add("$nonCompliantCount device(s) are currently noncompliant.")
        [void]$dynamicRecommendations.Add("Review the noncompliant device population and validate that compliance policies, remediation actions, and access controls are aligned.")
    }

    if ($devicesCount -gt 0 -and $encryptedCount -lt $devicesCount) {
        $unencryptedCount = $devicesCount - $encryptedCount
        [void]$topRisks.Add("$unencryptedCount device(s) are not reporting encryption enabled.")
        [void]$dynamicRecommendations.Add("Prioritize remediation for devices that are not encrypted and confirm BitLocker/FileVault enforcement is configured where applicable.")
    }

    if ($staleDeviceCount -gt 0) {
        [void]$topRisks.Add("$staleDeviceCount device(s) have not synced in more than $staleDeviceThresholdDays days.")
        [void]$dynamicRecommendations.Add("Review stale devices for cleanup, retirement, or troubleshooting to improve reporting accuracy and reduce management overhead.")
    }

    if ($unassignedConfigPolicyCount -gt 0) {
        $policyPhrase = if ($unassignedConfigPolicyCount -eq 1) { "policy is" } else { "policies are" }
        [void]$topRisks.Add("$unassignedConfigPolicyCount configuration $policyPhrase currently unassigned.")
        [void]$dynamicRecommendations.Add("Validate whether unassigned configuration policies are still needed; remove obsolete policies or assign required policies to the appropriate target groups.")
    }

    if ($endpointSecurityImplementedCategories -lt 6) {
        $missingCategories = @()
        foreach ($kv in $endpointSecurityCategoryCounts.GetEnumerator()) {
            if ($kv.Key -ne "Other Endpoint Security" -and $kv.Value -eq 0) {
                $missingCategories += $kv.Key
            }
        }
        if ($missingCategories.Count -gt 0) {
            [void]$topRisks.Add("Endpoint security coverage appears incomplete across recommended policy categories.")
            [void]$dynamicRecommendations.Add("Review missing endpoint security policy areas: $($missingCategories -join ', ').")
        }
    }

    $totalUpdateProfiles =
        @($updateRingPolicies).Count +
        @($featureUpdateProfiles).Count +
        @($qualityUpdateProfiles).Count +
        @($otherWindowsUpdatePolicies).Count

    if ($totalUpdateProfiles -eq 0) {
        [void]$topRisks.Add("No Windows Update governance policies were detected.")
        [void]$dynamicRecommendations.Add("Implement and validate Windows Update governance using update rings, feature update profiles, quality update profiles, or update-related configuration policies.")
    }

    if ($topRisks.Count -eq 0) {
        [void]$topRisks.Add("No critical risks were identified from the data collected by this report.")
    }
    if ($dynamicRecommendations.Count -eq 0) {
        [void]$dynamicRecommendations.Add("Maintain the current posture and continue periodic review of policies, device health, and update governance.")
    }

    # -----------------------------
    # Potential Policy Conflict / Overlap Analysis
    # -----------------------------
    $endpointOverlapMap = @{}
    foreach ($row in $endpointSecurityDisplayRows) {
        if (-not $endpointSecurityAssignmentMap.ContainsKey($row.Id)) { continue }

        $category = $endpointSecurityAssignmentMap[$row.Id].Category
        $targets  = @($endpointSecurityAssignmentMap[$row.Id].Names)

        foreach ($target in $targets) {
            if ([string]::IsNullOrWhiteSpace($target)) { continue }
            $key = "$category|$target"

            if (-not $endpointOverlapMap.ContainsKey($key)) {
                $endpointOverlapMap[$key] = New-Object System.Collections.Generic.List[string]
            }
            [void]$endpointOverlapMap[$key].Add($endpointSecurityAssignmentMap[$row.Id].Name)
        }
    }

    foreach ($key in $endpointOverlapMap.Keys) {
        $items = @($endpointOverlapMap[$key] | Sort-Object -Unique)
        if ($items.Count -gt 1) {
            $parts = $key -split '\|', 2
            $category = $parts[0]
            $target   = $parts[1]

            [void]$conflictFindings.Add([PSCustomObject]@{
                Area    = "Endpoint Security"
                Type    = "Potential overlap"
                Scope   = $target
                Details = "Multiple $category policies are assigned to the same target: " + ($items -join ", ")
            })
        }
    }

    $updateOverlapMap = @{}
    foreach ($ring in $updateRingPolicies) {
        if (-not $updateRingAssignmentMap.ContainsKey($ring.Id)) { continue }
        $targets = @($updateRingAssignmentMap[$ring.Id].Names)

        foreach ($target in $targets) {
            if ([string]::IsNullOrWhiteSpace($target)) { continue }
            $key = "Update Ring|$target"

            if (-not $updateOverlapMap.ContainsKey($key)) {
                $updateOverlapMap[$key] = New-Object System.Collections.Generic.List[string]
            }
            [void]$updateOverlapMap[$key].Add($ring.Name)
        }
    }

    foreach ($key in $updateOverlapMap.Keys) {
        $items = @($updateOverlapMap[$key] | Sort-Object -Unique)
        if ($items.Count -gt 1) {
            $parts = $key -split '\|', 2
            $scope = $parts[1]

            [void]$conflictFindings.Add([PSCustomObject]@{
                Area    = "Windows Update"
                Type    = "Potential overlap"
                Scope   = $scope
                Details = "Multiple classic update rings are assigned to the same target: " + ($items -join ", ")
            })
        }
    }

    $configNameMap = @{}
    foreach ($policy in $allAssessmentConfigPolicies) {
        $nameKey = Get-NormalizedNameKey -Name $policy.Name
        if ([string]::IsNullOrWhiteSpace($nameKey)) { continue }

        if (-not $configNameMap.ContainsKey($nameKey)) {
            $configNameMap[$nameKey] = New-Object System.Collections.Generic.List[string]
        }
        [void]$configNameMap[$nameKey].Add($policy.Name)
    }

    foreach ($key in $configNameMap.Keys) {
        $items = @($configNameMap[$key] | Sort-Object -Unique)
        if ($items.Count -gt 1) {
            [void]$conflictFindings.Add([PSCustomObject]@{
                Area    = "Configuration Policies"
                Type    = "Duplicate naming"
                Scope   = "Environment-wide"
                Details = "Multiple configuration policies appear to have highly similar names: " + ($items -join ", ")
            })
        }
    }

    foreach ($configPolicy in $allAssessmentConfigPolicies) {
        if ($endpointSecurityConfigPolicyIds.ContainsKey($configPolicy.Id)) { continue }

        $configTargets = @()
        if ($configPolicyAssignmentMap.ContainsKey($configPolicy.Id)) {
            $configTargets = @($configPolicyAssignmentMap[$configPolicy.Id].Names)
        }

        if ($configPolicy.Name -match 'antivirus|defender|bitlocker|encryption|firewall|edr|endpoint detection|attack surface reduction|asr|account protection|laps') {
            foreach ($row in $endpointSecurityDisplayRows) {
                $espTargets = @()
                $espName = $row.Name
                if ($endpointSecurityAssignmentMap.ContainsKey($row.Id)) {
                    $espTargets = @($endpointSecurityAssignmentMap[$row.Id].Names)
                    $espName = $endpointSecurityAssignmentMap[$row.Id].Name
                }

                $sharedTargets = $configTargets | Where-Object { $espTargets -contains $_ } | Sort-Object -Unique
                if (@($sharedTargets).Count -gt 0) {
                    [void]$conflictFindings.Add([PSCustomObject]@{
                        Area    = "Configuration + Endpoint Security"
                        Type    = "Cross-source overlap"
                        Scope   = ($sharedTargets -join ", ")
                        Details = "Configuration policy '$($configPolicy.Name)' may overlap with Endpoint Security policy '$espName' because both appear to target the same scope."
                    })
                }
            }
        }
    }

    $progressBar.Value = 82

    # -----------------------------
    # Platform coverage summary
    # -----------------------------
    $configPlatformCounts = New-PlatformCounter
    $compliancePlatformCounts = New-PlatformCounter
    $appPlatformCounts = New-PlatformCounter
    $devicePlatformCounts = New-PlatformCounter
    $endpointPlatformCounts = New-PlatformCounter

    foreach ($p in $allAssessmentConfigPolicies) {
        Add-PlatformCount -Map $configPlatformCounts -PlatformLabel $p.Platform
    }

    foreach ($p in $compliancePolicies) {
        Add-PlatformCount -Map $compliancePlatformCounts -PlatformLabel (Get-CompliancePolicyPlatformLabel -Policy $p)
    }

    foreach ($app in $apps) {
        $cls = Get-IntuneAppClassification -App $app
        Add-PlatformCount -Map $appPlatformCounts -PlatformLabel $cls.Platform
    }

    foreach ($device in $devices) {
        Add-PlatformCount -Map $devicePlatformCounts -PlatformLabel $device.OperatingSystem
    }

    foreach ($row in $endpointSecurityDisplayRows) {
        Add-PlatformCount -Map $endpointPlatformCounts -PlatformLabel $row.Platform
    }

    # -----------------------------
    # Client-ready HTML (with optional logo)
    # -----------------------------
    $logoHtml = ""
    if ($script:LogoDataUri) {
        $logoHtml = "<img class='logo' src='$($script:LogoDataUri)' alt='Logo'/>"
    }

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
  .score-grid {
    display:grid;
    grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    gap:12px;
    margin-top: 12px;
    margin-bottom: 18px;
  }
  .score-card {
    background:#ffffff;
    border:1px solid #dfe6e9;
    border-radius:8px;
    padding:12px;
  }
  .score-card .score-title {
    font-size:12px;
    color:#666;
    margin-bottom:6px;
    text-transform:uppercase;
    letter-spacing:.4px;
  }
  .score-card .score-value {
    font-size:26px;
    font-weight:bold;
    color:#2c3e50;
    line-height:1.1;
  }
  .score-card .score-label {
    margin-top:6px;
    font-size:12px;
    font-weight:bold;
    display:inline-block;
    padding:4px 8px;
    border-radius:999px;
  }
  .risk-good { background:#e8f8f1; color:#1e8449; }
  .risk-warn { background:#fff8e1; color:#b9770e; }
  .risk-bad  { background:#fdecea; color:#c0392b; }
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
    application deployment state, and platform coverage while identifying priority areas for optimization
    and risk reduction.
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
"@

    $html += @"
<div class='summary-box'>
  <h2>High-Level Summary</h2>
  <p>
    <strong>Total Devices:</strong> $(Encode-Html $devicesCount)<br/>
    <strong>Noncompliant Devices:</strong> $(Encode-Html $nonCompliantCount)<br/>
    <strong>Total Configuration Policies (All Types):</strong> $(Encode-Html $configPoliciesCount)<br/>
    <strong>Modern Configuration Policies:</strong> $(Encode-Html $modernConfigPoliciesCount)<br/>
    <strong>Classic Device Configurations:</strong> $(Encode-Html $classicDeviceConfigurationsCount)<br/>
    <strong>Unassigned Configuration Policies:</strong> $(Encode-Html $unassignedConfigPolicyCount)<br/>
    <strong>Configuration Assignment Retrieval Failures:</strong> $(Encode-Html $configAssignmentFailures)<br/>
    <strong>Total Endpoint Security Policies:</strong> $(Encode-Html $(@($endpointSecurityDisplayRows).Count))<br/>
    <strong>Total Windows Update Policies:</strong> $(Encode-Html $totalUpdateProfiles)<br/>
    <strong>Total Compliance Policies:</strong> $(Encode-Html $compliancePoliciesCount)<br/>
    <strong>Total Applications:</strong> $(Encode-Html $appsCount)
  </p>
  <p class='muted small'>Note: Some datasets use Microsoft Graph <em>/beta</em> endpoints.</p>

  <div class='score-grid'>
    <div class='score-card'>
      <div class='score-title'>Overall Posture Score</div>
      <div class='score-value'>$(Encode-Html $overallScore)%</div>
      <div class='score-label $($overallRisk.Css)'>$(Encode-Html $overallRisk.Label)</div>
    </div>
    <div class='score-card'>
      <div class='score-title'>Compliance Score</div>
      <div class='score-value'>$(Encode-Html $complianceScore)%</div>
      <div class='score-label $($complianceRisk.Css)'>$(Encode-Html $complianceRisk.Label)</div>
    </div>
    <div class='score-card'>
      <div class='score-title'>Encryption Score</div>
      <div class='score-value'>$(Encode-Html $encryptionScore)%</div>
      <div class='score-label $($encryptionRisk.Css)'>$(Encode-Html $encryptionRisk.Label)</div>
    </div>
    <div class='score-card'>
      <div class='score-title'>Config Hygiene Score</div>
      <div class='score-value'>$(Encode-Html $configHygieneScore)%</div>
      <div class='score-label $($configRisk.Css)'>$(Encode-Html $configRisk.Label)</div>
    </div>
    <div class='score-card'>
      <div class='score-title'>Endpoint Security Coverage</div>
      <div class='score-value'>$(Encode-Html $endpointSecurityScore)%</div>
      <div class='score-label $($endpointRisk.Css)'>$(Encode-Html $endpointRisk.Label)</div>
    </div>
  </div>

  <h3>Top Risks</h3>
  <ul>
"@

    foreach ($risk in $topRisks) {
        $html += "<li>$(Encode-Html $risk)</li>"
    }

    $html += @"
  </ul>

  <h3>Dynamic Recommendations</h3>
  <ul>
"@

    foreach ($rec in $dynamicRecommendations) {
        $html += "<li>$(Encode-Html $rec)</li>"
    }

    $html += @"
  </ul>
</div>
"@

    if ($collectionErrors.Count -gt 0) {
        $html += "<div class='warnings'><h2>Data Collection Warnings</h2><ul>"
        foreach ($e in $collectionErrors) { $html += "<li>$(Encode-Html $e)</li>" }
        $html += "</ul></div>"
    }

    # -----------------------------
    # Assessment Coverage by Platform
    # -----------------------------
    $html += "<h2>Assessment Coverage by Platform</h2>"
    $html += "<p class='section-note'>This section shows the number of discovered items per platform across the major workloads in the assessment. Platforms with zero counts are still shown so the reader can quickly identify gaps in management coverage.</p>"
    $html += "<table><tr><th>Platform</th><th>Devices</th><th>Configuration Policies</th><th>Compliance Policies</th><th>Endpoint Security Policies</th><th>Applications</th></tr>"

    foreach ($platform in $configPlatformCounts.Keys) {
        $html += "<tr>
          <td>$(Encode-Html $platform)</td>
          <td>$(Encode-Html $devicePlatformCounts[$platform])</td>
          <td>$(Encode-Html $configPlatformCounts[$platform])</td>
          <td>$(Encode-Html $compliancePlatformCounts[$platform])</td>
          <td>$(Encode-Html $endpointPlatformCounts[$platform])</td>
          <td>$(Encode-Html $appPlatformCounts[$platform])</td>
        </tr>"
    }
    $html += "</table>"

    # -----------------------------
    # Device Inventory
    # -----------------------------
    $html += "<h2>Device Inventory</h2>"
    $html += "<p class='section-note'>This section provides a list of devices currently managed by Microsoft Intune. It highlights each device’s platform, OS version, compliance status, whether disk encryption is enabled, and its last sync time.</p>"

    if ($devices -and $devices.Count -gt 0) {
        $html += "<table><tr><th>Name</th><th>Platform</th><th>OS Version</th><th>Compliance</th><th>Encryption</th><th>Last Sync</th></tr>"
        foreach ($device in $devices) {
            $enc = "Unknown"
            if ($null -ne $device.IsEncrypted) { $enc = if ($device.IsEncrypted) { "Encrypted" } else { "Not Encrypted" } }

            $lastSync = "Unknown"
            if ($device.LastSyncDateTime) {
                try { $lastSync = ([datetime]$device.LastSyncDateTime).ToString("yyyy-MM-dd HH:mm:ss") } catch { $lastSync = [string]$device.LastSyncDateTime }
            }

            $html += "<tr>
              <td>$(Encode-Html $device.DeviceName)</td>
              <td>$(Encode-Html $device.OperatingSystem)</td>
              <td>$(Encode-Html $device.OsVersion)</td>
              <td>$(Encode-Html $device.ComplianceState)</td>
              <td>$(Encode-Html $enc)</td>
              <td>$(Encode-Html $lastSync)</td>
            </tr>"
        }
        $html += "</table>"
    } else {
        $html += "<p class='muted'>No device data returned.</p>"
    }

    # -----------------------------
    # Device & Policy Hygiene Review
    # -----------------------------
    $html += "<h2>Device & Policy Hygiene Review</h2>"
    $html += "<p class='section-note'>This section highlights stale managed devices and configuration policies that currently have no assignments. These items are often strong indicators of cleanup opportunities, reporting inaccuracies, or incomplete deployment scope.</p>"

    $html += "<table><tr><th>Finding</th><th>Count</th><th>Details</th></tr>"

    $staleDeviceDetails = "None identified"
    if ($staleDevices -and $staleDevices.Count -gt 0) {
        $staleDeviceDetails = (($staleDevices | Select-Object -First 10 | ForEach-Object {
            $syncText = if ($_.LastSyncDateTime) { ([datetime]$_.LastSyncDateTime).ToString("yyyy-MM-dd HH:mm:ss") } else { "Unknown" }
            "$(Encode-Html $_.DeviceName) (Platform: $(Encode-Html $_.OperatingSystem), Last Sync: $(Encode-Html $syncText))"
        }) -join "<br/>")
        if ($staleDevices.Count -gt 10) {
            $staleDeviceDetails += "<br/><span class='muted'>Additional stale devices not shown.</span>"
        }
    }

    $unassignedPolicyDetails = "None identified"
    if ($unassignedConfigPolicies -and $unassignedConfigPolicies.Count -gt 0) {
        $unassignedPolicyDetails = (($unassignedConfigPolicies | Select-Object -First 10 | ForEach-Object {
            "$(Encode-Html $_.Name) (Platform: $(Encode-Html $_.Platform), Source: $(Encode-Html $_.PolicyStore))"
        }) -join "<br/>")
        if ($unassignedConfigPolicies.Count -gt 10) {
            $unassignedPolicyDetails += "<br/><span class='muted'>Additional unassigned policies not shown.</span>"
        }
    }

    $html += @"
<tr>
  <td>Stale Devices (>$staleDeviceThresholdDays days since last sync)</td>
  <td>$(Encode-Html $staleDeviceCount)</td>
  <td class='small'>$staleDeviceDetails</td>
</tr>
<tr>
  <td>Unassigned Configuration Policies</td>
  <td>$(Encode-Html $unassignedConfigPolicyCount)</td>
  <td class='small'>$unassignedPolicyDetails</td>
</tr>
</table>
"@

    $progressBar.Value = 86

    # -----------------------------
    # Modern Configuration Policies
    # -----------------------------
    $html += "<h2>Modern Configuration Policies</h2>"
    $html += "<p class='section-note'>This section includes general modern configuration policies such as Settings Catalog and template-based profiles. Endpoint Security-backed policies and update-related policies are shown in their respective sections later in the report.</p>"

    if ($modernConfigPoliciesForSection -and $modernConfigPoliciesForSection.Count -gt 0) {
        $html += "<table><tr><th>Name</th><th>Platform</th><th>Policy Store</th><th>Description</th><th>Last Modified</th><th>Assignments</th></tr>"

        foreach ($policy in ($modernConfigPoliciesForSection | Sort-Object Name)) {
            $assignmentsHtml = "<span class='muted'>None assigned</span>"
            if ($configPolicyAssignmentMap.ContainsKey($policy.Id)) {
                $assignmentsHtml = $configPolicyAssignmentMap[$policy.Id].Html
            }

            $html += "<tr>
              <td>$(Encode-Html $policy.Name)</td>
              <td>$(Encode-Html $policy.Platform)</td>
              <td>$(Encode-Html $policy.PolicyStore)</td>
              <td>$(Encode-Html $policy.Description)</td>
              <td>$(Encode-Html $policy.LastModified)</td>
              <td><div class='small muted'>$assignmentsHtml</div></td>
            </tr>"
        }

        $html += "</table>"
    } else {
        $html += "<p class='muted'>No modern configuration policy data returned.</p>"
    }

    # -----------------------------
    # Classic Device Configurations
    # -----------------------------
    $html += "<h2>Classic Device Configurations</h2>"
    $html += "<p class='section-note'>This section includes classic device configuration profiles collected from the legacy deviceConfigurations workload. Windows Update rings are excluded here and are shown separately in the Windows Update section to avoid duplication.</p>"

    if ($classicConfigPoliciesForSection -and $classicConfigPoliciesForSection.Count -gt 0) {
        $html += "<table><tr><th>Name</th><th>Platform</th><th>Policy Store</th><th>Description</th><th>Last Modified</th><th>Assignments</th></tr>"

        foreach ($policy in ($classicConfigPoliciesForSection | Sort-Object Name)) {
            $assignmentsHtml = "<span class='muted'>None assigned</span>"
            if ($configPolicyAssignmentMap.ContainsKey($policy.Id)) {
                $assignmentsHtml = $configPolicyAssignmentMap[$policy.Id].Html
            }

            $html += "<tr>
              <td>$(Encode-Html $policy.Name)</td>
              <td>$(Encode-Html $policy.Platform)</td>
              <td>$(Encode-Html $policy.PolicyStore)</td>
              <td>$(Encode-Html $policy.Description)</td>
              <td>$(Encode-Html $policy.LastModified)</td>
              <td><div class='small muted'>$assignmentsHtml</div></td>
            </tr>"
        }

        $html += "</table>"
    } else {
        $html += "<p class='muted'>No classic device configuration data returned.</p>"
    }

    # -----------------------------
    # Endpoint Security Policies
    # -----------------------------
    $html += "<h2>Endpoint Security Policies</h2>"
    $html += "<p class='section-note'>This section summarizes policies configured through Intune’s Endpoint Security workload, including legacy intents and modern configuration-policy-backed endpoint security controls.</p>"

    $html += "<table><tr><th>Category</th><th>Count</th></tr>"
    foreach ($kv in $endpointSecurityCategoryCounts.GetEnumerator()) {
        $html += "<tr><td>$(Encode-Html $kv.Key)</td><td>$(Encode-Html $kv.Value)</td></tr>"
    }
    $html += "</table>"

    if ($endpointSecurityDisplayRows -and @($endpointSecurityDisplayRows).Count -gt 0) {
        $assignedEndpointSecurityCount = (@($endpointSecurityDisplayRows | Where-Object { $_.Assigned -eq $true })).Count
        $unassignedEndpointSecurityCount = (@($endpointSecurityDisplayRows | Where-Object { $_.Assigned -ne $true })).Count

        $html += "<p class='section-note'><strong>Detected endpoint security-related policies:</strong> $(Encode-Html $(@($endpointSecurityDisplayRows).Count))<br/>"
        $html += "<strong>Assigned:</strong> $(Encode-Html $assignedEndpointSecurityCount) &nbsp;&nbsp; <strong>Not Assigned:</strong> $(Encode-Html $unassignedEndpointSecurityCount)</p>"

        $html += "<table><tr><th>Name</th><th>Platform</th><th>Category</th><th>Assigned</th><th>Assignments</th><th>Source</th><th>Last Modified</th></tr>"

        foreach ($row in $endpointSecurityDisplayRows) {
            $assignedStatus = "Not Assigned"
            $assignedCss = "risk-bad"
            if ($row.Assigned) {
                $assignedStatus = "Assigned"
                $assignedCss = "risk-good"
            }

            $assignedBadge = "<span class='score-label $assignedCss'>$assignedStatus</span>"

            $html += "<tr>
              <td>$(Encode-Html $row.Name)</td>
              <td>$(Encode-Html $row.Platform)</td>
              <td>$(Encode-Html $row.Category)</td>
              <td>$assignedBadge</td>
              <td><div class='small'>$($row.Assignments)</div></td>
              <td>$(Encode-Html $row.Source)</td>
              <td>$(Encode-Html $row.LastModified)</td>
            </tr>"
        }

        $html += "</table>"
    }
    else {
        $html += "<p class='muted'>No endpoint security policies detected.</p>"
    }

    # -----------------------------
    # Windows Update Policies
    # -----------------------------
    $html += "<h2>Windows Update Policies</h2>"
    $html += "<p class='section-note'>This section summarizes Windows Update governance including classic update rings, feature update profiles, quality update profiles, and update-related modern configuration policies.</p>"

    $html += "<table><tr><th>Policy Type</th><th>Platform</th><th>Count</th></tr>"
    $html += "<tr><td>Classic Update Rings</td><td>Windows</td><td>$(Encode-Html $(@($updateRingPolicies).Count))</td></tr>"
    $html += "<tr><td>Feature Update Profiles</td><td>Windows</td><td>$(Encode-Html $(@($featureUpdateProfiles).Count))</td></tr>"
    $html += "<tr><td>Quality Update Profiles</td><td>Windows</td><td>$(Encode-Html $(@($qualityUpdateProfiles).Count))</td></tr>"
    $html += "<tr><td>Other Windows Update Policies</td><td>Windows</td><td>$(Encode-Html $(@($otherWindowsUpdatePolicies).Count))</td></tr>"
    $html += "<tr><td><strong>Total Windows Update Policies</strong></td><td><strong>Windows</strong></td><td><strong>$(Encode-Html $totalUpdateProfiles)</strong></td></tr>"
    $html += "</table>"

    if ($updateRingPolicies -and @($updateRingPolicies).Count -gt 0) {
        $html += "<h3>Classic Update Rings</h3>"
        $html += "<table><tr><th>Name</th><th>Platform</th><th>Description</th><th>Assignments</th><th>Last Modified</th></tr>"
        foreach ($ring in ($updateRingPolicies | Sort-Object Name)) {
            $assignmentsHtml = "<span class='muted'>None assigned</span>"
            if ($updateRingAssignmentMap.ContainsKey($ring.Id)) {
                $assignmentsHtml = $updateRingAssignmentMap[$ring.Id].Html
            }

            $html += "<tr>
              <td>$(Encode-Html $ring.Name)</td>
              <td>Windows</td>
              <td>$(Encode-Html $ring.Description)</td>
              <td><div class='small'>$assignmentsHtml</div></td>
              <td>$(Encode-Html $ring.LastModified)</td>
            </tr>"
        }
        $html += "</table>"
    }

    if ($featureUpdateProfiles -and @($featureUpdateProfiles).Count -gt 0) {
        $html += "<h3>Feature Update Profiles</h3>"
        $html += "<table><tr><th>Name</th><th>Platform</th><th>Description</th><th>Last Modified</th></tr>"
        foreach ($policy in ($featureUpdateProfiles | Sort-Object displayName)) {
            $html += "<tr>
              <td>$(Encode-Html $policy.displayName)</td>
              <td>Windows</td>
              <td>$(Encode-Html $policy.description)</td>
              <td>$(Encode-Html $policy.lastModifiedDateTime)</td>
            </tr>"
        }
        $html += "</table>"
    }

    if ($qualityUpdateProfiles -and @($qualityUpdateProfiles).Count -gt 0) {
        $html += "<h3>Quality Update Profiles</h3>"
        $html += "<table><tr><th>Name</th><th>Platform</th><th>Description</th><th>Last Modified</th></tr>"
        foreach ($policy in ($qualityUpdateProfiles | Sort-Object displayName)) {
            $html += "<tr>
              <td>$(Encode-Html $policy.displayName)</td>
              <td>Windows</td>
              <td>$(Encode-Html $policy.description)</td>
              <td>$(Encode-Html $policy.lastModifiedDateTime)</td>
            </tr>"
        }
        $html += "</table>"
    }

    if ($otherWindowsUpdatePolicies -and @($otherWindowsUpdatePolicies).Count -gt 0) {
        $html += "<h3>Other Windows Update Policies</h3>"
        $html += "<table><tr><th>Name</th><th>Platform</th><th>Assignments</th><th>Last Modified</th></tr>"
        foreach ($policy in ($otherWindowsUpdatePolicies | Sort-Object Name)) {
            $assignmentsHtml = "<span class='muted'>None assigned</span>"
            if ($windowsUpdateAssignmentMap.ContainsKey($policy.Id)) {
                $assignmentsHtml = $windowsUpdateAssignmentMap[$policy.Id].Html
            }
            $html += "<tr>
              <td>$(Encode-Html $policy.Name)</td>
              <td>Windows</td>
              <td><div class='small'>$assignmentsHtml</div></td>
              <td>$(Encode-Html $policy.LastModified)</td>
            </tr>"
        }
        $html += "</table>"
    }

    if ($totalUpdateProfiles -eq 0) {
        $html += "<p class='muted'>No Windows Update policies detected.</p>"
    }

    $progressBar.Value = 90

    # -----------------------------
    # Compliance Policies
    # -----------------------------
    $html += "<h2>Compliance Policies</h2>"
    $html += "<p class='section-note'>Compliance policies set minimum security and health requirements that devices must meet. Devices that fail these checks may be marked noncompliant and restricted from accessing corporate resources.</p>"

    if ($compliancePolicies -and $compliancePolicies.Count -gt 0) {
        $html += "<table><tr><th>Name</th><th>Platform</th><th>Description</th><th>Last Modified</th></tr>"

        foreach ($cp in ($compliancePolicies | Sort-Object displayName)) {
            $name = $cp.displayName
            if ([string]::IsNullOrWhiteSpace($name)) {
                $name = "Unnamed Compliance Policy"
            }

            $description = if ($cp.description) { $cp.description } else { "No description" }
            $lastModified = if ($cp.lastModifiedDateTime) { $cp.lastModifiedDateTime } else { "Unknown" }
            $platform = Get-CompliancePolicyPlatformLabel -Policy $cp

            $html += "<tr>
              <td>$(Encode-Html $name)</td>
              <td>$(Encode-Html $platform)</td>
              <td>$(Encode-Html $description)</td>
              <td>$(Encode-Html $lastModified)</td>
            </tr>"
        }

        $html += "</table>"
    } else {
        $html += "<p class='muted'>No compliance policy data returned.</p>"
    }

    # -----------------------------
    # Applications
    # -----------------------------
    $html += "<h2>Applications</h2>"
    $html += "<p class='section-note'>This section summarizes applications deployed through Intune, including how they are sourced and maintained. It helps identify standardization opportunities and gaps in application management.</p>"

    if ($apps -and $apps.Count -gt 0) {
        $byKey = @{}
        foreach ($app in $apps) {
            $cls = Get-IntuneAppClassification -App $app
            $key = "$($cls.Platform) — $($cls.Type)"
            if ($cls.Platform -eq 'Windows' -and (-not [string]::IsNullOrWhiteSpace($cls.Source))) {
                $key = "$($cls.Platform) — $($cls.Type) — $($cls.Source)"
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

    # -----------------------------
    # Potential Policy Conflict / Overlap Analysis
    # -----------------------------
    $html += "<h2>Potential Policy Conflict / Overlap Analysis</h2>"
    $html += "<p class='section-note'>This section highlights potential overlap scenarios that may deserve engineering review. These findings are heuristic and are intended to identify areas where multiple policies with similar purpose may target the same population.</p>"

    if ($conflictFindings -and $conflictFindings.Count -gt 0) {
        $html += "<table><tr><th>Area</th><th>Type</th><th>Scope</th><th>Details</th></tr>"
        foreach ($finding in $conflictFindings) {
            $html += "<tr>
              <td>$(Encode-Html $finding.Area)</td>
              <td>$(Encode-Html $finding.Type)</td>
              <td>$(Encode-Html $finding.Scope)</td>
              <td>$(Encode-Html $finding.Details)</td>
            </tr>"
        }
        $html += "</table>"
    } else {
        $html += "<p class='muted'>No obvious potential policy overlap scenarios were identified by the report logic.</p>"
    }

    # -----------------------------
    # Actionable Recommendations
    # -----------------------------
    $html += @"
<div class='recommendations'>
  <h2>Actionable Recommendations</h2>
  <ul>
    <li>Review and remediate <strong>$(Encode-Html $nonCompliantCount) devices</strong> that are not meeting compliance requirements.</li>
    <li>Investigate any configuration policies where assignment details could not be retrieved to ensure proper targeting.</li>
    <li>Validate policy scope to ensure all required user/device populations are covered.</li>
    <li>Review both modern configuration policies and classic device configurations for redundancy and migration opportunities.</li>
    <li>Review application deployment strategy and confirm required apps are targeted appropriately.</li>
    <li>Review stale device objects and remove or troubleshoot records that have not checked in recently.</li>
    <li>Review endpoint security coverage and Windows update governance to ensure protections and patching are applied consistently.</li>
    <li>Review the Potential Policy Conflict / Overlap Analysis section for areas where multiple policies may target the same device populations.</li>
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
``