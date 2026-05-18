[CmdletBinding()]
param(
    [string]$ConfigPath = ".\config.json",
    [switch]$OpenReport
)

$ErrorActionPreference = "Stop"

# ==================================================
# GLOBALS
# ==================================================

$Script:Findings = New-Object System.Collections.Generic.List[object]
$Script:ReportRoot = Join-Path $PSScriptRoot "Reports"
$Script:Timestamp = Get-Date -Format "yyyyMMdd-HHmmss"

# ==================================================
# BANNER
# ==================================================

function Show-Banner {

    Clear-Host

    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "    M365 LICENSE OPTIMIZATION DASHBOARD v1.0.0" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
}

# ==================================================
# STATUS
# ==================================================

function Write-Status {

    param(
        [string]$Message,
        [string]$Level = "INFO"
    )

    $color = switch ($Level) {
        "OK" { "Green" }
        "WARN" { "Yellow" }
        "ERROR" { "Red" }
        default { "White" }
    }

    Write-Host "[$Level] $Message" -ForegroundColor $color
}

# ==================================================
# CONFIG
# ==================================================

function Load-Config {

    param([string]$Path)

    if (-not (Test-Path $Path)) {
        throw "Config file not found: $Path"
    }

    return Get-Content $Path -Raw | ConvertFrom-Json
}

# ==================================================
# REPORT FOLDER
# ==================================================

function Ensure-ReportFolder {

    if (-not (Test-Path $Script:ReportRoot)) {

        New-Item `
            -ItemType Directory `
            -Path $Script:ReportRoot | Out-Null
    }
}

# ==================================================
# MODULES
# ==================================================

function Import-RequiredModules {

    Write-Status "Loading required modules..."

    Import-Module Microsoft.Graph.Users
    Import-Module Microsoft.Graph.Reports
    Import-Module Microsoft.Graph.Identity.DirectoryManagement
}

# ==================================================
# CONNECTION
# ==================================================

function Connect-M365 {

    Write-Status "Connecting to Microsoft Graph..."

    Connect-MgGraph `
        -Scopes `
        "User.Read.All",
        "Directory.Read.All",
        "Reports.Read.All",
        "Organization.Read.All" `
        -NoWelcome
}

# ==================================================
# FINDINGS
# ==================================================

function Add-Finding {

    param(
        [string]$Category,
        [string]$Severity,
        [string]$User,
        [string]$UPN,
        [string]$Issue,
        [decimal]$MonthlyWaste,
        [decimal]$YearlyWaste,
        [string]$Recommendation
    )

    $Script:Findings.Add([pscustomobject]@{
        Category       = $Category
        Severity       = $Severity
        User           = $User
        UPN            = $UPN
        Issue          = $Issue
        MonthlyWaste   = $MonthlyWaste
        YearlyWaste    = $YearlyWaste
        Recommendation = $Recommendation
    })
}

# ==================================================
# LICENSE PRICING
# ==================================================

function Get-LicensePricing {

    param($Config)

    return $Config.LicensePricing
}

# ==================================================
# DATA COLLECTION
# ==================================================

function Get-M365Users {

    Write-Status "Collecting users..."

    return Get-MgUser -All `
        -Property `
        DisplayName,
        UserPrincipalName,
        AccountEnabled,
        AssignedLicenses,
        SignInActivity,
        Mail,
        CreatedDateTime
}

function Get-LicenseMap {

    Write-Status "Collecting tenant SKUs..."

    $skus = Get-MgSubscribedSku

    $map = @{}

    foreach ($sku in $skus) {

        $map[$sku.SkuId.Guid] = $sku.SkuPartNumber
    }

    return $map
}

# ==================================================
# ANALYSIS
# ==================================================

function Analyze-DisabledLicensedUsers {

    param(
        $Users,
        $SkuMap,
        $Pricing
    )

    foreach ($user in $Users) {

        if ($user.AccountEnabled -eq $false -and $user.AssignedLicenses.Count -gt 0) {

            foreach ($license in $user.AssignedLicenses) {

                $sku = $SkuMap[$license.SkuId.Guid]

                $monthly = 0

                if ($Pricing.$sku) {
                    $monthly = [decimal]$Pricing.$sku
                }

                Add-Finding `
                    -Category "Disabled User" `
                    -Severity "High" `
                    -User $user.DisplayName `
                    -UPN $user.UserPrincipalName `
                    -Issue "Disabled user still consuming license: $sku" `
                    -MonthlyWaste $monthly `
                    -YearlyWaste ($monthly * 12) `
                    -Recommendation "Remove unused licenses from disabled accounts."
            }
        }
    }
}

function Analyze-InactiveLicensedUsers {

    param(
        $Users,
        $SkuMap,
        $Pricing,
        [int]$InactiveDays
    )

    $cutoff = (Get-Date).AddDays(-$InactiveDays)

    foreach ($user in $Users) {

        if ($user.AssignedLicenses.Count -eq 0) {
            continue
        }

        $lastSignIn = $null

        if ($user.SignInActivity.LastSignInDateTime) {
            $lastSignIn = [datetime]$user.SignInActivity.LastSignInDateTime
        }

        if (-not $lastSignIn -or $lastSignIn -lt $cutoff) {

            foreach ($license in $user.AssignedLicenses) {

                $sku = $SkuMap[$license.SkuId.Guid]

                $monthly = 0

                if ($Pricing.$sku) {
                    $monthly = [decimal]$Pricing.$sku
                }

                Add-Finding `
                    -Category "Inactive User" `
                    -Severity "Medium" `
                    -User $user.DisplayName `
                    -UPN $user.UserPrincipalName `
                    -Issue "Licensed user inactive for over $InactiveDays days." `
                    -MonthlyWaste $monthly `
                    -YearlyWaste ($monthly * 12) `
                    -Recommendation "Review user activity and reclaim unused licensing."
            }
        }
    }
}

function Analyze-UnlicensedUsers {

    param($Users)

    foreach ($user in $Users) {

        if ($user.AccountEnabled -eq $true -and $user.AssignedLicenses.Count -eq 0) {

            Add-Finding `
                -Category "Unlicensed User" `
                -Severity "Low" `
                -User $user.DisplayName `
                -UPN $user.UserPrincipalName `
                -Issue "Active user has no assigned license." `
                -MonthlyWaste 0 `
                -YearlyWaste 0 `
                -Recommendation "Verify if this account should be licensed."
        }
    }
}

function Analyze-DuplicateLicensing {

    param(
        $Users,
        $SkuMap,
        $Pricing
    )

    foreach ($user in $Users) {

        if ($user.AssignedLicenses.Count -gt 1) {

            $assigned = @()

            foreach ($license in $user.AssignedLicenses) {

                $assigned += $SkuMap[$license.SkuId.Guid]
            }

            if ($assigned -contains "VISIOCLIENT" -and $assigned -contains "PROJECTPROFESSIONAL") {

                Add-Finding `
                    -Category "Duplicate Licensing" `
                    -Severity "Medium" `
                    -User $user.DisplayName `
                    -UPN $user.UserPrincipalName `
                    -Issue "User has both Visio and Project licenses." `
                    -MonthlyWaste 30 `
                    -YearlyWaste 360 `
                    -Recommendation "Verify whether both premium licenses are required."
            }
        }
    }
}

# ==================================================
# CSV EXPORT
# ==================================================

function Export-CsvReport {

    $path = Join-Path `
        $Script:ReportRoot `
        "LicenseOptimization-$Script:Timestamp.csv"

    $Script:Findings |
        Export-Csv `
        -Path $path `
        -NoTypeInformation

    return $path
}

# ==================================================
# HTML REPORT
# ==================================================

function Export-HtmlDashboard {

    $totalFindings = $Script:Findings.Count

    $monthlyWaste = (
        $Script:Findings |
        Measure-Object MonthlyWaste -Sum
    ).Sum

    $yearlyWaste = (
        $Script:Findings |
        Measure-Object YearlyWaste -Sum
    ).Sum

    $rows = foreach ($finding in $Script:Findings) {

        $severityColor = switch ($finding.Severity) {
            "High" { "#dc2626" }
            "Medium" { "#d97706" }
            "Low" { "#0284c7" }
            default { "#64748b" }
        }

@"
<tr>
<td><span style='color:white;background:$severityColor;padding:4px 10px;border-radius:999px'>$($finding.Severity)</span></td>
<td>$($finding.Category)</td>
<td>$($finding.User)</td>
<td>$($finding.UPN)</td>
<td>$($finding.Issue)</td>
<td>`$$($finding.MonthlyWaste)</td>
<td>`$$($finding.YearlyWaste)</td>
<td>$($finding.Recommendation)</td>
</tr>
"@
    }

$html = @"
<!DOCTYPE html>
<html>
<head>
<title>M365 License Optimization Dashboard</title>

<style>

body{
background:#0f172a;
color:#e2e8f0;
font-family:Segoe UI;
padding:30px;
}

h1{
color:#38bdf8;
}

.card-container{
display:grid;
grid-template-columns:repeat(3,1fr);
gap:20px;
margin-bottom:30px;
}

.card{
background:#1e293b;
padding:20px;
border-radius:16px;
}

.card-title{
color:#94a3b8;
font-size:14px;
margin-bottom:10px;
}

.card-value{
font-size:32px;
font-weight:bold;
}

table{
width:100%;
border-collapse:collapse;
background:#1e293b;
}

th{
background:#020617;
padding:12px;
text-align:left;
}

td{
padding:12px;
border-bottom:1px solid #334155;
}

tr:hover{
background:#273449;
}

</style>

</head>

<body>

<h1>M365 License Optimization Dashboard</h1>

<div class='card-container'>

<div class='card'>
<div class='card-title'>Total Findings</div>
<div class='card-value'>$totalFindings</div>
</div>

<div class='card'>
<div class='card-title'>Monthly Waste</div>
<div class='card-value'>`$$monthlyWaste</div>
</div>

<div class='card'>
<div class='card-title'>Yearly Waste</div>
<div class='card-value'>`$$yearlyWaste</div>
</div>

</div>

<table>

<thead>
<tr>
<th>Severity</th>
<th>Category</th>
<th>User</th>
<th>UPN</th>
<th>Issue</th>
<th>Monthly</th>
<th>Yearly</th>
<th>Recommendation</th>
</tr>
</thead>

<tbody>

$($rows -join "`n")

</tbody>

</table>

</body>
</html>
"@

    $path = Join-Path `
        $Script:ReportRoot `
        "Dashboard-$Script:Timestamp.html"

    Set-Content `
        -Path $path `
        -Value $html `
        -Encoding UTF8

    return $path
}

# ==================================================
# MAIN
# ==================================================

function Start-LicenseOptimization {

    Show-Banner

    Ensure-ReportFolder

    Write-Status "Loading configuration..."

    $config = Load-Config -Path $ConfigPath

    Import-RequiredModules

    Connect-M365

    $pricing = Get-LicensePricing -Config $config

    $skuMap = Get-LicenseMap

    $users = Get-M365Users

    Write-Status "Analyzing licensing..."

    Analyze-DisabledLicensedUsers `
        -Users $users `
        -SkuMap $skuMap `
        -Pricing $pricing

    Analyze-InactiveLicensedUsers `
        -Users $users `
        -SkuMap $skuMap `
        -Pricing $pricing `
        -InactiveDays $config.InactiveDays

    Analyze-UnlicensedUsers `
        -Users $users

    Analyze-DuplicateLicensing `
        -Users $users `
        -SkuMap $skuMap `
        -Pricing $pricing

    Write-Status "Generating reports..."

    $csv = Export-CsvReport
    $html = Export-HtmlDashboard

    Write-Host ""
    Write-Status "Analysis complete." "OK"
    Write-Host "CSV Report:  $csv"
    Write-Host "HTML Report: $html"

    if ($OpenReport) {
        Start-Process $html
    }
}

Start-LicenseOptimization
