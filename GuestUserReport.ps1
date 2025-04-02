<#
=============================================================================================
Name:           Export Office 365 Guest User and Membership Report using MS Graph PowerShell
=============================================================================================
#>

# Accept input parameters
Param (
    [Parameter(Mandatory = $false)]
    [int]$StaleGuests,
    [int]$RecentlyCreatedGuests,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
)

# Check if Microsoft Graph module is installed
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Host "Microsoft Graph module is not installed. Installing now..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force
    Write-Host "Microsoft Graph module installed successfully." -ForegroundColor Green
}

# Connect to Microsoft Graph
try {
    if ($TenantId -and $ClientId -and $CertificateThumbprint) {
        Connect-MgGraph -TenantId $TenantId -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -ErrorAction Stop | Out-Null
    } else {
        Connect-MgGraph -Scopes "Directory.Read.All" -ErrorAction Stop | Out-Null
    }
    Write-Host "Successfully connected to Microsoft Graph." -ForegroundColor Green
} catch {
    Write-Host "Error connecting to Microsoft Graph: $_" -ForegroundColor Red
    Exit
}

# Define output file
$ExportCSV = "./GuestUserReport_$((Get-Date -Format 'yyyy-MM-dd_HH-mm-ss')).csv"
Write-Host "Exporting report..." -ForegroundColor Cyan

# Get guest users
try {
    $GuestUsers = Get-MgUser -All -Filter "UserType eq 'Guest'" -ExpandProperty MemberOf
    if (-not $GuestUsers) {
        Write-Host "No guest users found." -ForegroundColor Red
        Exit
    }
} catch {
    Write-Host "Error fetching guest users: $_" -ForegroundColor Red
    Exit
}

# Process and export data
$Results = @()
$GuestCount = 0
$FilteredCount = 0
foreach ($User in $GuestUsers) {
    $GuestCount++
    Write-Progress -Activity "Processing users..." -Status "Processed: $GuestCount" -PercentComplete (($GuestCount / $GuestUsers.Count) * 100)
    
    # Handle null CreatedDateTime
    $AccountAge = if ($User.CreatedDateTime) { (New-TimeSpan -Start $User.CreatedDateTime).Days } else { 0 }
    
    # Apply filters
    if ($StaleGuests -and ($AccountAge -lt $StaleGuests)) { continue }
    if ($RecentlyCreatedGuests -and ($AccountAge -gt $RecentlyCreatedGuests)) { continue }
    
    # Handle null values
    $Company = if ($User.CompanyName) { $User.CompanyName } else { "-" }
    $GroupMembership = ($User.MemberOf | Select-Object -ExpandProperty DisplayName) -join ','
    if (-not $GroupMembership) { $GroupMembership = '-' }
    
    # Store result
    $Results += [PSCustomObject]@{
        DisplayName        = $User.DisplayName
        UserPrincipalName  = $User.UserPrincipalName
        Company           = $Company
        EmailAddress      = $User.Mail
        CreationTime      = $User.CreatedDateTime
        AccountAge_Days   = $AccountAge
        CreationType      = $User.CreationType
        InvitationAccepted = $User.ExternalUserState
        GroupMembership   = $GroupMembership
    }
    $FilteredCount++
}

# Export to CSV
$Results | Export-Csv -Path $ExportCSV -NoTypeInformation -Encoding UTF8
Write-Host "Report exported to: $ExportCSV" -ForegroundColor Green
Write-Host "Total guest users processed: $FilteredCount" -ForegroundColor Cyan

# Prompt to open file
if ($FilteredCount -gt 0 -and (Test-Path $ExportCSV)) {
    $Response = Read-Host "Do you want to open the report? (Y/N)"
    if ($Response -match "^[Yy]$") {
        Invoke-Item $ExportCSV
    }
} else {
    Write-Host "No guest users matched the filter criteria." -ForegroundColor Yellow
}

# Disconnect from Graph
Disconnect-MgGraph | Out-Null
Write-Host "Disconnected from Microsoft Graph." -ForegroundColor Green
