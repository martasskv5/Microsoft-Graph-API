# ============================================================
# CONFIGURATION - Change these variables for your tenant
# ============================================================
$TenantName     = "0z6sw"                          # e.g. contoso
$AdminUrl       = "https://$TenantName-admin.sharepoint.com"
$SchoolTitle    = "SPSITKNM"
$SchoolUrl      = "https://$TenantName.sharepoint.com/sites/$($SchoolTitle.ToLower())"
$YearTitle      = "2026-2027"
$YearUrl        = "https://$TenantName.sharepoint.com/sites/$($SchoolTitle.ToLower())-$($YearTitle.Replace('.','-'))"
$Subjects       = @("ANJ", "APM", "Appslab", "CEQ", "CIT", "CME", "CSX", "EKO", "ETS", "MAT", "SJL", "SXT", "USP")
$OwnerEmail     = "admin@$TenantName.onmicrosoft.com"   # Site collection admin

# ============================================================
# 1. CONNECT TO TENANT ADMIN (interactive login)
# ============================================================
Connect-PnPOnline -Url $AdminUrl -ClientId e33ff2e6-0d34-4565-9324-e2d80483c7e9

# ============================================================
# 2. CREATE TOP-LEVEL SCHOOL HUB (Communication Site)
# ============================================================
Write-Host "`n🏫 Creating School Hub..." -ForegroundColor Cyan
$SchoolSite = New-PnPSite `
    -Type CommunicationSite `
    -Title $SchoolTitle `
    -Url $SchoolUrl `
    -Owner $OwnerEmail `
    -SiteDesign Topic

Write-Host "   Site created at: $SchoolSite" -ForegroundColor Green

# Register it as a Hub Site
Register-PnPHubSite -Site $SchoolSite
Write-Host "   Registered as Hub Site" -ForegroundColor Green

# Optional: configure hub display name & restrict who can associate
Set-PnPHubSite -Identity $SchoolSite -Description "Main portal for all school years"
Grant-PnPHubSiteRights -Identity $SchoolSite -Principals $OwnerEmail

# ============================================================
# 3. CREATE SCHOOL YEAR HUB (Nested Hub)
# ============================================================
Write-Host "`n📅 Creating Year Hub: $YearTitle..." -ForegroundColor Cyan
$YearSite = New-PnPSite `
    -Type CommunicationSite `
    -Title $YearTitle `
    -Url $YearUrl `
    -Owner $OwnerEmail `
    -SiteDesign Topic

Set-PnPSite -Identity $YearUrl -NoScriptSite $false # Allow custom scripts for the year site

Write-Host "   Site created at: $YearSite" -ForegroundColor Green

# Register the year site as its own Hub Site
Register-PnPHubSite -Site $YearSite
Write-Host "   Registered as Hub Site" -ForegroundColor Green

# Associate the Year Hub to the School Hub (nested hub)
#Add-PnPHubSiteAssociation -Site $YearSite -HubSite $SchoolSite
Add-PnPHubToHubAssociation -SourceUrl $YearSite -TargetUrl $SchoolSite
Write-Host "   Associated to School Hub" -ForegroundColor Green

# ============================================================
# 4. CREATE SUBJECT SITES (Team Sites) UNDER THE YEAR HUB
# ============================================================
foreach ($Subject in $Subjects) {
    $Alias = "$($SchoolTitle.ToLower())-$($YearTitle.Replace('.','-'))-$($Subject.ToLower())"
    Write-Host "`n📚 Creating subject site: $Subject..." -ForegroundColor Cyan

    # Option A: M365 Group-connected Team Site (full collaboration)
    # $SubjectUrl = New-PnPSite -Type TeamSite -Title $Subject -Alias $Alias

    # Option B: Team Site WITHOUT M365 Group (leaner, no Teams/Outlook group)
    $SubjectUrl = New-PnPSite `
        -Type TeamSiteWithoutMicrosoft365Group `
        -Title $Subject `
        -Url "https://$TenantName.sharepoint.com/sites/$Alias" `
        -Owner $OwnerEmail

    Write-Host "   Site created at: $SubjectUrl" -ForegroundColor Green

    # Associate the subject site to the Year Hub
    Add-PnPHubSiteAssociation -Site $SubjectUrl -HubSite $YearSite
    Write-Host "   Associated to Year Hub" -ForegroundColor Green
}

# ============================================================
# 5. VERIFY THE STRUCTURE
# ============================================================
Write-Host "`n✅ Done! Verifying hub structure..." -ForegroundColor Cyan
Get-PnPHubSite | Format-Table SiteUrl, Title, Description

Write-Host "`n🔗 URLs Summary:" -ForegroundColor Yellow
Write-Host "   School Hub:  $SchoolSite"
Write-Host "   Year Hub:    $YearSite"
foreach ($Subject in $Subjects) {
    $Alias = "$($SchoolTitle.ToLower())-$($YearTitle.Replace('.','-'))-$($Subject.ToLower())"
    Write-Host "   $Subject`: https://$TenantName.sharepoint.com/sites/$Alias"
}