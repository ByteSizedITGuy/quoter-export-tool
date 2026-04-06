<#
.SYNOPSIS
    Exports all quotes from Quoter (ScalePad CPQ) as PDF files.

.DESCRIPTION
    Authenticates with the Quoter API, retrieves all quotes across all statuses,
    downloads each as a PDF, and organizes them into folders by status.

    Each quote revision is saved as a separate file. Filenames include the customer
    name, quote number, title, revision, and a unique identifier for safe re-runs.

.PARAMETER TenantDomain
    Your Quoter tenant domain (e.g., "yourcompany.quoter.com").

.PARAMETER ClientId
    OAuth client ID from Quoter (Account > API Keys).

.PARAMETER ClientSecret
    OAuth client secret from Quoter (Account > API Keys).

.PARAMETER OutputPath
    Directory to save PDFs. Defaults to ".\QuoterExport" in the current directory.

.PARAMETER CreatedAfter
    Only export quotes created after this date (e.g., "2025-01-01" or "2025-06-15T14:00:00").

.PARAMETER CreatedBefore
    Only export quotes created before this date.

.PARAMETER ModifiedAfter
    Only export quotes modified after this date. Useful for incremental backups — catches
    quotes that were updated (new revisions, status changes) since your last export.

.PARAMETER ModifiedBefore
    Only export quotes modified before this date.

.EXAMPLE
    .\Export-QuoterPDFs.ps1 -TenantDomain "acme.quoter.com" -ClientId "cid_xxx" -ClientSecret "xxx"

.EXAMPLE
    .\Export-QuoterPDFs.ps1 -TenantDomain "acme.quoter.com" -ClientId "cid_xxx" -ClientSecret "xxx" -OutputPath "D:\Backups\Quoter"

.EXAMPLE
    # Monthly backup: export quotes created in March 2026
    .\Export-QuoterPDFs.ps1 -TenantDomain "acme.quoter.com" -ClientId "cid_xxx" -ClientSecret "xxx" -CreatedAfter "2026-03-01" -CreatedBefore "2026-04-01"

.EXAMPLE
    # Weekly incremental: export anything modified in the last 7 days
    .\Export-QuoterPDFs.ps1 -TenantDomain "acme.quoter.com" -ClientId "cid_xxx" -ClientSecret "xxx" -ModifiedAfter (Get-Date).AddDays(-7).ToString("yyyy-MM-dd")

.NOTES
    Requires PowerShell 5.1+ (included with Windows 10/11).
    API rate limit: 5 requests/sec. The script throttles to stay under this.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Your Quoter tenant domain (e.g., yourcompany.quoter.com)")]
    [string]$TenantDomain,

    [Parameter(Mandatory = $true, HelpMessage = "OAuth client ID from Quoter")]
    [string]$ClientId,

    [Parameter(Mandatory = $true, HelpMessage = "OAuth client secret from Quoter")]
    [string]$ClientSecret,

    [Parameter(Mandatory = $false, HelpMessage = "Output directory for PDFs")]
    [string]$OutputPath = ".\QuoterExport",

    [Parameter(Mandatory = $false, HelpMessage = "Only export quotes created after this date (e.g., 2025-01-01)")]
    [string]$CreatedAfter,

    [Parameter(Mandatory = $false, HelpMessage = "Only export quotes created before this date")]
    [string]$CreatedBefore,

    [Parameter(Mandatory = $false, HelpMessage = "Only export quotes modified after this date")]
    [string]$ModifiedAfter,

    [Parameter(Mandatory = $false, HelpMessage = "Only export quotes modified before this date")]
    [string]$ModifiedBefore
)

$ErrorActionPreference = "Stop"

# ── Config ───────────────────────────────────────────────────────────────────
$ApiBase = "https://api.quoter.com/v1"
$PdfBase = "https://$TenantDomain/quote/download"
$PerPage = 100
$RequestDelayMs = 250  # 4 req/sec to stay safely under the 5/sec limit
$MaxRetries = 3

# ── Helpers ──────────────────────────────────────────────────────────────────
function Sanitize-Filename {
    param([string]$Name)
    $invalid = [IO.Path]::GetInvalidFileNameChars() -join ''
    $sanitized = $Name -replace "[$([regex]::Escape($invalid))]", '_'
    $sanitized = $sanitized -replace '\s+', ' '
    return $sanitized.Trim()
}

# ── Authentication ───────────────────────────────────────────────────────────
function Get-AccessToken {
    Write-Host "Authenticating with Quoter API..." -ForegroundColor Cyan

    $body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        grant_type    = "client_credentials"
    } | ConvertTo-Json

    try {
        $response = Invoke-RestMethod -Uri "$ApiBase/auth/oauth/authorize" `
            -Method Post `
            -ContentType "application/json" `
            -Body $body
    }
    catch {
        Write-Host "Authentication failed. Please check your Client ID and Client Secret." -ForegroundColor Red
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "`nYou can generate API keys in Quoter under Account > API Keys (requires Account Owner role)." -ForegroundColor Yellow
        exit 1
    }

    Write-Host "Authenticated successfully.`n" -ForegroundColor Green
    return $response.access_token
}

# ── Build Date Filters ──────────────────────────────────────────────────────
function Build-DateFilter {
    $filters = @()

    if ($CreatedAfter) {
        $dt = [datetime]::Parse($CreatedAfter).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        $filters += "filter[record_created_at]=gt:$dt"
    }
    if ($CreatedBefore) {
        $dt = [datetime]::Parse($CreatedBefore).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        $filters += "filter[record_created_at]=lt:$dt"
    }
    if ($ModifiedAfter) {
        $dt = [datetime]::Parse($ModifiedAfter).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        $filters += "filter[record_updated_at]=gt:$dt"
    }
    if ($ModifiedBefore) {
        $dt = [datetime]::Parse($ModifiedBefore).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        $filters += "filter[record_updated_at]=lt:$dt"
    }

    if ($filters.Count -gt 0) {
        return "&" + ($filters -join "&")
    }
    return ""
}

# ── Fetch a Page of Quotes ──────────────────────────────────────────────────
function Get-QuotePage {
    param(
        [string]$Token,
        [int]$Page
    )

    $dateFilter = Build-DateFilter
    $url = "$ApiBase/quotes?per_page=$PerPage&page=$Page$dateFilter"
    $headers = @{ Authorization = "Bearer $Token" }

    try {
        $response = Invoke-RestMethod -Uri $url -Method Get -Headers $headers
    }
    catch {
        if ($_.Exception.Response.StatusCode -eq 429) {
            Write-Host "  Rate limited on page $Page, waiting 2 seconds..." -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            return Get-QuotePage -Token $Token -Page $Page
        }
        throw
    }

    return $response
}

# ── Count Total Quotes ─────────────────────────────────────────────────────
function Get-TotalQuoteCount {
    param([string]$Token)

    $dateFilter = Build-DateFilter
    $url = "$ApiBase/quotes?per_page=1&page=1$dateFilter"
    $headers = @{ Authorization = "Bearer $Token" }
    $response = Invoke-RestMethod -Uri $url -Method Get -Headers $headers
    return $response.total_count
}

# ── Download a Single PDF ───────────────────────────────────────────────────
function Save-QuotePdf {
    param(
        [object]$Quote,
        [string]$FilePath
    )

    $url = "$PdfBase/$($Quote.uuid)"

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            # Download to a temp file first, then move to final path.
            # This avoids issues with special characters (brackets, etc.) in -OutFile.
            $tmpFile = [System.IO.Path]::GetTempFileName()
            Invoke-WebRequest -Uri $url -OutFile $tmpFile -UseBasicParsing | Out-Null

            # Verify it's actually a PDF
            $header = [System.IO.File]::ReadAllBytes($tmpFile)[0..4]
            if ($header -and $header.Count -ge 4) {
                $magic = [System.Text.Encoding]::ASCII.GetString($header[0..3])
                if ($magic -eq "%PDF") {
                    [System.IO.File]::Move($tmpFile, $FilePath)
                    return $true
                }
            }

            # Not a PDF - clean up and retry
            if (Test-Path -LiteralPath $tmpFile) { Remove-Item -LiteralPath $tmpFile -Force }

            if ($attempt -eq $MaxRetries) { return $false }
            Start-Sleep -Seconds 1
        }
        catch {
            if (Test-Path -LiteralPath $tmpFile) { Remove-Item -LiteralPath $tmpFile -Force }

            $status = $_.Exception.Response.StatusCode.value__
            if ($status -eq 429) {
                Write-Host "    Rate limited, waiting 2s..." -ForegroundColor Yellow
                Start-Sleep -Seconds 2
                continue
            }

            if ($attempt -eq $MaxRetries) {
                return $false
            }
            Start-Sleep -Seconds 1
        }
    }

    return $false
}

# ── Main ─────────────────────────────────────────────────────────────────────
Write-Host "`n============================================" -ForegroundColor White
Write-Host "  Quoter PDF Export Tool" -ForegroundColor White
Write-Host "  Tenant: $TenantDomain" -ForegroundColor Gray
if ($CreatedAfter)   { Write-Host "  Created After:  $CreatedAfter" -ForegroundColor Gray }
if ($CreatedBefore)  { Write-Host "  Created Before: $CreatedBefore" -ForegroundColor Gray }
if ($ModifiedAfter)  { Write-Host "  Modified After:  $ModifiedAfter" -ForegroundColor Gray }
if ($ModifiedBefore) { Write-Host "  Modified Before: $ModifiedBefore" -ForegroundColor Gray }
Write-Host "============================================`n" -ForegroundColor White

# Validate tenant domain format
if ($TenantDomain -notmatch '\.quoter\.com$') {
    Write-Host "Warning: Tenant domain '$TenantDomain' doesn't end with .quoter.com" -ForegroundColor Yellow
    Write-Host "Expected format: yourcompany.quoter.com`n" -ForegroundColor Yellow
}

# Authenticate
$token = Get-AccessToken

# Get total count first
$dateFilter = Build-DateFilter
if ($dateFilter) {
    Write-Host "Fetching quotes with date filters..." -ForegroundColor Cyan
}
else {
    Write-Host "Fetching quotes..." -ForegroundColor Cyan
}

$total = Get-TotalQuoteCount -Token $token
Start-Sleep -Milliseconds $RequestDelayMs

if ($total -eq 0) {
    Write-Host "No quotes found. Nothing to export." -ForegroundColor Yellow
    exit 0
}

Write-Host "Total quotes to process: $total`n" -ForegroundColor Green

# Set up output directory
$OutputPath = [IO.Path]::GetFullPath($OutputPath)

# Download PDFs page-by-page (fetching fresh UUIDs immediately before downloading)
Write-Host "Downloading PDFs to: $OutputPath`n" -ForegroundColor Cyan

$downloaded = 0
$failed = 0
$skipped = 0
$failures = @()
$page = 1
$hasMore = $true

while ($hasMore) {
    Write-Host "--- Fetching page $page ---" -ForegroundColor Cyan
    $response = Get-QuotePage -Token $token -Page $page
    $quotes = $response.data
    $hasMore = $response.has_more

    Write-Host "  Got $($quotes.Count) quotes, downloading immediately...`n" -ForegroundColor Gray

    foreach ($quote in $quotes) {
        $downloaded++

        # Ensure status directory exists
        $status = if ($quote.status) { $quote.status } else { "unknown" }
        $statusDir = Join-Path $OutputPath $status
        if (-not (Test-Path $statusDir)) {
            New-Item -ItemType Directory -Path $statusDir -Force | Out-Null
        }

        # Build filename: Customer - Number - Title (Rev X) [uuid-short].pdf
        $customer = if ($quote.billing_organization) { $quote.billing_organization } else { "Unknown Customer" }
        $title = if ($quote.name) { $quote.name } else { "Untitled" }
        $revSuffix = if ([int]$quote.revision -gt 1) { " (Rev $($quote.revision))" } else { "" }
        $uuidShort = $quote.uuid.Substring(0, 8)

        $filename = Sanitize-Filename "$customer - $($quote.number) - $title$revSuffix [$uuidShort].pdf"
        $filePath = Join-Path $statusDir $filename

        # Skip if already downloaded (resume support)
        if ((Test-Path -LiteralPath $filePath) -and (Get-Item -LiteralPath $filePath).Length -gt 0) {
            $skipped++
            Write-Host "  [$downloaded/$total] SKIP $filename" -ForegroundColor DarkGray
            continue
        }

        $success = Save-QuotePdf -Quote $quote -FilePath $filePath

        if ($success) {
            $sizeMB = [math]::Round((Get-Item -LiteralPath $filePath).Length / 1MB, 2)
            Write-Host "  [$downloaded/$total] $filename ($sizeMB MB)" -ForegroundColor Green
        }
        else {
            $failed++
            $failures += [PSCustomObject]@{
                Filename = $filename
                Status   = $status
                QuoteId  = $quote.id
                Number   = $quote.number
                UUID     = $quote.uuid
            }
            Write-Host "  [$downloaded/$total] FAIL $filename" -ForegroundColor Red
        }

        Start-Sleep -Milliseconds $RequestDelayMs
    }

    $page++
    Write-Host ""
}

# Summary
Write-Host "============================================" -ForegroundColor White
Write-Host "  Export Complete" -ForegroundColor White
Write-Host "============================================" -ForegroundColor White
Write-Host "  Downloaded : $($total - $failed - $skipped)" -ForegroundColor Green
Write-Host "  Skipped    : $skipped" -ForegroundColor DarkGray
Write-Host "  Failed     : $failed" -ForegroundColor $(if ($failed -gt 0) { "Red" } else { "Green" })
Write-Host "  Output     : $OutputPath" -ForegroundColor Cyan
Write-Host ""

if ($failures.Count -gt 0) {
    $failPath = Join-Path $OutputPath "failures.csv"
    $failures | Export-Csv -Path $failPath -NoTypeInformation
    Write-Host "Failed downloads saved to: $failPath" -ForegroundColor Yellow
    Write-Host ""
    foreach ($f in $failures) {
        Write-Host "  - $($f.Filename)" -ForegroundColor Red
    }
    Write-Host ""
}

Write-Host "Done!`n" -ForegroundColor Green
