# Quoter PDF Export Tool

A PowerShell script that bulk-exports all quotes from Quoter as PDF files for backup or archival.

Built for **MSPs and IT service providers** who want to maintain local backups of their quoting data.

## Features

- Exports **every quote** across all statuses (pending, accepted, fulfilled, lost)
- Downloads **every revision** as a separate PDF
- Organizes PDFs into folders by status
- Filenames include customer name, quote number, title, and revision
- **Date filtering** — export by created or modified date for incremental/scheduled backups
- **Resume-safe** — re-run the script and it skips already-downloaded files
- Respects Quoter's API rate limits
- Zero dependencies — just PowerShell 5.1+ (pre-installed on Windows 10/11)

## Prerequisites

1. **A Quoter account** with Account Owner access
2. **API credentials** — generate these in Quoter:
   - Log in to Quoter
   - Go to **Account > API Keys**
   - Click **Create API Key**
   - Copy the **Client ID** and **Client Secret**
3. **Your tenant domain** — this is the subdomain you use to access Quoter (e.g., `yourcompany.quoter.com`). You can find it in your browser's address bar when logged in.

## Quick Start

1. Download `Export-QuoterPDFs.ps1` to a folder on your computer.

2. Open PowerShell and navigate to that folder:
   ```powershell
   cd C:\path\to\folder
   ```

3. If you haven't run PowerShell scripts before, you may need to allow script execution:
   ```powershell
   Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
   ```

4. Run the export:
   ```powershell
   .\Export-QuoterPDFs.ps1 -TenantDomain "yourcompany.quoter.com" -ClientId "cid_xxxxx" -ClientSecret "your_secret_here"
   ```

5. PDFs will be saved to a `QuoterExport` folder in your current directory.

## Parameters

| Parameter | Required | Description |
|---|---|---|
| `-TenantDomain` | Yes | Your Quoter tenant domain (e.g., `yourcompany.quoter.com`) |
| `-ClientId` | Yes | OAuth Client ID from Quoter API Keys |
| `-ClientSecret` | Yes | OAuth Client Secret from Quoter API Keys |
| `-OutputPath` | No | Directory to save PDFs (default: `.\QuoterExport`) |
| `-CreatedAfter` | No | Only export quotes created after this date (e.g., `2025-01-01`) |
| `-CreatedBefore` | No | Only export quotes created before this date |
| `-ModifiedAfter` | No | Only export quotes modified after this date |
| `-ModifiedBefore` | No | Only export quotes modified before this date |

## Examples

**Basic export:**
```powershell
.\Export-QuoterPDFs.ps1 -TenantDomain "acme.quoter.com" -ClientId "cid_abc123" -ClientSecret "mysecret"
```

**Export to a specific folder:**
```powershell
.\Export-QuoterPDFs.ps1 -TenantDomain "acme.quoter.com" -ClientId "cid_abc123" -ClientSecret "mysecret" -OutputPath "D:\Backups\Quoter"
```

**Resume a previous export** (just run the same command again — already-downloaded files are skipped):
```powershell
.\Export-QuoterPDFs.ps1 -TenantDomain "acme.quoter.com" -ClientId "cid_abc123" -ClientSecret "mysecret"
```

**Monthly backup — export quotes created in March 2026:**
```powershell
.\Export-QuoterPDFs.ps1 -TenantDomain "acme.quoter.com" -ClientId "cid_abc123" -ClientSecret "mysecret" -CreatedAfter "2026-03-01" -CreatedBefore "2026-04-01"
```

**Weekly incremental — export anything modified in the last 7 days:**
```powershell
.\Export-QuoterPDFs.ps1 -TenantDomain "acme.quoter.com" -ClientId "cid_abc123" -ClientSecret "mysecret" -ModifiedAfter (Get-Date).AddDays(-7).ToString("yyyy-MM-dd")
```

**Catch up on recent activity — quotes created or updated this year:**
```powershell
.\Export-QuoterPDFs.ps1 -TenantDomain "acme.quoter.com" -ClientId "cid_abc123" -ClientSecret "mysecret" -ModifiedAfter "2026-01-01"
```

## Output Structure

```
QuoterExport/
  accepted/     # Signed/accepted quotes
  fulfilled/    # Fulfilled quotes
  lost/         # Lost quotes
  pending/      # Pending/open quotes
  failures.csv  # Any downloads that failed (if applicable)
```

**Filename format:**
```
Customer Name - 2010701 - Quote Title (Rev 2) [2525-abc].pdf
```

## How Long Does It Take?

The Quoter API has a rate limit of 5 requests per second. The script throttles to 4/sec to stay safe. Expect roughly **3 seconds per quote** including the PDF download.

| Quotes | Estimated Time |
|---|---|
| 50 | ~3 minutes |
| 100 | ~5 minutes |
| 300 | ~15 minutes |
| 500 | ~25 minutes |

## Scheduled Backups

You can use Windows Task Scheduler to run this automatically on a schedule.

**Example: Weekly backup every Monday at 6 AM**

1. Open Task Scheduler and create a new task.
2. Set the trigger to **Weekly**, on **Monday** at **6:00 AM**.
3. Set the action to **Start a program**:
   - Program: `powershell.exe`
   - Arguments:
     ```
     -ExecutionPolicy Bypass -File "C:\Scripts\Export-QuoterPDFs.ps1" -TenantDomain "acme.quoter.com" -ClientId "cid_abc123" -ClientSecret "mysecret" -ModifiedAfter (Get-Date).AddDays(-7).ToString("yyyy-MM-dd") -OutputPath "D:\Backups\Quoter"
     ```

Since the script skips already-downloaded files (matched by unique ID in the filename), overlapping date ranges won't create duplicates.

## Troubleshooting

**"Authentication failed"**
- Verify your Client ID and Client Secret are correct
- API keys are generated in Quoter under **Account > API Keys**
- You need the **Account Owner** role to create API keys

**"Running scripts is disabled on this system"**
- Run: `Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned`

**Some quotes show as FAIL**
- Some draft quotes or deleted revisions may not have PDFs available on Quoter's servers
- Failed downloads are logged to `failures.csv` in the output folder

**Script was interrupted**
- Just run it again — it will skip already-downloaded files and pick up where it left off

## Security Notes

- **Never commit your API credentials to source control.** Pass them as parameters or use environment variables.
- The script only **reads** data from Quoter — it does not create, modify, or delete anything.
- PDF download URLs are unauthenticated (they use a unique UUID), which is how Quoter's built-in quote sharing works.

## License

0BSD — Use however you want, no attribution required, no warranty.
