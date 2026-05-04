# --- CONFIGURATION ---
$ExcelFile = "D:\PCToGitHub\PendingPOapproved.xlsx"
$RepoPath = "D:\PCToGitHub"
$MyToken = "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"
$RemoteUrl = "https://abbasg2005:$($MyToken)@github.com/abbasg2005/Excel.git"

# --- STEP 1: REFRESH EXCEL ---
Write-Host "Starting Excel..." -ForegroundColor Cyan
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $true  # We want to see it for this test
$Workbook = $Excel.Workbooks.Open($ExcelFile)

Write-Host "Refreshing SQL Data... Please wait." -ForegroundColor Yellow
$Workbook.RefreshAll()

# Wait 45 seconds to ensure the local SQL query finishes
Start-Sleep -Seconds 45 

$Workbook.Save()
$Workbook.Close()
$Excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
Write-Host "Excel saved and closed." -ForegroundColor Green

# --- STEP 2: PUSH TO GITHUB ---
Set-Location $RepoPath
Write-Host "Pushing to GitHub..." -ForegroundColor Cyan

git add .
git commit -m "Auto-refresh: $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
git push $RemoteUrl main --force