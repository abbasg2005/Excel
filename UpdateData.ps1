# --- CONFIGURATION ---
$ExcelFile = "D:\PCToGitHub\PendingPOapproved.xlsx"
$RepoPath = "D:\PCToGitHub"
$MyToken = "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"

# --- STEP 1: REFRESH EXCEL ---
Write-Host "Starting Excel..." -ForegroundColor Cyan
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $true 
$Excel.DisplayAlerts = $false
$Excel.AlertBeforeOverwriting = $false

$Workbook = $Excel.Workbooks.Open($ExcelFile)
Write-Host "Refreshing SQL Data..." -ForegroundColor Yellow
$Workbook.RefreshAll()

# Wait for local SQL query to finish
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

# This direct URL bypasses the 403 Forbidden error
git push "https://abbasg2005:$($MyToken)@github.com/abbasg2005/Excel.git" main --force