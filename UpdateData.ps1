# --- CONFIGURATION ---
$ExcelFile = "D:\PCToGitHub\PendingPOapproved.xlsx"  # <--- Change to your actual file name
$RepoPath = "D:\PCToGitHub"

# --- STEP 1: REFRESH EXCEL DATA ---
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$Workbook = $Excel.Workbooks.Open($ExcelFile)

# This triggers the SQL query update
$Workbook.RefreshAll()

# Wait 30 seconds for the SQL data to finish downloading
Start-Sleep -Seconds 30 

$Workbook.Save()
$Excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)

# --- STEP 2: PUSH TO GITHUB ---
Set-Location $RepoPath
git add .
git commit -m "Auto-refresh SQL data: $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
git push origin main