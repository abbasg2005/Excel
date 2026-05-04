# --- CONFIGURATION ---
$ExcelFile = "D:\PCToGitHub\PendingPOapproved.xlsx"
$RepoPath = "D:\PCToGitHub"
# Your verified token
$MyToken = "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"

# --- STEP 1: REFRESH EXCEL ---
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false 
$Excel.DisplayAlerts = $false
$Workbook = $Excel.Workbooks.Open($ExcelFile)
$Workbook.RefreshAll()
Start-Sleep -Seconds 45 
$Workbook.Save() 
$Workbook.Close()
$Excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)

# --- STEP 2: PUSH TO GITHUB ---
Set-Location $RepoPath
git add .
git commit -m "First upload of PendingPOapproved: $(Get-Date -Format 'yyyy-MM-dd HH:mm')"

# We use 'abbasg2005' here because that is the name in your URL
git push "https://abbasg2005:$($MyToken)@github.com/abbasg2005/Excel.git" main --force