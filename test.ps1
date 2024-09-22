Set-ExecutionPolicy unrestricted

$loginUrl = 'https://login.flexera.com/oidc/token'
$endpoint = 'https://api.flexera.com'
$refreshToken = 'MwD8ZXNqG4lKT04Ldn9Nu2QWA-FLWtMNzOHQe_Xeg34'
$customReport = 'Contract Data - Shell Tickets'
$excelFilePath = "C:\Loveneesh\Powershell\workbook.xlsx"
$sheetName = "Shell Ticket Data"
$tableName = "Table1"

# Generate access token based on input org ID and refresh token.
try {
    Write-Host "Generating access token..."
    $postParams = @{
        grant_type='refresh_token';
        refresh_token=$refreshToken
    }
    $resp = Invoke-WebRequest -Uri $loginUrl -Method POST -Body $postParams -ContentType 'application/x-www-form-urlencoded'
    $accessToken = ($resp.Content | ConvertFrom-Json).access_token
    $resp = Invoke-WebRequest -Uri "https://api.flexera.com/fnms/v1/orgs/27816/reports" -Method GET -Headers @{Authorization="Bearer $accessToken"}
    Write-Host "Access token generated!"
    if (-not [string]::IsNullOrEmpty($customReport)) {
        Write-Host "Getting reportID based on report name: $customReport"
        $reports = $resp | ConvertFrom-Json
        $reportID = $reports | Where-Object {$_.title -eq "$customReport"}
        $reportID = $reportID.id
    }
}
catch {
    Write-Error $_
    exit
}

# Execute the custom report using the ID we found based on the name.
Write-Host "Executing custom report with ID:$reportID. Complex reports may take a while." -fore white -back black
$nextPage = $null
$resp = Invoke-WebRequest -Uri "$endpoint/fnms/v1/orgs/27816/reports/$reportID/execute" -Method GET -Headers @{Authorization="Bearer $accessToken"}
if ($resp.StatusCode -eq '200') {
    # Suppressing error for this approach to getting the nextPage
    $ErrorActionPreference = "SilentlyContinue"
    $nextPage = $resp.Content | ConvertFrom-Json | Select-Object -expand "nextPage"
    $ErrorActionPreference = "Continue"
}

# Converting json output to PowerShell
$respJson = $resp.Content | ConvertFrom-Json
$results = $respJson.values

# Create an instance of Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

# Open the existing workbook
$workbook = $excel.Workbooks.Open($excelFilePath)
$sheet = $workbook.Sheets.Item($sheetName)

# Clear existing data (optional)
$sheet.Cells.Clear()

# Write headers to the sheet
$headers = $results[0].PSObject.Properties.Name
for ($i = 0; $i -lt $headers.Count; $i++) {
    $sheet.Cells.Item(1, $i + 1) = $headers[$i]
}

# Write data to the sheet
$row = 2
foreach ($result in $results) {
    $col = 1
    foreach ($key in $result.PSObject.Properties.Name) {
        $sheet.Cells.Item($row, $col) = $result.$key
        $col++
    }
    $row++
}

# Define the range for the table
$lastColumn = $headers.Count
$lastRow = $results.Count + 1
$tableRange = $sheet.Range("A1").Resize($lastRow, $lastColumn)

# Check if the table already exists and delete it
try {
    $table = $sheet.ListObjects.Item($tableName)
    $table.Delete()
}
catch {
    # Table does not exist, no action needed
}

# Create a new table
$table = $sheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $tableRange, $null, [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
$table.Name = $tableName
$table.TableStyle = "TableStyleMedium2" # Change the style as needed

# Save and close the workbook
$workbook.Save()
$workbook.Close()

# Quit Excel
$excel.Quit()

# Release COM objects
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host "Process Complete!" -fore white -back black