# Start measuring the script runtime
$startTime = Get-Date

# Excel file path and details
$baseExcelFilePath = ""
$sheetName = "Hostnames"
$hostnameColumn = "C"
$resultColumn = "D"

# Initialize a counter for failed pings
$failedCount = 0
$totalCount = 0
$failedHosts = @()

# Open the base Excel file as read only
$excel = New-Object -ComObject Excel. Application
$excel. Visible = $false
$workbook = $excel.Workbooks.Open($baseExcelFflePath, [System. Type]::Missing, $true) # Open as read-only
$worksheet = $workbook.Sheets.Item($sheetName)

# Create a new workbook to save the results
$newWorkbook = $excel.Workbooks.Add()
$newWorksheet = $newWorkbook.Sheets.Item(1)

# Copy the content from the base workbook to the new workbook
$worksheet.UsedRange.Copy($newworksheet.Range("A1"))

# Ping hostnames and update the new Excel File
$rowIndex = 2
while ($newworksheet.Range("$hostnameColumn$rowIndex*).Text -ne "") {
    $hostname = $newWorksheet.Range("$hostnameColumn$rowIndex").Text
    $result = Test-Connection -ComputerName $hostname -Count 1 -Quiet
    $newWorksheet.Range("$resultColumn$rowIndex").Value2 = if ($result) ("Success" ) else {
    "Failed"
    $failedcount++
    $rowdata = @()
    for ($col = 1; $col -le $newWorksheet.UsedRange.Columns.Count; $col++) {
    $rowdata += $newWorksheet.Cells.Item($rowIndex, $col).Text
}
$failedHosts += ,$rowData
}
$totalCount++
$rowIndex++

# Debugging: Output the total and failed counts
Write-Output "Total Hosts: $totalCount" 
Write-Output "Failed Hosts: $failedCount"

Define destination directory
destinationDirectory = ""

# Get current date and time in a filename-friendly format
$dateString = Get-Date -Format "dd-MMM-yyyy"

# Create base file name
$baseFileName = "Desired_Ping_ Result_$($dateString)*
$fileExtension = ".xlsx"

# Generate a unique file name by adding an incremental value if the file already exists
$counter = 1
$newFileName = "$baseFileName$fileExtension"
$newFilePath = Join-Path -Path $destinationDirectory -ChildPath $newFileName
$counter++

while (Test-Path $newFilePath) {
$newFileName = "$baseFileName' ($counter')$fileExtension"
$newFilePath = Join-Path -Path $destinationDirectory -ChildPath $newFileName
$counter++
}

# Save the new workbook
$newWorkbook.SaveAs($newFilePath)

# Output the new file path for confirmation
Write-Output "File saved as: $newFilePath"

# Close Excel workbooks
$workbook.Saved = $true # Mark the base workbook as saved to avoid prompt
$workbook.Close($false) # Close the base workbook without saving
$newworkbook.Close($true) # Close the new workbook and save changes
$excel.Quit()

# Release COM objects (Excel)
[System. Runtime.InteropServices-Marshal]::ReleaseComObject($worksheet) | Out -Null 
[System.Runtime.InteropServices.Marshal]:: ReleaseComObject($workbook) | Out-Null [System.Runtime.InteropServices.Marshal]::ReleaseComObject($newWorksheet) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($newworkbook) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

Step 2: Read the data from the new Excel file and filter for failed hosts
Sexcel - New-Object -ComObject Excel.Application
Sexcel. Visible - $false
Sexcel.DisplayAlerts = $false
Sworkbook = Sexcel.Workbooks.Open($newFilePath)
Sworksheet = Sworkbook.Sheets. Item(1)
SusedRange - $worksheet.UsedRange
tailedhosts = 00)

# Get headers
for ($col = 1; $col -le SusedRange.Columns. Count; $col++) {
Sheaders += SusedRange.Cells. Item(1, $col). Text
}
# Get failed hosts
For ($row = 2; Srow -le SusedRange.Rows.Count; Srow++) {# Assuming the first row is the header
$status = SusedRange. Cells. ItemSrow, 4).Text # Assuming the status is in the fourth column
if ($status -eq "Failed") (
$rowData = @()
for $col = 1; Scol - le SusedRange.Columns.Count; $col++) (
SrowData += SusedRange. Cells. ItemSrow, Scol). Text
§failedHosts +*, SrowData
# Close the Excel file
Sworkbook. Close($false) Sexcel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseCom0bject(Sworksheet) | Out-Null [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null [System.Runtime. Interopservices.Marshal]::ReleaseComObject(Sexcel) | Out-Null
# Calculate failure percentage
§failurePercentage - If ($totalCount -ne 0) { [math]:: Round((SfailedCount / StotalCount) * 100, 2) } else ( 0 }
# Debugging: Output the failure percentage
Write-Output "Failure Percentage: $failurePercentage%"
# Step 3: Convert the filtered data to HTML
ShtmiTable - "<table border='1'><tr>" foreach (Sheader in Sheaders) (
ShtulTable += *<th›$header</th›*
ShtmlTable +- *</tr>*
foreach (Srow in $failedHosts) {
ShtmlTable += "<tr>"
foreach (Scell in Srow) {
ShtmiTable += *<td›$cell</td›*
}
ShtmlTable += *</tr>*
ShtmiTable +- "‹/table>*
# Debugging: Output the HTML table to the console
Write-Output "HTML Table: ShtmlTable"
alculate the script runtime

# Calculate the script runtime
SendTime = Get-Date
Sruntime - SendTime - SstartTime
SruntimeFormatted - Sruntime. ToString(hh\=mm\: 5s*)
# Step 4: Send the HTML content in an email
$smtpServer - *smtp.example.com"
§smtpFrom = *your_email@example.com"
SsmtpTo = "receiver_email@example.com"
SmessageSubject - "Scheduled PMD & BBG Checkout Results | SdateString | Failure count - $failedCount-
SmessageBody = 0*
chiml>
< body>
<h2>Failed Hosts Repont</h2> (table border='1*›
‹tr›
<th›Total Hosts</th› <th›Failed Hosts</th›
<th›Failure Percentages/th>
</tr>
‹tr›
<td›$totalCount</td›
(td›$failedCount</td>
ctd>$failurePercentage%</td›
</tr>
‹/table›
‹br>
ShtmlTable
<p›The ping test results have been saved as 'SnewFileName' under location: SdestinationDirectory.‹/p›
‹p›Script Runtime: SruntimeFormatted</p›
</body> </html>
"e
Send-MailMessage -From SsitpFron -To $smtpTo -Subject SmessageSubject -Body SmessageBody -BodyÄsHtml -SutpServer SsatpServer
# Send email with Out look
# Create a new Outlook application instance
Soutlook - New-Object -ConObject Outlook Application
# Create a new mail item
Small = Soutlook.Createltem(®) # 0 corresponds to olMailItem
# Set subject and body
Smail. Subject - HYD Opel Ping Check Results | SdateString | Failure count - $failedCount"
Smail.HTMLBody = SmessageBody
POethe eletan. salni8gs.com*, "prashanth, hregs, com", "kishorekumargovda.sles. com*, *us-est-hyde. com)

ста›та eacount</ta>
‹td›$failurePercentage%</td›
</tr>
‹/table>
‹br>
ShtmlTable
‹p›The ping test results have been saved as 'SnewFileName' under location: SdestinationDirectory.‹/p›
‹p›Script Runtime: SruntimeFormatted</p›
</body> </html>
Send-MailMessage -From SsmtpFrom -To SsmtpTo -Subject SmessageSubject -Body SmessageBody -BodyAsHtml -SmtpServer SsmtpServer
# Send email with Outlook
# Create a new Outlook application instance
Sout look = New-Object -ComObject Outlook. Application
# Create a new mail item
Small - Soutlook.Createltem(0) # 0 corresponds to olMailItem
# Set subject and body
Small. Subject - "HYD Opel Ping Check Results | SdateString | Failure count - SfailedCount" Smail.HTMLBody - SmessageBody
# Define the recipients
recipients - @(*chetan.saini@gs.com*, "prashanth.hr@gs.com", "kishorekumargowda.sßgs.com", "gs-osd-hyd@gs-com*)
# Add and resolve each recipient
foreach (Srecipient in recipients) {
SmailRecipient - Smail.Recipients.Add (Srecipient)
if (-not SmailRecipient.Resolve()) f
Write-Warning "Could not resolve recipient: Srecipient"
# Check if there are resolved recipients before sending
if (Smail.Recipients.Count -eq 0) (
Write-Error "No valid recipients were added. Email will not be sent."
} else f
# Send the email
Smail. Send(
}
# Release COM objects (Outlook)
# Cleanup
[System. GC]: :Collect)