# Define parameters
param(
    [string]$ManagerEmail,  # Input for Manager Email
    [string]$OutputCsv = "awmsa.csv"  # Output file
)

# List of input CSV files
$csvFiles = @("LowScoringUsers.csv", "UserIncompleteSessions.csv", "UserPhishingFailures.csv", "UserIncompleteRemediations.csv")

# Initialize an empty array to collect data
$allData = @()

# Loop through each CSV file
foreach ($csvFile in $csvFiles) {
    # Extract the ReportType from the file name (without .csv)
    $reportType = [System.IO.Path]::GetFileNameWithoutExtension($csvFile)

    # Import the CSV data
    $data = Import-Csv -Path $csvFile

    # Filter the data based on the Manager Email
    $filteredData = $data | Where-Object { $_.'Manager Email' -eq $ManagerEmail }

    # Add the ReportType column
    $filteredData | ForEach-Object { $_ | Add-Member -MemberType NoteProperty -Name ReportType -Value $reportType }

    # Append the filtered data to the array
    $allData += $filteredData
}

# Sort the data by First Name
$sortedData = $allData | Sort-Object -Property 'First Name'

# Export the sorted data to the output CSV file
$sortedData | Export-Csv -Path $OutputCsv -NoTypeInformation

Write-Host "Report generated: $OutputCsv"

