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

    # Add "Status", "Quiz Score", and "Clicked" columns (if missing) and push manager-related columns down
    $filteredData | ForEach-Object {
        # Add Status (if it doesn't already exist)
        if (-not $_.PSObject.Properties['Status']) { $_ | Add-Member -MemberType NoteProperty -Name Status -Value $null }
        # Add Quiz Score (if it doesn't already exist)
        if (-not $_.PSObject.Properties['Quiz Score']) { $_ | Add-Member -MemberType NoteProperty -Name 'Quiz Score' -Value $null }
        # Add Clicked (if it doesn't already exist)
        if (-not $_.PSObject.Properties['Clicked']) { $_ | Add-Member -MemberType NoteProperty -Name Clicked -Value $null }
    }

    # Append the filtered data to the array
    $allData += $filteredData
}

# Sort the data by ReportType and then by First Name
$sortedData = $allData | Sort-Object -Property ReportType, 'First Name'

# Reorder columns, making ReportType the first column and placing "Status", "Quiz Score", and "Clicked" before the Manager columns
$sortedData = $sortedData | Select-Object ReportType, 'First Name', 'Last Name', Email, 'Sent Date (UTC)', Title, Status, 'Quiz Score', Clicked, 'Manager First Name', 'Manager Last Name', 'Manager Email'

# Export the sorted data to the output CSV file
$sortedData | Export-Csv -Path $OutputCsv -NoTypeInformation

Write-Host "Report generated: $OutputCsv"
