<# 
.SYNOPSIS
This script offers a choice to either delete a previously run search or to initiate a new process for deleting meeting requests.

.DESCRIPTION
The user is prompted at the beginning of the script to choose between deleting a previously run search or starting the process of deleting specific meeting requests. Depending on the choice, the script either proceeds directly to the deletion of a previous search or begins the process for a new search and deletion operation.

Author: Mohd Azhar uddin
.NOTES
Version: 2.0
Date: 10th April 2024
Version: 1.0
Author: Mohd Azhar uddin
Date: 10th April 2024

.EXAMPLE
PS> .\Remove-ExEmployeeMeetings.ps1
#>

function Delete-PreviouslyRunSearch {
    $searchName = Read-Host "Enter the name of the search to delete"
    # Assume Delete-ComplianceSearch is a placeholder for the actual cmdlet or script that deletes the search
    New-ComplianceSearchAction -SearchName $searchName -Purge -PurgeType HardDelete
    Write-Host "Deletion action initiated for search: $searchName" -ForegroundColor Green
}

function Start-NewDeletionProcess {
# Check and Install Required Module for Exchange Online
$requiredModule = "ExchangeOnlineManagement"
if (-not (Get-Module -ListAvailable -Name $requiredModule)) {
    Install-Module -Name $requiredModule -Force -AllowClobber
    Write-Host "Installed module: $requiredModule" -ForegroundColor Green
} else {
    Write-Host "Module already installed: $requiredModule" -ForegroundColor Yellow
}

# Connect to Exchange Online
try {
    Connect-ExchangeOnline
    Write-Host "Connected to Exchange Online." -ForegroundColor Green
} catch {
    Write-Error "Error connecting to Exchange Online: $_"
    exit
}

# Connect to Security & Compliance Center
try {
    Connect-IPPSSession
    Write-Host "Connected to Security & Compliance Center." -ForegroundColor Green
} catch {
    Write-Error "Error connecting to Security & Compliance Center: $_"
    exit
}

# Define the search query parameters
$searchName = Read-Host "Enter a name for the search"
$startDate = Read-Host "Enter the start date for the search (YYYY-MM-DD)"
$subject = Read-Host "Enter the subject of the meeting requests to search for"

# Adjust the date query
$searchDate = (Get-Date $startDate).AddDays(-1).ToString("yyyy-MM-dd")

# Construct the search query
$searchQuery = "(c:c)(date>$searchDate)(subject:`"$subject`")(ItemClass=IPM.Appointment)(ItemClass=IPM.Schedule)"
Write-Host "Search query constructed: $searchQuery" -ForegroundColor Yellow

# Create a new compliance search
try {
    New-ComplianceSearch -Name "$searchName" -ExchangeLocation All -ContentMatchQuery $searchQuery
    Write-Host "Search created with name: $searchName-$startDate" -ForegroundColor Green
} catch {
    Write-Error "Error creating compliance search: $_"
    exit
}

# Start the compliance search
try {
    Start-ComplianceSearch -Identity "$searchName-$startDate"
    Write-Host "Search started: $searchName-$startDate" -ForegroundColor Green
} catch {
    Write-Error "Error starting compliance search: $_"
    exit
}

# Wait for the search to complete (polling)
$searchCompleted = $false
while (-not $searchCompleted) {
    Start-Sleep -Seconds 30
    $searchStatus = Get-ComplianceSearch -Identity "$searchName-$startDate"
    if ($searchStatus.Status -eq "Completed") {
        $searchCompleted = $true
        Write-Host "Search completed: $searchName-$startDate" -ForegroundColor Green
    } else {
        Write-Host "Waiting for search to complete..." -ForegroundColor Yellow
    }
}

# Wait for an additional 30 seconds to ensure the search results are fully available
Start-Sleep -Seconds 30

# Retrieve and display the search statistics
try {
    $searchResults = Get-ComplianceSearch -Identity "$searchName-$startDate"
    Write-Host "Search results statistics:" -ForegroundColor Green
    Write-Host "Items found: $($searchResults.Items)" -ForegroundColor Yellow
    Write-Host "Unindexed items: $($searchResults.UnindexedItems)" -ForegroundColor Yellow
    Write-Host "Mailboxes: $($searchResults.Mailboxes)" -ForegroundColor Yellow
} catch {
    Write-Error "Error retrieving search results statistics: $_"
    exit
}

# Export the search results
try {
    New-ComplianceSearchAction -SearchName "$searchName-$startDate" -Export -Format FxStream -Confirm:$false > $null
    Write-Host "Export action initiated for search: $searchName-$startDate" -ForegroundColor Green
} catch {
    Write-Error "Error exporting search results: $_"
    exit
}

Write-Host "Search and export completed successfully. Follow the instructions below to review and delete the search results:" -ForegroundColor Yellow
Write-Host "1. Wait for the notification email that the export is ready." -ForegroundColor Yellow
Write-Host "2. Follow the link in the email to access the Security & Compliance Center." -ForegroundColor Yellow
Write-Host "3. Navigate to 'Search > Content search' and find your search name: $searchName-$startDate." -ForegroundColor Yellow
Write-Host "4. Click on 'Export' to download the results as PST or individual messages." -ForegroundColor Yellow
Write-Host "5. Review the export to ensure it contains the expected results." -ForegroundColor Yellow

# Ask for confirmation to proceed with deletion
$validation = Read-Host "Once you have reviewed the search results, do you want to proceed with deletion? (Y/N)"
if ($validation -eq "Y") {
    try {
        $purgeAction = New-ComplianceSearchAction -SearchName "$searchName-$startDate" -Purge -PurgeType HardDelete
        Write-Host "Deletion action initiated for search: $searchName-$startDate" -ForegroundColor Green
        # Note: Retrieving the exact number of items purged might depend on your environment's configuration.
    } catch {
        Write-Error "Error during deletion process: $_"
    }
} else {
    Write-Host "Deletion canceled by user." -ForegroundColor Yellow
}

# Informing about the search and export names
Write-Host "The search was named: $searchName-$startDate" -ForegroundColor Yellow
Write-Host "The export was named: Export-$searchName-$startDate" -ForegroundColor Yellow

# Disconnect the session
Disconnect-ExchangeOnline -Confirm:$false
}

# User choice
Write-Host "Select an operation:" -ForegroundColor Magenta
Write-Host "1. Delete a previously run search" -ForegroundColor Green
Write-Host "2. Start a new meeting request deletion process" -ForegroundColor Green
$choice = Read-Host "Enter your choice (1 or 2)"

switch ($choice) {
    "1" {
        Delete-PreviouslyRunSearch
    }
    "2" {
        Start-NewDeletionProcess
    }
    default {
        Write-Host "Invalid choice. Please enter a valid option (1 or 2)."
    }
}
