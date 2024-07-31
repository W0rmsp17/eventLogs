#Copy script onto users computer 
#Will scan
#It will copy any issue related to that thing 
# Event Things:
# Bitlocker, Printer... 
#

param (
    [string]$SearchLog,
    [int]$Days = 7,  
    [string]$Level = "All", 
    [switch]$EnableKeywords,  
    [switch]$Verbose,
    [switch]$Debug,
    [switch]$help
)

$helpMessage = @"
Usage: .\printSweeper.ps1 -SearchLog <SearchTerm> [-Days <NumberOfDays>] [-Level <EventLevel>] [-EnableKeywords] [-Verbose] [-Debug] [--help]

Options:
    -SearchLog <SearchTerm>    Specify the search term (e.g., "Printer", "Bitlocker", "DiskIssues", "Authentication", "Microsoft365Apps", "Teams", "OneDrive", "ShutdownCrash").
    -Days <NumberOfDays>       Optional. Specify the number of days to look back. Default is 7 days.
    -Level <EventLevel>        Optional. Specify the event level (e.g., "Critical", "Warning", "Error", "Information"). Default is "All".
    -EnableKeywords            Optional. Enable filtering based on keywords.
    -Verbose                   Optional. Enable verbose output.
    -Debug                     Optional. Enable debug output to see input details.
    --help                     Display this help message.
"@


if ($help) {
    Write-Host $helpMessage
    exit
}

$WebFeedURL = "https://W0rmsp17.github.io/eventLogs/eventLogs.json"  

function Get-LogDetailsFromWeb {
    param (
        [string]$URL,
        [string]$SearchTerm
    )
    try {
        $response = Invoke-WebRequest -Uri $URL -UseBasicParsing
        if ($response.StatusCode -eq 200) {
            $logDetails = $response.Content | ConvertFrom-Json
            if ($logDetails.PSObject.Properties.Name -contains $SearchTerm) {
                return $logDetails.$SearchTerm
            } else {
                Write-Host "No details found for search term: $SearchTerm"
                return $null
            }
        } else {
            Write-Host "Failed to retrieve log details from web feed. Status Code: $($response.StatusCode)"
            return $null
        }
    } catch {
        Write-Host "Error accessing web feed: $_"
        return $null
    }
}


function Check-LogExists {
    param (
        [string]$LogName
    )
    try {
        Get-WinEvent -ListLog $LogName -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}


if ($Verbose) { Write-Host "Fetching log details from web feed..." }
$logDetails = Get-LogDetailsFromWeb -URL $WebFeedURL -SearchTerm $SearchLog

if (-not $logDetails) {
    Write-Host "No log details retrieved. Exiting script."
    exit
}

$Logs = $logDetails.logs
$EventIDs = $logDetails.eventIDs
$Keywords = $logDetails.keywords
$TimeWindow = $logDetails.timeWindow

if ($Verbose) { Write-Host "Log details retrieved successfully." }


$StartDate = (Get-Date).AddDays(-$Days) 
$EndDate = Get-Date


$LevelMapping = @{
    "Critical" = 1
    "Error" = 2
    "Warning" = 3
    "Information" = 4
    "Verbose" = 5
    "All" = 0
}

$LevelValue = $LevelMapping[$Level]

if ($Debug) {
    Write-Host "Debug Information:"
    Write-Host "SearchLog: $SearchLog"
    Write-Host "Days: $Days"
    Write-Host "Level: $Level ($LevelValue)"
    Write-Host "EnableKeywords: $EnableKeywords"
    Write-Host "Verbose: $Verbose"
    Write-Host "Logs: $Logs"
    Write-Host "EventIDs: $EventIDs"
    Write-Host "Keywords: $Keywords"
    Write-Host "TimeWindow: $TimeWindow"
}

$events = @()
foreach ($log in $Logs) {
    if (-not (Check-LogExists -LogName $log)) {
        Write-Host "Log $log does not exist on this computer."
        continue
    }
    if ($Verbose) { Write-Host "Checking log: $log" }
    try {
        $logEvents = Get-WinEvent -FilterHashtable @{
            LogName = $log;
            StartTime = $StartDate;
            EndTime = $EndDate
        } -ErrorAction Stop


        $filteredEvents = $logEvents | Where-Object { $_.Id -in $EventIDs -and ($LevelValue -eq 0 -or $_.Level -eq $LevelValue) }
        
        if ($EnableKeywords -and $Keywords) {
            $filteredEvents = $filteredEvents | Where-Object { $Keywords | ForEach-Object { $_ -match $_.Message } }
        }

        if ($filteredEvents) {
            $events += $filteredEvents
            if ($Verbose) { Write-Host "Found $($filteredEvents.Count) events in log: $log" }
        } else {
            if ($Verbose) { Write-Host "No matching events found in log: $log" }
        }
    } catch {
        Write-Host "No events found in log: $log"
    }
}


if ($events.Count -gt 0) {
    $relatedEvents = @()
    foreach ($event in $events) {
        $eventTime = $event.TimeCreated
        foreach ($log in $Logs) {
            if (-not (Check-LogExists -LogName $log)) {
                Write-Host "Log $log does not exist on this computer."
                continue
            }
            try {
                $relatedLogEvents = Get-WinEvent -FilterHashtable @{
                    LogName = $log
                    StartTime = $eventTime.AddMinutes(-$TimeWindow)
                    EndTime = $eventTime.AddMinutes($TimeWindow)
                } -ErrorAction Stop

                $relatedEvents += $relatedLogEvents | Where-Object { $_.LogName -ne $event.LogName }
            } catch {
                Write-Host "No related events found in log: $log"
            }
        }
    }

    $events += $relatedEvents | Select-Object -Unique
}

if ($events.Count -gt 0) {
    if ($Verbose) { Write-Host "Events found:" }
    $events | Select-Object TimeCreated, Id, LevelDisplayName, Message | Format-Table -AutoSize
} else {
    Write-Host "No events found for search term '$SearchLog' in the specified date range."
}


$ExcelFile = "C:\Temp\${SearchLog}IssuesLog.xlsx"


function Save-ToExcel {
    param (
        [string]$FilePath,
        [array]$Data
    )

    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    $Workbook = $Excel.Workbooks.Add()
    $Sheet = $Workbook.Worksheets.Item(1)

    $Headers = @("TimeCreated", "Id", "LevelDisplayName", "Message")
    $Sheet.Cells.Item(1, 1).Resize(1, $Headers.Length).Value = $Headers

    $row = 2
    foreach ($event in $Data) {
        $Sheet.Cells.Item($row, 1).Value = $event.TimeCreated
        $Sheet.Cells.Item($row, 2).Value = $event.Id
        $Sheet.Cells.Item($row, 3).Value = $event.LevelDisplayName
        $Sheet.Cells.Item($row, 4).Value = $event.Message
        $row++
    }

    $Workbook.SaveAs($FilePath)
    $Workbook.Close()
    $Excel.Quit()
}

if ($events.Count -gt 0) {
    Save-ToExcel -FilePath $ExcelFile -Data $events
    Write-Host "Print issue logs have been saved to $ExcelFile"
} else {
    Write-Host "No print issue logs to save."
}
