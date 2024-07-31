param (
    [string]$SearchLog,
    [int]$Days = 7,  
    [switch]$Verbose
)


$helpMessage = @"
Use: .\printSweeper.ps1 -SearchLog <SearchTerm> [-Days <NumberOfDays>] [-Verbose] [--help]
exmaple: .\printSweeper.ps1 -SearchLog "Printer" -Days 3 -Verbose

Options:
    -SearchLog <SearchTerm>    Specify the search term (e.g., "Printer", "Bitlocker").
    -Days <NumberOfDays>       Optional. Specify the number of days to look back. Default is 7 days.
    -Verbose                   Optional. Enable verbose output.
    --help                     Display this help message.
"@


if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey("help")) {
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
            if ($logDetails.ContainsKey($SearchTerm)) {
                return $logDetails[$SearchTerm]
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


if ($Verbose) { Write-Host "Fetching log details from web feed..." }
$logDetails = Get-LogDetailsFromWeb -URL $WebFeedURL -SearchTerm $SearchLog

if (-not $logDetails) {
    Write-Host "No log details retrieved. Exiting script."
    exit
}

$Logs = $logDetails.logs
$EventIDs = $logDetails.eventIDs

if ($Verbose) { Write-Host "Log details retrieved successfully." }


$StartDate = (Get-Date).AddDays(-$Days)
$EndDate = Get-Date

$events = @()
foreach ($log in $Logs) {
    if ($Verbose) { Write-Host "Checking log: $log" }
    try {
        $logEvents = Get-WinEvent -FilterHashtable @{
            LogName = $log;
            StartTime = $StartDate;
            EndTime = $EndDate
        } | Where-Object { $_.Id -in $EventIDs }
        if ($logEvents) {
            $events += $logEvents
        } else {
            if ($Verbose) { Write-Host "No events found in log: $log" }
        }
    } catch {
        Write-Host "Error retrieving events from log: $log. $_"
    }
}


if ($events) {
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
    [void]$Sheet.Cells.Item(1, 1).Resize(1, $Headers.Length).Value = $Headers

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

if ($events) {
    Save-ToExcel -FilePath $ExcelFile -Data $events
    Write-Host "Print issue logs have been saved to $ExcelFile"
} else {
    Write-Host "No print issue logs to save."
}
