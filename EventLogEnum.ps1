param (
    [string]$OutputFilePath = "C:\Temp\EventLogsCollection.json", 
    [switch]$Verbose,
    [switch]$help
)


$helpMessage = @"
Usage: .\enumerateLogs.ps1 [-OutputFilePath <Path>] [-Verbose] [--help]

Options:
    -OutputFilePath <Path>     Optional. Specify the output file path. Default is C:\Temp\EventLogsCollection.json.
    -Verbose                   Optional. Enable verbose output.
    --help                     Display this help message.
"@


if ($help) {
    Write-Host $helpMessage
    exit
}
function Get-ErrorEvents {
    param (
        [string]$LogPath
    )

    $errorEvents = @()
    try {
        $events = Get-WinEvent -LogName $LogPath -ErrorAction Stop
        $errorEvents = $events | Where-Object { $_.LevelDisplayName -eq "Error" } | Select-Object -ExpandProperty Id -Unique
    } catch {
        if ($Verbose) {
            Write-Host "Error retrieving events from log: $LogPath"
        }
    }

    return $errorEvents
}


$logPaths = Get-WinEvent -ListLog * | Where-Object { $_.LogName -like "Microsoft-Windows-*/Admin" -or $_.LogName -like "Microsoft-Windows-*/Operational" } | Select-Object -ExpandProperty LogName

$logsCollection = @{}

foreach ($logPath in $logPaths) {
    if ($Verbose) { Write-Host "Processing log: $logPath" }
    $errorEvents = Get-ErrorEvents -LogPath $logPath
    if ($errorEvents.Count -gt 0) {
        $logsCollection[$logPath] = $errorEvents
    }
}


$logsCollection | ConvertTo-Json -Depth 4 | Out-File -FilePath $OutputFilePath -Force

Write-Host "Logs collection saved to $OutputFilePath"
