Import-Module ImportExcel

#TODO:Add Exception Handling -- Trought It's not really needed, since any exception is printed to the 
#   terminal and require manual fixing

$backupPath = $PSScriptRoot + "\Tiago\DB\backup"
$dailyFilesPath = $PSScriptRoot + "\Tiago\DB\Cognos\Carga Diaria D-2 (Simplificado).xlsx"
$dealerContactsPath = $PSScriptRoot + "\Tiago\QUERIES\Contatos dos Concessionarios.xlsx"
$dailyTrackerPath = $PSScriptRoot + "\Tiago\DB\Cognos\D-2 Tracker.xlsx"
$requestsMadePath = $PSScriptRoot + "\Tiago\DB\CargaDiariaV1\Monitorias Realizadas.xlsx"
$dailyReportPath = $PSScriptRoot + "\Relatórios\Relatório Carga Diária\Relatorio Carga Diaria.xlsx"
$proactivityTrackerPath = $PSScriptRoot + "\Tiago\DB\CargaDiariaV1\proatividade.xlsx"
$monitoredDealersPath = $PSScriptRoot + "\Tiago\DB\CargaDiariaV1\Dealers Monitorados.xlsx"
$allPresentFilesPath = $PSScriptRoot + "\Tiago\DB\Cognos\Carga Arquivos Presentes (Simplificado).xlsx"
$allPresentFilesBackupPath = $PSScriptRoot + "\Tiago\DB\Cognos\Carga Arquivos Presentes\currentFiles.xlsx"
$allPresentFilesBackupPathOlder = $PSScriptRoot + "\Tiago\DB\Cognos\Carga Arquivos Presentes\currentFilesOlder.xlsx"

#Haven't yet figured out how to filter a Import-Excel object based on another Import-Excel object 
#Like so: $filteredDealerList = $dealerList | Where-Object {$_.Type -in $typesList}
#So, I'm hardcoding the dealer types:
$monitoredTypes = @{
    PS_1 = ("CONC", "FCOM", "SRSV", "UNSV", "DPEC")
    RO_2 = ("CONC", "FCOM", "SRSV", "UNSV")
}

$dailySheetsNames = @("PS_1", "RO_2")

function Refresh-Connections {
    param (
        [Parameter(Mandatory, Position=0,
        HelpMessage="Path to the Excel Workbook")]
        [string]$path,

        [Parameter(Position = 1,
        HelpMessage="Should the workbook be kept openned")]
        [bool]$keepOpen = $False
    )
    
    $excelApp = New-Object -ComObject Excel.Application
    $workbook = $excelApp.Workbooks.Open($path)
    $excelApp.Visible = $keepOpen
    $connections = $workbook.Connections
    $workbook.RefreshAll()
    while ($connections | ForEach-Object {if($_.OLEDBConnection.Refreshing){$True}}) {
        Start-Sleep -Seconds 1    
    }

    $workbook.DisplayAlerts = $false;
    $workbook.Save()
    $workbook.DisplayAlerts = $True
    if (!$keepOpen) {
        $workbook.Close()
    }

    $excelApp.Visible = $True
    $excelApp.Quit()
}

function Backup(){
    param (
        [Parameter(Mandatory, Position=0,
        HelpMessage="Path to the Excel Workbook")]
        [string]$path,

        [Parameter(Position=1,
        HelpMessage="Delete the old file")]
        [bool]$delete=$true
    )


    if ($delete) {
        Move-Item -Path $path -Destination $backupPath -Force
    } else {
        Copy-Item -Path $path -Destination $backupPath
    }
    
}

function Get-Dates(){
    param (
        [Parameter(Mandatory, Position = 0)]
        [string]$workbookPath
    )

    $dates = Import-Excel -Path $workbookPath | Select-Object "Date" | Sort-Object -Unique

    return $dates
}

function Get-DailyProactivity(){
    $dailyProactivity = Import-Excel `
        -Path $dailyReportPath -WorksheetName "effectivityBackend" -StartRow 2 -EndRow 2 -StartColumn 4 -EndColumn 4 -NoHeader

    return $dailyProactivity.P1
}

function Set-DailyProactivity(){
    param (
        [Parameter(Mandatory, Position = 0)]
        [double]$value
    )

    $header = "Dia", "%"
    $today = Get-Date -Format "yyyy-MM-dd"
    $proactivity = ConvertFrom-Csv -InputObject @("$today,$value") -Header $header

    Export-Excel -Path $proactivityTrackerPath -InputObject $proactivity -WorksheetName "Sheet1" -Append

}

function Update-DailyTracker(){
    Backup $dailyTrackerPath $false

    foreach ($name in $dailySheetsNames) {
        $data = Import-Excel $dailyFilesPath -WorksheetName $name     
        Export-Excel -Path $dailyTrackerPath -InputObject $data -WorksheetName $name -Append
    }

    Backup $dailyFilesPath
}

function Update-RequestsMade(){
    param (
        [Parameter(Mandatory, Position=0,
        HelpMessage="Date to update in format 'yyyy-mm-dd'")]
        [string]$date
    )

    $dealerList = Import-Excel $monitoredDealersPath 
    $sheetName = "Monitorias Realizadas"

    Backup $requestsMadePath $false

    foreach($key in $monitoredTypes.Keys){
        $filteredDealerList = $dealerList | Where-Object {$_.Type -in $monitoredTypes.Item($key)}
        $filteredDealerList = $filteredDealerList | Select-Object "Gm Code"
        
        $data = @("Date,Type,Code")

        foreach($dealer in $filteredDealerList) {
            $fileType = $key.Substring(0,2)
            $code = $dealer.'GM Code'

            $data += "$date,$fileType,$code"
        }

        $data = ConvertFrom-Csv $data
        

        Export-Excel -Path $requestsMadePath -InputObject $data -Append -WorksheetName $sheetName  
    } 
}


function Update-DailyReport(){
    Move-Item -Path $allPresentFilesBackupPath $allPresentFilesBackupPathOlder -Force
    Move-Item -Path $allPresentFilesPath -Destination $allPresentFilesBackupPath -Force

    $actualProactivity = Get-DailyProactivity
    Set-DailyProactivity ($actualProactivity)
    Refresh-Connections ($dailyReportPath)

    
}

function main(){
    Write-Progress -Activity "Running Script" -Status "Updating dealer contacts" -PercentComplete 10

    Refresh-Connections $dealerContactsPath
    Write-Output "Done updating dealer contacts"
    
    Write-Progress -Activity "Running Script" -Status "Getting dates to update" -PercentComplete 30
    $dates = Get-Dates $dailyFilesPath

    Write-Progress -Activity "Running Script" -Status "Updating daily tracker" -PercentComplete 50
    Update-DailyTracker $date 
    Write-Output "Done updating daily tracker"

    Write-Progress -Activity "Running Script" -Status "Updating requests made" -PercentComplete 70
    foreach($date in $dates) {
        Write-Output "Adding date $date"
        Update-RequestsMade $date.Date
    }
    Write-Output "Done updating requests made"

    $updateReportAnswer = Read-Host "Update the report as well? (y/N)"
    if ($updateReportAnswer == "y") {
        Write-Progress -Activity "Running Script" -Status "Updating the report" -PercentComplete 80
        Update-DailyReport
    }

    Write-Progress -Activity "Running Script" -Completed
    Write-Output "All done!"
}


main