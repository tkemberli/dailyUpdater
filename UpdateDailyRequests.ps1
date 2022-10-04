Import-Module ImportExcel

$mainPath = ""
$backupPath = "D:\Downloads\Atualizador\BackupPath"
$dailyFilesPath = "D:\Downloads\Atualizador\Carga Diaria D-2 (Simplificado).xlsx"
$dealerContactsPath = "D:\Downloads\Atualizador\Sheets\QUERIES\Contatos dos Concessionarios.xlsx"
$dailyTrackerPath = "D:\Downloads\Atualizador\D-2 Tracker.xlsx"
$requestsMadePath = "D:\Downloads\Atualizador\Monitorias Realizadas.xlsx"
$dailyReportPath = "D:\Downloads\Atualizador\Relatorio Carga Diaria.xlsx"
$proactivityTrackerPath = "D:\Downloads\Atualizador\proactivity.xlsx"
$monitoredDealersPath = "D:\Downloads\Atualizador\Dealers Monitorados.xlsx"

#Haven't figured out how to filter a Import-Excel object based on another Import-Excel object yet
#Like so: $filteredDealerList = $dealerList | Where-Object {$_.Type -in $typesList}
#Hardcoding the dealer types:
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

    $workbook.Save()
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
    $actualProactivity = Get-DailyProactivity
    Set-DailyProactivity ($actualProactivity)
}

Write-Output "Updating dealer contacts..."
#Refresh-Connections $dealerContactsPath
Write-Output "Done"

$dates = Get-Dates $dailyFilesPath

Write-Output "Updating daily tracker..."
Update-DailyTracker $date 
Write-Output "Done"

Write-Output "Updating requests made..."
foreach($date in $dates) {
    Write-Output "Adding date $date"
    Update-RequestsMade $date.Date
}
Write-Output "Done"