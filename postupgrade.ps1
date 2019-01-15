<#
.DESCRIPTION
    Removing Windows 10 Built-in from a list stored in CSV of XLSX file. And checks for AD collection.
.PARAMETER
    None
.EXAMPLE
    .\postupgrade.ps1
.NOTES
    Script name: postupgrade.ps1
    Version:     0.1
    Author:      Andrey Afanasyev
    Contact:     @clientmgmt
    DateCreated: Friday 11 january 2019
    LastUpdate:  Tuesday 15 january 2019
    #>


Write-Host " "
Write-Host "This script allow to uninstall Windows Built-In Application from XLSX or CSV file" -ForegroundColor Yellow
Write-Host " "

$ExcelFile = "RemovedApps.xlsx"
$CSVFile = "RemovedApps.csv"
$objExcel = $null
$ExcelFileExists = $null

# Check existence of Microsoft Excel application to create a new com object


$objExcel = New-Object -ComObject Excel.Application -ErrorAction SilentlyContinue 

$ExcelFileExists = Test-Path -Path $ExcelFile
$CSVFileExists = Test-Path -Path $CSVFile


Write-Host "Checking existance of a Microsoft Excel Application..." -ForegroundColor Yellow
if (!($objExcel)) {
    # module is not loaded
    Write-Host "Error: Microsoft Excel Application doesnot exist on this system!" -ForegroundColor Red
    
    Write-Host "Checking existance of a $CSVFile file..." -ForegroundColor Yellow
    if(!($CSVFileExists)) {
        Write-Host "Error: File $CSVFile does not exist! Please create this file. It contains one column. Each cell of column has a name of a Windows 10 Built-in app" -ForegroundColor Red
        break
    } else {
        Write-Host "File $CSVFile exist." -ForegroundColor Green
        $csv = Get-Content $CSVFile
        for ($i = 0; $i -lt $csv.Length; $i++){
            $AppName = [string]$csv[$i]
        # Check if names are correct. Delete app if yes.
            if(!(Get-AppxPackage *$AppName* | Select-Object Name)){
                Write-Host "< $AppName > The application name is not correct. Apparently, name spelled wrongly or App is not installed." -ForegroundColor Red
                Write-Host "Use command 'Get-AppxPackage *AppName* | Select Name' and name variations to find a correct name."
                Write-Host "Please fix RemovedApps.xlsx with the correct name." -ForegroundColor Yellow
            }else{
                Write-Host "Deleting: $AppName"
                #Get-AppxPackage *$AppName* | Remove-AppxPackage
            }
        }
    }

}else {
    Write-Host "Microsoft Excel Application exist on this system." -ForegroundColor Green

    Write-Host "Checking existance of a $ExcelFile file..." -ForegroundColor Yellow
    if(!($ExcelFileExists)) {
        # module is not loaded
        Write-Host "Error: File $ExcelFile does not exist!" -ForegroundColor Red
    }else {
        Write-Host "File $ExcelFile exist." -ForegroundColor Green

        # Look in Excel File
        $WorkBook = $objExcel.Workbooks.Open($(Resolve-Path -Path $ExcelFile))
        
        if (!($Workbook.Sheets.Count -eq 1)) {
            Write-Host "File $ExcelFile contains more than 1 sheet!" -ForegroundColor Red

        }else {
            Write-Host "File $ExcelFile contains exactly 1 sheet." -ForegroundColor Green
            $WorkBookSheet = $WorkBook.sheets.Item($Workbook.Sheets.Count)
            $intRowMax = ($WorkBookSheet.UsedRange.Rows).count
            # Count rows in a column
            $intRowMax = ($WorkBookSheet.UsedRange.Rows).count
            $intColMax = ($WorkBookSheet.UsedRange.column).count
            for ($intRow = 1; $intRow -le $intRowMax; $intRow++) {
                $AppName = $WorkBookSheet.Cells.Item($intRow, $intColMax).Value2
            # Check if names are correct. Delete app if yes.
                if(!(Get-AppxPackage *$AppName* | Select-Object Name)){
                    Write-Host "< $AppName > The application name is not correct. Apparently, name spelled wrongly or App is not installed." -ForegroundColor Red
                    Write-Host "Use command 'Get-AppxPackage *AppName* | Select Name' and name variations to find a correct name."
                    Write-Host "Please fix RemovedApps.xlsx with the correct name." -ForegroundColor Yellow
                }else{
                    Write-Host "Deleting: $AppName"
                    #Get-AppxPackage *$AppName* | Remove-AppxPackage
                }
            }
        }
    }
}

