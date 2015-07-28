param([string]$macro)

$confirm = Read-Host "This will run $macro, overwriting local changes. Enter 'y' to proceed"

if ($confirm -eq 'y'){

$excel = new-object -comobject excel.application
$excelFile = Join-Path $PWD "build manager.xlsm"

$workbook = $excel.workbooks.open($excelFile)
$excel.Run($macro)
$workbook.close()

$excel.quit()
}