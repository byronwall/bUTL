param([string]$macro)
$excel = new-object -comobject excel.application
$excelFile = Join-Path $PWD "build manager.xlsm"

$workbook = $excel.workbooks.open($excelFile)
$excel.Run($macro)
$workbook.close()

$excel.quit()