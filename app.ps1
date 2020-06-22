# param (
#     [Parameter(Mandatory=$true)][string]$excelFile = "appConvertToJSON2.xlsm",
#     [Parameter(Mandatory=$true)][string]$macro = "ConsoMain.Batch",
#     [Parameter(Mandatory=$true)][string]$formName = "TRANSFORM_RECORD",
#     [Parameter(Mandatory=$true)][boolean]$moveOn = $false
# )

$excelFile = "appConvertToJSON2.xlsm"
$macro = "ConsoMain.Batch"
$formName = "TRANSFORM_RECORD"
$moveOn = $false
##
$curFolder = pwd 
$fullpath = Join-Path $curFolder.Path $excelFile
$excel = new-object -comobject excel.application
$excel.Visible = $false
$workbook = $excel.workbooks.open($fullpath)
$excel.Run($macro, $formName, $moveOn)
$workbook.close()
