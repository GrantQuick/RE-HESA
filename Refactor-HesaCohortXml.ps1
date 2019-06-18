$config = Get-Content ".\config.json" | ConvertFrom-Json

$cohortFile = ($config | Select-Object -Property "cohortFileForXmlRefactor").cohortFileForXmlRefactor
$outputFile = ($config | Select-Object -Property "destinationForRefactoredCohortFile").destinationForRefactoredCohortFile

[xml]$cohortXML = Get-Content -Path $cohortFile

$rows = $cohortXML.GOREPORT.GOROWS.GOROW

$concise = $rows | Select-Object OWNSTU,HUSID,FNAMES,SURNAME

# Outputting to Excel file requires the ImportExcel Module available here: https://github.com/dfinke/ImportExcel
$concise | Export-Excel -Autosize -NoNumberConversion * -Path $outputFile

# If you cannot export directly to Excel, then export to CSV
# $concise | Export-Csv -NoClobber -NoTypeInformation -Path $outputFile