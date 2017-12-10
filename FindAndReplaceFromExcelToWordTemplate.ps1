<#
 # Inspired by Scripting Guys blog
 # https://blogs.technet.microsoft.com/heyscriptingguy/2014/09/14/weekend-scripter-manipulating-word-and-excel-with-powershell/
#>

Param(

    [parameter(Mandatory=$true)]
    [string]
    $pathToSaveWordDocuments = "C:\.....\myDirectory\",

    [parameter(Mandatory=$true)]
    [string]
    $pathToDataExcelWorksheet = "C:\.....\myExcelWorksheet.xlsx",

    [parameter(Mandatory=$true)]
    [string]
    $pathToTemplateWordDocument = "C:\.....\myWordTemplate.docx",

    [parameter(Mandatory=$true)]
    [string]
    $columnAfind = "{{{ Word A}}}",

    [parameter(Mandatory=$true)]
    [string]
    $columnBfind = "{{{ Word B}}}"
)

function FindAndReplace ($Document, $findtext, $replacewithtext) {
  $FindReplace=$Document.ActiveWindow.Selection.Find
  $matchCase = $true;
  $matchWholeWord = $true;
  $matchWildCards = $false;
  $matchSoundsLike = $false;
  $matchAllWordForms = $false;
  $forward = $true;
  $format = $false;
  $matchKashida = $false;
  $matchDiacritics = $false;
  $matchAlefHamza = $false;
  $matchControl = $false;
  $read_only = $false;
  $visible = $true;
  $replace = 2;
  $wrap = 1;
  $FindReplace.Execute($findText, $matchCase, $matchWholeWord, $matchWildCards, `
                        $matchSoundsLike, $matchAllWordForms, $forward, $wrap, $format, `
                        $replaceWithText, $replace, $matchKashida ,$matchDiacritics, `
                        $matchAlefHamza, $matchControl)
}

$Excel = New-Object -ComObject Excel.Application
$Word = New-Object –ComObject Word.Application

$Workbook = $Excel.workbooks.open($pathToDataExcelWorksheet)

Do {
    $Row = 2
    $columnAData = $Workbook.Activesheet.Range("A$Row").text
    $columnBData = $Workbook.Activesheet.Range("B$Row").text
    
    $Doc = $Word.documents.open($pathToTemplateWordDocument)

    FindAndReplace -Document $Doc -findtext $columnAfind -replacewithtext $columnAData
    FindAndReplace -Document $Doc -findtext $columnBfind -replacewithtext $columnBData

    $filename = "$pathToSaveWordDocuments\$columnBData.docx"
    Write-Host "Saving file to: "
    $filename

    $Doc.saveas([REF]$Filename)
    $Doc.close()

    $Row++

} while (($columnAData.Length –ne 0) -and ($columnBData.Length -ne 0))

$Workbook.close()
$Word.quit()
$Excel.quit()