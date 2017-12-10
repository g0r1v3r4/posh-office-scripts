<#
 # Inspired by Scripting Guys blog
 # https://blogs.technet.microsoft.com/heyscriptingguy/2014/09/14/weekend-scripter-manipulating-word-and-excel-with-powershell/

 # Usage:

 .\CreateWordDocumentsFromWordTemplateAndExcelData.ps1 -pathToSaveWordDocuments C:\...\directory `
    -pathToDataExcelWorksheet C:\...\myExcelWorksheet.xlsx `
    -pathToTemplateWordDocument C:\...\myWordTemplate.docx `
    -columnAfind "{{{ Search This }}}" `
    -columnBfind "{{{ Word B Find }}}"
#>

Param(

    [parameter(Mandatory=$true)]
    [string]
    $pathToSaveWordDocuments = "C:\...\directory",

    [parameter(Mandatory=$true)]
    [string]
    $pathToDataExcelWorksheet = "C:\...\myExcelWorksheet.xlsx",

    [parameter(Mandatory=$true)]
    [string]
    $pathToTemplateWordDocument = "C:\...\myWordTemplate.docx",

    [parameter(Mandatory=$true)]
    [string]
    $columnAfind = "{{{ Search This }}}",

    [parameter(Mandatory=$true)]
    [string]
    $columnBfind = "{{{ Search That }}}"
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
$Word.Visible = $false
$Workbook = $Excel.workbooks.open($pathToDataExcelWorksheet)

# Assuming there is a "header" row
$Row = 2

Do {
    # Data in column A can be a 'complex' string
    # 'complex' means that there can be multiple values in one cell
    # values will be delimited by ';'
    $columnAData = $Workbook.Activesheet.Range("A$Row").text
    
    # Puts a newline and carriage return by replacing the ";" semi-colon
    # characters in the read string from "A$Row"
    $editedString = $columnAData.replace(";", "`r`n")

    # Data in column B is guaranteed to be a 'simple' string
    # 'simple' means a qualified file name
    $columnBData = $Workbook.Activesheet.Range("B$Row").text
  
    if($columnBData.length -ne 0){   
        $Doc = $Word.documents.open($pathToTemplateWordDocument)

        FindAndReplace -Document $Doc -findtext $columnAfind -replacewithtext $editedString
   
        FindAndReplace -Document $Doc -findtext $columnBfind -replacewithtext $columnBData

        $filename = "$pathToSaveWordDocuments\$columnBData.docx"
        Write-Host "Saving file to: "
        $filename

        $Doc.saveas([REF]$Filename)
        $Doc.close()
    }

    $Row++

} while (($columnAData.Length –ne 0) -and ($columnBData.Length -ne 0))

$Workbook.close()
$Word.quit()
$Excel.quit()