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
    $pathToTemplateWordDocument = "C:\.....\myWordTemplate.docx"
)

$Excel = New-Object -ComObject Excel.Application
$Word = New-Object –ComObject Word.Application

function OpenExcelBook ($FileName) {
    return $Excel.workbooks.open($Filename)
}

function CloseExcelBook ($Workbook) {
    #$Workbook.save()
    $Workbook.close()
}

function ReadCellData ($Workbook, $Cell) {
    $Worksheet = $Workbook.Activesheet
    return $Worksheet.Range($Cell).text
}

function SearchAWord ($Document, $findtext, $replacewithtext) {
  $FindReplace=$Document.ActiveWindow.Selection.Find
  $matchCase = $false;
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

function SaveAsWordDoc ($Document, $FileName) {
    $Document.Saveas([REF]$Filename)
    $Document.close()
}

function OpenWordDoc ($Filename) {
    return $Word.documents.open($Filename)
}

function OutPut () {
    $Workbook = OpenExcelBook –Filename $pathToDataExcelWorksheet

    $Row = 2

    Do {

        $Data = ReadCellData -Workbook $Workbook -Cell "A$Row"
        Write-Host $Data
        $FirstName = $Data


        If ($Data.length –ne 0) {
            $Doc = OpenWordDoc -Filename $pathToTemplateWordDocument
            SearchAWord -Document $Doc -findtext '{{{- FirstName -}}}' -replacewithtext $Data

            $Data = ReadCellData -Workbook $Workbook -Cell "B$Row"
            $LastName = $Data
            SearchAWord -Document $Doc -findtext '{{{- LastName -}}}' -replacewithtext $Data

            $SaveName="$pathToSaveWordDocuments\$LastName.docx"

            Write-Host $SaveName
            SaveAsWordDoc –document $Doc –Filename $SaveName

            $Row++
        }

    } while ($Data.length -ne 0)

    
    CloseExcelBook –workbook $Workbook
    return $cellData
}

OutPut