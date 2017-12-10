Param(

    [parameter(Mandatory=$true)]
    [string]
    $pathToWordDocuments = "C:\.....\myDirectory\",

    [parameter(Mandatory=$true)]
    [string]
    $wordToSearchFor = [ref]"2016",

    [parameter(Mandatory=$true)]
    [string]
    $wordToReplaceWith = [ref]"2017"
)

function FindAndReplace ($Document, $findtext, $replacewithtext) {
  $FindReplace = $Document.ActiveWindow.Selection.Find
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

$word = New-Object -ComObject Word.Application
$word.visible = $false
$files = Get-ChildItem -Path $pathToWordDocuments -Filter *.docx

for ($i = 0; $i -lt $files.Count; $i++) {
    $filename = $files[$i].FullName 
    $doc = $word.Documents.Open($filename)
    
    $isFound = FindAndReplace -Document $doc -findtext $wordToSearchFor -replacewithtext $wordToReplaceWith

    if ($isFound -eq $false) {
        Write-Host "Could not find search word in file :"
        $filename
    }

    $doc.Save()
    $doc.close()
}

$word.quit()
