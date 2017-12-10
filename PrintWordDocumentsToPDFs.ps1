Param(
    #Path to where the documents are saved
    [parameter(Mandatory=$True)]
    [string]
    $DocumentsPath = 'C:\....\Folder'
    )

$Word = New-Object -ComObject Word.Application
$Word.Visible = $False

# This filter will find .doc as well as .docx documents
Get-ChildItem -Path $DocumentsPath -Filter *.doc? | ForEach-Object {

    $Document = $Word.Documents.Open($_.FullName)

    $PdfFilename = "$($_.DirectoryName)\$($_.BaseName).pdf"

    $Word.ActiveWindow.View.ShowRevisionsAndComments = $False
    $Word.ActiveWindow.View.RevisionsView = 0

    $Word.Options.WarnBeforeSavingPrintingSendingMarkup = $false

    $Document.SaveAs([ref] $PdfFilename, [ref] 17)

    $Document.Close()
}

$Word.Quit()