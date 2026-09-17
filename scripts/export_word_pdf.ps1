param(
    [Parameter(Mandatory=$true)][string]$InputDocx,
    [Parameter(Mandatory=$true)][string]$OutputPdf
)
$ErrorActionPreference = 'Stop'
$source = (Resolve-Path -LiteralPath $InputDocx).Path
$destination = [System.IO.Path]::GetFullPath($OutputPdf)
if ([System.IO.Path]::GetExtension($source) -ne '.docx') { throw 'Input must be DOCX.' }
if ([System.IO.Path]::GetExtension($destination) -ne '.pdf') { throw 'Output must be PDF.' }
if (Test-Path -LiteralPath $destination) { throw 'Output already exists; choose a new PDF path.' }
[System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($destination)) | Out-Null
$word = $null
$document = $null
$owned = $false
$previous = $null
try {
    $word = New-Object -ComObject Word.Application
    # Never close or alter a pre-existing user session if COM returns one.
    if ($word.Documents.Count -ne 0) { throw 'Word instance contains existing documents; leaving it untouched.' }
    $owned = $true
    $previous = @($word.AutomationSecurity, $word.Options.UpdateLinksAtOpen, $word.Options.UpdateFieldsAtPrint, $word.Options.UpdateLinksAtPrint)
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3
    $word.Options.UpdateLinksAtOpen = $false
    $word.Options.UpdateFieldsAtPrint = $false
    $word.Options.UpdateLinksAtPrint = $false
    $document = $word.Documents.Open($source, $false, $true, $false)
    $document.Repaginate()
    $pages = $document.ComputeStatistics(2)
    # wdExportFormatPDF = 17. Do not save the read-only source.
    $document.ExportAsFixedFormat($destination, 17)
    if (-not (Test-Path -LiteralPath $destination)) { throw 'Word did not produce the PDF.' }
    [pscustomobject]@{engine='Microsoft Word'; version=$word.Version; pages=$pages; output=$destination} | ConvertTo-Json
} finally {
    if ($null -ne $document) {
        $document.Close([ref]0)
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
    }
    if ($null -ne $word) {
        if ($null -ne $previous) {
            $word.AutomationSecurity = $previous[0]
            $word.Options.UpdateLinksAtOpen = $previous[1]
            $word.Options.UpdateFieldsAtPrint = $previous[2]
            $word.Options.UpdateLinksAtPrint = $previous[3]
        }
        if ($owned) { $word.Quit([ref]0) }
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
    }
}
