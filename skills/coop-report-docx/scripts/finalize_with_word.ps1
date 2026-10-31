# Opens a .docx in Microsoft Word (Windows only), refreshes the table of contents
# and every other field, saves it, and optionally exports a PDF for checking.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File finalize_with_word.ps1 -Docx report.docx [-Pdf report.pdf]
#
# If Word opens the file read-only (e.g. Office is not activated), fields cannot be
# refreshed; the script says so and still exports the PDF. The .docx is built with
# "update fields on open", so Word asks the student to refresh the TOC on first open.
param(
    [Parameter(Mandatory = $true)][string]$Docx,
    [string]$Pdf
)
$ErrorActionPreference = 'Stop'
$docxPath = (Resolve-Path $Docx).Path
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open($docxPath, $false, $false)
    if ($doc.ReadOnly) {
        Write-Output "warning: Word opened the file read-only (Office may not be activated) - TOC not refreshed"
    }
    else {
        # Two passes: the first fills the TOCs, which shifts pages; the second fixes page numbers.
        for ($pass = 0; $pass -lt 2; $pass++) {
            foreach ($toc in $doc.TablesOfContents) { $toc.Update() }
            $doc.Fields.Update() | Out-Null
        }
        $doc.Save()
        Write-Output "updated: $docxPath"
    }
    if ($Pdf) {
        $pdfPath = [System.IO.Path]::GetFullPath($Pdf)
        $doc.ExportAsFixedFormat($pdfPath, 17)  # 17 = wdExportFormatPDF
        Write-Output "pdf: $pdfPath"
    }
    $doc.Close($false)
}
finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
