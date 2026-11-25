# Run in Windows PowerShell with desktop Microsoft Word installed.
param([Parameter(Mandatory=$true)][string]$Document)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$tempFile = Join-Path ([IO.Path]::GetTempPath()) ("coursework-" + [guid]::NewGuid() + ".docx")
Copy-Item -LiteralPath $Document -Destination $tempFile
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open($tempFile)
    $toc = $doc.TablesOfContents.Item(1)
    $toc.Update()
    $doc.Repaginate()
    $toc.UpdatePageNumbers()
    $doc.Repaginate()
    $doc.RemoveDocumentInformation(4) # personal information
    $doc.Save()
    $doc.Close(0)
    $doc = $word.Documents.Open($tempFile, $false, $true)
    $doc.Repaginate()
    "Word $($word.Version), pages $($doc.ComputeStatistics(2)), TOCs $($doc.TablesOfContents.Count)"
    foreach ($p in $doc.Paragraphs) {
        if ($p.OutlineLevel -le 2) {
            $r = $p.Range.Duplicate
            $r.Collapse(1)
            "Page $($r.Information(3)) | $($p.Range.Text.Trim())"
        }
    }
    foreach ($t in $doc.Tables) {
        $r = $t.Range.Duplicate
        $r.Collapse(1)
        $start = $r.Information(3)
        $r = $t.Range.Duplicate
        $r.End--
        $r.Collapse(0)
        "Table pages $start-$($r.Information(3)), rows $($t.Rows.Count)"
    }
    $doc.Close(0)
    $doc = $null
    Copy-Item -LiteralPath $tempFile -Destination $Document -Force
} finally {
    if ($doc) { $doc.Close(0) }
    if ($word) { $word.Quit() }
    Remove-Item -LiteralPath $tempFile
}
