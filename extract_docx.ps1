$base = "C:\Users\Joaquin.Suazo\Documents\SGP-Produccion\Documentos\"
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-DocxText {
    param($path)
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
        $entry = $zip.Entries | Where-Object { $_.FullName -eq "word/document.xml" }
        if (-not $entry) { $zip.Dispose(); return "NO ENCONTRADO" }
        $stream = $entry.Open()
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
        $xml = $reader.ReadToEnd()
        $reader.Close()
        $zip.Dispose()
        $doc = [xml]$xml
        $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
        $ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
        $paras = $doc.SelectNodes("//w:p", $ns)
        $lines = @()
        foreach ($p in $paras) {
            $runs = $p.SelectNodes(".//w:t", $ns)
            $line = ""
            foreach ($r in $runs) { $line += $r.InnerText }
            if ($line.Trim() -ne "") { $lines += $line }
        }
        return $lines -join "`n"
    } catch { return "ERROR: $_" }
}

$files = Get-ChildItem $base -Filter "*.docx" | Sort-Object Name
foreach ($f in $files) {
    Write-Host "=== $($f.Name) ==="
    $text = Get-DocxText $f.FullName
    $preview = $text.Substring(0, [Math]::Min(2000, $text.Length))
    Write-Host $preview
    Write-Host ""
}
