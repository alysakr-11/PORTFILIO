$src = "C:\Users\alysa\Downloads\WEB\Aly_Sakr_CV.docx"
$tmp = "C:\Users\alysa\AppData\Local\Temp\cv_edit"
Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Path $tmp -Force | Out-Null

# Rename .docx to .zip and extract
$zipPath = $src -replace '\.docx$','.zip'
Copy-Item -LiteralPath $src -Destination $zipPath -Force
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $tmp)

# Read document.xml
$xmlPath = Join-Path $tmp "word\document.xml"
[xml]$xml = Get-Content -Path $xmlPath -Raw
$ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
$ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')

# Find paragraph with github link at end of contact line
$paras = $xml.SelectNodes('//w:p', $ns)
$target = $null
foreach($p in $paras){
    $txt = ($p.SelectNodes('.//w:t', $ns) | ForEach-Object { $_.InnerText }) -join ''
    if($txt -match 'github\.com/alysakr-11\s*$'){ $target = $p; break }
}

if(-not $target){ Write-Output "ERROR: target not found"; exit 1 }

$parent = $target.ParentNode
$nsW = $ns.LookupNamespace('w')

# Create new paragraph
$newP = $xml.CreateElement('w:p', $nsW)
$pPr = $xml.CreateElement('w:pPr', $nsW)
$newP.AppendChild($pPr) | Out-Null

# Copy pPr from target to keep formatting
$targetPPr = $target.SelectSingleNode('w:pPr', $ns)
if($targetPPr){ $pPr.InnerXml = $targetPPr.InnerXml }

# Create run with text
$r = $xml.CreateElement('w:r', $nsW)
$rPr = $xml.CreateElement('w:rPr', $nsW)
$rFonts = $xml.CreateElement('w:rFonts', $nsW)
$rFonts.SetAttribute('ascii', 'Calibri')
$rFonts.SetAttribute('hAnsi', 'Calibri')
$rPr.AppendChild($rFonts) | Out-Null
$sz = $xml.CreateElement('w:sz', $nsW)
$sz.SetAttribute('val', '20')
$rPr.AppendChild($sz) | Out-Null
$r.AppendChild($rPr) | Out-Null
$t = $xml.CreateElement('w:t', $nsW)
$t.InnerText = "   Portfolio: https://alysakr-11.github.io/NEW-REPO-/"
$r.AppendChild($t) | Out-Null
$newP.AppendChild($r) | Out-Null

# Insert after the target paragraph
$parent.InsertAfter($newP, $target) | Out-Null

# Save
$xml.Save($xmlPath)
Write-Output "XML updated OK"

# Rebuild .docx
Remove-Item -LiteralPath $zipPath -Force -ErrorAction SilentlyContinue
[System.IO.Compression.ZipFile]::CreateFromDirectory($tmp, $zipPath, [System.IO.Compression.CompressionLevel]::Optimal, $false)
Copy-Item -LiteralPath $zipPath -Destination $src -Force
Remove-Item -LiteralPath $zipPath -Force -ErrorAction SilentlyContinue
Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
Write-Output "CV updated successfully"
