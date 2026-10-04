# ==============================================================================
# convert_md_to_docx.ps1 - Native PowerShell Markdown to Word (.docx) Converter
# Zero external dependencies required - uses .NET System.IO.Compression
# Repository: https://github.com/StephSMITH-hub/theos
# ==============================================================================

param(
    [Parameter(Position=0)]
    [string]$InputPath = "",
    [Parameter(Position=1)]
    [string]$OutputPath = "",
    [switch]$All,
    [switch]$Recursive,
    [string]$Directory = "",
    [string]$FontFamily = "Calibri"
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$RepoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -Path $RepoRoot

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Escape-XmlText([string]$text) {
    if ([string]::IsNullOrEmpty($text)) { return "" }
    return $text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace('"', "&quot;").Replace("'", "&apos;")
}

function Parse-InlineMarkdown([string]$text) {
    if ([string]::IsNullOrEmpty($text)) { return @() }
    
    # Remove transcript escape characters
    $text = $text.Replace('\"', '"').Replace("\'", "'")
    
    $tokens = @()
    $pattern = '(`(?<code_val>[^`]+)`)' +
               '|(\*\*\*(?<bi_val>[^*]+)\*\*\*)' +
               '|(___(?<bi2_val>[^_]+)___)' +
               '|(\*\*(?<b_val>[^*]+)\*\*)' +
               '|(__(?<b2_val>[^_]+)__)' +
               '|(\*(?<i_val>[^*]+)\*)' +
               '|(_(?<i2_val>[^_]+)_)' +
               '|(\[(?<link_txt>[^\]]+)\]\((?<link_url>[^\)]+)\))'

    $regex = New-Object System.Text.RegularExpressions.Regex($pattern)
    $lastIdx = 0
    $matches = $regex.Matches($text)

    foreach ($m in $matches) {
        if ($m.Index -gt $lastIdx) {
            $tokens += [PSCustomObject]@{
                Text = $text.Substring($lastIdx, $m.Index - $lastIdx)
                Bold = $false; Italic = $false; Code = $false; Link = $null
            }
        }
        if ($m.Groups['code_val'].Success) {
            $tokens += [PSCustomObject]@{ Text = $m.Groups['code_val'].Value; Bold = $false; Italic = $false; Code = $true; Link = $null }
        } elseif ($m.Groups['bi_val'].Success -or $m.Groups['bi2_val'].Success) {
            $val = if ($m.Groups['bi_val'].Success) { $m.Groups['bi_val'].Value } else { $m.Groups['bi2_val'].Value }
            $tokens += [PSCustomObject]@{ Text = $val; Bold = $true; Italic = $true; Code = $false; Link = $null }
        } elseif ($m.Groups['b_val'].Success -or $m.Groups['b2_val'].Success) {
            $val = if ($m.Groups['b_val'].Success) { $m.Groups['b_val'].Value } else { $m.Groups['b2_val'].Value }
            $tokens += [PSCustomObject]@{ Text = $val; Bold = $true; Italic = $false; Code = $false; Link = $null }
        } elseif ($m.Groups['i_val'].Success -or $m.Groups['i2_val'].Success) {
            $val = if ($m.Groups['i_val'].Success) { $m.Groups['i_val'].Value } else { $m.Groups['i2_val'].Value }
            $tokens += [PSCustomObject]@{ Text = $val; Bold = $false; Italic = $true; Code = $false; Link = $null }
        } elseif ($m.Groups['link_txt'].Success) {
            $tokens += [PSCustomObject]@{ Text = $m.Groups['link_txt'].Value; Bold = $false; Italic = $false; Code = $false; Link = $m.Groups['link_url'].Value }
        }
        $lastIdx = $m.Index + $m.Length
    }

    if ($lastIdx -lt $text.Length) {
        $tokens += [PSCustomObject]@{
            Text = $text.Substring($lastIdx)
            Bold = $false; Italic = $false; Code = $false; Link = $null
        }
    }
    return $tokens
}

function Build-RunsXml($tokens, [string]$defaultFont = "Calibri", [string]$defaultColor = "222222", [int]$szVal = 22) {
    $runsXml = New-Object System.Text.StringBuilder
    foreach ($t in $tokens) {
        $escaped = Escape-XmlText $t.Text
        $bTag = if ($t.Bold) { "<w:b/>" } else { "" }
        $iTag = if ($t.Italic) { "<w:i/>" } else { "" }
        $fontName = if ($t.Code) { "Consolas" } else { $defaultFont }
        $color = if ($t.Code) { "990000" } elseif ($t.Link) { "1B365D" } else { $defaultColor }
        [void]$runsXml.Append("<w:r><w:rPr><w:rFonts w:ascii=""$fontName"" w:hAnsi=""$fontName""/>$bTag$iTag<w:sz w:val=""$szVal""/><w:color w:val=""$color""/></w:rPr><w:t xml:space=""preserve"">$escaped</w:t></w:r>")
    }
    return $runsXml.ToString()
}

function Convert-SingleMdToDocx([string]$mdFile, [string]$outFile) {
    if (-not (Test-Path $mdFile)) {
        Write-Host "Error: Input file not found: $mdFile" -ForegroundColor Red
        return $false
    }
    if ([string]::IsNullOrWhiteSpace($outFile)) {
        $outFile = [System.IO.Path]::ChangeExtension($mdFile, ".docx")
    }

    $outDir = Split-Path -Parent $outFile
    if ($outDir -and -not (Test-Path $outDir)) {
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    }

    $fileName = Split-Path -Leaf $mdFile
    $outLeaf = Split-Path -Leaf $outFile
    Write-Host "[*] Converting: $fileName -> $outLeaf..." -ForegroundColor Yellow

    $content = Get-Content -Path $mdFile -Encoding UTF8
    $bodyXml = New-Object System.Text.StringBuilder

    foreach ($rawLine in $content) {
        $line = $rawLine.Trim()
        if ([string]::IsNullOrWhiteSpace($line)) { continue }

        # Heading 1
        if ($line.StartsWith("# ")) {
            $hText = Escape-XmlText ($line.Substring(2).Trim())
            [void]$bodyXml.AppendLine(@"
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/><w:spacing w:before="260" w:after="120"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="$FontFamily" w:hAnsi="$FontFamily"/><w:b/><w:sz w:val="36"/><w:color w:val="1F497D"/></w:rPr><w:t>$hText</w:t></w:r>
    </w:p>
"@)
        }
        # Heading 2
        elseif ($line.StartsWith("## ")) {
            $hText = Escape-XmlText ($line.Substring(3).Trim())
            [void]$bodyXml.AppendLine(@"
    <w:p>
      <w:pPr><w:pStyle w:val="Heading2"/><w:spacing w:before="220" w:after="100"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="$FontFamily" w:hAnsi="$FontFamily"/><w:b/><w:sz w:val="28"/><w:color w:val="2C5E8A"/></w:rPr><w:t>$hText</w:t></w:r>
    </w:p>
"@)
        }
        # Heading 3
        elseif ($line.StartsWith("### ")) {
            $hText = Escape-XmlText ($line.Substring(4).Trim())
            [void]$bodyXml.AppendLine(@"
    <w:p>
      <w:pPr><w:pStyle w:val="Heading3"/><w:spacing w:before="180" w:after="80"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="$FontFamily" w:hAnsi="$FontFamily"/><w:b/><w:sz w:val="25"/><w:color w:val="3B6998"/></w:rPr><w:t>$hText</w:t></w:r>
    </w:p>
"@)
        }
        # Blockquote / Scripture (> quote)
        elseif ($line.StartsWith(">")) {
            $qText = $line -replace '^>\s?', ''
            $tokens = Parse-InlineMarkdown $qText
            $runs = Build-RunsXml -tokens $tokens -defaultFont $FontFamily -defaultColor "4A5568" -szVal 21
            [void]$bodyXml.AppendLine(@"
    <w:p>
      <w:pPr><w:ind w:left="720"/><w:spacing w:before="80" w:after="120"/><w:pBdr><w:left w:val="single" w:sz="12" w:space="8" w:color="CBD5E0"/></w:pBdr></w:pPr>
      $runs
    </w:p>
"@)
        }
        # Scripture Verse Citation ([16] And he that stealeth...)
        elseif ($line -match '^(\[\d+\])\s*(.*)$') {
            $verseNum = Escape-XmlText $matches[1]
            $verseText = $matches[2]
            $tokens = Parse-InlineMarkdown $verseText
            $bodyRuns = Build-RunsXml -tokens $tokens -defaultFont $FontFamily -defaultColor "1B365D" -szVal 21
            [void]$bodyXml.AppendLine(@"
    <w:p>
      <w:pPr><w:ind w:left="500"/><w:spacing w:before="60" w:after="100"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="$FontFamily" w:hAnsi="$FontFamily"/><w:b/><w:sz w:val="21"/><w:color w:val="1F497D"/></w:rPr><w:t xml:space="preserve">$verseNum </w:t></w:r>
      $bodyRuns
    </w:p>
"@)
        }
        # Bullet list item (- item, * item)
        elseif ($line -match '^[-*+]\s+(.*)$') {
            $itemText = $matches[1]
            $tokens = Parse-InlineMarkdown $itemText
            $runs = Build-RunsXml -tokens $tokens -defaultFont $FontFamily -defaultColor "222222" -szVal 22
            [void]$bodyXml.AppendLine(@"
    <w:p>
      <w:pPr><w:ind w:left="720" w:hanging="360"/><w:spacing w:after="100"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/><w:sz w:val="20"/></w:rPr><w:t>&#183; </w:t></w:r>
      $runs
    </w:p>
"@)
        }
        # Standard Paragraph
        else {
            $tokens = Parse-InlineMarkdown $line
            $runs = Build-RunsXml -tokens $tokens -defaultFont $FontFamily -defaultColor "222222" -szVal 22
            [void]$bodyXml.AppendLine(@"
    <w:p>
      <w:pPr><w:spacing w:after="140" w:line="276" w:lineRule="auto"/></w:pPr>
      $runs
    </w:p>
"@)
        }
    }

    $documentXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
$($bodyXml.ToString())
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>
"@

    $contentTypesXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>
"@

    $relsXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"@

    $stylesXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="$FontFamily" w:hAnsi="$FontFamily" w:cs="$FontFamily"/>
        <w:sz w:val="22"/>
        <w:color w:val="222222"/>
      </w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal" w:default="1">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:rPr>
      <w:b/><w:sz w:val="36"/><w:color w:val="1F497D"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:rPr>
      <w:b/><w:sz w:val="28"/><w:color w:val="2C5E8A"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:rPr>
      <w:b/><w:sz w:val="25"/><w:color w:val="3B6998"/>
    </w:rPr>
  </w:style>
</w:styles>
"@

    # Remove destination if already exists to overwrite cleanly
    if (Test-Path $outFile) {
        Remove-Item -Path $outFile -Force
    }

    # Create ZIP archive (.docx)
    $zip = [System.IO.Compression.ZipFile]::Open($outFile, [System.IO.Compression.ZipArchiveMode]::Create)
    
    function Add-ZipEntry($zipArchive, $entryName, $stringContent) {
        $entry = $zipArchive.CreateEntry($entryName, [System.IO.Compression.CompressionLevel]::Optimal)
        $writer = New-Object System.IO.StreamWriter($entry.Open(), [System.Text.Encoding]::UTF8)
        $writer.Write($stringContent)
        $writer.Flush()
        $writer.Close()
    }

    Add-ZipEntry $zip "[Content_Types].xml" $contentTypesXml
    Add-ZipEntry $zip "_rels/.rels" $relsXml
    Add-ZipEntry $zip "word/document.xml" $documentXml
    Add-ZipEntry $zip "word/styles.xml" $stylesXml
    $zip.Dispose()

    $sizeKb = (Get-Item $outFile).Length / 1KB
    Write-Host "[+] SUCCESS: Created '$outLeaf' ($([Math]::Round($sizeKb, 1)) KB)" -ForegroundColor Green
    return $true
}

# --- Execution Entry Point ---
if ($All) {
    $files = Get-ChildItem -Path $RepoRoot -Filter "*.md" -Recurse | Where-Object { $_.Name -notmatch "README|CHAT_AND_UPDATES_LOG" }
    Write-Host "Found $($files.Count) markdown files across workspace. Converting..." -ForegroundColor Cyan
    foreach ($f in $files) {
        Convert-SingleMdToDocx -mdFile $f.FullName -outFile ($f.FullName -replace '\.md$', '.docx')
    }
} elseif ($Directory) {
    $recurseFlag = if ($Recursive) { $true } else { $false }
    $files = Get-ChildItem -Path $Directory -Filter "*.md" -Recurse:$recurseFlag
    Write-Host "Found $($files.Count) markdown files in $Directory. Converting..." -ForegroundColor Cyan
    foreach ($f in $files) {
        Convert-SingleMdToDocx -mdFile $f.FullName -outFile ($f.FullName -replace '\.md$', '.docx')
    }
} elseif ($InputPath) {
    Convert-SingleMdToDocx -mdFile $InputPath -outFile $OutputPath
} else {
    Write-Host ""
    Write-Host "=====================================================================" -ForegroundColor Cyan
    Write-Host "      THEOS - MARKDOWN TO DOCX CONVERTER (NATIVE ENGINE)" -ForegroundColor Cyan
    Write-Host "=====================================================================" -ForegroundColor Cyan
    Write-Host "  [1] Convert a single Markdown file (.md)" -ForegroundColor White
    Write-Host "  [2] Convert all .md files in a specific folder" -ForegroundColor White
    Write-Host "  [3] Convert all .md files in the entire Theos workspace" -ForegroundColor White
    Write-Host "  [0] Exit" -ForegroundColor Gray
    Write-Host "=====================================================================" -ForegroundColor Cyan
    $choice = Read-Host "Enter choice (0-3)"

    if ($choice -eq "1") {
        $p = Read-Host "`nEnter path to .md file (or drag & drop here)"
        if ($p) {
            Convert-SingleMdToDocx -mdFile $p.Trim('"').Trim("'")
        }
    } elseif ($choice -eq "2") {
        $d = Read-Host "`nEnter folder path (or press Enter for current folder)"
        if ([string]::IsNullOrWhiteSpace($d)) { $d = $RepoRoot }
        $files = Get-ChildItem -Path $d.Trim('"').Trim("'") -Filter "*.md"
        Write-Host "Found $($files.Count) markdown files in $d. Converting..." -ForegroundColor Cyan
        foreach ($f in $files) {
            Convert-SingleMdToDocx -mdFile $f.FullName -outFile ($f.FullName -replace '\.md$', '.docx')
        }
    } elseif ($choice -eq "3") {
        $files = Get-ChildItem -Path $RepoRoot -Filter "*.md" -Recurse | Where-Object { $_.Name -notmatch "README|CHAT_AND_UPDATES_LOG" }
        Write-Host "Found $($files.Count) markdown files across workspace. Converting..." -ForegroundColor Cyan
        foreach ($f in $files) {
            Convert-SingleMdToDocx -mdFile $f.FullName -outFile ($f.FullName -replace '\.md$', '.docx')
        }
    } else {
        Write-Host "Goodbye!" -ForegroundColor Gray
    }
}

