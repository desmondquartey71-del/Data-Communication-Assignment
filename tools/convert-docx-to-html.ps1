param(
  [Parameter(Mandatory = $true)]
  [string]$InputDocx,

  [Parameter(Mandatory = $true)]
  [string]$OutputHtml
)

Add-Type -AssemblyName System.IO.Compression.FileSystem

function HtmlEncode([string]$value) {
  return [System.Net.WebUtility]::HtmlEncode($value)
}

$docxPath = (Resolve-Path -LiteralPath $InputDocx).Path
$outputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputHtml)

$archive = [System.IO.Compression.ZipFile]::OpenRead($docxPath)
try {
  $entry = $archive.GetEntry("word/document.xml")
  if (-not $entry) {
    throw "word/document.xml was not found in $docxPath"
  }

  $stream = $entry.Open()
  try {
    $reader = New-Object System.IO.StreamReader($stream)
    $xmlText = $reader.ReadToEnd()
  }
  finally {
    $stream.Dispose()
  }

  [xml]$xml = $xmlText
  $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
  $ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

  $parts = New-Object System.Collections.Generic.List[string]
  $paragraphs = $xml.SelectNodes("//w:body/w:p", $ns)

  foreach ($paragraph in $paragraphs) {
    $texts = New-Object System.Collections.Generic.List[string]

    foreach ($node in $paragraph.SelectNodes(".//w:t|.//w:tab|.//w:br", $ns)) {
      if ($node.LocalName -eq "t") {
        $text = $node.InnerText.Trim()
        if ($text.Length -gt 0) {
          $texts.Add($text)
        }
      }
      elseif ($node.LocalName -eq "tab") {
        $texts.Add(" ")
      }
      elseif ($node.LocalName -eq "br") {
        $texts.Add("<br>")
      }
    }

    if ($texts.Count -eq 0) {
      continue
    }

    $joined = [string]::Join(" ", $texts)
    $joined = $joined -replace "\s+([,.;:!?])", '$1'
    $joined = $joined -replace "\(\s+", "("
    $joined = $joined -replace "\s+\)", ")"
    $joined = $joined.Trim()
    if ($joined.Length -eq 0) {
      continue
    }

    $encoded = HtmlEncode $joined
    $encoded = $encoded -replace "&lt;br&gt;", "<br>"

    $outline = $paragraph.SelectSingleNode("./w:pPr/w:outlineLvl", $ns)
    $style = $paragraph.SelectSingleNode("./w:pPr/w:pStyle", $ns)
    $isNumbered = $null -ne $paragraph.SelectSingleNode("./w:pPr/w:numPr", $ns)

    if ($outline) {
      $level = [int]$outline.GetAttribute("val", $ns.LookupNamespace("w"))
      if ($level -le 0) {
        $parts.Add("          <h2>$encoded</h2>")
      }
      else {
        $parts.Add("          <h3>$encoded</h3>")
      }
    }
    elseif ($style -and $style.GetAttribute("val", $ns.LookupNamespace("w")) -match "Title") {
      $parts.Add("          <h2>$encoded</h2>")
    }
    elseif ($isNumbered) {
      $parts.Add("          <p class=`"doc-note`">$encoded</p>")
    }
    else {
      $parts.Add("          <p>$encoded</p>")
    }
  }

  $body = [string]::Join([Environment]::NewLine, $parts)
  $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Assignment Viewer - Communications Satellite</title>
  <link rel="stylesheet" href="../styles.css">
</head>
<body>
  <header class="topbar">
    <a class="brand" href="../index.html" aria-label="Go to home">
      <span class="brand-mark">DC</span>
      <span>Signal Studio</span>
    </a>
    <nav class="site-menu assignment-menu" aria-label="Assignment navigation">
      <a href="../index.html#about">Profile</a>
      <a href="assignment.html">Assignment</a>
      <a href="../index.html#video">Video</a>
    </nav>
  </header>

  <main>
    <section class="section document-section">
      <div class="section-heading">
        <p class="eyebrow">Assignment Viewer</p>
        <h1>Communications Satellite</h1>
      </div>
      <article class="document-view">
$body
      </article>
    </section>
  </main>

  <footer class="footer">
    <p>PUIS/24110034 | DESMOND QUARTEY</p>
  </footer>
</body>
</html>
"@

  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllText($outputPath, $html, $utf8NoBom)
}
finally {
  $archive.Dispose()
}
