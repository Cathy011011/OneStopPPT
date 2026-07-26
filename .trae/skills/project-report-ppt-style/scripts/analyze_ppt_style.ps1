param(
    [Parameter(Mandatory = $true)][string]$PptxPath,
    [Parameter(Mandatory = $true)][string]$OutputPath
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Read-ZipXml($zip, [string]$name) {
    $entry = $zip.Entries | Where-Object { $_.FullName -eq $name } | Select-Object -First 1
    if ($null -eq $entry) { return $null }
    $reader = [IO.StreamReader]::new($entry.Open())
    try {
        $xml = [xml]$reader.ReadToEnd()
        return $xml
    }
    finally { $reader.Dispose() }
}

function Add-Count($map, $key) {
    if ([string]::IsNullOrWhiteSpace($key)) { return }
    if (-not $map.ContainsKey($key)) { $map[$key] = 0 }
    $map[$key] += 1
}

$resolvedPptx = [IO.Path]::GetFullPath($PptxPath)
$resolvedOutput = [IO.Path]::GetFullPath($OutputPath)
$zip = [IO.Compression.ZipFile]::OpenRead($resolvedPptx)
try {
    $presentationXml = Read-ZipXml $zip "ppt/presentation.xml"
    $ns = [Xml.XmlNamespaceManager]::new($presentationXml.NameTable)
    $ns.AddNamespace("p", "http://schemas.openxmlformats.org/presentationml/2006/main")
    $ns.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
    $ns.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart")
    $sizeNode = $presentationXml.SelectSingleNode("/p:presentation/p:sldSz", $ns)
    $cx = [double]$sizeNode.cx
    $cy = [double]$sizeNode.cy

    $themeXml = Read-ZipXml $zip "ppt/theme/theme1.xml"
    $themeNs = [Xml.XmlNamespaceManager]::new($themeXml.NameTable)
    $themeNs.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
    $colors = [ordered]@{}
    foreach ($node in $themeXml.SelectNodes("/a:theme/a:themeElements/a:clrScheme/*", $themeNs)) {
        $colorNode = $node.SelectSingleNode("a:srgbClr|a:sysClr", $themeNs)
        if ($null -ne $colorNode) { $colors[$node.LocalName] = if ($colorNode.val) { $colorNode.val } else { $colorNode.lastClr } }
    }
    $fonts = [ordered]@{}
    foreach ($kind in @("majorFont", "minorFont")) {
        $fontNode = $themeXml.SelectSingleNode("/a:theme/a:themeElements/a:fontScheme/a:$kind", $themeNs)
        $fonts[$kind] = [ordered]@{
            latin = $fontNode.SelectSingleNode("a:latin", $themeNs).typeface
            eastAsian = $fontNode.SelectSingleNode("a:ea", $themeNs).typeface
            complexScript = $fontNode.SelectSingleNode("a:cs", $themeNs).typeface
        }
    }

    $fontSizes = @{}
    $fontFaces = @{}
    $textColors = @{}
    $alignments = @{}
    $slides = @()
    $slideEntries = @($zip.Entries | Where-Object { $_.FullName -match '^ppt/slides/slide\d+\.xml$' } | Sort-Object { [int]([regex]::Match($_.FullName, 'slide(\d+)\.xml').Groups[1].Value) })
    foreach ($entry in $slideEntries) {
        $slideXml = Read-ZipXml $zip $entry.FullName
        $texts = @($slideXml.SelectNodes("//a:t", $ns) | ForEach-Object { $_.'#text' })
        $runs = @($slideXml.SelectNodes("//a:rPr", $ns))
        foreach ($run in $runs) {
            if ($run.sz) { Add-Count $fontSizes ([string]([double]$run.sz / 100)) }
            if ($run.typeface) { Add-Count $fontFaces ([string]$run.typeface) }
            $color = $run.SelectSingleNode("a:solidFill/a:srgbClr", $ns)
            if ($null -ne $color) { Add-Count $textColors ([string]$color.val) }
        }
        foreach ($paragraph in @($slideXml.SelectNodes("//a:pPr", $ns))) { if ($paragraph.algn) { Add-Count $alignments ([string]$paragraph.algn) } }
        $slides += [pscustomobject][ordered]@{
            number = [int]([regex]::Match($entry.FullName, 'slide(\d+)\.xml').Groups[1].Value)
            shapes = @($slideXml.SelectNodes("//p:sp", $ns)).Count
            textShapes = @($slideXml.SelectNodes("//p:sp[p:txBody]", $ns)).Count
            pictures = @($slideXml.SelectNodes("//p:pic", $ns)).Count
            connectors = @($slideXml.SelectNodes("//p:cxnSp", $ns)).Count
            groups = @($slideXml.SelectNodes("//p:grpSp", $ns)).Count
            tables = @($slideXml.SelectNodes("//a:tbl", $ns)).Count
            charts = @($slideXml.SelectNodes("//c:chart", $ns)).Count
            textCharacters = (($texts -join "").Length)
        }
    }

    $result = [ordered]@{
        source = $resolvedPptx
        slideSize = [ordered]@{ emuWidth = [int64]$cx; emuHeight = [int64]$cy; ratio = [math]::Round($cx / $cy, 6); aspect = "16:9" }
        masters = @($presentationXml.SelectNodes("/p:presentation/p:sldMasterIdLst/p:sldMasterId", $ns)).Count
        slides = $slides
        totals = [ordered]@{
            slides = $slides.Count
            shapes = ($slides | Measure-Object -Property shapes -Sum).Sum
            textShapes = ($slides | Measure-Object -Property textShapes -Sum).Sum
            pictures = ($slides | Measure-Object -Property pictures -Sum).Sum
            connectors = ($slides | Measure-Object -Property connectors -Sum).Sum
            tables = ($slides | Measure-Object -Property tables -Sum).Sum
            charts = ($slides | Measure-Object -Property charts -Sum).Sum
            mediaFiles = @($zip.Entries | Where-Object { $_.FullName -match '^ppt/media/' }).Count
        }
        theme = [ordered]@{ name = $themeXml.theme.name; colors = $colors; fonts = $fonts }
        typography = [ordered]@{ fontSizes = $fontSizes; fontFaces = $fontFaces; textColors = $textColors; alignments = $alignments }
    }
    $parent = [IO.Path]::GetDirectoryName($resolvedOutput)
    if ($parent) { New-Item -ItemType Directory -Force -Path $parent | Out-Null }
    $result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $resolvedOutput -Encoding utf8
    Write-Output $resolvedOutput
}
finally { $zip.Dispose() }
