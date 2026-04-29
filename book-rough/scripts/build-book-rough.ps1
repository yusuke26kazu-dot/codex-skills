param(
  [Parameter(Mandatory = $true)]
  [string]$TemplatePath,

  [Parameter(Mandatory = $true)]
  [string]$OutputPath,

  [Parameter(Mandatory = $true)]
  [string]$ShopName,

  [Parameter(Mandatory = $true)]
  [string]$RightImagePath,

  [Parameter(Mandatory = $true)]
  [string]$LeftTopImagePath,

  [Parameter(Mandatory = $true)]
  [string]$LeftLowerImagePath,

  [int]$BaseSlideIndex = 0,

  [string]$RenderDir = ""
)

$ErrorActionPreference = "Stop"

function Resolve-ExistingPath {
  param([string]$Path, [string]$Label)
  if (-not (Test-Path -LiteralPath $Path)) {
    throw "$Label not found: $Path"
  }
  return (Resolve-Path -LiteralPath $Path).Path
}

function Release-ComObject {
  param($Object)
  if ($null -ne $Object) {
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Object)
  }
}

$template = Resolve-ExistingPath $TemplatePath "Template"
$rightImage = Resolve-ExistingPath $RightImagePath "Right image"
$leftTopImage = Resolve-ExistingPath $LeftTopImagePath "Left-top image"
$leftLowerImage = Resolve-ExistingPath $LeftLowerImagePath "Left-lower image"
$output = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
$outputDir = Split-Path -Parent $output
if ($outputDir -and -not (Test-Path -LiteralPath $outputDir)) {
  New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
}

if ((Resolve-Path -LiteralPath $template).Path -ne $output) {
  Copy-Item -LiteralPath $template -Destination $output -Force
}

$powerPoint = $null
$presentation = $null
$slide = $null

try {
  $powerPoint = New-Object -ComObject PowerPoint.Application
  $powerPoint.DisplayAlerts = 1
  $presentation = $powerPoint.Presentations.Open($output, $false, $false, $false)

  if ($presentation.Slides.Count -lt 1) {
    throw "Template has no slides."
  }

  if ($BaseSlideIndex -le 0) {
    $BaseSlideIndex = $presentation.Slides.Count
  }
  if ($BaseSlideIndex -lt 1 -or $BaseSlideIndex -gt $presentation.Slides.Count) {
    throw "BaseSlideIndex $BaseSlideIndex is outside the slide range 1-$($presentation.Slides.Count)."
  }

  for ($i = $presentation.Slides.Count; $i -ge 1; $i--) {
    if ($i -ne $BaseSlideIndex) {
      $presentation.Slides.Item($i).Delete()
    }
  }

  $slide = $presentation.Slides.Item(1)

  $groupCount = 0
  for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
    if ($slide.Shapes.Item($i).Type -eq 6) {
      $groupCount++
    }
  }
  if ($groupCount -lt 1) {
    throw "No template group was found on the base slide. Use the reference-example BOOKラフ deck, not a flattened screenshot-only slide."
  }

  for ($i = $slide.Shapes.Count; $i -ge 1; $i--) {
    $shape = $slide.Shapes.Item($i)
    if ($shape.Type -ne 6) {
      $shape.Delete()
    }
  }

  # Coordinates are PowerPoint points for the corrected BOOKラフ v2 format.
  $right = $slide.Shapes.AddPicture($rightImage, 0, -1, 287.3, 64.8, 398.6, 385.5)
  $right.ZOrder(1) | Out-Null

  $slide.Shapes.AddPicture($leftTopImage, 0, -1, 66.1, 88.3, 162.3, 91.0) | Out-Null
  $slide.Shapes.AddPicture($leftLowerImage, 0, -1, 152.8, 156.4, 130.6, 73.3) | Out-Null

  $title = $slide.Shapes.AddTextbox(1, 66.1, 263.0, 420.0, 44.0)
  $title.TextFrame.MarginLeft = 0
  $title.TextFrame.MarginRight = 0
  $title.TextFrame.MarginTop = 0
  $title.TextFrame.MarginBottom = 0
  $title.TextFrame.WordWrap = 0
  $title.TextFrame.AutoSize = 1
  $title.TextFrame.TextRange.Text = $ShopName
  $title.TextFrame.TextRange.Font.Name = "HG明朝E"
  $title.TextFrame.TextRange.Font.NameFarEast = "HG明朝E"
  $title.TextFrame.TextRange.Font.Size = 28
  $title.TextFrame.TextRange.Font.Bold = 0
  $title.TextFrame.TextRange.Font.Color.RGB = 16777215
  $title.TextFrame.TextRange.ParagraphFormat.Alignment = 1

  $extension = [System.IO.Path]::GetExtension($output).ToLowerInvariant()
  if ($extension -eq ".pptx") {
    $presentation.SaveAs($output, 24)
  } else {
    $presentation.SaveAs($output, 1)
  }

  if ($RenderDir) {
    $renderPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($RenderDir)
    New-Item -ItemType Directory -Force -Path $renderPath | Out-Null
    $slide.Export((Join-Path $renderPath "slide1_960x720.png"), "PNG", 960, 720) | Out-Null
    $slide.Export((Join-Path $renderPath "slide1_1600x1200.png"), "PNG", 1600, 1200) | Out-Null
  }
}
finally {
  if ($null -ne $presentation) {
    $presentation.Close()
  }
  if ($null -ne $powerPoint) {
    $powerPoint.Quit()
  }
  Release-ComObject $slide
  Release-ComObject $presentation
  Release-ComObject $powerPoint
}

Write-Output "Created: $output"
if ($RenderDir) {
  Write-Output "Rendered previews: $RenderDir"
}
