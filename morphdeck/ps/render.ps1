# Export every slide of a .pptx to PNG through PowerPoint.
param([string]$Src, [string]$Out, [int]$Width = 1280)
New-Item -ItemType Directory -Force -Path $Out | Out-Null
try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application") }
catch { $app = New-Object -ComObject PowerPoint.Application }
foreach ($open in @($app.Presentations)) { if ($open.FullName -eq $Src) { $open.Close() } }
$p = $app.Presentations.Open($Src, $true, $false, $false)
$h = [int]($Width * $p.PageSetup.SlideHeight / $p.PageSetup.SlideWidth)
$i = 1
foreach ($s in $p.Slides) { $s.Export((Join-Path $Out ("s{0:d2}.png" -f $i)), "PNG", $Width, $h); $i++ }
$p.Close()
