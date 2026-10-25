# Record chosen slides of a .pptx to MP4 through PowerPoint; each click advances after $Advance seconds.
# PowerPoint cannot export while the same file is open for editing, so work on a copy.
param([string]$Src, [string]$Out, [string]$Keep, [int]$Advance = 1, [int]$Height = 720)
try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application") }
catch { $app = New-Object -ComObject PowerPoint.Application }
foreach ($o in @($app.Presentations)) { if ($o.FullName -eq $Src) { $o.Close() } }
$p = $app.Presentations.Open($Src, $false, $false, $false)
if ($Keep) {
  $keep = $Keep.Split(",") | ForEach-Object { [int]$_ }
  for ($i = $p.Slides.Count; $i -ge 1; $i--) { if ($keep -notcontains $i) { $p.Slides($i).Delete() } }
}
foreach ($s in $p.Slides) { $s.SlideShowTransition.AdvanceOnTime = -1; $s.SlideShowTransition.AdvanceTime = $Advance }
$p.CreateVideo($Out, $true, $Advance, $Height, 25, 85)
# CreateVideoStatus: 0 not started yet, 1 running, 2 queued, 3 done, 4 failed
$t0 = Get-Date
do { Start-Sleep 2; $st = $p.CreateVideoStatus } while ($st -lt 3 -and ((Get-Date) - $t0).TotalSeconds -lt 600)
"status=$st"
$p.Close()
if ($st -ne 3) { exit 1 }
