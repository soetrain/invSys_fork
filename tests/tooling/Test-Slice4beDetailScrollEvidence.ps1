# Detector calibration only. Synthetic images contain no operator/runtime data.
[CmdletBinding()]
param([string]$RepoRoot='.',[ValidateSet('RED','GREEN')][string]$Phase='GREEN')
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
$repo=(Resolve-Path $RepoRoot).Path
. (Join-Path $repo 'tests/tooling/Slice4beDetailScrollEvidence.ps1')
Add-Type -AssemblyName System.Drawing
$root=Join-Path $repo ('reports/runtime/detail-scroll-detector/'+[Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$results=[Collections.Generic.List[object]]::new()
function Check([string]$Name,[bool]$Passed){
 $results.Add([pscustomobject]@{Check=$Name;Passed=$Passed})
 Write-Output ($Name+': '+$(if($Passed){'PASS'}else{'FAIL'}))
}
function New-ScrollFixture([string]$Path,[double]$Scale,[bool]$Moved,[bool]$OutsideNoise){
 $bitmap=[Drawing.Bitmap]::new([int](1000*$Scale),[int](600*$Scale));$drawing=$null
 try {
  $drawing=[Drawing.Graphics]::FromImage($bitmap);$drawing.Clear([Drawing.Color]::White)
  # Physical rendering of one horizontal scrollbar, with a thumb at either end.
  $drawing.FillRectangle([Drawing.Brushes]::LightGray,[single](20*$Scale),[single](380*$Scale),[single](960*$Scale),[single](20*$Scale))
  $left=if($Moved){550}else{45}
  $drawing.FillRectangle([Drawing.Brushes]::DarkGray,[single]($left*$Scale),[single](382*$Scale),[single](370*$Scale),[single](16*$Scale))
  # Change only unrelated content above the real scrollbar. It must not count.
  if($OutsideNoise){$drawing.FillRectangle([Drawing.Brushes]::Black,44,384,880,12)}
  $bitmap.Save($Path,[Drawing.Imaging.ImageFormat]::Png)
 }finally{if($null -ne $drawing){$drawing.Dispose()};$bitmap.Dispose()}
}
$geometry='960|300|20|920|290|screen=1040,490|list=120,200,1080,500|form=100,100,1100,700|dpi=96'
foreach($scale in @(1,1.5,2)){
 $name=([int]($scale*100)).ToString()
 $before=Join-Path $root ($name+'-before.png');$after=Join-Path $root ($name+'-after.png')
 New-ScrollFixture $before $scale $false $false
 New-ScrollFixture $after $scale $true $false
 Check ('Detector.Moving.'+$name) (Measure-DetailScrollMotion $before $after $geometry).MovementObserved
 Check ('Detector.Stationary.'+$name) (-not (Measure-DetailScrollMotion $before $before $geometry).MovementObserved)
 if($scale -gt 1){
  $noise=Join-Path $root ($name+'-noise.png');New-ScrollFixture $noise $scale $false $true
  Check ('Detector.UnrelatedContentIgnored.'+$name) (-not (Measure-DetailScrollMotion $before $noise $geometry).MovementObserved)
 }
}
$results.ToArray()|ConvertTo-Json|Set-Content (Join-Path $root ($Phase.ToLowerInvariant()+'.json'))
Write-Output ('REPORT_ROOT='+$root)
if(@($results|Where-Object {-not $_.Passed}).Count){exit 1}
