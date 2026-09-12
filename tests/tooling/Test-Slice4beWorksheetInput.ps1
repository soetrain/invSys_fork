$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before disposable input calibration.'}
$reportRoot=Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) 'reports/runtime/slice4be-receiving-activity'
$run=Join-Path $reportRoot ('native-calibration-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $run | Out-Null
. (Join-Path $PSScriptRoot 'Slice4beWorksheetInput.ps1')
$excel=$null;$book=$null;$owner=$null;$component=$null;$sheet=$null;$button=$null;$other=$null
$checks=[ordered]@{}
try {
 $excel=New-Object -ComObject Excel.Application
 $owner=Get-Process -Id ([WorksheetInput]::Owner([IntPtr]$excel.Hwnd))
 if($owner.ProcessName -cne 'EXCEL'){throw 'Owned Excel identity absent.'}
 $birth=$owner.StartTime.ToUniversalTime().Ticks
 $excel.DisplayAlerts=$false;$excel.AutomationSecurity=1;$excel.EnableEvents=$false
 $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
 $component=$book.VBProject.VBComponents.Add(1);$component.Name='modInputProbe'
 $component.CodeModule.AddFromString(@'
Public Entries As Long
Public ShapeCaller As Boolean
Public BoundBook As Boolean
Public Sub ButtonAction()
 Entries = Entries + 1
 ShapeCaller = False
 If VarType(Application.Caller) = vbString Then ShapeCaller = (Application.Caller = "btnInputProbe")
 BoundBook = (ActiveWorkbook Is ThisWorkbook)
End Sub
Public Function ReadEvidence() As String
 ReadEvidence = CStr(Entries) & "|" & CStr(ShapeCaller) & "|" & CStr(BoundBook)
End Function
'@)
 $button=$sheet.Shapes.AddFormControl(0,80,70,180,32);$button.Name='btnInputProbe'
 $button.TextFrame.Characters().Text='Disposable input calibration'
 $button.OnAction="'"+$book.Name+"'!modInputProbe.ButtonAction"
 $excel.Visible=$true;$excel.WindowState=-4137;$book.Activate();$sheet.Activate()
 $excel.ActiveWindow.ScrollRow=1;$excel.ActiveWindow.ScrollColumn=1
 $checks.ForegroundOwned=[WorksheetInput]::Activate([IntPtr]$book.Windows.Item(1).Hwnd)
 if(-not $checks.ForegroundOwned){throw 'Foreground calibration unavailable.'}
 Start-Sleep -Milliseconds 500
 $point=Get-WorksheetButtonPoint $excel $button $book $sheet
 $x=$point.X;$y=$point.Y
 Add-Type -AssemblyName System.Drawing
 $bounds=New-Object WorksheetInput+Rect
 if(-not [WorksheetInput]::GetWindowRect([IntPtr]$excel.Hwnd,[ref]$bounds)){throw 'Owned window bounds unavailable.'}
 $capture=New-Object Drawing.Bitmap(($bounds.Right-$bounds.Left),($bounds.Bottom-$bounds.Top))
 $graphics=[Drawing.Graphics]::FromImage($capture)
 try {
  $graphics.CopyFromScreen($bounds.Left,$bounds.Top,0,0,$capture.Size)
  $graphics.DrawEllipse([Drawing.Pens]::Red,$x-$bounds.Left-5,$y-$bounds.Top-5,10,10)
  $capture.Save((Join-Path $run 'fixture.png'))
 } finally {$graphics.Dispose();$capture.Dispose()}
 [WorksheetInput]::Click([IntPtr]$book.Windows.Item(1).Hwnd,$x,$y)
 Start-Sleep -Milliseconds 700
 $evidence=[string]$excel.Run(("'"+$book.Name+"'!modInputProbe.ReadEvidence"))
 $parts=$evidence.Split('|')
 Write-Output ('Native evidence: '+$evidence)
 $checks.NativeEntry=($parts[0] -ceq '1')
 $checks.ExactShapeCaller=($parts[1] -ceq 'True')
 $checks.FixtureBound=($parts[2] -ceq 'True')
 [void]$excel.Run(("'"+$book.Name+"'!modInputProbe.ButtonAction"))
 $evidence=[string]$excel.Run(("'"+$book.Name+"'!modInputProbe.ReadEvidence"))
 $checks.ProgrammaticCallDistinct=($evidence -ceq '2|False|True')
 Write-Output ('Programmatic evidence: '+$evidence)
 $other=$excel.Workbooks.Add()
 $other.Activate();$null=[WorksheetInput]::Activate([IntPtr]$other.Windows.Item(1).Hwnd)
 Start-Sleep -Milliseconds 350
 $checks.OtherWorkbookWindowRejected=$false
 try {[WorksheetInput]::Click([IntPtr]$book.Windows.Item(1).Hwnd,$point.X,$point.Y)} catch {$checks.OtherWorkbookWindowRejected=$_.Exception.GetBaseException().Message -ceq 'Input target is not foreground owned Excel.'}
 $checks.InactiveFixtureRejected=$false
 try {$null=Get-WorksheetButtonPoint $excel $button $book $sheet} catch {$checks.InactiveFixtureRejected=$_.Exception.Message -ceq 'Native fixture workbook and sheet are not active.'}
 $excel.Visible=$false;$book.Activate();$excel.Visible=$true;$book.Activate();$sheet.Activate()
 if(-not [WorksheetInput]::Activate([IntPtr]$book.Windows.Item(1).Hwnd)){throw 'Owned fixture foreground unavailable after workbook switch.'}
 Start-Sleep -Milliseconds 350
 $point=Get-WorksheetButtonPoint $excel $button $book $sheet
 [WorksheetInput]::Click([IntPtr]$book.Windows.Item(1).Hwnd,$point.X,$point.Y)
 Start-Sleep -Milliseconds 500
 $checks.NativeAfterWorkbookSwitch=([string]$excel.Run(("'"+$book.Name+"'!modInputProbe.ReadEvidence")) -ceq '3|True|True')
 $checks | ConvertTo-Json
 if($checks.Values -contains $false){throw 'Native input calibration failed.'}
} finally {
 $checks | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $run 'checks.json')
 if($null -ne $book){try{$book.Close($false)}catch{}}
 if($null -ne $other){try{$other.Close($false)}catch{}}
 if($null -ne $excel){try{$excel.Quit()}catch{}}
 foreach($obj in @($button,$sheet,$component,$book,$other,$excel)){if($null -ne $obj){try{$null=[Runtime.InteropServices.Marshal]::ReleaseComObject($obj)}catch{}}}
 if($null -ne $owner){
  if(-not $owner.WaitForExit(10000)){
   if($owner.StartTime.ToUniversalTime().Ticks -ne $birth -or $owner.ProcessName -cne 'EXCEL'){throw 'Owner identity changed; preserve process.'}
   $owner.Kill();if(-not $owner.WaitForExit(10000)){throw 'Owned Excel has not exited.'}
  }
  $owner.Dispose()
 }
}
