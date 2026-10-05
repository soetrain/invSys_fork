# Native preview observation closes preview only; it never invokes Print.
. (Join-Path $PSScriptRoot 'Slice4beProductionPrintAccessibility.ps1')
function Start-ProductionPrintPreviewObserver([int]$ProcessId,[long]$Window,[string]$ImagePath,[string]$StopPath){
    $capture=${function:Initialize-SettingsCapture}.ToString()
    $accessibility=${function:Initialize-ProductionPrintAccessibility}.ToString()
    Start-Job -ArgumentList $ProcessId,$Window,$ImagePath,$StopPath,$capture,$accessibility -ScriptBlock {
        param($OwnedProcess,$OwnedWindow,$ImagePath,$StopPath,$CaptureDefinition,$AccessibilityDefinition)
        $ErrorActionPreference='Stop'
        Add-Type -AssemblyName UIAutomationClient
        Add-Type -AssemblyName UIAutomationTypes
        & ([scriptblock]::Create($CaptureDefinition))
        & ([scriptblock]::Create($AccessibilityDefinition))
        $until=[DateTime]::UtcNow.AddSeconds(40)
        $found=$false;$captured=$false;$closed=$false;$occluded=$false
        try{
            while([DateTime]::UtcNow -lt $until -and -not (Test-Path -LiteralPath $StopPath)){
                $root=[Windows.Automation.AutomationElement]::FromHandle([IntPtr]$OwnedWindow)
                if($null -eq $root -or $root.Current.ProcessId -ne $OwnedProcess){throw 'Owned preview window unavailable.'}
                $name=New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::NameProperty,'Close Print Preview')
                $kind=New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ControlTypeProperty,[Windows.Automation.ControlType]::Button)
                $condition=New-Object Windows.Automation.AndCondition($name,$kind)
                $buttons=$root.FindAll([Windows.Automation.TreeScope]::Descendants,$condition)
                $button=$null;$legacy=$null
                if($buttons.Count -eq 1 -and $buttons.Item(0).Current.IsEnabled -and -not $buttons.Item(0).Current.IsOffscreen){$button=$buttons.Item(0)}
                if($null -eq $button){$legacy=[InvSysPrintPreviewButton]::Find([IntPtr]$OwnedWindow,[uint32]$OwnedProcess)}
                if($null -ne $button -or $null -ne $legacy){
                    if($null -ne $button -and $button.Current.ProcessId -ne $OwnedProcess){throw 'Preview close ownership mismatch.'}
                    $found=$true
                    try{[InvSysSettingsCapture]::SaveVisibleWindow([IntPtr]$OwnedWindow,$ImagePath);$captured=$true}
                    catch{
                        $occluded=$true
                        $front=[InvSysSettingsCapture]::OwnedVisibleForm('Production',[IntPtr]$OwnedWindow)
                        if($front -ne [IntPtr]::Zero){try{[InvSysSettingsCapture]::SaveVisibleWindow($front,$ImagePath)}catch{}}
                    }
                    if($null -ne $button){$button.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()}else{$legacy.Close()}
                    $closed=$true
                    break
                }
                Start-Sleep -Milliseconds 200
            }
            [pscustomobject]@{PreviewCloseObserved=$found;PreviewCaptured=$captured;CloseInvoked=$closed;PreviewOccluded=$occluded;PrintInvoked=$false}
        }catch{
            [pscustomobject]@{PreviewCloseObserved=$found;PreviewCaptured=$captured;CloseInvoked=$closed;PreviewOccluded=$occluded;PrintInvoked=$false;ObserverFailed=$true}
        }
    }
}

function Install-ProductionPrintNativeProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $module=$project.VBComponents.Item('modProductionRecallReport').CodeModule
    $start=$module.ProcStartLine('Preview',0);$end=$start+$module.ProcCountLines('Preview',0)
    $hits=@(for($i=$start;$i -lt $end;$i++){if($module.Lines($i,1).Trim() -ceq 'TestProductionDesigner.PrintPreviewForTest wsReport'){$i}})
    if($hits.Count -ne 1){throw 'Print native seam anchor unavailable; not product RED.'}
    $module.ReplaceLine($hits[0],@'
    If TestProductionDesigner.PrintNativeEnabled Then
        TestProductionDesigner.PrintNativeBefore
        wsReport.PrintOut Preview:=True
        TestProductionDesigner.PrintNativeAfter
    Else
        TestProductionDesigner.PrintPreviewForTest wsReport
    End If
'@)
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    $form.AddFromString(@'
Public Sub PrintNativeCloseFormForTest()
    mBtnClose_Click
End Sub
'@)
    $adapter=$project.VBComponents.Item('TestProductionDesigner').CodeModule
    $adapter.InsertLines(1,@'
Private mPrintNativeEnabled As Boolean, mPrintNativeEntered As Boolean, mPrintNativeReturned As Boolean
Private mPrintNativeCloseForm As Boolean, mPrintNativeClosed As Boolean
'@)
    $start=$adapter.ProcStartLine('PrintPreviewForTest',0);$last=$start+$adapter.ProcCountLines('PrintPreviewForTest',0)-1
    while($adapter.Lines($last,1).Trim() -cne 'End Sub'){$last--;if($last -le $start){throw 'Print preview end anchor missing.'}}
    $adapter.InsertLines($last,'    PrintNativeCloseBoundary')
    $adapter.AddFromString(@'
Public Sub PrintNativeArm(ByVal native As Boolean, ByVal closeForm As Boolean)
    mPrintNativeEnabled = native: mPrintNativeEntered = False: mPrintNativeReturned = False
    mPrintNativeCloseForm = closeForm: mPrintNativeClosed = False
End Sub
Public Function PrintNativeEnabled() As Boolean
    PrintNativeEnabled = mPrintNativeEnabled
End Function
Public Sub PrintNativeBefore()
    mPrintNativeEntered = True
End Sub
Public Sub PrintNativeAfter()
    mPrintNativeReturned = True
End Sub
Public Sub PrintNativeCloseBoundary()
    If Not mPrintNativeCloseForm Then Exit Sub
    mPrintNativeCloseForm = False
    mForm.PrintNativeCloseFormForTest
    mPrintNativeClosed = Not modOperationsFormLifetime.IsLoaded(mForm)
End Sub
Public Function PrintNativeFact(ByVal fact As String) As Boolean
    Select Case fact
        Case "Entered": PrintNativeFact = mPrintNativeEntered
        Case "Returned": PrintNativeFact = mPrintNativeReturned
        Case "Closed": PrintNativeFact = mPrintNativeClosed
    End Select
End Function
'@)
}

function Test-ProductionPrintNative($Fixture,$Other,$Book,$Sheet,$Decoy){
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Hash([string]$Path){$stream=[IO.File]::Open($Path,'Open','Read','ReadWrite');try{(Get-FileHash -InputStream $stream).Hash}finally{$stream.Dispose()}}
    function Fingerprint($Worksheet){ConvertTo-Json -Compress -Depth 10 -InputObject @($Worksheet.UsedRange.Formula)}
    function AuthorityPins($Target){$pins=@{};foreach($file in Get-ChildItem -LiteralPath $Target.Root -Recurse -File|Where-Object {$_.Name -notlike '~$*' -and $_.Extension -ceq '.xlsb'}){$pins[$file.FullName]=Hash $file.FullName};return $pins}
    function Same($Before,$After){if($Before.Count -ne $After.Count){return $false};foreach($path in $Before.Keys){if(-not $After.ContainsKey($path) -or $After[$path] -cne $Before[$path]){return $false}};return $true}
    $source=Fingerprint $Sheet;$foreign=Fingerprint $Decoy.Worksheets.Item('Production')
    foreach($case in @('Preview','CloseForm')){
        $observer=$null;$noticeObserver=$null;$stop='';$priorVisible=$excel.Visible
        try{
            SelectTarget $Fixture 'config-producer'
            $excel.Visible=$true
            [void](Probe 'OpenDesigner' @($Book.Name));[void](Probe 'RunLocalShowAndCapture' @($Book.Name,'PRINT'))
            [void](Probe 'PrintYieldArm' @('',''));[void](Probe 'ResetPrintPreviewForTest')
            [void](Probe 'PrintNativeArm' @(($case -ceq 'Preview'),($case -ceq 'CloseForm')))
            $before=AuthorityPins $Fixture;$otherPins=AuthorityPins $Other
            $stop=Join-Path $runRoot ('print-native-'+$case+'-stop')
            $processes=@(Get-Process EXCEL);if($processes.Count -ne 1){throw 'Isolated Excel required.'}
            $noticeObserver=Start-DialogCaptureAndDismiss -ExcelProcessId $processes[0].Id -TimeoutSeconds 45 -StopPath $stop
            if($case -ceq 'Preview'){
                $Book.Activate()
                $observer=Start-ProductionPrintPreviewObserver $processes[0].Id ([long]$Book.Windows.Item(1).Hwnd) (Join-Path $reportRoot 'print-native-preview.png') $stop
            }
            $label='PrintNative.'+$case
            $returned=[bool](Probe 'PrintYieldAct')
            [IO.File]::WriteAllText($stop,'')
            Check ($label+'.ActualHandlerReturned') $returned
            Check ($label+'.OneOwnerEntry') ([int](Probe 'PrintOwnerEntries') -eq 1)
            Check ($label+'.OneReportRead') ([int](Probe 'PrintReportReads') -eq 1)
            if($case -ceq 'Preview'){
                Wait-Job $observer -Timeout 45|Out-Null
                $receipt=@(Receive-Job $observer -ErrorAction SilentlyContinue)
                if($observer.State -ne 'Completed' -or $receipt.Count -ne 1){throw 'Native preview observer did not complete; not product RED.'}
                $receipt[0]|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'print-native-preview.json')
                Check ($label+'.OriginalNativeCallEntered') ([bool](Probe 'PrintNativeFact' @('Entered')))
                Check ($label+'.OriginalNativeCallReturned') ([bool](Probe 'PrintNativeFact' @('Returned')))
                Check ($label+'.PreviewCloseObserved') $receipt[0].PreviewCloseObserved
                Check ($label+'.PreviewVisibleAndCaptured') $receipt[0].PreviewCaptured
                Check ($label+'.OnlyPreviewCloseInvoked') ($receipt[0].CloseInvoked -and -not $receipt[0].PrintInvoked)
                Check ($label+'.TruthfulStatus') ([string](Probe 'PrintStatus') -ceq 'Print preview closed.')
                Check ($label+'.GuardsRestored') ([bool](Probe 'PrintYieldFact' @('GuardsRestored')))
                CaptureOwnedFormByCaptionEvidence 'Production' 'print-native-return.png'
            }else{
                Check ($label+'.ActualCloseHandlerDismissedForm') ([bool](Probe 'PrintNativeFact' @('Closed')))
                Check ($label+'.NoReinitialization') ([bool](Probe 'PrintYieldFact' @('NoReinitialization')))
                Check ($label+'.FormRemainsUnloaded') (-not [bool](Probe 'PrintYieldFact' @('Loaded')))
                Check ($label+'.WorkbookRemainsOpen') ($Book.Name -cin @($excel.Workbooks|ForEach-Object Name))
            }
            Check ($label+'.SourceAndExactKeyPreserved') ((Fingerprint $Sheet) -ceq $source)
            Check ($label+'.DecoyPreserved') ((Fingerprint $Decoy.Worksheets.Item('Production')) -ceq $foreign)
            Check ($label+'.AuthorityWorkbookBytesPreserved') (Same $before (AuthorityPins $Fixture))
            Check ($label+'.OtherWarehousePreserved') (Same $otherPins (AuthorityPins $Other))
        }finally{
            if($stop){[IO.File]::WriteAllText($stop,'')}
            foreach($job in @($observer,$noticeObserver)){
                if($null -ne $job){Wait-Job $job -Timeout 45|Out-Null;if($job.State -ne 'Completed'){Stop-Job $job};Receive-Job $job -ErrorAction SilentlyContinue|Out-Null;Remove-Job $job}
            }
            [void](Probe 'PrintNativeArm' @($false,$false));[void](Probe 'RunLocalSafeClose')
            $excel.Visible=$priorVisible
        }
    }
}
