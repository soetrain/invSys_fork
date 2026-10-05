# Native captured-workbook closure through the actual Complete Run handler.
# Reuse the existing read-return observer; never recreate a dismissed form.
. (Join-Path $PSScriptRoot 'Slice4beProductionCheckInClosed.ps1')
. (Join-Path $PSScriptRoot 'Slice4beProductionCheckInClosedYield.ps1')
function Install-ProductionCompleteClosedProbe {
    Install-ProductionCheckInClosedProbe
    Install-ProductionCheckInClosedYieldProbe
}

function Test-ProductionCompleteClosed($Fixture,$Other,[string]$Path,$Decoy,[string]$Canary,[switch]$ResourceDiagnostic){
    # Probe, Owner and Hash are supplied by the completion baseline scope.
    # Read-only counters locate native resource exhaustion without querying forms.
    . (Join-Path $PSScriptRoot 'Slice4beGuideResourceTrace.ps1')
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class CompleteClosureResources {
    [DllImport("kernel32.dll")] public static extern IntPtr OpenProcess(uint access, bool inherit, int process);
    [DllImport("kernel32.dll")] public static extern bool CloseHandle(IntPtr handle);
    [DllImport("user32.dll")] public static extern uint GetGuiResources(IntPtr process, uint flags);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
}
'@
    [uint32]$processId=0
    [void][CompleteClosureResources]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$processId)
    function MarkClosure([string]$Stage){
        $handle=[CompleteClosureResources]::OpenProcess(1024,$false,[int]$processId)
        if($handle -eq [IntPtr]::Zero){throw 'Closure resource process unavailable.'}
        try{
            $savedHosts=0;$visibleSavedHosts=0
            foreach($openBook in $excel.Workbooks){
                if(-not $openBook.IsAddin -and $openBook.Path -ne ''){
                    $savedHosts++
                    if(@($openBook.Windows|Where-Object Visible).Count){$visibleSavedHosts++}
                }
            }
            $sample=[pscustomobject]@{
                Utc=[DateTime]::UtcNow.ToString('o');Boundary=$boundary;Stage=$Stage
                Gdi=[CompleteClosureResources]::GetGuiResources($handle,0)
                GdiPeak=[CompleteClosureResources]::GetGuiResources($handle,2)
                User=[CompleteClosureResources]::GetGuiResources($handle,1)
                XLMAIN=[InvSysGuideResources]::ExcelWindowCount($processId)
                SavedHosts=$savedHosts;VisibleSavedHosts=$visibleSavedHosts
                Visible=$excel.Visible;ScreenUpdating=$excel.ScreenUpdating
                EnableEvents=$excel.EnableEvents;DisplayAlerts=$excel.DisplayAlerts
                Calculation=$excel.Calculation;CutCopyMode=$excel.CutCopyMode
            }
            $sample|ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'complete-closure-resources.jsonl')
            # Diagnostic bound only; retain the failed trace and ordinary cleanup.
            if($ResourceDiagnostic -and ($sample.Gdi -ge 3000 -or $sample.XLMAIN -ge 80)){
                throw 'Diagnostic resource bound reached; not product acceptance.'
            }
        }finally{[void][CompleteClosureResources]::CloseHandle($handle)}
    }
    $book=$null
    try{
        foreach($boundary in @('Entry','CompletePending','AvailableQuantity','EntityKind')){
            MarkClosure 'BeforeReset'
            [void](Probe 'CheckYieldReset')
            MarkClosure 'BeforeSelectTarget'
            SelectTarget $Fixture 'config-producer'
            MarkClosure 'TargetSelected'
            $casePath=Join-Path (Split-Path $Path) ('complete-closed-'+$boundary.ToLowerInvariant()+'.xlsb')
            if(Test-Path -LiteralPath $casePath){throw 'Preserve existing closure fixture.'}
            Copy-Item -LiteralPath $Path -Destination $casePath
            $book=$excel.Workbooks.Open($casePath,0,$false)
            MarkClosure 'WorkbookOpened'
            [void](Probe 'RunLocalReopen' @($book.Name))
            MarkClosure 'FormCreated'
            if(-not [bool](Probe 'CompleteBaselinePrepare' @($true))){throw 'Real selected Check In prerequisite unavailable; not product RED.'}
            MarkClosure 'Prepared'
            $book.Save();$saved=Hash $casePath
            Initialize-SettingsCapture
            $page=[int](Probe 'RunLocalShowAndCapture' @($book.Name,'CHECK_IN'))
            MarkClosure 'Shown'
            $visibleBefore=([InvSysSettingsCapture]::OwnedVisibleForm('Production',[IntPtr]$excel.Hwnd) -ne [IntPtr]::Zero)
            if(-not $visibleBefore -or $page -ne 3){throw 'Native completion surface unavailable; not product RED.'}
            $Decoy.Activate()
            $ownerBefore=[string](Probe 'RunLocalOwnerState')
            $before=@(Get-Slice4beActivityFiles $Fixture);$otherBefore=@(Get-Slice4beActivityFiles $Other)
            [void](Probe 'CompleteEntryReset' @($false))
            # This also resets the native form-initialization counter.
            [void](Probe 'CheckClosedYieldArm' @($boundary,1,$Decoy.Name))
            $started=[DateTimeOffset]::UtcNow.ToString('o')
            $countsBefore=$excel.Workbooks.Count;$capturedName=$book.Name
            $returned=$true;$invoked=$true;$later=0;$ownerSame=$true
            if($boundary -ceq 'Entry'){
                $projectionBefore=[string](Probe 'RunLocalState')
                $book.Close($false);$book=$null
                $visibleAfter=([InvSysSettingsCapture]::OwnedVisibleForm('Production',[IntPtr]$excel.Hwnd) -ne [IntPtr]::Zero)
                $invoked=$visibleAfter
                if($invoked){$returned=[bool](Probe 'CompleteEntryAct' @(''))}
                $ownerSame=([string](Probe 'RunLocalOwnerState') -ceq $ownerBefore)
            }else{
                MarkClosure 'BeforeHandler'
                $returned=[bool](Probe 'CompleteEntryAct' @(''))
                MarkClosure 'HandlerReturned'
                $evidence=([string](Probe 'CheckYieldEvidence')).Split('|')
                $native=([string](Probe 'CheckClosedYieldReceipt')).Split('|')
                if($evidence.Count -ne 4 -or $evidence[0] -cne 'True' -or $evidence[1] -cne 'True' -or $evidence[3] -cne 'True' -or $native.Count -ne 7 -or $native[0] -cne 'True' -or $native[1] -cne 'True' -or $native[4] -cne 'True' -or $native[5] -cne 'True'){
                    throw ('Actual completion boundary/closure prerequisite unavailable: '+$boundary+'; not product RED.')
                }
                $book=$null;$later=[int]$evidence[2]
                $countsBefore=[int]$native[2]
                $ownerSame=[bool](Probe 'CheckYieldOwnerPreserved')
                $visibleAfter=([InvSysSettingsCapture]::OwnedVisibleForm('Production',[IntPtr]$excel.Hwnd) -ne [IntPtr]::Zero)
            }
            $names=@(foreach($openBook in $excel.Workbooks){[string]$openBook.Name})
            if($capturedName -cin $names -or $Decoy.Name -cnotin $names){throw 'Native captured close or live decoy prerequisite unavailable; not product RED.'}
            $entries=[int](Probe 'CompleteEntryCount')
            if($entries -ne $(if($invoked){1}else{0})){throw 'Actual Complete Run Click entry not established; not product RED.'}
            $projection=$true;$guards=$true;$refusal=$true
            # Query form controls only while the native window still exists.
            if($visibleAfter){
                $projection=if($boundary -ceq 'Entry'){[string](Probe 'RunLocalState') -ceq $projectionBefore}else{[bool](Probe 'CheckClosedYieldFormProjectionPreserved')}
                $guards=[bool](Probe 'CompleteEntryFact' @('GuardsRestored'))
                $refusal=[bool](Probe 'CheckBaselineContextRefused')
                if($CaptureEvidence){CaptureOwnedFormByCaptionEvidence 'Production' ('complete-closed-'+$boundary.ToLowerInvariant()+'.png')}
            }
            $native=([string](Probe 'CheckClosedYieldReceipt')).Split('|')
            $expectedOwners=if($boundary -cin @('Entry','CompletePending')){'0|0'}else{'1|0'}
            $label='CompleteClosed.'+$boundary
            Check ($label+'.NativeWorkbookRemoved') ($countsBefore -eq $excel.Workbooks.Count+1)
            Check ($label+'.NativeActionOpportunityProtected') ($returned -and $refusal)
            Check ($label+'.NoLaterReadAttempts') ($later -eq 0)
            Check ($label+'.CompletionOwnerBoundary') ([string](Probe 'CompleteBaselineOwners') -ceq $expectedOwners)
            Check ($label+'.OwnerAtBoundaryPreserved') $ownerSame
            Check ($label+'.CapturedBindingRejected') (-not [bool](Probe 'RunLocalClosedBindingCurrent'))
            Check ($label+'.SurvivingProjectionPreserved') $projection
            Check ($label+'.SurvivingGuardsRestored') $guards
            Check ($label+'.NoFormReinitialization') ($native.Count -eq 7 -and [int]$native[6] -eq 0)
            Check ($label+'.ExactInputBalancesPreserved') ([bool](Owner 'CompleteBaselineBalancesForTest' @($false)))
            Check ($label+'.SavedOperatorBytesPreserved') ((Hash $casePath) -ceq $saved)
            Check ($label+'.DecoyPreserved') ($Decoy.Worksheets.Count -eq 1 -and $Decoy.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq $Canary)
            Check ($label+'.NoActivityOrRedirectedRecords') ((@(Get-Slice4beActivityFiles $Fixture) -join '|') -ceq ($before -join '|') -and (@(Get-Slice4beActivityFiles $Other) -join '|') -ceq ($otherBefore -join '|'))
            [pscustomobject]@{Boundary=$boundary;StartUTC=$started;EndUTC=[DateTimeOffset]::UtcNow.ToString('o');VisibleBefore=$visibleBefore;VisibleAfter=$visibleAfter;HandlerInvoked=$invoked;HandlerEntries=$entries;HandlerReturned=$returned;WorkbooksBefore=$countsBefore;WorkbooksAfter=$excel.Workbooks.Count;CapturedStillOpen=$false;DecoyStillOpen=$true;LaterReadAttempts=$later;FormInitializationsAfterArm=[int]$native[6];DismissedFormControlsQueried=$false;SurvivingControlAssertionsConditional=$true}|ConvertTo-Json|Set-Content (Join-Path $reportRoot ('complete-closed-'+$boundary.ToLowerInvariant()+'.json'))
            MarkClosure 'AssertionsComplete'
            [void](Probe 'CheckYieldReset')
            MarkClosure 'BeforeSafeClose'
            [void](Probe 'RunLocalSafeClose')
            MarkClosure 'SafeCloseReturned'
        }
    }finally{
        [void](Probe 'CheckYieldReset');[void](Probe 'RunLocalSafeClose')
        if($null -ne $book){try{$book.Close($false)}catch{}}
        SelectTarget $Fixture 'config-producer'
    }
}
