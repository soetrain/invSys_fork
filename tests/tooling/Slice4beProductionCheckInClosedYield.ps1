# Close the actual captured workbook inside the already-running operator handler.
# Unsaved probes observe native lifetime; they never recreate a dismissed form.
function Install-ProductionCheckInClosedYieldProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $adapter=$project.VBComponents.Item('TestProductionDesigner').CodeModule
    $start=$adapter.ProcStartLine('CheckYieldReturned',0);$end=$start+$adapter.ProcCountLines('CheckYieldReturned',0)
    $hits=@(for($line=$start;$line -lt $end;$line++){if($adapter.Lines($line,1).Trim() -ceq 'If mCheckYieldInterruption = "Permission" Then'){$line}})
    if($hits.Count -ne 1){throw 'Native Check In interruption anchor changed; not product RED.'}
    $adapter.InsertLines($hits[0],@'
    If mCheckYieldInterruption = "ClosedWorkbook" Then
        CheckClosedYieldClose
        mCheckYieldValid = mCheckClosedYieldRemoved And mCheckClosedYieldDecoyOpen And _
            Not modProductionDesignerActions.ContextIsCurrent(mRunLocalCapturedContextForTest, mRunLocalCapturedBookForTest)
        Exit Sub
    End If
'@)
    $adapter.InsertLines(1,@'
Private mCheckClosedYieldDecoy As String, mCheckClosedYieldRemoved As Boolean
Private mCheckClosedYieldDecoyOpen As Boolean, mCheckClosedYieldVisible As Boolean
Private mCheckClosedYieldBefore As Long, mCheckClosedYieldAfter As Long
Private mCheckClosedYieldContext As Boolean, mCheckClosedYieldInitializations As Long
Private mCheckClosedYieldFormProjection As String, mCheckClosedYieldWorksheetProjection As String
'@)
    $adapter.AddFromString(@'
Public Sub CheckClosedYieldInitialized()
    mCheckClosedYieldInitializations = mCheckClosedYieldInitializations + 1
End Sub
Public Sub CheckClosedYieldArm(ByVal boundary As String, ByVal ordinal As Long, ByVal decoyName As String)
    CheckYieldArm boundary, ordinal, "ClosedWorkbook", ""
    CheckClosedReset
    mCheckClosedYieldDecoy = decoyName: mCheckClosedYieldRemoved = False
    mCheckClosedYieldDecoyOpen = False: mCheckClosedYieldVisible = False
    mCheckClosedYieldBefore = 0: mCheckClosedYieldAfter = 0
    mCheckClosedYieldContext = False: mCheckClosedYieldInitializations = 0
    mCheckClosedYieldFormProjection = "": mCheckClosedYieldWorksheetProjection = ""
End Sub
Private Sub CheckClosedYieldClose()
    Dim wb As Workbook, capturedName As String
    capturedName = mRunLocalCapturedBookForTest.Name
    mCheckClosedYieldVisible = mForm.Visible
    mCheckClosedYieldContext = modProductionDesignerActions.ContextIsCurrent( _
        mRunLocalCapturedContextForTest, mRunLocalCapturedBookForTest)
    mCheckClosedYieldBefore = Application.Workbooks.Count
    mCheckClosedYieldFormProjection = mForm.CheckClosedYieldFormProjectionForTest()
    mCheckClosedYieldWorksheetProjection = mForm.CheckBaselineWorksheetStateForTest()
    mRunLocalCapturedBookForTest.Close False
    mCheckClosedYieldAfter = Application.Workbooks.Count
    mCheckClosedYieldRemoved = True
    For Each wb In Application.Workbooks
        If StrComp(wb.Name, capturedName, vbBinaryCompare) = 0 Then mCheckClosedYieldRemoved = False
        If StrComp(wb.Name, mCheckClosedYieldDecoy, vbBinaryCompare) = 0 Then mCheckClosedYieldDecoyOpen = True
    Next wb
End Sub
Public Function CheckClosedYieldReceipt() As String
    CheckClosedYieldReceipt = CStr(mCheckClosedYieldVisible) & "|" & CStr(mCheckClosedYieldContext) & "|" & _
        CStr(mCheckClosedYieldBefore) & "|" & CStr(mCheckClosedYieldAfter) & "|" & _
        CStr(mCheckClosedYieldRemoved) & "|" & CStr(mCheckClosedYieldDecoyOpen) & "|" & _
        CStr(mCheckClosedYieldInitializations)
End Function
Public Function CheckClosedYieldFormProjectionPreserved() As Boolean
    CheckClosedYieldFormProjectionPreserved = (mForm.CheckClosedYieldFormProjectionForTest() = mCheckClosedYieldFormProjection)
End Function
Public Function CheckClosedYieldProjectionEvidence() As String
    CheckClosedYieldProjectionEvidence = CStr(mCheckClosedYieldWorksheetProjection <> "") & "|" & _
        CStr(mForm.CheckBaselineWorksheetStateForTest() = "") & "|" & CStr(CheckYieldProjectionPreserved())
End Function
'@)
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    $projection=$form.Lines($form.ProcStartLine('CheckYieldProjectionForTest',0),$form.ProcCountLines('CheckYieldProjectionForTest',0))
    $composite='RunLocalStateForTest() & "|" & CheckBaselineWorksheetStateForTest() & "|" & _'
    if([regex]::Matches($projection,[regex]::Escape($composite)).Count -ne 1){throw 'Form/worksheet projection separation anchor changed.'}
    # The closed worksheet is protected by saved-byte checks, not by a surviving-control query.
    $projection=$projection.Replace($composite,'RunLocalStateForTest() & "|" & _').Replace('CheckYieldProjectionForTest','CheckClosedYieldFormProjectionForTest')
    $form.AddFromString($projection)
    $line=$form.ProcBodyLine('UserForm_Initialize',0)
    $form.InsertLines($line+1,'    TestProductionDesigner.CheckClosedYieldInitialized')
}

function Test-ProductionCheckInClosedYield($Fixture,$Other,[string]$Path,$Decoy,[string]$Canary,[switch]$Routed){
    # Probe and Hash are supplied by the existing native-entry fixture scope.
    $cases=@(foreach($boundary in @('AvailableQuantity','EntityKind','RunPalette','ManagerCheck')){@{Mode='Reusable';Boundary=$boundary;Ordinal=1}})
    $cases+=@(@{Mode='Worksheet';Boundary='ResolveKey';Ordinal=1},@{Mode='Worksheet';Boundary='ResolveKey';Ordinal=2})
    foreach($boundary in @('InventoryPicker','DefaultLocation','IngredientChoices')){$cases+=@{Mode='Worksheet';Boundary=$boundary;Ordinal=1}}
    if($Routed){$cases=@(foreach($ordinal in 1,2){@{Mode='Routed';Boundary='AvailableQuantity';Ordinal=$ordinal}})}
    $book=$null
    try{
        foreach($case in $cases){
            [void](Probe 'CheckYieldReset')
            SelectTarget $Fixture 'config-producer'
            $casePath=Join-Path (Split-Path $Path) ('check-in-closed-yield-'+$case.Mode+'-'+$case.Boundary+'-'+$case.Ordinal+'.xlsb')
            if(Test-Path -LiteralPath $casePath){throw 'Native closure fixture destination already exists.'}
            Copy-Item -LiteralPath $Path -Destination $casePath
            $book=$excel.Workbooks.Open($casePath,0,$false)
            $Decoy.Activate()
            if($Routed){[void](Probe 'OpenDesigner' @($book.Name))}else{[void](Probe 'RunLocalReopen' @($book.Name))}
            if($case.Mode -ceq 'Worksheet'){
                if(-not [bool](Probe 'RunWorksheetPrepare')){throw 'Native read-return worksheet fixture unavailable; not product RED.'}
                $selectedKey=[string](Probe 'RunWorksheetKey')
                if(-not $selectedKey){throw 'Native read-return exact key unavailable; not product RED.'}
                $ready=[bool](Probe 'CheckBaselineWorksheetStage' @($selectedKey,$Canary))
                # Exercise both genuine Domain identity reads, without a local shortcut.
                foreach($local in @($book.Worksheets)){if($local.Name -ceq 'InventoryManagement'){$local.Name='CheckYieldLocalSource'}}
                [void](Probe 'CheckYieldResetCache')
            }elseif($Routed){$ready=[bool](Probe 'CheckRoutedStage')}
            else{$ready=[bool](Probe 'CheckBaselineReusableStage' @('Selected'))}
            if(-not $ready){throw 'Native read-return staging unavailable; not product RED.'}
            $book.Save();$saved=Hash $casePath
            Initialize-SettingsCapture
            $page=[int](Probe 'RunLocalShowAndCapture' @($book.Name,'CHECK_IN'))
            $visibleBefore=([InvSysSettingsCapture]::OwnedVisibleForm('Production',[IntPtr]$excel.Hwnd) -ne [IntPtr]::Zero)
            if(-not $visibleBefore -or $page -ne 3){throw 'Native Check In surface unavailable; not product RED.'}
            $before=@(Get-Slice4beActivityFiles $Fixture);$otherBefore=@(Get-Slice4beActivityFiles $Other)
            [void](Probe 'CheckClosedYieldArm' @($case.Boundary,$case.Ordinal,$Decoy.Name))
            $started=[DateTimeOffset]::UtcNow.ToString('o')
            $returned=[bool](Probe 'CheckBaselineAct' @(''))
            $evidence=([string](Probe 'CheckYieldEvidence')).Split('|')
            $native=([string](Probe 'CheckClosedYieldReceipt')).Split('|')
            if($evidence.Count -ne 4 -or $evidence[0] -cne 'True' -or $evidence[1] -cne 'True' -or $evidence[3] -cne 'True' -or $native.Count -ne 7 -or $native[0] -cne 'True' -or $native[1] -cne 'True' -or $native[4] -cne 'True' -or $native[5] -cne 'True'){
                throw ('Native Check In real read/closure fixture unavailable: '+$case.Mode+'/'+$case.Boundary+'; not product RED.')
            }
            $book=$null
            $visibleAfter=([InvSysSettingsCapture]::OwnedVisibleForm('Production',[IntPtr]$excel.Hwnd) -ne [IntPtr]::Zero)
            $entries=[int](Probe 'CheckClosedEntries')
            if($entries -ne 1){throw 'Native Check In handler entry not proved; not product RED.'}
            $projection=$true;$guards=$true;$refusal=$true;$projectionEvidence=@()
            # Never query controls on a dismissed form: such a query can reinitialize it.
            if($visibleAfter){
                $projection=[bool](Probe 'CheckClosedYieldFormProjectionPreserved')
                $projectionEvidence=([string](Probe 'CheckClosedYieldProjectionEvidence')).Split('|')
                if($projectionEvidence.Count -ne 3){throw 'Native projection diagnostic envelope unavailable.'}
                $guards=[bool](Probe 'CheckBaselineGuards')
                $refusal=[bool](Probe 'CheckBaselineContextRefused')
                if($CaptureEvidence -and $case.Ordinal -eq 1 -and $case.Boundary -cin @('AvailableQuantity','ResolveKey')){
                    CaptureOwnedFormByCaptionEvidence 'Production' ('check-in-closed-read-'+$case.Mode.ToLowerInvariant()+'.png')
                }
            }
            $label='CheckInClosedYield.'+$case.Mode+'.'+$case.Boundary+'.'+$case.Ordinal
            Check ($label+'.ActualHandlerReturned') $returned
            Check ($label+'.NativeWorkbookRemoved') ([int]$native[2] -eq [int]$native[3]+1)
            Check ($label+'.NoLaterReadAttempts') ([int]$evidence[2] -eq 0)
            Check ($label+'.OwnerAtBoundaryPreserved') ([bool](Probe 'CheckYieldOwnerPreserved'))
            Check ($label+'.CapturedBindingRejected') (-not [bool](Probe 'RunLocalClosedBindingCurrent'))
            Check ($label+'.SurvivingProjectionPreserved') $projection
            Check ($label+'.SurvivingGuardsRestored') $guards
            Check ($label+'.DismissedOrVisibleRefusal') $refusal
            Check ($label+'.SavedOperatorBytesPreserved') ((Hash $casePath) -ceq $saved)
            Test-ProductionCheckInJournal $Fixture $Other $before $otherBefore 'FAILED' $label $Canary
            [pscustomobject]@{Mode=$case.Mode;Boundary=$case.Boundary;Ordinal=$case.Ordinal;StartUTC=$started;EndUTC=[DateTimeOffset]::UtcNow.ToString('o');VisibleBefore=$visibleBefore;VisibleAtReadReturn=$true;VisibleAfter=$visibleAfter;HandlerEntries=$entries;HandlerAlreadyRunningBeforeClosure=$true;HandlerReturned=$returned;WorkbooksAtReadReturn=[int]$native[2];WorkbooksAfterClose=[int]$native[3];CapturedStillOpen=$false;DecoyStillOpen=$true;LaterReadAttempts=[int]$evidence[2];FormInitializationsAfterArm=[int]$native[6];DismissedFormControlsQueried=$false;SurvivingControlAssertionsConditional=$true;FormProjectionPreserved=$projection;WorksheetProjectionBeforeNonempty=($projectionEvidence.Count -eq 3 -and $projectionEvidence[0] -ceq 'True');WorksheetProjectionAfterEmpty=($projectionEvidence.Count -eq 3 -and $projectionEvidence[1] -ceq 'True');OriginalCompositeProjectionEqual=($projectionEvidence.Count -eq 3 -and $projectionEvidence[2] -ceq 'True')}|ConvertTo-Json|Set-Content (Join-Path $reportRoot ('check-in-closed-yield-'+$case.Mode.ToLowerInvariant()+'-'+$case.Boundary.ToLowerInvariant()+'-'+$case.Ordinal+'.json'))
            [void](Probe 'CheckYieldReset');[void](Probe 'RunLocalSafeClose')
        }
    }finally{
        [void](Probe 'CheckYieldReset');[void](Probe 'RunLocalSafeClose')
        if($null -ne $book){try{$book.Close($false)}catch{}}
        SelectTarget $Fixture 'config-producer'
    }
}
