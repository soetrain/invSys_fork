# Install only into unsaved packages, before any form is constructed.
function Install-ProductionLifecycleNativeProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $project.VBComponents.Item('frmProduction').CodeModule.AddFromString(@'
Public Function NativeLifecyclePresentForTest(ByVal designer As String, ByVal action As String) As Boolean
    Dim button As MSForms.CommandButton
    If designer = "Process" Then
        mPages.Value = 0
        If action = "Release" Then Set button = mBtnProcessRelease
        If action = "Obsolete" Then Set button = mBtnProcessObsolete
    ElseIf designer = "Recipe" Then
        mPages.Value = 1
        If action = "Release" Then Set button = mBtnRecipeRelease
        If action = "Obsolete" Then Set button = mBtnRecipeObsolete
    End If
    If button Is Nothing Then Err.Raise 5
    If Not Me.Visible Then Me.Show vbModeless
    Me.Repaint: DoEvents
    NativeLifecyclePresentForTest = Me.Visible And mPages.Visible And button.Visible And button.Enabled
End Function
Public Function NativeLifecycleCancelForTest(ByVal designer As String, ByVal action As String) As String
    Dim previousTest As Boolean
    previousTest = mReusableActionTestInProgress
    On Error GoTo Failed
    mReusableActionTestInProgress = False
    If designer = "Process" And action = "Release" Then
        mBtnProcessRelease_Click
    ElseIf designer = "Process" And action = "Obsolete" Then
        mBtnProcessObsolete_Click
    ElseIf designer = "Recipe" And action = "Release" Then
        mBtnRecipeRelease_Click
    ElseIf designer = "Recipe" And action = "Obsolete" Then
        mBtnRecipeObsolete_Click
    Else
        Err.Raise 5
    End If
    NativeLifecycleCancelForTest = "RETURNED"
Done:
    mReusableActionTestInProgress = previousTest
    Exit Function
Failed:
    NativeLifecycleCancelForTest = "HANDLER_ERROR"
    Resume Done
End Function
Public Function NativeLifecycleBoundForTest(ByVal wb As Workbook) As Boolean
    NativeLifecycleBoundForTest = (mOperatorWorkbook Is wb)
End Function
'@)
    $project.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Function NativePresent(ByVal designer As String, ByVal action As String) As Boolean
    NativePresent = mForm.NativeLifecyclePresentForTest(designer, action)
End Function
Public Function NativeCancel(ByVal designer As String, ByVal action As String) As String
    NativeCancel = mForm.NativeLifecycleCancelForTest(designer, action)
End Function
Public Function NativeBound(ByVal workbookName As String) As Boolean
    NativeBound = mForm.NativeLifecycleBoundForTest(Application.Workbooks(workbookName))
End Function
'@)
}

function Test-ProductionLifecycleNative($Fixture) {
    . (Join-Path $PSScriptRoot 'Slice4beProductionNativeChoice.ps1')
    function Probe([string]$Method,[object[]]$Values=@()) {
        Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values
    }
    function SourceWire {
        $wire=[string](Run 'invSys.Designs.Domain.xlam' 'modDesignsBridgeApi.ReadDesignsQueryBridgeResult' @('PUBLICATION_EVENTS',$Fixture.Warehouse,$Fixture.Root))
        if($wire -cnotmatch '^EVTSRC1\tDesigns\tAvailable\t'){throw 'Owning source prerequisite unavailable; not behavioral RED.'}
        return $wire
    }
    function ApplyPrerequisite([string]$Designer,[string]$Action,[string]$State) {
        $delivered=[string](Probe 'LifecycleAct' @($Designer,$Action))
        $valid=$delivered -ceq 'RETURNED' -and [string](Probe 'LifecycleStatus' @($Designer)) -ceq $State
        Check ('NativeCancel.Setup.'+$Designer+'.'+$Action) $valid
        if(-not $valid){throw 'Owning definition prerequisite unavailable; not cancellation RED.'}
    }
    function Decline([string]$Designer,[string]$Action,[string]$State) {
        $label='NativeCancel.'+$Designer+'.'+$Action
        $id='PRODUCTION_'+$Designer.ToUpperInvariant()+'_'+$Action.ToUpperInvariant()
        $ready=[bool](Probe 'NativePresent' @($Designer,$Action))
        Check ($label+'.ActualControlVisibleAndEnabled') $ready
        if(-not $ready){throw 'Native owning control is unreachable; inspect the fixture before claiming RED.'}
        $draft=[string](Probe 'State' @($Designer));$source=SourceWire
        $before=@(Get-Slice4beActivityFiles $Fixture)
        $decoy.Activate()
        $boundBefore=[bool](Probe 'NativeBound' @($book.Name))
        $evidence=Invoke-ProductionNativeCancellation $Designer $Action ($Designer.ToLowerInvariant()+'-'+$Action.ToLowerInvariant()+'-cancel.png')
        Check ($label+'.ExactNativeQuestionAndNo') ($evidence.HandlerReturned -and $evidence.ExactQuestion -and $evidence.NoClickDelivered)
        Check ($label+'.DefaultChoiceIsNo') $evidence.DefaultNo
        Check ($label+'.VisibleQuestionCaptured') $evidence.Captured
        Check ($label+'.NoOwningSourceChange') ((SourceWire) -ceq $source -and [string](Probe 'LifecycleStatus' @($Designer)) -ceq $State)
        Check ($label+'.LocalDraftPreserved') ([string](Probe 'State' @($Designer)) -ceq $draft)
        $raw=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {[IO.File]::ReadAllText($_)})
        $rows=@($raw|ForEach-Object {$_|ConvertFrom-Json})
        $attempt=@($rows|Where-Object OutcomeCode -CEQ 'REQUESTED');$cancel=@($rows|Where-Object OutcomeCode -CEQ 'CANCELLED')
        $pair=$rows.Count -eq 2 -and $attempt.Count -eq 1 -and $cancel.Count -eq 1
        Check ($label+'.RequestedAndCancelled') $pair
        $linked=$pair;$facts=$pair;$safe=$pair
        if($pair){
            $linked=$attempt[0].ActivityId -cne '' -and $attempt[0].ActivityId -ceq $cancel[0].ActivityId -and $attempt[0].RecordId -cne $cancel[0].RecordId
            $facts=$attempt[0].DataEffect -ceq 'Unknown' -and $cancel[0].DataEffect -ceq 'Unchanged' -and $cancel[0].Severity -ceq 'Notice' -and $cancel[0].EventCode -ceq ($id+'_CANCELLED')
            foreach($row in $rows){$facts=$facts -and $row.ControlId -ceq $id -and $row.OwnerId -ceq 'PRODUCTION_DESIGN_LIFECYCLE' -and $row.UserId -ceq 'config-producer' -and $row.WarehouseId -ceq $Fixture.Warehouse -and $row.CatalogVersion -eq 13 -and @($row.SourceEventRefs).Count -eq 0}
        }
        foreach($text in $raw){foreach($value in @($canary,$Fixture.Secret,(CredentialHash $Fixture.Secret),$Fixture.Root,'mBtn','PayloadJson')){if($text.Contains($value)){$safe=$false}}}
        Check ($label+'.DistinctCorrelatedRecords') $linked
        Check ($label+'.ExactCancelledFactsWithoutReferences') $facts
        Check ($label+'.NoEnteredDataOrSecrets') $safe
        Check ($label+'.CancellationNotice') ([string](Probe 'LifecycleNotice') -ceq 'Design lifecycle action cancelled.')
        Check ($label+'.CapturedWorkbookAndUnknownColumnPreserved') ($boundBefore -and [bool](Probe 'NativeBound' @($book.Name)) -and $sheet.Cells.Item(1,1).Value2 -ceq 'Operator Extra' -and $sheet.Cells.Item(2,1).Value2 -ceq $canary -and $decoy.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'Decoy preserved')
    }
    $canary='NATIVE'+[guid]::NewGuid().ToString('N');$book=$null;$decoy=$null
    try {
        SelectTarget $Fixture
        try {
            [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
            if(-not [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveSettings' @('DesignsEnabled','TRUE'))){throw 'Explicit Designs setup failed.'}
        } finally {[void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')}
        SelectTarget $Fixture 'config-producer'
        $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
        $sheet.Cells.Item(1,1).Value2='Operator Extra';$sheet.Cells.Item(2,1).Value2=$canary
        $book.SaveAs((Join-Path $runRoot 'native-cancellation-operator.xlsb'),50)
        $decoy=$excel.Workbooks.Add();$decoy.Worksheets.Item(1).Cells.Item(1,1).Value2='Decoy preserved'
        [void](Probe 'OpenDesigner' @($book.Name))
        $ready=[string](Probe 'LifecyclePrepare' @($canary))
        if($ready -cnotmatch '^READY\|([^|]+)\|([^|]+)$'){throw 'Process draft prerequisite unavailable.'}
        $processId=$Matches[1];$processVersion=$Matches[2]
        ApplyPrerequisite 'Process' 'Save' 'DRAFT'
        Decline 'Process' 'Release' 'DRAFT'
        ApplyPrerequisite 'Process' 'Release' 'RELEASED'
        Decline 'Process' 'Obsolete' 'RELEASED'
        if(-not [bool](Probe 'ReleasedRecipe' @($processId,$processVersion,$canary))){throw 'Released Process prerequisite unavailable.'}
        ApplyPrerequisite 'Recipe' 'Save' 'DRAFT'
        Decline 'Recipe' 'Release' 'DRAFT'
        ApplyPrerequisite 'Recipe' 'Release' 'RELEASED'
        Decline 'Recipe' 'Obsolete' 'RELEASED'
    } finally {
        [void](Probe 'CloseDesigner')
        if($null -ne $book){$book.Close($false)}
        if($null -ne $decoy){$decoy.Close($false)}
        SelectTarget $Fixture
    }
}
