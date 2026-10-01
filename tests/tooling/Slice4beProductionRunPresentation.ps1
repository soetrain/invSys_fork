# Unsaved fixture adapters exercise the packaged Click handlers unchanged.
# This gate covers only the presentation pair, not all nine catalog24 controls.
function Install-ProductionRunPresentationProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $project.VBComponents.Item('frmProduction').CodeModule.AddFromString(@'
Public Function RunPresentationStageForTest(ByVal mode As String, ByVal canary As String) As String
    Dim i As Long, prior As Boolean, stage As String
    On Error GoTo Failed
    prior = mLoading: mLoading = True
    stage = "ClearPalette"
    mLstRunPalette.Clear
    stage = "ClearTree": mLstRunTree.Clear
    stage = "ClearOutput": mLstManagerOutput.Clear
    stage = "ClearProcess": mCmbRunProcess.Clear
    stage = "ClearTreeProcess": mCmbTreeRunProcess.Clear
    stage = "TreeState"
    EnsureRunTreeState
    mRunTreeCollapsed.RemoveAll
    If mode <> "Empty" Then
        For i = 0 To 1
            stage = "PaletteValues"
            mLstRunPalette.AddItem "RUN-FIXTURE"
            mLstRunPalette.List(i, 1) = "1"
            mLstRunPalette.List(i, 2) = canary
            mLstRunPalette.List(i, 3) = "RUN-KEY-" & CStr(i)
            mLstRunPalette.List(i, 4) = canary
            mLstRunPalette.List(i, 5) = "25"
            mLstRunPalette.List(i, 6) = "1"
            mLstRunPalette.List(i, 7) = "EA"
            mLstRunPalette.List(i, 8) = "8"
            mLstRunPalette.List(i, 9) = "RUN-LOCATION"
            stage = "ProcessMap": StoreRunProcess mLstRunPalette, i, canary
        Next i
        If mode = "Collapsed" Then mRunTreeCollapsed(RunTreeGroupKey(mLstRunPalette, 0)) = True
        If mode = "ProcessCollapsed" Then mRunTreeCollapsed("PROC|" & ProcessKey(canary)) = True
    End If
    stage = "BuildTree": BuildRunTreeFromPaletteList
    mLoading = prior
    RunPresentationStageForTest = "READY"
    Exit Function
Failed:
    mLoading = prior
    RunPresentationStageForTest = "FIXTURE_FAILED|" & stage & "|" & CStr(Err.Number)
End Function
Public Function RunPresentationActForTest(ByVal action As String, ByVal guard As String) As String
    Dim priorLoading As Boolean, priorBusy As Boolean
    On Error GoTo Failed
    priorLoading = mLoading: priorBusy = mDesignerActionInProgress
    If guard = "Loading" Then mLoading = True
    If guard = "Busy" Then mDesignerActionInProgress = True
    Select Case action
        Case "TREE_EXPAND": mBtnRunTreeExpandAll_Click
        Case "TREE_COLLAPSE": mBtnRunTreeCollapseAll_Click
        Case Else: Err.Raise 5
    End Select
    RunPresentationActForTest = mTxtStatus.Text
Finished:
    mLoading = priorLoading: mDesignerActionInProgress = priorBusy
    Exit Function
Failed:
    RunPresentationActForTest = "HANDLER_ERROR|" & CStr(Err.Number)
    Resume Finished
End Function
Public Function RunPresentationStateForTest(ByVal paletteOnly As Boolean) As String
    Dim i As Long, c As Long, key As Variant, result As String
    For i = 0 To mLstRunPalette.ListCount - 1
        For c = 0 To mLstRunPalette.ColumnCount - 1
            result = result & "|" & NzStr(mLstRunPalette.List(i, c))
        Next c
    Next i
    If Not paletteOnly Then
        For Each key In mRunTreeCollapsed.Keys
            result = result & "|COLLAPSED|" & CStr(key)
        Next key
        For i = 0 To mLstRunTree.ListCount - 1
            result = result & "|TREE|" & NzStr(mLstRunTree.List(i, 0)) & "|" & NzStr(mLstRunTree.List(i, 2))
        Next i
    End If
    RunPresentationStateForTest = result
End Function
Public Function RunPresentationShapeForTest() As String
    RunPresentationShapeForTest = CStr(mLstRunPalette.ListCount) & "|" & CStr(mLstRunTree.ListCount) & "|" & CStr(mRunTreeCollapsed.Count)
End Function
'@)
    $project.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Function RunPresentationStage(ByVal mode As String, ByVal canary As String) As String
    On Error GoTo Failed
    RunPresentationStage = mForm.RunPresentationStageForTest(mode, canary)
    Exit Function
Failed:
    RunPresentationStage = "FIXTURE_FAILED|FormBinding|" & CStr(Err.Number)
End Function
Public Function RunPresentationAct(ByVal action As String, ByVal guard As String) As String
    RunPresentationAct = mForm.RunPresentationActForTest(action, guard)
End Function
Public Function RunPresentationState(ByVal paletteOnly As Boolean) As String
    RunPresentationState = mForm.RunPresentationStateForTest(paletteOnly)
End Function
Public Function RunPresentationShape() As String
    RunPresentationShape = mForm.RunPresentationShapeForTest()
End Function
'@)
    $packages['invSys.Core.xlam'].VBProject.VBComponents.Item('TestShippingCatalog').CodeModule.AddFromString(@'
Public Function RunPresentationPolicy(ByVal navigation As Boolean) As Boolean
    Dim context As String, version As Long, request As String, report As String, model As Object, row As Variant
    context = modActivity.CaptureContext()
    If Not modTrackingPolicySettings.ReadEditor(context, version, request, report) Then Exit Function
    Set model = modTrackingPolicyModel.Defaults(True)
    For Each row In model("Controls")
        If CStr(row("ControlId")) = "PRODUCTION_RUN_TREE_EXPAND" Or CStr(row("ControlId")) = "PRODUCTION_RUN_TREE_COLLAPSE" Then row("Collect") = navigation
    Next row
    request = modTrainingJson.EncodeObject(model)
    RunPresentationPolicy = modTrackingPolicySettings.SavePolicy(context, version, request, report)
End Function
Public Function RunPresentationTerminal(ByVal recordJson As String) As Boolean
    Dim record As Object
    Set record = modTrainingJson.DecodeObject(recordJson)
    RunPresentationTerminal = modEvaluationMatches.CommandCompleted(record)
End Function
'@)
}

function Test-ProductionRunPresentation($Fixture,$Other) {
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Files {@(Get-Slice4beActivityFiles $Fixture)}
    function Hash([string]$Path){$s=[IO.File]::Open($Path,'Open','Read','ReadWrite');try{(Get-FileHash -InputStream $s).Hash}finally{$s.Dispose()}}
    function Stage([string]$Mode){
        $ready=[string](Probe 'RunPresentationStage' @($Mode,$canary))
        if($ready -cne 'READY'){
            if($ready -notmatch '^FIXTURE_FAILED\|[A-Za-z]+\|-?[0-9]+$'){$ready='Unavailable'}
            throw ('Run presentation fixture unavailable: '+$ready+'; not product RED.')
        }
    }
    function State([bool]$PaletteOnly=$false){[string](Probe 'RunPresentationState' @($PaletteOnly))}
    function Pair([string[]]$Before,[string]$Action,[string]$Case){
        $raw=@(Files|Where-Object{$_ -cnotin $Before}|ForEach-Object{[IO.File]::ReadAllText($_)})
        $records=@($raw|ForEach-Object{$_|ConvertFrom-Json})
        $first=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$last=@($records|Where-Object OutcomeCode -CEQ 'PRESENTED')
        $paired=$records.Count -eq 2 -and $first.Count -eq 1 -and $last.Count -eq 1
        $context=$paired;$safe=$paired;$integrity=$paired;$linked=$false;$facts=$false;$terminal=$false
        $control='PRODUCTION_RUN_'+$Action
        foreach($r in $records){
            $context=$context -and $r.ControlId -ceq $control -and $r.OwnerId -ceq 'PRODUCTION_RUN_LOCAL' -and $r.UserId -ceq 'config-producer' -and $r.WarehouseId -ceq $Fixture.Warehouse -and $r.StationId -ceq 'S1' -and $r.CatalogVersion -eq 24
            $safe=$safe -and @($r.SourceEventRefs).Count -eq 0
        }
        foreach($value in $raw){
            foreach($secret in @($canary,$Fixture.Secret,(CredentialHash $Fixture.Secret),$Fixture.Root,'mBtn','PinHash','RUN-KEY-','RUN-LOCATION','RUN-FIXTURE')){
                $encoded=ConvertTo-Json -InputObject $secret -Compress
                if($value.Contains($secret) -or $value.Contains($encoded.Substring(1,$encoded.Length-2))){$safe=$false}
            }
            $match=[regex]::Match($value,',"ContentSha256":"([a-f0-9]{64})"\}$')
            if(-not $match.Success){$integrity=$false;continue}
            $sha=[Security.Cryptography.SHA256]::Create()
            try{$hash=[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($value.Substring(0,$match.Index)+'}'))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
            $integrity=$integrity -and $hash -ceq $match.Groups[1].Value
        }
        if($paired){
            $linked=$first[0].ActivityId -cne '' -and $first[0].ActivityId -ceq $last[0].ActivityId -and $first[0].RecordId -cne $last[0].RecordId
            $facts=$first[0].DataEffect -ceq 'Unknown' -and $first[0].Severity -ceq 'Info' -and $last[0].DataEffect -ceq 'Unchanged' -and $last[0].Severity -ceq 'Info' -and $last[0].EventCode -ceq ($control+'_PRESENTED')
            $terminal=-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.RunPresentationTerminal' @(($first[0]|ConvertTo-Json -Depth 20 -Compress))) -and [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.RunPresentationTerminal' @(($last[0]|ConvertTo-Json -Depth 20 -Compress)))
        }
        Check ($Case+'.AttemptAndOutcome') $paired
        Check ($Case+'.ExactContext') $context
        Check ($Case+'.NoEnteredDataOrSources') $safe
        Check ($Case+'.Integrity') $integrity
        Check ($Case+'.LinkedDistinctRecords') $linked
        Check ($Case+'.OwnerFact') $facts
        Check ($Case+'.ExactCommandTerminal') $terminal
    }
    $canary='RUN'+[guid]::NewGuid().ToString('N');$book=$null;$decoy=$null;$pins=@{};$recordPins=@{}
    SelectTarget $Fixture
    if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.RunPresentationPolicy' @($true))){throw 'Authorized Run policy fixture unavailable; not product RED.'}
    try{
        SelectTarget $Fixture 'config-producer'
        $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
        $sheet.Cells.Item(1,1).Value2='Operator Extra';$sheet.Cells.Item(2,1).Value2=$canary;$sheet.Cells.Item(2,2).Formula='=1+2'
        $path=Join-Path $runRoot 'run-presentation-operator.xlsb';$book.SaveAs($path,50);$book.Close($false)
        $bookPin=Hash $path;$book=$excel.Workbooks.Open($path,0,$false);$sheet=$book.Worksheets.Item(1)
        $decoy=$excel.Workbooks.Add();$decoy.Activate();[void](Probe 'OpenDesigner' @($book.Name))
        foreach($file in Get-ChildItem -LiteralPath $Fixture.Root -Recurse -File|Where-Object{$_.Extension -in '.xlsb','.xlsm' -and $_.Name -notlike '~$*'}){$pins[$file.FullName]=Hash $file.FullName}
        foreach($file in Files){$recordPins[$file]=Hash $file}
        foreach($action in @('TREE_EXPAND','TREE_COLLAPSE')){
            $json=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @(('PRODUCTION_RUN_'+$action),24));$def=if($json){$json|ConvertFrom-Json}else{$null}
            $caption=if($action -ceq 'TREE_EXPAND'){'Expand'}else{'Collapse'}
            Check ('Run.'+$action+'.FixedMetadata') ($null -ne $def -and $def.Caption -ceq $caption -and $def.OwnerId -ceq 'PRODUCTION_RUN_LOCAL' -and $def.Class -ceq 'Navigation' -and $def.Role -ceq 'Production' -and $def.Capability -ceq 'PROD_POST' -and $def.Surface -ceq 'Operations > Production > Production Run - Tree')
            foreach($mode in @('Expanded','Collapsed','ProcessCollapsed','Empty')){
                Stage $mode;$palette=State $true
                foreach($repeat in 1..2){
                    $before=@(Files);$notice=[string](Probe 'RunPresentationAct' @($action,''));$label='Run.'+$action+'.'+$mode+'.'+$repeat
                    $shape=if($mode -ceq 'Empty'){'0|0|0'}elseif($action -ceq 'TREE_EXPAND'){'2|4|0'}elseif($mode -ceq 'ProcessCollapsed'){'2|1|2'}else{'2|2|1'}
                    Check ($label+'.ExistingTreeShape') ([string](Probe 'RunPresentationShape') -ceq $shape)
                    Check ($label+'.PalettePreserved') ((State $true) -ceq $palette)
                    Check ($label+'.ExistingStatus') ($notice -ceq $(if($action -ceq 'TREE_EXPAND'){'All ingredient choices shown.'}else{'All ingredient choices hidden.'}))
                    Pair $before $action $label
                }
            }
            foreach($guard in @('Loading','Busy')){
                Stage $(if($action -ceq 'TREE_EXPAND'){'Collapsed'}else{'Expanded'})
                $state=State;$before=@(Files);[void](Probe 'RunPresentationAct' @($action,$guard))
                Check ('Run.'+$action+'.'+$guard+'.NoMutationOrActivity') ((State) -ceq $state -and @(Files).Count -eq $before.Count)
            }
        }
        Check 'Run.UnknownValuesAndFormula' ($sheet.Cells.Item(1,1).Value2 -ceq 'Operator Extra' -and $sheet.Cells.Item(2,1).Value2 -ceq $canary -and $sheet.Cells.Item(2,2).Formula -ceq '=1+2')
        $same=$true;foreach($file in $pins.Keys){$same=$same -and (Hash $file) -ceq $pins[$file]};Check 'Run.SavedAuthorityPreserved' $same
        foreach($guard in @('Target','Session','SignedOut','ClosedWorkbook')){
            SelectTarget $Fixture 'config-producer';[void](Probe 'OpenDesigner' @($book.Name));$decoy.Activate()
            if($guard -ceq 'Target'){SelectTarget $Other 'config-producer'}
            if($guard -ceq 'Session'){SelectTarget $Fixture 'config-producer'}
            if($guard -ceq 'SignedOut'){[void](Run 'invSys.Core.xlam' 'modAuth.SignOut')}
            if($guard -ceq 'ClosedWorkbook'){$book.Close($false);$book=$null}
            $before=@(Files);$otherBefore=@(Get-Slice4beActivityFiles $Other)
            foreach($action in @('TREE_EXPAND','TREE_COLLAPSE')){
                Stage $(if($action -ceq 'TREE_EXPAND'){'Collapsed'}else{'Expanded'});$state=State
                $notice=[string](Probe 'RunPresentationAct' @($action,''))
                Check ('Run.Guard.'+$guard+'.'+$action+'.NoMutation') ((State) -ceq $state)
                Check ('Run.Guard.'+$guard+'.'+$action+'.VisibleRefusal') ($notice.StartsWith('Session, warehouse, or captured workbook changed.') -and $notice.Contains('Reopen Production'))
            }
            Check ('Run.Guard.'+$guard+'.NoRedirectedActivity') (@(Files).Count -eq $before.Count -and @(Get-Slice4beActivityFiles $Other).Count -eq $otherBefore.Count)
        }
        [void](Probe 'CloseDesigner');Check 'Run.OperatorBytesPreserved' ((Hash $path) -ceq $bookPin)
        $same=$true;foreach($file in $recordPins.Keys){$same=$same -and (Hash $file) -ceq $recordPins[$file]};Check 'Run.OlderRecordsImmutable' $same
    }finally{[void](Probe 'CloseDesigner');if($null -ne $book){$book.Close($false)};if($null -ne $decoy){$decoy.Close($false)};SelectTarget $Fixture}
}
