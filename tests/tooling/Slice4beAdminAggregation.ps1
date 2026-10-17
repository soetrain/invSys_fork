# Actual packaged Ribbon/form diagnostic. Discovery roots and modal presentation
# are adapted in disposable copies; discovery, authorization and writes are real.
# Output contains counts/booleans only, never credentials, paths or row values.
function Install-AdminAggregationProbe {
    $admin=$packages['invSys.Admin.xlam'].VBProject
    $core=$packages['invSys.Core.xlam'].VBProject
    $ribbon=$admin.VBComponents.Add(2);$ribbon.Name='TestAggregationRibbon'
    $ribbon.CodeModule.AddFromString(@'
Option Explicit
Implements Office.IRibbonControl
Private Property Get IRibbonControl_Id() As String
    IRibbonControl_Id = "btnAdminHqAggregator"
End Property
Private Property Get IRibbonControl_Tag() As String
    IRibbonControl_Tag = ""
End Property
Private Property Get IRibbonControl_Context() As Object
    Set IRibbonControl_Context = Application.ActiveWindow
End Property
'@)
    $probe=$admin.VBComponents.Add(1);$probe.Name='TestAggregation'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mHost As Workbook, mOwnerHost As Workbook, mCalls As Long, mSucceeded As Boolean, mPhase As String
Private mForm As frmAggregationSources
Public Sub Phase(ByVal value As String)
    mPhase = value
End Sub
Public Function LastPhase() As String
    LastPhase = mPhase
End Function
Public Sub OwnerEntered(ByVal wb As Workbook)
    mCalls = mCalls + 1: Set mOwnerHost = wb
End Sub
Public Sub OwnerSucceeded()
    mSucceeded = True
End Sub
Public Function OwnerCalls() As Long
    OwnerCalls = mCalls
End Function
Public Function OwnerSuccess() As Boolean
    OwnerSuccess = mSucceeded
End Function
Public Function SetFeedForComparison(ByVal feed As String, ByVal warehouse As String) As Boolean
    Dim report As String
    SetFeedForComparison = modConfig.UpdateConfigValue("PathSharePointRoot", feed, report, warehouse, "S1")
End Function
Public Function FeedbackRedacted(ByVal first As String, ByVal second As String) As Boolean
    FeedbackRedacted = frmAggregationSources.FeedbackRedactedForTest(first, second)
End Function
Public Function LegacyAggregate(ByVal bookName As String, ByVal warehouse As String) As Boolean
    Dim report As String, wb As Workbook
    Set wb = Application.Workbooks(bookName)
    TestAggregationRoots.ResetError
    LegacyAggregate = modAdminConsole.RunHQAggregationFromAdmin("config-admin", warehouse, wb, report)
End Function
Public Function BoundHostPreserved() As Boolean
    BoundHostPreserved = (mOwnerHost Is mHost)
End Function
Public Function PreparePublished(ByVal warehouse As String, ByVal bookName As String) As Boolean
    Dim report As String, wb As Workbook
    Set wb = Application.Workbooks(bookName)
    If Not modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(wb, report) Then Exit Function
    report = modAdminConsole.SeedDemoInventoryForAutomation(warehouse, "S1", "config-admin")
    If Left$(report, 3) <> "OK|" Then Exit Function
    PreparePublished = modAdminConsole.PublishWarehouseArtifacts("config-admin", warehouse, Nothing, wb, report)
End Function
Public Function OpenForm(ByVal bookName As String) As Boolean
    Dim control As New TestAggregationRibbon, item As Object
    Set mHost = Application.Workbooks(bookName)
    mHost.Activate
    RibbonOnActionAdmin control
    DoEvents
    For Each item In VBA.UserForms
        If TypeName(item) = "frmAggregationSources" Then
            Set mForm = item: OpenForm = item.Visible
        End If
    Next item
End Function
Public Function SelectSource(ByVal warehouse As String) As Long
    If Not mForm Is Nothing Then SelectSource = mForm.SelectSourceForTest(warehouse)
End Function
Public Function AddSource(ByVal warehouse As String) As Long
    If Not mForm Is Nothing Then AddSource = mForm.AddSourceForTest(warehouse)
End Function
Public Function SourceIs(ByVal bookName As String) As Boolean
    If Not mForm Is Nothing Then SourceIs = mForm.SourceIsForTest(bookName)
End Function
Public Function NoInventedRetry() As Boolean
    If Not mForm Is Nothing Then NoInventedRetry = mForm.NoInventedRetryForTest()
End Function
Public Sub ConfirmRetryReview(ByVal confirmed As Boolean)
    If Not mForm Is Nothing Then mForm.ConfirmRetryReviewForTest confirmed
End Sub
Public Sub Aggregate()
    Dim item As Object, present As Boolean
    mCalls = 0: mSucceeded = False: Set mOwnerHost = Nothing
    mPhase = "Entry": TestAggregationRoots.ResetError
    If mForm Is Nothing Then Exit Sub
    For Each item In VBA.UserForms
        If item Is mForm Then present = True
    Next item
    If Not present Then Exit Sub
    If mForm.Visible Then mForm.AggregateForTest
End Sub
Public Sub CloseForm()
    Dim item As Object
    For Each item In VBA.UserForms
        If TypeName(item) = "frmAggregationSources" Then item.CloseForTest: Exit For
    Next item
    Set mForm = Nothing: Set mHost = Nothing: Set mOwnerHost = Nothing
End Sub
'@)
    $form=$admin.VBComponents.Item('frmAggregationSources').CodeModule
    $modernAuthority=$form.Lines(1,$form.CountOfLines) -match 'Private mSourceContext'
    $form.AddFromString(@'
Private mConfirmRetryForTest As Boolean
Public Sub ConfirmRetryReviewForTest(ByVal confirmed As Boolean)
    mConfirmRetryForTest = confirmed
End Sub
Public Function SourceIsForTest(ByVal bookName As String) As Boolean
    Dim expected As Workbook
    On Error GoTo Unavailable
    Set expected = Application.Workbooks(bookName)
    SourceIsForTest = (mSourceWorkbook Is expected) And (modAggregationCommands.BoundWorkbookName(mSourceContext) = bookName)
Unavailable:
End Function
Public Function NoInventedRetryForTest() As Boolean
    NoInventedRetryForTest = (InStr(1, CStr(mLblStatus.Caption), "previous attempt", vbTextCompare) = 0)
End Function
Public Function SelectSourceForTest(ByVal warehouse As String) As Long
    Dim i As Long
    For i = 0 To mLstSources.ListCount - 1
        mLstSources.Selected(i) = (CStr(mLstSources.List(i, 0)) = warehouse)
        If mLstSources.Selected(i) Then SelectSourceForTest = SelectSourceForTest + 1
    Next i
    mLstSources_Click
End Function
Public Function AddSourceForTest(ByVal warehouse As String) As Long
    Dim i As Long
    For i = 0 To mLstSources.ListCount - 1
        If CStr(mLstSources.List(i, 0)) = warehouse Then mLstSources.Selected(i) = True
        If mLstSources.Selected(i) Then AddSourceForTest = AddSourceForTest + 1
    Next i
    mLstSources_Click
End Function
Public Function FeedbackRedactedForTest(ByVal first As String, ByVal second As String) As Boolean
    FeedbackRedactedForTest = (InStr(1, CStr(mLblStatus.Caption), first, vbTextCompare) = 0 And _
        InStr(1, CStr(mLblStatus.Caption), second, vbTextCompare) = 0)
End Function
Public Sub AggregateForTest()
    mBtnAggregate_Click
End Sub
Public Sub CloseForTest()
    mBtnClose_Click
End Sub
Private Function MsgBox(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String = "") As VbMsgBoxResult
    MsgBox = vbOK
    If (Buttons And 7) = vbYesNo Then
        If mConfirmRetryForTest Then MsgBox = vbYes Else MsgBox = vbNo
    End If
End Function
'@)
    if(-not $modernAuthority){
        foreach($target in @(@($form,'SourceIsForTest'),@($probe.CodeModule,'SourceIs'))){
            $first=$target[0].ProcStartLine($target[1],0);$count=$target[0].ProcCountLines($target[1],0)
            $target[0].DeleteLines($first,$count)
        }
    }
    if(-not $AdminAggregationRecording){
        # Adapt the new optional notice at its native dialog boundary too.
        # Recording probes supply their own counted notice adapter.
        foreach($component in $admin.VBComponents){
            if($component.Name -cne 'cProvisioningActivity'){continue}
            $component.CodeModule.AddFromString(@'
Private Function MsgBox(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String = "") As VbMsgBoxResult
    MsgBox = vbOK
End Function
'@)
            break
        }
    }
    $entry=$admin.VBComponents.Item('modAdmin').CodeModule
    $start=$entry.ProcStartLine('Admin_AggregateGlobalSnapshot_Click',0);$count=$entry.ProcCountLines('Admin_AggregateGlobalSnapshot_Click',0)
    $body=$entry.Lines($start,$count)
    if([regex]::Matches($body,'frmAggregationSources.Show\b').Count -ne 1){throw 'Aggregation presentation boundary changed.'}
    $entry.DeleteLines($start,$count);$entry.InsertLines($start,$body.Replace('frmAggregationSources.Show','frmAggregationSources.Show vbModeless'))
    $owner=$admin.VBComponents.Item('modAdminConsole').CodeModule
    $start=$owner.ProcStartLine('RunHQAggregationFromSourceSet',0);$count=$owner.ProcCountLines('RunHQAggregationFromSourceSet',0)
    $body=$owner.Lines($start,$count)
    $resolve='    Set auditWb = ResolveAdminWorkbook(adminWb)'
    $success='        RunHQAggregationFromSourceSet = True'
    if(-not($body.Contains($resolve) -and $body.Contains($success))){throw 'Real aggregation observation boundaries unavailable.'}
    $body=$body.Replace($resolve,$resolve+"`r`n"+'    TestAggregation.OwnerEntered auditWb').Replace($success,$success+"`r`n"+'        TestAggregation.OwnerSucceeded')
    $join='    outputPath = sharePointRoot & "\\Global\\invSys.Global.InventorySnapshot.xlsb"'
    if($body.Contains($join)){
        $body=$body.Replace($join,$join+"`r`n"+'    If TestAggregationRoots.UseNormalizedPath() Then outputPath = TestAggregationRoots.NormalizedFeed(sharePointRoot)')
    }elseif(-not $body.Contains('modAggregationCommands.Aggregate')){throw 'Aggregation ownership boundary unavailable.'}
    foreach($point in @(@('    If Not EnsureAdminContext','SourceConfigAuth'),@('    If Not RequireAdminMaintenance','ProvisionedPermission'),@('    If Not ValidateAggregationSourceFilesAdmin','SourceValidation'),@('    sharePointRoot =','OutputConfig'),@('    If modHqAggregator.GenerateGlobalSnapshotFromFiles','Aggregate'))){
        $body=$body.Replace($point[0],('    TestAggregation.Phase "'+$point[1]+'"'+"`r`n"+$point[0]))
    }
    $owner.DeleteLines($start,$count);$owner.InsertLines($start,$body)
    $start=$owner.ProcStartLine('RunHQAggregationFromAdmin',0);$count=$owner.ProcCountLines('RunHQAggregationFromAdmin',0)
    $body=$owner.Lines($start,$count)
    if($body.Contains($join)){
        $body=$body.Replace($join,$join+"`r`n"+'    If TestAggregationRoots.UseNormalizedPath() Then outputPath = TestAggregationRoots.NormalizedFeed(sharePointRoot)')
    }elseif(-not $body.Contains('modAggregationCommands.CaptureForm')){throw 'Manual aggregation ownership boundary unavailable.'}
    $owner.DeleteLines($start,$count);$owner.InsertLines($start,$body)
    $gate=$core.VBComponents.Add(1);$gate.Name='TestAggregationRoots'
    $gate.CodeModule.AddFromString(@'
Option Explicit
Private mFirst As String, mSecond As String, mError As Long, mStage As String, mCalls As Long
Private mNormalized As Boolean, mExpected As String, mOutput As String
Public Sub ResetError()
    mError = 0: mStage = "": mCalls = 0: mOutput = ""
End Sub
Public Sub Entered(ByVal outputPath As String)
    mCalls = mCalls + 1: mOutput = outputPath
End Sub
Public Sub ObservedDestination(ByVal path As String)
    mOutput = path
End Sub
Public Sub PathComparison(ByVal normalized As Boolean, ByVal expected As String)
    mNormalized = normalized: mExpected = expected
End Sub
Public Function UseNormalizedPath() As Boolean
    UseNormalizedPath = mNormalized
End Function
Public Function CanonicalDestination() As Boolean
    CanonicalDestination = (StrComp(mExpected, mOutput, vbTextCompare) = 0)
End Function
Public Function NormalizedFeed(ByVal root As String) As String
    Dim prefix As String
    root = Replace$(Trim$(root), "/", "\")
    If Left$(root, 2) = "\\" Then prefix = "\\": root = Mid$(root, 3)
    Do While InStr(root, "\\") > 0: root = Replace$(root, "\\", "\"): Loop
    Do While Right$(root, 1) = "\": root = Left$(root, Len(root) - 1): Loop
    NormalizedFeed = prefix & root & "\Global\invSys.Global.InventorySnapshot.xlsb"
End Function
Public Function Calls() As Long
    Calls = mCalls
End Function
Public Sub Stage(ByVal value As String)
    mStage = value
End Sub
Public Function LastStage() As String
    LastStage = mStage
End Function
Public Function NormalizedControl(ByVal sourcePath As String, ByVal outputPath As String) As Boolean
    Dim report As String
    NormalizedControl = modHqAggregator.GenerateGlobalSnapshotFromFiles(sourcePath, outputPath, report)
End Function
Public Function Fingerprint(ByVal path As String) As String
    Fingerprint = modTrainingWire.FileSha256(path)
End Function
Public Sub ObserveError(ByVal number As Long)
    mError = number
End Sub
Public Function ErrorNumber() As Long
    ErrorNumber = mError
End Function
Public Sub SetRoots(ByVal first As String, ByVal second As String)
    mFirst = first: mSecond = second
End Sub
Public Function Roots() As Collection
    Dim result As New Collection
    If mFirst <> "" Then result.Add mFirst
    If mSecond <> "" Then result.Add mSecond
    Set Roots = result
End Function
'@)
    if($core.VBComponents.Item('modTrainingWire').CodeModule.Lines(1,$core.VBComponents.Item('modTrainingWire').CodeModule.CountOfLines) -notmatch 'Public Function FileSha256'){
        $first=$gate.CodeModule.ProcStartLine('Fingerprint',0);$count=$gate.CodeModule.ProcCountLines('Fingerprint',0)
        $gate.CodeModule.DeleteLines($first,$count)
    }
    $aggregation=$core.VBComponents.Item('modHqAggregator').CodeModule
    $start=$aggregation.ProcStartLine('RunHQAggregation',0);$count=$aggregation.ProcCountLines('RunHQAggregation',0)
    $body=$aggregation.Lines($start,$count)
    $body=$body.Replace('    Dim snapshotsFolder As String','    Dim snapshotsFolder As String'+"`r`n"+'    TestAggregationRoots.Entered outputPath')
    $aggregation.DeleteLines($start,$count);$aggregation.InsertLines($start,$body)
    $start=$aggregation.ProcStartLine('GenerateGlobalSnapshotFromFiles',0);$count=$aggregation.ProcCountLines('GenerateGlobalSnapshotFromFiles',0)
    $body=$aggregation.Lines($start,$count)
    if(-not $body.Contains('FailGenerate:')){throw 'Aggregation error observation boundary missing.'}
    $body=$body.Replace('    On Error GoTo FailGenerate',"    On Error GoTo FailGenerate`r`n    TestAggregationRoots.Entered outputPath").Replace('FailGenerate:',"FailGenerate:`r`n    TestAggregationRoots.ObserveError Err.Number")
    if($body.Contains('commandContext')){
        $body=$body.Replace('    TestAggregationRoots.Entered outputPath','    TestAggregationRoots.Entered outputPath'+"`r`n"+'    If commandContext <> "" Then TestAggregationRoots.ObservedDestination CStr(modAggregationAuthority.Entry(commandContext)("Destination"))')
    }
    $aggregation.DeleteLines($start,$count);$aggregation.InsertLines($start,$body)
    $start=$aggregation.ProcStartLine('WriteGlobalSnapshotWorkbook',0);$count=$aggregation.ProcCountLines('WriteGlobalSnapshotWorkbook',0)
    $body=$aggregation.Lines($start,$count)
    foreach($point in @(@('    EnsureFolderForFileHq','OutputFolder'),@('    CloseWorkbookByFullNameHq','ExistingOutput'),@('    Set wb = Application.Workbooks.Add','NewWorkbook'),@('    Set loSnap =','SnapshotTable'),@('    For Each key In globalRows.Keys','SnapshotRows'),@('    Set wsStatus =','StatusSheet'),@('    Set loStatus =','StatusTable'),@('    wb.SaveAs','SaveAs'),@('    wb.Close','Close'))){
        $body=$body.Replace($point[0],('    TestAggregationRoots.Stage "'+$point[1]+'"'+"`r`n"+$point[0]))
    }
    $aggregation.DeleteLines($start,$count);$aggregation.InsertLines($start,$body)
    # Input boundary only: prevent discovery of the operator's remembered servers.
    $nas=$core.VBComponents.Item('modNasConnection').CodeModule
    $start=$nas.ProcStartLine('GetKnownWarehouseTargetRoots',0);$count=$nas.ProcCountLines('GetKnownWarehouseTargetRoots',0)
    $nas.DeleteLines($start,$count)
    $nas.InsertLines($start,"Public Function GetKnownWarehouseTargetRoots() As Collection`r`nSet GetKnownWarehouseTargetRoots = TestAggregationRoots.Roots()`r`nEnd Function")
}

function Test-AdminAggregation($Fixture,$Other) {
    function Invoke([string]$Name,[object[]]$Values=@()){Run 'invSys.Admin.xlam' ('TestAggregation.'+$Name) $Values}
    function Hash([string]$Path){$stream=[IO.File]::Open($Path,'Open','Read','ReadWrite');try{(Get-FileHash -InputStream $stream).Hash}finally{$stream.Dispose()}}
    function SourcePins($f){
        $map=@{}
        foreach($path in @($f.Config,(Join-Path $f.Root ($f.Warehouse+'.invSys.Auth.xlsb')),(Join-Path $f.Root ($f.Warehouse+'.invSys.Data.Inventory.xlsb')),$f.Snapshot)){$map[$path]=Hash $path}
        return $map
    }
    function Preserved($pins){foreach($path in $pins.Keys){if((Hash $path) -cne $pins[$path]){return $false}};return $true}
    function ReadEntities([string]$Path,[string]$TableName){
        $wb=$excel.Workbooks.Open($Path,0,$true)
        try{
            $table=Table $wb $TableName;$map=[Collections.Generic.Dictionary[string,double]]::new([StringComparer]::Ordinal)
            foreach($row in $table.ListRows){
                $warehouse=[string]$row.Range.Cells.Item(1,$table.ListColumns.Item('WarehouseId').Index).Value2
                $key=[string]$row.Range.Cells.Item(1,$table.ListColumns.Item('System_Key').Index).Value2
                $quantity=[double]$row.Range.Cells.Item(1,$table.ListColumns.Item('QtyOnHand').Index).Value2
                $identity=$warehouse+'|'+$key
                if($map.ContainsKey($identity)){throw 'Duplicate exact entity in generated snapshot.'}
                $map[$identity]=$quantity
            }
            return $map
        }finally{$wb.Close($false)}
    }
    $hostBook=$excel.Workbooks.Add();$otherBook=$excel.Workbooks.Add()
    $hostBook.Windows.Item(1).Visible=$true;$otherBook.Windows.Item(1).Visible=$true;$excel.Visible=$true
    $otherBook.Worksheets.Item(1).Cells.Item(1,1).Value2='Preserve unrelated workbook'
    $observations=@()
    try{
        foreach($f in @($Fixture,$Other)){
            SelectTarget $f
            if(-not [bool](Invoke 'PreparePublished' @($f.Warehouse,$hostBook.Name))){throw 'Generated seed/publication prerequisite failed; not behavioral RED.'}
            $share=Join-Path $runRoot ($(if($f -eq $Fixture){'a'}else{'b'})+'-share')
            $snapshot=Join-Path $share ('Snapshots/'+$f.Warehouse+'.invSys.Snapshot.Inventory.xlsb')
            if(-not(Test-Path -LiteralPath $snapshot)){throw 'Owning publication did not provide the expected snapshot.'}
            $f|Add-Member -NotePropertyName Snapshot -NotePropertyValue $snapshot
            $f|Add-Member -NotePropertyName Output -NotePropertyValue (Join-Path $share 'Global/invSys.Global.InventorySnapshot.xlsb')
        }
        $expected=ReadEntities $Fixture.Snapshot 'tblInventorySnapshot'
        Check 'Aggregation.Fixture.OwnerCreatedEntities' ($expected.Count -gt 0)
        if(-not $expected.Count){throw 'No owner-created inventory to aggregate.'}
        [void](Run 'invSys.Core.xlam' 'TestAggregationRoots.SetRoots' @($Fixture.Root,$Other.Root))
        $sourcePins=SourcePins $Fixture;$otherPins=SourcePins $Other
        if($AdminAggregationRecording){
            Test-AggregationRecording $Fixture $Other $hostBook $expected
            return
        }
        if($AdminAggregationYieldGroup){
            Test-AggregationYields $Fixture $Other $hostBook
            Check ('Aggregation.YieldGroup.'+$AdminAggregationYieldGroup+'.SourceAuthoritiesAndSnapshotsPreserved') ((Preserved $sourcePins) -and (Preserved $otherPins))
            return
        }
        if($AdminAggregationRestartOnly){
            Test-AggregationRestart $Fixture $Other $hostBook
            Check 'Aggregation.Restart.SourceAuthoritiesAndSnapshotsPreserved' ((Preserved $sourcePins) -and (Preserved $otherPins))
            return
        }
        if($AdminAggregationCompletionOnly){
            $module=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modHqAggregator').CodeModule
            if($module.Lines(1,$module.CountOfLines) -match 'commandContext'){
                Test-AggregationIntegrity $Fixture $Other $hostBook
            }
            Test-AggregationCompletion $Fixture $Other $hostBook
            Check 'Aggregation.Completion.SourceAuthoritiesAndSnapshotsPreserved' ((Preserved $sourcePins) -and (Preserved $otherPins))
            return
        }
        $owned=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
        # Supplemental regression for the observed missing-destination creation:
        # the reader must not materialize a file, even before the public action.
        $wire=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modTrainingWire').CodeModule
        if($wire.Lines(1,$wire.CountOfLines) -match 'Public Function FileSha256'){
            $missing=Join-Path $runRoot 'absent-fingerprint-input.xlsb'
            $hash=[string](Run 'invSys.Core.xlam' 'TestAggregationRoots.Fingerprint' @($missing))
            Check 'Aggregation.Fingerprint.MissingInputRemainsAbsent' ($hash -ceq '' -and -not(Test-Path -LiteralPath $missing))
            Check 'Aggregation.Fingerprint.ExactBinarySnapshotHash' ([string](Run 'invSys.Core.xlam' 'TestAggregationRoots.Fingerprint' @($Fixture.Snapshot)) -ieq (Hash $Fixture.Snapshot))
        }
        if($AdminAggregationPathComparison){
            $driveFeed=Split-Path -Parent (Split-Path -Parent $Fixture.Output)
            if($driveFeed -notmatch '^[A-Za-z]:\\'){throw 'Drive comparison requires a generated local drive fixture.'}
            $uncFeed='\\localhost\'+$driveFeed.Substring(0,1)+'$\'+$driveFeed.Substring(3)
            if(-not(Test-Path -LiteralPath $uncFeed)){throw 'Disposable UNC fixture unavailable; not behavioral RED.'}
            foreach($kind in @('Drive','UNC')){
                $feed=if($kind -ceq 'Drive'){$driveFeed}else{$uncFeed}
                SelectTarget $Fixture
                if(-not [bool](Invoke 'SetFeedForComparison' @(($feed+'\\'),$Fixture.Warehouse))){throw 'Owning Config command did not establish the path fixture.'}
                $expectedOutput=Join-Path $feed 'Global\invSys.Global.InventorySnapshot.xlsb'
                $comparisonPins=SourcePins $Fixture
                foreach($normalized in @($false,$true)){
                    $variant=if($normalized){'Normalized'}else{'Original'}
                    SelectTarget $Fixture
                    if(Test-Path -LiteralPath $expectedOutput){Remove-Item -LiteralPath $expectedOutput}
                    [void](Run 'invSys.Core.xlam' 'TestAggregationRoots.PathComparison' @($normalized,$expectedOutput))
                    if(-not [bool](Invoke 'OpenForm' @($hostBook.Name))){throw 'Path comparison did not open the packaged form.'}
                    if([int](Invoke 'SelectSource' @($Fixture.Warehouse)) -ne 1){throw 'Path comparison source set was not identical.'}
                    [void](Invoke 'Aggregate')
                    $success=[bool](Invoke 'OwnerSuccess');$persisted=Test-Path -LiteralPath $expectedOutput
                    $exact=$false
                    if($persisted){
                        $actual=ReadEntities $expectedOutput 'tblGlobalInventorySnapshot';$exact=$actual.Count -eq $expected.Count
                        foreach($id in $expected.Keys){$exact=$exact -and $actual.ContainsKey($id) -and $actual[$id] -eq $expected[$id]}
                    }
                    $canonical=[bool](Run 'invSys.Core.xlam' 'TestAggregationRoots.CanonicalDestination')
                    Check ('Aggregation.Path.'+$kind+'.'+$variant+'.CanonicalDestination') $canonical
                    Check ('Aggregation.Path.'+$kind+'.'+$variant+'.PackagedCompletePersistence') ($success -and $persisted -and $exact)
                    Check ('Aggregation.Path.'+$kind+'.'+$variant+'.SourceBytesPreserved') (Preserved $comparisonPins)
                    $observations+=[pscustomobject]@{Case=('Path'+$kind+$variant);CanonicalDestination=$canonical;OwnerSuccess=$success;SourceOutput=$persisted;ExactEntities=$exact;CoreErrorNumber=[int](Run 'invSys.Core.xlam' 'TestAggregationRoots.ErrorNumber');LastCoreStage=[string](Run 'invSys.Core.xlam' 'TestAggregationRoots.LastStage')}
                    [void](Invoke 'CloseForm')
                }
            }
            SelectTarget $Fixture
            if(-not [bool](Invoke 'SetFeedForComparison' @($driveFeed,$Fixture.Warehouse))){throw 'Path fixture feed restoration failed.'}
            [void](Run 'invSys.Core.xlam' 'TestAggregationRoots.PathComparison' @($false,$Fixture.Output))
            $sourcePins=SourcePins $Fixture
        }
        SelectTarget $Fixture
        $normal=[bool](Run 'invSys.Core.xlam' 'TestAggregationRoots.NormalizedControl' @($Fixture.Snapshot,$Fixture.Output))
        Check 'Aggregation.Control.NormalizedPathServiceSucceeds' ($normal -and (Test-Path -LiteralPath $Fixture.Output))
        if($normal){
            $actual=ReadEntities $Fixture.Output 'tblGlobalInventorySnapshot';$exact=$actual.Count -eq $expected.Count
            foreach($id in $expected.Keys){$exact=$exact -and $actual.ContainsKey($id) -and $actual[$id] -eq $expected[$id]}
            Check 'Aggregation.Control.ExactSelectedEntitiesAndQuantities' $exact
        }
        Check 'Aggregation.Control.SourcesPreserved' ((Preserved $sourcePins) -and (Preserved $otherPins))
        if($AdminAggregationBoundaryMatrix){
            Test-AggregationBoundaries $Fixture $Other $hostBook $otherBook ([ref]$observations)
            if($AdminAggregationBoundaryGroup -ne 'All'){
                Check ('Aggregation.BoundaryGroup.'+$AdminAggregationBoundaryGroup+'.SourceAuthoritiesAndSnapshotsPreserved') ((Preserved $sourcePins) -and (Preserved $otherPins))
                return
            }
            # Deliberate boundary mutations must not pre-prepare the decoy used
            # by the retained original ordered workbook-switch regression.
            $otherBook.Close($false);$otherBook=$excel.Workbooks.Add()
            $otherBook.Windows.Item(1).Visible=$true
            $otherBook.Worksheets.Item(1).Cells.Item(1,1).Value2='Preserve unrelated workbook'
        }
        $authorityCases=if($AdminAggregationWriteBoundariesOnly){@('NewSession')}else{@('Admin','SwitchWorkbook','SignedOut','SignedOutHint','Reader','OtherWarehouseAdmin','NewSession')}
        foreach($case in $authorityCases){
            SelectTarget $Fixture
            foreach($f in @($Fixture,$Other)){
                $output=[IO.Path]::GetFullPath($f.Output)
                if(-not $output.StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Output is outside this generated fixture.'}
                if(Test-Path -LiteralPath $output){Remove-Item -LiteralPath $output}
            }
            if(-not [bool](Invoke 'OpenForm' @($hostBook.Name))){throw 'Actual Aggregation callback did not open the form.'}
            Check ('Aggregation.'+$case+'.ExactDiscoveredSourceSelected') ([int](Invoke 'SelectSource' @($Fixture.Warehouse)) -eq 1)
            $otherSheetCount=$otherBook.Worksheets.Count
            switch($case){
                'SwitchWorkbook' {$otherBook.Activate()}
                'SignedOut' {[void](Run 'invSys.Core.xlam' 'modAuth.SignOut')}
                'SignedOutHint' {[void](Run 'invSys.Core.xlam' 'modAuth.SignOut');[void](Run 'invSys.Core.xlam' 'modRoleEventWriter.SetCurrentUserId' @('config-admin'))}
                'Reader' {SelectTarget $Fixture 'config-reader'}
                'OtherWarehouseAdmin' {SelectTarget $Other}
                'NewSession' {SelectTarget $Fixture}
            }
            $signed=[bool](Run 'invSys.Core.xlam' 'modAuth.IsSignedIn')
            $user=[string](Run 'invSys.Core.xlam' 'modAuth.GetCurrentUserId')
            $target=[string](Run 'invSys.Core.xlam' 'modNasConnection.GetCurrentTargetWarehouseId')
            $capable=[bool](Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ADMIN_MAINT',$user,$Fixture.Warehouse,'S1'))
            $valid=switch($case){
                'SignedOut' {-not $signed -and -not $capable}
                'SignedOutHint' {-not $signed -and -not $capable}
                'Reader' {$signed -and -not $capable}
                'OtherWarehouseAdmin' {$signed -and $target -cne $Fixture.Warehouse -and -not $capable}
                default {$signed -and $target -ceq $Fixture.Warehouse -and $capable}
            }
            Check ('Aggregation.'+$case+'.AuthorityPrecondition') $valid
            if(-not $valid){throw 'Authority precondition unavailable; not behavioral RED.'}
            [void](Invoke 'Aggregate')
            $success=[bool](Invoke 'OwnerSuccess');$bound=[bool](Invoke 'BoundHostPreserved');$calls=[int](Invoke 'OwnerCalls')
            $sourceOutput=Test-Path -LiteralPath $Fixture.Output;$otherOutput=Test-Path -LiteralPath $Other.Output
            $authorized=$case -cin @('Admin','SwitchWorkbook')
            $coreCalls=[int](Run 'invSys.Core.xlam' 'TestAggregationRoots.Calls')
            Check ('Aggregation.'+$case+'.AuthorizedCoreDispatchOnly') ($coreCalls -eq $(if($authorized){1}else{0}))
            Check ('Aggregation.'+$case+'.NoRepeatedOwner') ($calls -le 1 -and (-not $authorized -or $calls -eq 1))
            Check ('Aggregation.'+$case+'.OriginalWorkbookBound') ($bound -or (-not $authorized -and $calls -eq 0))
            Check ('Aggregation.'+$case+'.ExpectedWriteBoundary') $(if($authorized){$success -and $sourceOutput -and -not $otherOutput}else{-not $success -and -not $sourceOutput -and -not $otherOutput})
            if($authorized -and $sourceOutput){
                $actual=ReadEntities $Fixture.Output 'tblGlobalInventorySnapshot';$exact=$actual.Count -eq $expected.Count
                foreach($id in $expected.Keys){$exact=$exact -and $actual.ContainsKey($id) -and $actual[$id] -eq $expected[$id]}
                Check ('Aggregation.'+$case+'.ExactSelectedEntitiesAndQuantities') $exact
            }
            Check ('Aggregation.'+$case+'.SourceAuthorityAndSnapshotPreserved') (Preserved $sourcePins)
            Check ('Aggregation.'+$case+'.OtherAuthorityAndSnapshotPreserved') (Preserved $otherPins)
            Check ('Aggregation.'+$case+'.OtherWorkbookPreserved') ($otherBook.Worksheets.Count -eq $otherSheetCount -and $otherBook.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'Preserve unrelated workbook')
            $phase=[string](Invoke 'LastPhase');$errorNumber=[int](Run 'invSys.Core.xlam' 'TestAggregationRoots.ErrorNumber')
            $stage=[string](Run 'invSys.Core.xlam' 'TestAggregationRoots.LastStage')
            $observations+=[pscustomobject]@{Case=$case;SignedIn=$signed;SourceCapability=$capable;OwnerCalls=$calls;CoreDispatches=$coreCalls;OwnerSuccess=$success;OriginalWorkbook=$bound;SourceOutput=$sourceOutput;OtherOutput=$otherOutput;LastOwnerPhase=$phase;CoreErrorNumber=$errorNumber;LastCoreStage=$stage}
            [void](Invoke 'CloseForm')
        }
    }catch{
        $line=$_.InvocationInfo.ScriptLineNumber
        [pscustomobject]@{ExceptionType=$_.Exception.GetType().FullName;HResult=$_.Exception.HResult;ScriptLine=$line}|
            ConvertTo-Json|Set-Content (Join-Path $reportRoot 'aggregation-harness-failure.json')
        throw
    }finally{
        [void](Invoke 'CloseForm');$hostBook.Close($false);$otherBook.Close($false);SelectTarget $Fixture
        [pscustomobject]@{DiagnosticOnly=$true;Acceptance=$false;Cases=$observations}|ConvertTo-Json -Depth 4|Set-Content (Join-Path $reportRoot 'aggregation-authority-observations.json')
    }
}
