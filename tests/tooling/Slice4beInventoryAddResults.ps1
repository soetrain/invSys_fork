# Draft next D13 gate: actual Add handler, owning generated keys and saved outcomes.
function Install-InventoryAddResultProbe {
    $project=$packages['invSys.Admin.xlam'].VBProject
    $project.VBComponents.Item('TestInventoryCommands').CodeModule.AddFromString(@'
Public Function AddResultShowsForTest() As Long
    AddResultShowsForTest = mShows
End Function
Public Function AddResultPartialForTest() As Boolean
    AddResultPartialForTest = (InStr(1, mLastNotice, "partial", vbTextCompare) > 0 And _
        InStr(1, mLastNotice, "Event:", vbTextCompare) > 0 And _
        InStr(1, mLastNotice, "saved", vbTextCompare) > 0)
End Function
'@)
    $probe=$project.VBComponents.Add(1);$probe.Name='TestInventoryAddResult'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mLose As Boolean, mReturned As Boolean, mKey As String
Public Sub Arm(ByVal lose As Boolean)
    mLose = lose: mReturned = False: mKey = ""
End Sub
Public Sub RecordKey(ByVal key As String)
    mKey = key
End Sub
Public Sub AfterProcessor()
    mReturned = True
    If mLose Then modAuth.SignOut
End Sub
Public Function Returned() As Boolean
    Returned = mReturned
End Function
Public Function KeyMatches(ByVal key As String) As Boolean
    KeyMatches = (mKey <> "" And mKey = key)
End Function
'@)
    $owner=$project.VBComponents.Item('modAdmin').CodeModule
    $start=$owner.ProcStartLine('AddInventoryItemForWarehouse',0);$count=$owner.ProcCountLines('AddInventoryItemForWarehouse',0)
    $body=$owner.Lines($start,$count)
    foreach($point in @(
        @('    systemKey = modRoleEventWriter.CreateSystemKey()','    TestInventoryAddResult.RecordKey systemKey'),
        @('    processedCount = modProcessor.RunBatch(warehouseId, 0, batchReport, eventIdOut, applied, True)','    TestInventoryAddResult.AfterProcessor')
    )){
        $pattern='(?im)^[\t ]*'+[regex]::Escape($point[0].Trim())+'[\t ]*\r?$'
        if([regex]::Matches($body,$pattern).Count -ne 1){throw 'Add result observation boundary unavailable.'}
        $mark=$point[1]
        $body=[regex]::Replace($body,$pattern,[Text.RegularExpressions.MatchEvaluator]{param($m) $m.Value.TrimEnd("`r")+"`r`n"+$mark})
    }
    $owner.DeleteLines($start,$count);$owner.InsertLines($start,$body)
}

function Test-InventoryAddResults($Fixture,$Other) {
    foreach($case in @('Authorized','AuthorityLost')){
        SelectTarget $Fixture
        $hostBook=$excel.Workbooks.Add();$hostBook.SaveAs((Join-Path $runRoot ('add-result-'+$case+'.xlsx')),51)
        try{
            $before=Get-AdminInventoryEntryFacts $Fixture -Persisted
            [void](Run 'invSys.Admin.xlam' 'TestInventoryAddResult.Arm' @(($case -ceq 'AuthorityLost')))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.SetSelection' @($Other.Root,''))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Execute' @($hostBook.Name,'Add','Authorized','ModalReturn',$Fixture.Root,$Fixture.Warehouse,$Fixture.Secret))
            $facts=([string](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Facts')).Split('|')
            $id='InventoryAddResult.'+$case
            $ready=$facts.Count -eq 5 -and $facts[0] -ceq 'True' -and $facts[1] -ceq 'True' -and $facts[2] -ceq '1'
            Check ($id+'.ActualAddHandlerReached') $ready
            if(-not $ready){throw 'Actual Add action unavailable; not product RED.'}
            Check ($id+'.OriginalOperatorCaptured') ($facts[3] -ceq 'True')
            Check ($id+'.ProcessorReturned') ([bool](Run 'invSys.Admin.xlam' 'TestInventoryAddResult.Returned'))
            $shows=[long](Run 'invSys.Admin.xlam' 'TestInventoryCommands.AddResultShowsForTest')
            Check ($id+'.ExpectedContinuation') ($shows -eq $(if($case -ceq 'Authorized'){2}else{1}))
            $after=Get-AdminInventoryEntryFacts $Fixture -Persisted
            $newKeys=@($after.EntityQuantities.Keys|Where-Object {-not $before.EntityQuantities.ContainsKey($_)})
            $exact=$newKeys.Count -eq 1 -and $after.EntityQuantities[$newKeys[0]] -eq 0 -and $after.LogRows -eq $before.LogRows+1 -and $after.AppliedRows -eq $before.AppliedRows+1
            Check ($id+'.ExactNewZeroEntitySaved') $exact
            if($newKeys.Count -ne 1){throw 'Expected owning creation output unavailable.'}
            Check ($id+'.OwningGeneratedKeyPreserved') ([bool](Run 'invSys.Admin.xlam' 'TestInventoryAddResult.KeyMatches' @($newKeys[0])))
            if($case -ceq 'AuthorityLost'){
                Check ($id+'.HonestSavedPartialReported') ([bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.AddResultPartialForTest'))
            }
        }finally{$hostBook.Close($false)}
    }
}
