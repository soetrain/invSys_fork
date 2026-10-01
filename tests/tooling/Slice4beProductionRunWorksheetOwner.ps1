# Real LOCAL read-model returns and explicit Clear cleanup; adapters remain unsaved.
function Install-ProductionRunWorksheetOwnerProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $owner=$project.VBComponents.Item('mProduction').CodeModule
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    function Replace-RunOwnerLine($Module,[string]$Procedure,[string]$Old,[string]$New){
        $start=$Module.ProcStartLine($Procedure,0);$end=$start+$Module.ProcCountLines($Procedure,0)
        $hits=@(for($i=$start;$i -lt $end;$i++){if($Module.Lines($i,1).Trim() -ceq $Old){$i}})
        if($hits.Count -ne 1){throw ('Worksheet owner seam changed: '+$Procedure+'; not product RED.')}
        $Module.DeleteLines($hits[0],1)
        $Module.InsertLines($hits[0],$New)
    }
    Replace-RunOwnerLine $owner 'RefreshProductionInventoryReadModelForWorkbookResult' 'If report = "" Then report = IIf(refreshed, "OK", "Inventory read-model refresh failed.")' '    TestProductionDesigner.RunSheetReadReturned refreshed, report, targetWb.Name
    If report = "" Then report = IIf(refreshed, "OK", "Inventory read-model refresh failed.")'
    Replace-RunOwnerLine $owner 'BtnClearRecipeChooser' 'MsgBox "Recipe Chooser cleared.", vbInformation' '    TestProductionDesigner.RunSheetNotice "Recipe Chooser cleared.", vbInformation'
    Replace-RunOwnerLine $owner 'BtnClearRecipeChooser' 'If wsProd Is Nothing Then Exit Sub' '    TestProductionDesigner.RunSheetClearResolved wsProd
    If wsProd Is Nothing Then Exit Sub'
    foreach($proc in @('mBtnLoaderRefresh_Click','mBtnManagerRefresh_Click')){
        Replace-RunOwnerLine $form $proc 'MsgBox refreshReport, vbExclamation, "Production Inventory Refresh"' '        TestProductionDesigner.RunSheetNotice refreshReport, vbExclamation, "Production Inventory Refresh"'
    }
    $form.AddFromString(@'
Public Function RunSheetOwnerPrepareForTest() As Boolean
    Dim report As String, lo As ListObject
    If Not RunWorksheetPrepareForTest() Then Exit Function
    If Not modOperationsPrimitiveBridge.EnsureProductionWorkbookSurface(mOperatorWorkbook.Name, report) Then Exit Function
    Set lo = mOperatorWorkbook.Worksheets("InventoryManagement").ListObjects("invSys")
    If lo.ListRows.Count = 0 Then lo.ListRows.Add
    lo.ListColumns("System_Key").DataBodyRange.Cells(1, 1).Value2 = mRunSheetKey
    lo.ListColumns("ITEM_CODE").DataBodyRange.Cells(1, 1).Value2 = "DEMO-RAW-BLACK-TEA"
    lo.ListColumns("UOM").DataBodyRange.Cells(1, 1).Value2 = mRunSheetUom
    lo.ListColumns("LOCATION").DataBodyRange.Cells(1, 1).Value2 = mRunSheetLocation
    lo.ListColumns("QtyOnHand").DataBodyRange.Cells(1, 1).Value2 = mRunSheetAvailable
    lo.ListColumns("QtyAvailable").DataBodyRange.Cells(1, 1).Value2 = mRunSheetAvailable
    lo.ListColumns.Add.Name = "Operator Extra"
    lo.ListColumns.Add.Name = "Operator Formula"
    lo.ListColumns("Operator Extra").DataBodyRange.Cells(1, 1).Value2 = "RUN-LOCAL-EXTRA"
    lo.ListColumns("Operator Formula").DataBodyRange.Cells(1, 1).Formula = "=1+2"
    RunSheetOwnerPrepareForTest = True
End Function
Public Function RunSheetOwnerStageForTest(ByVal action As String, ByVal mode As String, ByVal canary As String) As String
    Dim ws As Worksheet, lo As ListObject, name As Variant, phase As String
    On Error GoTo Failed
    phase = "Restore"
    RunSheetOwnerRestoreForTest
    modProductionReusableRun.ClearReusableRun
    mLoading = True
    phase = "Page"
    If action = "MANAGER_REFRESH" Then mPages.Value = 4 Else mPages.Value = 3
    If action = "CLEAR" Then
        phase = "Production"
        Set ws = mOperatorWorkbook.Worksheets("Production")
        For Each name In Array("RC_RecipeChoose", "RecipeChooser_generated", "InventoryPalette_generated", "ProductionOutput", "Prod_invSys_Check")
            phase = CStr(name)
            Set lo = ws.ListObjects(CStr(name))
            ' Use the surface's reserved blank row, without shifting neighboring tables.
            If lo.ListRows.Count = 0 Then lo.ListRows.Add AlwaysInsert:=False
            Select Case CStr(name)
                Case "RC_RecipeChoose": lo.ListColumns("RECIPE").DataBodyRange.Cells(1, 1).Value2 = canary
                Case "RecipeChooser_generated", "ProductionOutput": lo.ListColumns("PROCESS").DataBodyRange.Cells(1, 1).Value2 = canary
                Case "InventoryPalette_generated"
                    lo.ListColumns("ITEM_CODE").DataBodyRange.Cells(1, 1).Value2 = "DEMO-RAW-BLACK-TEA"
                    lo.ListColumns("System_Key").DataBodyRange.Cells(1, 1).Value2 = mRunSheetKey
                Case "Prod_invSys_Check": lo.ListColumns("System_Key").DataBodyRange.Cells(1, 1).Value2 = mRunSheetKey
            End Select
        Next name
        phase = "RenameProduction"
        If mode = "Missing" Then ws.Name = "RunOwnerUnavailable"
    ElseIf mode = "Missing" Then
        phase = "RenameInventory"
        mOperatorWorkbook.Worksheets("InventoryManagement").ListObjects("invSys").Name = "RunOwnerUnavailableInv"
    End If
    TestProductionDesigner.RunSheetReset mOperatorWorkbook.Name
    mLoading = False
    RunSheetOwnerStageForTest = "READY"
    Exit Function
Failed:
    mLoading = False
    RunSheetOwnerStageForTest = "FIXTURE_FAILED|" & phase & "|" & CStr(Err.Number)
End Function
Public Sub RunSheetOwnerRestoreForTest()
    Dim ws As Worksheet, lo As ListObject
    For Each ws In mOperatorWorkbook.Worksheets
        If ws.Name = "RunOwnerUnavailable" Then ws.Name = "Production"
        For Each lo In ws.ListObjects
            If lo.Name = "RunOwnerUnavailableInv" Then lo.Name = "invSys"
        Next lo
    Next ws
End Sub
Public Function RunSheetOwnerResultForTest(ByVal action As String, ByVal mode As String) As Boolean
    Dim ws As Worksheet, lo As ListObject, name As Variant
    If action <> "CLEAR" Then
        RunSheetOwnerResultForTest = TestProductionDesigner.RunSheetReadResult(mode, mOperatorWorkbook.Name)
        Exit Function
    End If
    If mode = "Missing" Then
        Set ws = mOperatorWorkbook.Worksheets("RunOwnerUnavailable")
        Set lo = ws.ListObjects("RC_RecipeChoose")
        RunSheetOwnerResultForTest = CStr(lo.ListColumns("RECIPE").DataBodyRange.Cells(1, 1).Value2) <> "" And _
            TestProductionDesigner.RunSheetNoticeCount() = 0
        Exit Function
    End If
    Set ws = mOperatorWorkbook.Worksheets("Production")
    For Each lo In ws.ListObjects
        If lo.Name = "RecipeChooser_generated" Or lo.Name = "InventoryPalette_generated" Then Exit Function
    Next lo
    For Each name In Array("RC_RecipeChoose", "ProductionOutput", "Prod_invSys_Check")
        Set lo = ws.ListObjects(CStr(name))
        If Not lo.DataBodyRange Is Nothing Then
            If Application.WorksheetFunction.CountA(lo.DataBodyRange) <> 0 Then Exit Function
        End If
    Next name
    RunSheetOwnerResultForTest = TestProductionDesigner.RunSheetNoticeCount() = 1
End Function
Public Function RunSheetOwnerPreservedForTest() As Boolean
    Dim lo As ListObject, row As Long
    Set lo = Nothing
    On Error Resume Next
    Set lo = mOperatorWorkbook.Worksheets("InventoryManagement").ListObjects("invSys")
    If lo Is Nothing Then Set lo = mOperatorWorkbook.Worksheets("InventoryManagement").ListObjects("RunOwnerUnavailableInv")
    On Error GoTo 0
    If lo Is Nothing Then Exit Function
    For row = 1 To lo.ListRows.Count
        If StrComp(CStr(lo.ListColumns("System_Key").DataBodyRange.Cells(row, 1).Value2), mRunSheetKey, vbBinaryCompare) = 0 Then
            RunSheetOwnerPreservedForTest = lo.ListColumns("Operator Extra").DataBodyRange.Cells(row, 1).Value2 = "RUN-LOCAL-EXTRA" And _
                lo.ListColumns("Operator Formula").DataBodyRange.Cells(row, 1).Formula = "=1+2"
            Exit Function
        End If
    Next row
End Function
'@)
    $adapter=$project.VBComponents.Item('TestProductionDesigner').CodeModule
    $adapter.InsertLines(1,@'
Private mRunSheetReads As Long, mRunSheetReadOK As Boolean, mRunSheetReadReport As String
Private mRunSheetReadBook As String, mRunSheetNotices As Long
Private mRunSheetExpectedBook As String, mRunSheetClearTarget As String
'@)
    $adapter.AddFromString(@'
Public Sub RunSheetReset(ByVal expectedBook As String)
    mRunSheetReads = 0: mRunSheetNotices = 0: mRunSheetReadReport = "": mRunSheetReadBook = ""
    mRunSheetExpectedBook = expectedBook: mRunSheetClearTarget = "NotEntered"
End Sub
Public Sub RunSheetClearResolved(ByVal ws As Worksheet)
    If ws Is Nothing Then
        mRunSheetClearTarget = "Unavailable"
    ElseIf ws.Parent.Name = mRunSheetExpectedBook Then
        mRunSheetClearTarget = "Captured"
    ElseIf ws.Parent.IsAddin Then
        mRunSheetClearTarget = "OtherAddin"
    Else
        mRunSheetClearTarget = "OtherWorkbook"
    End If
End Sub
Public Function RunSheetClearTarget() As String
    RunSheetClearTarget = mRunSheetClearTarget
End Function
Public Sub RunSheetReadReturned(ByVal ok As Boolean, ByVal report As String, ByVal bookName As String)
    mRunSheetReads = mRunSheetReads + 1: mRunSheetReadOK = ok
    mRunSheetReadReport = report: mRunSheetReadBook = bookName
End Sub
Public Sub RunSheetNotice(ByVal message As String, ByVal style As VbMsgBoxStyle, Optional ByVal title As String = "")
    mRunSheetNotices = mRunSheetNotices + 1
End Sub
Public Function RunSheetNoticeCount() As Long
    RunSheetNoticeCount = mRunSheetNotices
End Function
Public Function RunSheetReadResult(ByVal mode As String, ByVal bookName As String) As Boolean
    If mRunSheetReads <> 1 Or mRunSheetReadBook <> bookName Then Exit Function
    If mode = "Missing" Then
        RunSheetReadResult = Not mRunSheetReadOK And mRunSheetReadReport = "invSys table not found." And mRunSheetNotices = 1
    Else
        RunSheetReadResult = mRunSheetReadOK And mRunSheetNotices = 0
    End If
End Function
Public Function RunSheetOwnerPrepare() As Boolean
    RunSheetOwnerPrepare = mForm.RunSheetOwnerPrepareForTest()
End Function
Public Function RunSheetOwnerStage(ByVal action As String, ByVal mode As String, ByVal canary As String) As String
    RunSheetOwnerStage = mForm.RunSheetOwnerStageForTest(action, mode, canary)
End Function
Public Sub RunSheetOwnerRestore()
    mForm.RunSheetOwnerRestoreForTest
End Sub
Public Function RunSheetOwnerResult(ByVal action As String, ByVal mode As String) As Boolean
    RunSheetOwnerResult = mForm.RunSheetOwnerResultForTest(action, mode)
End Function
Public Function RunSheetOwnerPreserved() As Boolean
    RunSheetOwnerPreserved = mForm.RunSheetOwnerPreservedForTest()
End Function
'@)
}

function Test-ProductionRunWorksheetOwner($Fixture) {
    if(-not [bool](Probe 'RunSheetOwnerPrepare')){throw 'Supported worksheet owner fixture unavailable; not product RED.'}
    Check 'RunWorksheetOwner.SupportedSurfaceAndSeedIdentity' $true
    foreach($action in @('LOADER_REFRESH','MANAGER_REFRESH','CLEAR')){
        foreach($mode in @('Missing','Normal')){
            $before=@(Files);$ready=[string](Probe 'RunSheetOwnerStage' @($action,$mode,$canary))
            if($ready -cne 'READY'){
                if($ready -notmatch '^FIXTURE_FAILED\|[A-Za-z_]+\|-?[0-9]+$'){$ready='Unavailable'}
                throw ('Worksheet owner staging unavailable: '+$ready+'; not product RED.')
            }
            try{
                $decoy.Activate();$label='RunWorksheetOwner.'+$action+'.'+$mode
                Check ($label+'.SetupNotUserAction') (@(Files).Count -eq $before.Count)
                $notice=[string](Probe 'RunFaultAct' @($action))
                Check ($label+'.NoUnhandledError') (-not $notice.StartsWith('HANDLER_ERROR|'))
                Check ($label+'.GuardsRestoredWithoutAdapterReset') ([bool](Probe 'RunFaultGuards'))
                Check ($label+'.RealOwnerResultAndNotificationCount') ([bool](Probe 'RunSheetOwnerResult' @($action,$mode)))
                if($action -ceq 'CLEAR'){
                    $target=[string](Probe 'RunSheetClearTarget')
                    if($target -cnotin @('NotEntered','Unavailable','Captured','OtherAddin','OtherWorkbook')){throw 'Unknown Clear diagnostic category.'}
                    Write-Output ('Run Clear '+$mode+' resolved target: '+$target)
                    $bound=if($mode -ceq 'Missing'){$target -cin @('NotEntered','Unavailable')}else{$target -ceq 'Captured'}
                    Check ($label+'.OwnerUsesCapturedWorkbook') $bound
                }
                Check ($label+'.CustomColumnsFormulaAndExactKey') ([bool](Probe 'RunSheetOwnerPreserved'))
                Check ($label+'.CanonicalSourcePreserved') ([bool](Probe 'RunWorksheetSource'))
                $outcome=if($mode -ceq 'Missing'){'FAILED'}elseif($action -ceq 'CLEAR'){'STAGED'}else{'REFRESHED'}
                Pair $before $action $outcome $label
            }finally{[void](Probe 'RunSheetOwnerRestore')}
        }
    }
}
