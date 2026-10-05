# Project two real owner-created entity keys; never manufacture inventory identity.
function Install-ProductionOutputIdentityProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    $form.InsertLines(1,@'
Private mOutputIdentityOriginal As Variant, mOutputIdentitySorted As Variant
Private mOutputIdentityCreated As String, mOutputIdentityBatch As String, mOutputIdentityTotal As String
'@)
    $form.AddFromString(@'
Public Function OutputIdentityForTest(ByVal action As String, ByVal canary As String) As Boolean
    Dim lo As ListObject, row As Long, key As String, extra As String, priorLoading As Boolean
    On Error GoTo Failed
    Set lo = ProductionTable(TABLE_MANAGER_OUTPUT)
    Select Case action
        Case "Blank"
            mBtnManagerRefresh_Click
            If mLstManagerOutput.ListCount <> 1 Then Exit Function
            OutputIdentityForTest = CellByHeader(lo, 1, "System_Key") = "" And NzStr(mLstManagerOutput.List(0, 8)) = "" And _
                OutputIdentityPickForTest(0, "", 1#)
        Case "Prepare"
            If lo.ListRows.Count <> 1 Or modProductionReusableRun.ReusableRunIsLoaded() Then Exit Function
            mOutputIdentityCreated = CellByHeader(lo, 1, "System_Key")
            If mOutputIdentityCreated = "" Or mRunSheetKey = "" Or mOutputIdentityCreated = mRunSheetKey Then Exit Function
            mOutputIdentityOriginal = lo.DataBodyRange.Formula
            mOutputIdentityBatch = NzStr(mLstManagerOutput.List(0, 4))
            mOutputIdentityTotal = NzStr(mLstManagerOutput.List(0, 6))
            lo.ListRows.Add
            lo.ListRows(2).Range.Formula = lo.ListRows(1).Range.Formula
            SetCellByHeader lo, 2, "System_Key", mRunSheetKey
            SetCellByHeader lo, 1, "REAL OUTPUT", 11#
            SetCellByHeader lo, 2, "REAL OUTPUT", 22#
            SetCellByHeader lo, 2, "Operator Extra", canary & "-second"
            OutputIdentityForTest = CellByHeader(lo, 1, "ITEM_CODE") = CellByHeader(lo, 2, "ITEM_CODE")
        Case "Refresh"
            mBtnManagerRefresh_Click
            OutputIdentityForTest = mLstManagerOutput.ListCount = 2
        Case "Display"
            If mLstManagerOutput.ListCount <> lo.ListRows.Count Then Exit Function
            For row = 1 To lo.ListRows.Count
                If StrComp(NzStr(mLstManagerOutput.List(row - 1, 8)), CellByHeader(lo, row, "System_Key"), vbBinaryCompare) <> 0 Then Exit Function
            Next row
            OutputIdentityForTest = True
        Case "Second", "CachedSorted"
            OutputIdentityForTest = OutputIdentityPickForTest(1, mRunSheetKey, 22#)
        Case "Sort"
            With lo.Sort
                .SortFields.Clear
                .SortFields.Add Key:=lo.ListColumns(ProductionColumnIndex(lo, "REAL OUTPUT")).Range, Order:=xlDescending
                .Header = xlYes
                .Apply
            End With
            mOutputIdentitySorted = lo.DataBodyRange.Formula
            OutputIdentityForTest = CellByHeader(lo, 1, "System_Key") = mRunSheetKey And CellByHeader(lo, 2, "System_Key") = mOutputIdentityCreated
        Case "CreatedSorted"
            OutputIdentityForTest = OutputIdentityPickForTest(1, mOutputIdentityCreated, 11#)
        Case "Custom"
            For row = 1 To lo.ListRows.Count
                key = CellByHeader(lo, row, "System_Key")
                extra = canary: If key = mRunSheetKey Then extra = extra & "-second"
                If CellByHeader(lo, row, "Operator Extra") <> extra Then Exit Function
                If lo.DataBodyRange.Cells(row, ProductionColumnIndex(lo, "Operator Formula")).Formula <> "=1+2" Then Exit Function
            Next row
            OutputIdentityForTest = True
        Case "History"
            For row = 0 To 1
                If NzStr(mLstManagerOutput.List(row, 4)) <> mOutputIdentityBatch Or NzStr(mLstManagerOutput.List(row, 6)) <> mOutputIdentityTotal Then Exit Function
                If NextOutputBatchNumberForListIndex(row) <> CLng(mOutputIdentityBatch) + 1 Then Exit Function
            Next row
            OutputIdentityForTest = True
        Case "Missing"
            lo.ListRows(2).Delete
            OutputIdentityForTest = OutputIdentityPickForTest(1, "MISSING", 0#)
        Case "RestoreSorted"
            If lo.ListRows.Count = 1 Then lo.ListRows.Add
            lo.DataBodyRange.Formula = mOutputIdentitySorted
            mBtnManagerRefresh_Click
            OutputIdentityForTest = CellByHeader(lo, 2, "System_Key") = mOutputIdentityCreated
        Case "MissingHeader"
            lo.ListColumns(ProductionColumnIndex(lo, "System_Key")).Name = "Identity Probe Hidden"
            OutputIdentityForTest = OutputIdentityPickForTest(1, "MISSING", 0#)
            lo.ListColumns("Identity Probe Hidden").Name = "System_Key"
        Case "Duplicate"
            SetCellByHeader lo, 2, "System_Key", mRunSheetKey
            mBtnManagerRefresh_Click
            OutputIdentityForTest = OutputIdentityPickForTest(1, "MISSING", 0#)
        Case "Restore"
            If ProductionColumnIndex(lo, "System_Key") = 0 Then lo.ListColumns("Identity Probe Hidden").Name = "System_Key"
            Do While lo.ListRows.Count > 1: lo.ListRows(lo.ListRows.Count).Delete: Loop
            lo.DataBodyRange.Formula = mOutputIdentityOriginal
            mBtnManagerRefresh_Click
            OutputIdentityForTest = CellByHeader(lo, 1, "System_Key") = mOutputIdentityCreated And CellByHeader(lo, 1, "Operator Extra") = canary
    End Select
    Exit Function
Failed:
    OutputIdentityForTest = False
End Function
Private Function OutputIdentityPickForTest(ByVal index As Long, ByVal key As String, ByVal quantity As Double) As Boolean
    Dim priorLoading As Boolean, row As Long, lo As ListObject
    priorLoading = mLoading: mLoading = True
    mLstManagerOutput.ListIndex = index
    mTxtOutputReal.Text = "identity sentinel"
    mLoading = priorLoading
    mLstManagerOutput_Click
    row = SelectedProductionOutputTableRow()
    If key = "MISSING" Then
        OutputIdentityPickForTest = row = 0 And mTxtOutputReal.Text = "identity sentinel"
    ElseIf row > 0 Then
        Set lo = ProductionTable(TABLE_MANAGER_OUTPUT)
        OutputIdentityPickForTest = StrComp(CellByHeader(lo, row, "System_Key"), key, vbBinaryCompare) = 0 And Val(mTxtOutputReal.Text) = quantity
    End If
End Function
'@)
    $project.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Function OutputIdentity(ByVal action As String, ByVal canary As String) As Boolean
    OutputIdentity = mForm.OutputIdentityForTest(action, canary)
End Function
'@)
}

function Test-ProductionOutputIdentity($Book,$Decoy,[string]$Canary){
    $label='CompleteActivity.Worksheet.Identity'
    if(-not [bool](Probe 'OutputIdentity' @('Prepare',$Canary))){throw 'Two existing distinct entity keys with shared SKU unavailable; not product RED.'}
    try{
        $Decoy.Activate()
        foreach($case in @('Refresh','Display','Second','Custom','History')){Check ($label+'.'+$case) ([bool](Probe 'OutputIdentity' @($case,$Canary)))}
        if(-not [bool](Probe 'OutputIdentity' @('Sort',$Canary))){throw 'Actual projection sort unavailable; not product RED.'}
        foreach($case in @('CachedSorted','Refresh','Display','CreatedSorted','Custom','History')){Check ($label+'.Sorted.'+$case) ([bool](Probe 'OutputIdentity' @($case,$Canary)))}
        CaptureOwnedFormByCaptionEvidence 'Production' 'complete-worksheet-exact-output-identities.png'
        Check ($label+'.MissingKeyNoRetarget') ([bool](Probe 'OutputIdentity' @('Missing',$Canary)))
        if(-not [bool](Probe 'OutputIdentity' @('RestoreSorted',$Canary))){throw 'Projection fixture restore failed.'}
        Check ($label+'.MissingHeaderNoRetarget') ([bool](Probe 'OutputIdentity' @('MissingHeader',$Canary)))
        Check ($label+'.DuplicateKeyNoRetarget') ([bool](Probe 'OutputIdentity' @('Duplicate',$Canary)))
    }finally{
        if(-not [bool](Probe 'OutputIdentity' @('Restore',$Canary))){throw 'Original output projection restore failed.'}
    }
    foreach($fact in @('ExactInput','FreshOutput','OutputCustom','CheckCustom','Log')){Check ($label+'.Preserved.'+$fact) ([bool](Probe 'CompleteWorksheetFact' @($fact,$Canary)))}
    Check ($label+'.DecoyPreserved') ($Decoy.Worksheets.Count -eq 1 -and $Decoy.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq $Canary)
}
