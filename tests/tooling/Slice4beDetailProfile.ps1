# D18 profile UI observations through the real packaged Settings form.
# No Config values, entered data or runtime identity leave the disposable fixture.
function Install-Slice4beDetailProfileProbe($TestModule,$FormCode) {
    $TestModule.CodeModule.AddFromString(@'
Public Function DetailProfileSurface() As String
    Dim page As Object, family As Object, fields As Object, enabled As Object, preview As Object
    Set page = TrackingPage(mForm, "Event Tracking")
    Set family = TrackingControl(page, "ComboBox", "", "cmbDetailFamily")
    Set fields = TrackingControl(page, "ListBox", "", "lstDetailFields")
    Set enabled = TrackingControl(page, "CheckBox", "", "chkDetailEnabled")
    Set preview = TrackingControl(page, "TextBox", "", "txtDetailPreview")
    DetailProfileSurface = CStr(Not family Is Nothing) & "|" & _
        CStr(Not fields Is Nothing And Not enabled Is Nothing) & "|" & _
        CStr(TrackingButton(page, "Move Up") And TrackingButton(page, "Move Down")) & "|" & _
        CStr(TrackingButton(page, "Save Detail Profile") And _
             Not TrackingControl(page, "CommandButton", "Reload", "btnReloadDetailProfile") Is Nothing And _
             Not TrackingControl(page, "CommandButton", "Reset to Default", "btnResetDetailProfile") Is Nothing)
    If preview Is Nothing Then
        DetailProfileSurface = DetailProfileSurface & "|False"
    Else
        DetailProfileSurface = DetailProfileSurface & "|" & CStr(preview.Locked And preview.MultiLine)
    End If
End Function
Public Function DetailProfileSelectSection() As Boolean
    Dim page As Object, sections As Object, child As Object
    Set page = TrackingPage(mForm, "Event Tracking")
    Set sections = TrackingControl(page, "MultiPage", "", "mpEventTracking")
    If sections Is Nothing Then Exit Function
    TrackingSettingsSelectPage "Event Tracking"
    For Each child In sections.Pages
        If child.Caption = "Event Detail" Then
            sections.Value = child.Index
            mForm.Show vbModeless
            mForm.Repaint
            DetailProfileSelectSection = True
            Exit Function
        End If
    Next child
End Function
'@)
    $detail=$null
    try {$detail=$packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('cAdminEventDetail').CodeModule} catch {}
    if($null -eq $detail){return}
    Install-DetailProfileReadDiagnostics
    # Fixed numeric diagnostics preserve the failing call without exporting values.
    $FormCode.AddFromString(@'
Private mDetailGeneralError As Long
Private mDetailGeneralStage As Long
Public Function DetailGeneralDiagnostic() As String
    DetailGeneralDiagnostic = CStr(mDetailGeneralError) & "|" & CStr(mDetailGeneralStage)
End Function
'@)
    $start=$FormCode.ProcStartLine('D5TestSave',0)
    $count=$FormCode.ProcCountLines('D5TestSave',0)
    $body=$FormCode.Lines($start,$count)
    $body=$body.Replace('    Dim i As Long',"    Dim i As Long`r`n    On Error GoTo DetailGeneralFailed`r`n    mDetailGeneralError = 0: mDetailGeneralStage = 0")
    $body=$body.Replace('            mLstConfig.ListIndex = i',"            mDetailGeneralStage = 1`r`n            mLstConfig.ListIndex = i")
    $body=$body.Replace('            mLstConfig_Click',"            mDetailGeneralStage = 2`r`n            mLstConfig_Click")
    $body=$body.Replace('            mTxtConfigValue.Value = valueText',"            mDetailGeneralStage = 3`r`n            mTxtConfigValue.Value = valueText")
    $body=$body.Replace('            mBtnSaveConfig_Click',"            mDetailGeneralStage = 4`r`n            mBtnSaveConfig_Click")
    $body=$body.Replace('            D5TestSave = mLblStatus.Caption',"            mDetailGeneralStage = 5`r`n            D5TestSave = mLblStatus.Caption")
    $body=$body.Replace('End Function',"    Exit Function`r`nDetailGeneralFailed:`r`n    mDetailGeneralError = Err.Number`r`n    D5TestSave = `"`"`r`nEnd Function")
    $FormCode.DeleteLines($start,$count)
    $FormCode.InsertLines($start,$body)
    $detail.AddFromString(@'
Private mDetailTestEntries As Long
Public Function DetailTestRequest() As String
    DetailTestRequest = mRequest
End Function
Public Function DetailTestEntries() As Long
    DetailTestEntries = mDetailTestEntries
End Function
Public Sub DetailTestChoose(ByVal family As String, ByVal field As String)
    Dim index As Long
    mFamily.Value = family
    mFamily_Change
    For index = 0 To mRows.ListCount - 1
        If mRows.List(index, 0) = field Then
            mRows.ListIndex = index
            mRows_Change
            Exit Sub
        End If
    Next index
    Err.Raise 5
End Sub
Public Sub DetailTestToggle(ByVal enabled As Boolean)
    mEnabled.Value = enabled
    mEnabled_Click
End Sub
Public Sub DetailTestMoveUp()
    mUp_Click
End Sub
Public Function DetailTestRequiredLocked() As Boolean
    DetailTestRequiredLocked = Not mEnabled.Enabled And CBool(mEnabled.Value)
End Function
Public Function DetailTestSyntheticPreview() As Boolean
    DetailTestSyntheticPreview = InStr(1, mPreview.Value, "SYNTHETIC PREVIEW - no live event data", vbBinaryCompare) = 1 And _
        InStr(1, mPreview.Value, "EXAMPLE-ACTOR", vbBinaryCompare) > 0
End Function
'@)
    $detail.InsertLines($detail.ProcBodyLine('SaveProfile',0)+2,'    mDetailTestEntries = mDetailTestEntries + 1')
    $FormCode.AddFromString(@'
Public Function DetailTestRequest() As String
    DetailTestRequest = mDetail.DetailTestRequest()
End Function
Public Function DetailTestSave() As Boolean
    DetailTestSave = mDetail.SaveProfile()
End Function
Public Function DetailTestEntries() As Long
    DetailTestEntries = mDetail.DetailTestEntries()
End Function
Public Sub DetailTestChoose(ByVal family As String, ByVal field As String)
    mDetail.DetailTestChoose family, field
End Sub
Public Sub DetailTestToggle(ByVal enabled As Boolean)
    mDetail.DetailTestToggle enabled
End Sub
Public Sub DetailTestMoveUp()
    mDetail.DetailTestMoveUp
End Sub
Public Function DetailTestRequiredLocked() As Boolean
    DetailTestRequiredLocked = mDetail.DetailTestRequiredLocked()
End Function
Public Function DetailTestSyntheticPreview() As Boolean
    DetailTestSyntheticPreview = mDetail.DetailTestSyntheticPreview()
End Function
Public Sub DetailTestReset()
    mDetail.ResetProfile
End Sub
Public Sub DetailTestReload()
    mDetail.ReloadProfile
End Sub
'@)
    $TestModule.CodeModule.AddFromString(@'
Public Function DetailGeneralDiagnostic() As String
    If mForm Is Nothing Then
        DetailGeneralDiagnostic = "91|-1"
    Else
        DetailGeneralDiagnostic = mForm.DetailGeneralDiagnostic()
    End If
End Function
Public Function DetailProfileRequest() As String
    DetailProfileRequest = mForm.DetailTestRequest()
End Function
Public Function DetailProfileSave() As Boolean
    DetailProfileSave = mForm.DetailTestSave()
End Function
Public Function DetailProfileEntries() As Long
    DetailProfileEntries = mForm.DetailTestEntries()
End Function
Public Sub DetailProfileChoose(ByVal family As String, ByVal field As String)
    mForm.DetailTestChoose family, field
End Sub
Public Sub DetailProfileToggle(ByVal enabled As Boolean)
    mForm.DetailTestToggle enabled
End Sub
Public Sub DetailProfileMoveUp()
    mForm.DetailTestMoveUp
End Sub
Public Function DetailProfileRequiredLocked() As Boolean
    DetailProfileRequiredLocked = mForm.DetailTestRequiredLocked()
End Function
Public Function DetailProfileSyntheticPreview() As Boolean
    DetailProfileSyntheticPreview = mForm.DetailTestSyntheticPreview()
End Function
Public Sub DetailProfileReset()
    mForm.DetailTestReset
End Sub
Public Sub DetailProfileReload()
    mForm.DetailTestReload
End Sub
Public Function DetailProfileSaveDirect(ByVal context As String, ByVal version As Long, ByVal request As String) As Boolean
    Dim report As String
    DetailProfileSaveDirect = modEventDetailSettings.SaveProfile(context, version, request, report)
End Function
'@)
}
function Test-Slice4beDetailProfile($Fixture,$Other) {
    $before=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    $flags=([string](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileSurface')).Split('|')
    if($flags.Count -ne 5 -or @($flags|Where-Object {$_ -cnotin @('True','False')}).Count){throw 'Invalid profile surface observation'}
    $names=@('FamilySelector','FieldEditor','FieldOrderActions','SeparateProfileActions','ReadOnlyPreviewSurface')
    for($i=0;$i -lt $names.Count;$i++){Check ('DetailProfile.'+$names[$i]) ($flags[$i] -ceq 'True')}
    $selected=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileSelectSection')
    Check 'DetailProfile.SectionReachableAndFits' ($selected -and [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingSettingsLayoutFits'))
    Check 'DetailProfile.OpenDoesNotWriteConfig' ($before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    if($selected -and @($flags|Where-Object {$_ -eq 'False'}).Count -eq 0){Test-DetailProfileActions $Fixture $Other}
    if($CaptureEvidence -and $selected){
        $wasVisible=$excel.Visible
        try {
            $excel.Visible=$true
            [void](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileSelectSection')
            Start-Sleep -Milliseconds 300
            CaptureFormEvidence 'invSys Settings' 'detail-profile-editor.png'
        }
        finally {$excel.Visible=$wasVisible}
    }
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingSettingsSelectPage' @('General'))
}
function Get-DetailProfileRequest {[string](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileRequest')}
function Get-DetailProfileVersion($Fixture) {
    $book=$excel.Workbooks.Open($Fixture.Config,0,$true)
    try {
        $latest=0
        foreach($sheet in $book.Worksheets){foreach($table in $sheet.ListObjects){
            if($table.Name -eq 'tblEventDetailProfiles'){foreach($row in $table.ListRows){$latest=[Math]::Max($latest,[int]$row.Range.Cells.Item(1,$table.ListColumns.Item('ProfileVersion').Index).Value2)}}
        }}
        return $latest
    } finally {$book.Close($false)}
}
function Test-DetailProfileActions($Fixture,$Other) {
    $request=Get-DetailProfileRequest
    $model=$request|ConvertFrom-Json
    Check 'DetailProfile.ModelLoaded' ($model.SchemaVersion -eq 1 -and $model.Fields.Count -gt 30)
    $before=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileChoose' @('Receiving','System_Key'))
    Check 'DetailProfile.RequiredIdentityCannotBeDisabled' ([bool](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileRequiredLocked'))
    $original=$model.Fields|Where-Object {$_.EventFamily -ceq 'Receiving' -and $_.FieldId -ceq 'Reference'}
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileChoose' @('Receiving','Reference'))
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileToggle' @($false))
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileMoveUp')
    $staged=Get-DetailProfileRequest|ConvertFrom-Json
    $changed=$staged.Fields|Where-Object {$_.EventFamily -ceq 'Receiving' -and $_.FieldId -ceq 'Reference'}
    Check 'DetailProfile.ToggleAndOrderStageOnly' (-not $changed.Enabled -and $changed.DisplayOrder -eq $original.DisplayOrder-1 -and $before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    Check 'DetailProfile.OtherFamilyUnaffected' (@($staged.Fields|Where-Object {$_.EventFamily -ceq 'Shipping' -and $_.FieldId -ceq 'Reference' -and $_.Enabled -and $_.DisplayOrder -eq $original.DisplayOrder}).Count -eq 1)
    Check 'DetailProfile.PreviewUsesSyntheticValues' ([bool](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileSyntheticPreview'))
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileReset')
    Check 'DetailProfile.ResetStagesDefaultsOnly' ($request -ceq (Get-DetailProfileRequest) -and $before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    $context=[string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext')
    foreach($case in @('UnknownEnvelopeField','InvalidSchema','UnknownField','UnknownFamily','DuplicateField','DuplicateOrder','InvalidBoolean','RequiredDisabled','FractionalOrder','MissingField')){
        $bad=$request|ConvertFrom-Json
        switch($case){
            UnknownEnvelopeField {$bad|Add-Member NoteProperty Unsupported $true}
            InvalidSchema {$bad.SchemaVersion='1'}
            UnknownField {$bad.Fields[0].FieldId='RawPayload'}
            UnknownFamily {$bad.Fields[0].EventFamily='Unregistered'}
            DuplicateField {$bad.Fields[1].FieldId=$bad.Fields[0].FieldId}
            DuplicateOrder {$bad.Fields[1].DisplayOrder=$bad.Fields[0].DisplayOrder}
            InvalidBoolean {$bad.Fields[0].Enabled='true'}
            RequiredDisabled {$bad.Fields[0].Enabled=$false}
            FractionalOrder {$bad.Fields[0].DisplayOrder=1.5}
            MissingField {$bad.Fields=@($bad.Fields|Select-Object -Skip 1)}
        }
        $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileSaveDirect' @($context,0,($bad|ConvertTo-Json -Depth 8 -Compress)))
        Check ('DetailProfile.Reject'+$case) (-not $ok -and $before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    }
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileSaveDirect' @($context,1,$request))
    Check 'DetailProfile.WrongExpectedVersionDenied' (-not $ok -and $before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileSaveDirect' @($context,0,$request))
    Check 'DetailProfile.SignedOutDenied' (-not $ok -and $before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    SelectTarget $Other
    $otherBefore=(Get-FileHash -LiteralPath $Other.Config).Hash
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileSave')
    Check 'DetailProfile.CapturedTargetDenied' (-not $ok -and $before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash -and $otherBefore -ceq (Get-FileHash -LiteralPath $Other.Config).Hash)
    SelectTarget $Fixture 'config-reader'
    $context=[string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext')
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileSaveDirect' @($context,0,$request))
    Check 'DetailProfile.MissingCapabilityDenied' (-not $ok -and $before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    SelectTarget $Fixture
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileSave')
    Check 'DetailProfile.StaleSessionDenied' (-not $ok -and $before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileSelectSection')
    $context=[string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext')
    foreach($readOnly in @($true,$false)){
        $book=$excel.Workbooks.Open($Fixture.Config,0,$readOnly)
        $table=Table $book 'tblWarehouseConfig'
        $cell=$table.ListColumns.Item('WarehouseName').DataBodyRange.Cells.Item(1,1)
        if(-not $readOnly){$cell.Value2='unsaved profile fixture marker'}
        $value=$cell.Value2
        $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileSaveDirect' @($context,0,$request))
        $name=if($readOnly){'ReadOnly'}else{'Dirty'}
        $preserved=(-not $ok -and $cell.Value2 -ceq $value -and ($readOnly -or -not $book.Saved))
        $book.Close($false)
        Check ('DetailProfile.'+$name+'ConfigPreserved') ($preserved -and $before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    }
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileChoose' @('Receiving','Reference'))
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileToggle' @($false))
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileMoveUp')
    $entries=[int](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileEntries')
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileSave')
    Check 'DetailProfile.RealSaveActionEntered' ([int](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileEntries') -eq $entries+1)
    Check 'DetailProfile.AuthorizedSavePublishesVersion' ($ok -and (Get-DetailProfileVersion $Fixture) -eq 1)
    if(-not $ok){return}
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileReload')
    $saved=(Get-DetailProfileRequest|ConvertFrom-Json).Fields|Where-Object {$_.EventFamily -ceq 'Receiving' -and $_.FieldId -ceq 'Reference'}
    $reloadMatches=(-not $saved.Enabled -and $saved.DisplayOrder -eq $original.DisplayOrder-1)
    Check 'DetailProfile.SavedChoiceAndOrderReload' $reloadMatches
    if(-not $reloadMatches){
        $diagnostic=[string](Run 'invSys.Core.xlam' 'modEventDetailSettings.DetailReadDiagnostic')
        if($diagnostic -notmatch '^-?[0-9]+\|-?[0-9]+\|-?[0-9]+\|-?[0-9]+$'){throw 'Invalid fixed profile reader diagnostic'}
        Write-Output ('Profile read stage|error|decode stage|row: '+$diagnostic)
        return
    }
    $before=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileSaveDirect' @($context,0,$request))
    Check 'DetailProfile.StaleVersionCannotAppend' (-not $ok -and $before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    $book=$excel.Workbooks.Open($Fixture.Config,0,$false)
    $headers=Table $book 'tblEventDetailProfiles'; $fields=Table $book 'tblEventDetailFields'
    $column=$headers.ListColumns.Add(1); $column.Name='Profile Extra'; $column.DataBodyRange.Value2='header fixture value'
    $column=$fields.ListColumns.Add(1); $column.Name='Profile Extra'; $column.DataBodyRange.Value2='field fixture value'
    $priorHeaders=$headers.DataBodyRange.Value2; $priorFields=$fields.DataBodyRange.Value2; $fieldCount=$fields.ListRows.Count
    $policyHeaders=Table $book 'tblEventTrackingPolicies'; $policyFields=Table $book 'tblEventTrackingControls'
    $policyHeaderValues=$policyHeaders.DataBodyRange.Value2; $policyFieldValues=$policyFields.DataBodyRange.Value2
    $book.Save(); $book.Close($false)
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileChoose' @('Receiving','Reference'))
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileToggle' @($true))
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileSave')
    Check 'DetailProfile.SecondSaveAppendsVersion' ($ok -and (Get-DetailProfileVersion $Fixture) -eq 2)
    $book=$excel.Workbooks.Open($Fixture.Config,0,$true)
    try {
        $headers=Table $book 'tblEventDetailProfiles'; $fields=Table $book 'tblEventDetailFields'
        $afterHeaders=$headers.DataBodyRange.Value2; $afterFields=$fields.DataBodyRange.Value2
        $retained=($headers.ListRows.Count -eq 2 -and $fields.ListRows.Count -eq $fieldCount*2)
        for($c=1;$c -le $priorHeaders.GetLength(1);$c++){$retained=$retained -and $afterHeaders.GetValue(1,$c) -ceq $priorHeaders.GetValue(1,$c)}
        for($r=1;$r -le $fieldCount;$r++){for($c=1;$c -le $priorFields.GetLength(1);$c++){$retained=$retained -and $afterFields.GetValue($r,$c) -ceq $priorFields.GetValue($r,$c)}}
        Check 'DetailProfile.PriorVersionsAndUnknownColumnsPreserved' $retained
        $policyHeaders=Table $book 'tblEventTrackingPolicies'; $policyFields=Table $book 'tblEventTrackingControls'
        $afterPolicyHeaders=$policyHeaders.DataBodyRange.Value2; $afterPolicyFields=$policyFields.DataBodyRange.Value2
        $policyRetained=($afterPolicyHeaders.GetLength(0) -eq $policyHeaderValues.GetLength(0) -and $afterPolicyFields.GetLength(0) -eq $policyFieldValues.GetLength(0))
        for($r=1;$r -le $policyHeaderValues.GetLength(0);$r++){for($c=1;$c -le $policyHeaderValues.GetLength(1);$c++){$policyRetained=$policyRetained -and $afterPolicyHeaders.GetValue($r,$c) -ceq $policyHeaderValues.GetValue($r,$c)}}
        for($r=1;$r -le $policyFieldValues.GetLength(0);$r++){for($c=1;$c -le $policyFieldValues.GetLength(1);$c++){$policyRetained=$policyRetained -and $afterPolicyFields.GetValue($r,$c) -ceq $policyFieldValues.GetValue($r,$c)}}
        Check 'DetailProfile.TrackingPolicyScopeUnchanged' $policyRetained
    } finally {$book.Close($false)}
    $savedRequest=Get-DetailProfileRequest
    $before=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileReset')
    $staged=Get-DetailProfileRequest
    $eventsEnabled=$excel.EnableEvents
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyCancelSave' @($Fixture.Config))
    try {
        $excel.EnableEvents=$true
        $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileSave')
        $cancelled=[int](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyCancelledCount')
    } finally {
        $excel.EnableEvents=$eventsEnabled
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyStopCancelling')
    }
    Check 'DetailProfile.ExcelSaveCancellationObserved' ($cancelled -eq 1)
    Check 'DetailProfile.CancelledSavePreservesVersionAndEdits' (-not $ok -and $before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash -and (Get-DetailProfileVersion $Fixture) -eq 2 -and $staged -ceq (Get-DetailProfileRequest))
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyCloseAction')
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
    Check 'DetailProfile.CloseDiscardsStagedProfile' ($savedRequest -ceq (Get-DetailProfileRequest) -and $before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    $book=$excel.Workbooks.Open($Fixture.Config,0,$false)
    $fields=Table $book 'tblEventDetailFields'
    $fields.ListColumns.Item('Enabled').DataBodyRange.Cells.Item($fieldCount+1,1).Value2=$false
    $book.Save(); $book.Close($false)
    $malformedHash=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileReload')
    Check 'DetailProfile.MalformedLatestUsesSafeDefaults' ($request -ceq (Get-DetailProfileRequest) -and $malformedHash -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileSaveDirect' @($context,2,$request))
    Check 'DetailProfile.MalformedLatestCannotBeOverwritten' (-not $ok -and $malformedHash -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    $book=$excel.Workbooks.Open($Fixture.Config,0,$false)
    $fields=Table $book 'tblEventDetailFields'
    $fields.ListColumns.Item('Enabled').DataBodyRange.Cells.Item($fieldCount+1,1).Value2=$true
    $book.Save(); $book.Close($false)
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailProfileReload')
}
function Install-DetailProfileReadDiagnostics {
    $model=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modEventDetailModel').CodeModule
    $body=$model.Lines(1,$model.CountOfLines)
    $rules=@(
        @('    Set model = modTrainingJson.DecodeObject(request)','    mDetailDecodeStage = 1: mDetailDecodeRow = 0'),
        @('    If Not ExactFields(model,','    mDetailDecodeStage = 2'),
        @('    If Not IntegerInRange(model(','    mDetailDecodeStage = 3'),
        @('    If TypeName(model(','    mDetailDecodeStage = 4'),
        @('        If Not IsObject(row)','        mDetailDecodeStage = 10: mDetailDecodeRow = mDetailDecodeRow + 1'),
        @('        If Not ExactFields(row,','        mDetailDecodeStage = 11'),
        @('        If VarType(row("EventFamily"))','        mDetailDecodeStage = 12'),
        @('        If Not families.Exists','        mDetailDecodeStage = 13'),
        @('        If VarType(row("Enabled"))','        mDetailDecodeStage = 14'),
        @('        If Not IntegerInRange(row(','        mDetailDecodeStage = 15'),
        @('        If definition("Required")','        mDetailDecodeStage = 16'),
        @('        If fieldsSeen.Exists','        mDetailDecodeStage = 17'),
        @('    If model("Fields").Count','    mDetailDecodeStage = 18'),
        @('    Set Decode = model','    mDetailDecodeStage = 19')
    )
    foreach($rule in $rules){$body=[regex]::Replace($body,'(?im)^'+[regex]::Escape($rule[0]),($rule[1]+"`r`n"+$rule[0]))}
    $model.DeleteLines(1,$model.CountOfLines);$model.AddFromString($body)
    $model.AddFromString(@'
Private mDetailDecodeStage As Long
Private mDetailDecodeRow As Long
Public Function DetailDecodeDiagnostic() As String
    DetailDecodeDiagnostic = CStr(mDetailDecodeStage) & "|" & CStr(mDetailDecodeRow)
End Function
'@)
    $store=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modEventDetailStore').CodeModule
    $body=$store.Lines(1,$store.CountOfLines)
    $rules=@(
        @('    version = 0','    mDetailReadStage = 1: mDetailReadError = 0: mDetailReadDecode = "0|0"'),
        @('    Set wb = ConfigWorkbook','    mDetailReadStage = 2'),
        @('    FindTables wb,','    mDetailReadStage = 3'),
        @('    Set seen =','    mDetailReadStage = 4'),
        @('    If Not modEventDetailModel.IntegerInRange(modTrackingPolicyModel.TableValue(headers, selected,','    mDetailReadStage = 5'),
        @('    If Not modTrainingWire.ValidUtcTimestamp','    mDetailReadStage = 6'),
        @('    If CStr(modTrackingPolicyModel.TableValue(headers, selected,','    mDetailReadStage = 7'),
        @('    For row = 1 To fields.ListRows.Count','    mDetailReadStage = 8'),
        @('    Set model = modEventDetailModel.Decode','    mDetailReadStage = 9'),
        @('    If model Is Nothing','    mDetailReadDecode = modEventDetailModel.DetailDecodeDiagnostic()'),
        @('    version = latest','    mDetailReadStage = 10'),
        @('    ReadProfile = False','    mDetailReadError = Err.Number')
    )
    foreach($rule in $rules){$body=[regex]::Replace($body,'(?im)^'+[regex]::Escape($rule[0]),($rule[1]+"`r`n"+$rule[0]))}
    $store.DeleteLines(1,$store.CountOfLines);$store.AddFromString($body)
    $store.AddFromString(@'
Private mDetailReadStage As Long
Private mDetailReadError As Long
Private mDetailReadDecode As String
Public Function DetailReadDiagnostic() As String
    DetailReadDiagnostic = CStr(mDetailReadStage) & "|" & CStr(mDetailReadError) & "|" & mDetailReadDecode
End Function
'@)
    $packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modEventDetailSettings').CodeModule.AddFromString(@'
Public Function DetailReadDiagnostic() As String
    DetailReadDiagnostic = modEventDetailStore.DetailReadDiagnostic()
End Function
'@)
}
