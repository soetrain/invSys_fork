# Opt-in, unsaved diagnostic for a failing real Settings constructor. The trace
# contains attempt ordinal, fixed constructor stage and numeric VBA error only.
# 0=start, 1=context, 2=layout, 3=config rows, 4=connection policy, 5=carriers,
# 6=UOMs, 7=constructor returned. An interrupted trace is not acceptance evidence.
function Install-Slice4beSettingsOpenTrace($TestModule,$FormCode,[string]$OutputPath) {
    $module=$TestModule.CodeModule
    $start=$module.ProcStartLine('OpenSettings',0)
    $count=$module.ProcCountLines('OpenSettings',0)
    $module.DeleteLines($start,$count)
    $module.InsertLines($start,@'
Public Sub OpenSettings()
    Dim failure As Long
    On Error GoTo Failed
    mSettingsTraceAttempt = mSettingsTraceAttempt + 1
    SettingsOpenStageForTest 0
    Set mForm = New frmAdminSettings
    SettingsOpenStageForTest 7
    Exit Sub
Failed:
    failure = Err.Number
    WriteSettingsOpenTraceForTest failure
    Err.Raise failure, "Settings fixture", "Settings construction failed."
End Sub
'@)
    $helpers=@'
Private mSettingsTraceAttempt As Long
Private mSettingsTraceStage As Long
Public Sub SettingsOpenStageForTest(ByVal stage As Long)
    mSettingsTraceStage = stage
    WriteSettingsOpenTraceForTest 0
End Sub
Private Sub WriteSettingsOpenTraceForTest(ByVal failure As Long)
    Dim file As Integer, opened As Boolean
    On Error GoTo Failed
    file = FreeFile
    Open "__TRACE_FILE__" For Append As #file
    opened = True
    Print #file, CStr(mSettingsTraceAttempt) & "," & CStr(mSettingsTraceStage) & "," & CStr(failure)
Failed:
    On Error Resume Next
    If opened Then Close #file
End Sub
'@
    $module.AddFromString($helpers.Replace('__TRACE_FILE__',$OutputPath.Replace('"','""')))
    $first=$FormCode.ProcBodyLine('UserForm_Initialize',0)
    $last=$FormCode.ProcStartLine('UserForm_Initialize',0)+$FormCode.ProcCountLines('UserForm_Initialize',0)-1
    $stages=@('CaptureTargetContext','BuildLayout','LoadConfigRows','LoadConnectionPolicy','LoadCarriers','LoadUoms')
    $insertions=@()
    for($line=$first;$line -le $last;$line++) {
        $statement=$FormCode.Lines($line,1).Trim()
        for($index=0;$index -lt $stages.Count;$index++) {
            if($statement -ieq $stages[$index]){$insertions+=,[pscustomobject]@{Line=$line;Stage=$index+1}}
        }
    }
    if($insertions.Count -ne 6 -or @($insertions|Group-Object Stage|Where-Object Count -NE 1).Count) {
        throw 'Settings constructor stages are unavailable or ambiguous.'
    }
    foreach($entry in $insertions|Sort-Object Line -Descending) {
        $FormCode.InsertLines($entry.Line,('    TestD5Commands.SettingsOpenStageForTest '+$entry.Stage))
    }
    # Fixed layout milestones distinguish native control construction from
    # editor initialization without recording captions, values or context.
    $first=$FormCode.ProcBodyLine('BuildLayout',0)
    $last=$FormCode.ProcStartLine('BuildLayout',0)+$FormCode.ProcCountLines('BuildLayout',0)-1
    $anchors=@(
        'Me.Caption = "invSys Settings"',
        'Set mPages = Me.Controls.Add("Forms.MultiPage.1", "mpSettings", True)',
        'Set mControlParent = mPages.Pages(0)',
        'Set mLstConfig = AddListBox("lstConfig", 12, 58, 680, 225)',
        'AddLabel "lblServerConnection", "Server Connection", 12, 338, 150, 18, True',
        'AddLabel "lblSection", "Shipping Carriers", 12, 420, 150, 18, True',
        'AddLabel "lblUomSection", "Recipe UOM Catalog", 365, 420, 170, 18, True',
        'Set mTrackingSections = mPages.Pages(1).Controls.Add("Forms.MultiPage.1", "mpEventTracking", True)',
        'Set mControlParent = mTrackingSections.Pages(0)',
        'mTracking.Initialize mControlParent, mActivityContext',
        'Set mControlParent = mTrackingSections.Pages(1)',
        'mDetail.Initialize mControlParent, mActivityContext',
        'Set mControlParent = mTrackingSections.Pages(2)',
        'mPreference.Initialize mControlParent, mActivityContext',
        'Set mControlParent = Me',
        'Set mControlParent = Nothing'
    )
    $insertions=@()
    for($line=$first;$line -le $last;$line++) {
        $statement=$FormCode.Lines($line,1).Trim()
        for($index=0;$index -lt $anchors.Count;$index++) {
            if($statement -ieq $anchors[$index]){$insertions+=,[pscustomobject]@{Line=$line;Stage=100+$index}}
        }
    }
    if($insertions.Count -ne $anchors.Count -or @($insertions|Group-Object Stage|Where-Object Count -NE 1).Count) {
        throw 'Settings layout milestones are unavailable or ambiguous.'
    }
    foreach($entry in $insertions|Sort-Object Line -Descending) {
        $FormCode.InsertLines($entry.Line,('    TestD5Commands.SettingsOpenStageForTest '+$entry.Stage))
    }
}
