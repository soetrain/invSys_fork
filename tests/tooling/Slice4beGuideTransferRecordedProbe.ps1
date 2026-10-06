# Disposable hooks at the real picker and terminal observation boundaries.
# They inject fixture faults; ordinary owners, recording and evaluation stay real.
function Install-GuideTransferRecordedProbe {
    if(-not $script:RecordingProbeInstalled){throw 'Existing recording/evaluation probes must precede transfer fault probes.'}
    $core=$packages['invSys.Core.xlam'].VBProject
    $probe=$core.VBComponents.Add(1);$probe.Name='TestGuideTransferRecorded'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mKind As String, mConfig As String, mCandidate As String, mStore As String, mAuth As String
Private mHits As Long, mSame As Boolean
Public Sub Arm(ByVal kind As String, ByVal config As String, ByVal candidate As String, ByVal store As String, ByVal auth As String)
    mKind = kind: mConfig = config: mCandidate = candidate: mStore = store: mAuth = auth
    mHits = 0: mSame = False
End Sub
Public Function Hits() As Long
    Hits = mHits
End Function
Public Function SameContext() As Boolean
    SameContext = mSame
End Function
Public Sub OwnerBoundary()
    Dim book As Workbook, table As ListObject, row As ListRow, target As WarehouseTarget
    Dim context As String, changed As Long
    If mKind <> "Permission" Then Exit Sub
    context = modActivity.CaptureContext(): mKind = "": mHits = mHits + 1
    Set target = modNasConnection.GetCurrentTarget()
    Set book = Application.Workbooks.Open(mAuth, UpdateLinks:=0, ReadOnly:=False)
    Set table = book.Worksheets("Capabilities").ListObjects("tblCapabilities")
    For Each row In table.ListRows
        If CStr(row.Range.Cells(1, table.ListColumns("UserId").Index).Value2) = "config-admin" And _
           CStr(row.Range.Cells(1, table.ListColumns("Capability").Index).Value2) = "ACTION_PATH_MAINT" Then
            row.Range.Cells(1, table.ListColumns("Status").Index).Value2 = "Inactive"
            changed = changed + 1
        End If
    Next row
    book.Save: book.Close SaveChanges:=False
    If changed <> 1 Then Err.Raise 5, , "Expected one disposable maintenance grant."
    modAuth.LoadAuth target.WarehouseId
    mSame = (context <> "" And context = modActivity.CaptureContext())
End Sub
Public Sub TerminalBoundary()
    Dim fs As Object, stream As Object, kind As String, context As String
    If mKind = "" Or mKind = "Permission" Then Exit Sub
    kind = mKind: mKind = "": mHits = mHits + 1
    context = modActivity.CaptureContext(): Set fs = CreateObject("Scripting.FileSystemObject")
    If kind = "PolicyChanged" Or kind = "PolicyUnreadable" Then
        fs.CopyFile mCandidate, mConfig, True
    ElseIf kind = "StoreUnavailable" Then
        fs.MoveFolder mStore, mStore & "-transfer-recorded-held"
        Set stream = fs.CreateTextFile(mStore, False)
        stream.Write "Disposable transfer terminal fault": stream.Close
    Else
        Err.Raise 5, , "Unknown transfer fault fixture."
    End If
    mSame = (context <> "" And context = modActivity.CaptureContext())
End Sub
'@)
    $module=$core.VBComponents.Item('modActivity').CodeModule
    $start=$module.ProcStartLine('FinishAction',0);$end=$start+$module.ProcCountLines('FinishAction',0)
    $lines=@(for($i=$start;$i -lt $end;$i++){
        if($module.Lines($i,1).Trim() -ieq 'If Not modActivityPolicy.ReadPolicy(target, action("ControlId"), version, collect, visible, notice) Then GoTo CleanExit'){$i}
    })
    if($lines.Count -ne 1){throw 'Transfer terminal observation seam changed; not product RED.'}
    $module.InsertLines($lines[0],'    If action("ControlId") = "VIEWER_GUIDE_EXPORT" Or action("ControlId") = "VIEWER_GUIDE_IMPORT" Then TestGuideTransferRecorded.TerminalBoundary')
    $picker=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('modGuideTransferUi').CodeModule
    $start=$picker.ProcStartLine('SelectFileForTransfer',0)
    $picker.InsertLines($start+1,'    Application.Run "''invSys.Core.xlam''!TestGuideTransferRecorded.OwnerBoundary"')
}
