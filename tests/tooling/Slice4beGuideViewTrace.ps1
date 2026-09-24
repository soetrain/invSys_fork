# Fixed procedure labels and clock readings only, in disposable compiled packages.
# No arguments, operational values, DoEvents, extra refreshes or policy bypasses.
function Install-Slice4beGuideViewTrace {
    $targets=@{
        Operations=@{
            frmActionPathView=@('RefreshView','ApplyLayout','UserForm_Activate','UserForm_Layout')
            frmActionPaths=@('ReadSelection','ReadViewEvaluation','UserForm_Activate')
            modInventoryViewer=@('GuideDraftControlForTest')
        }
        Core=@{
            modPathPresentation=@('Capture','Read','ReadPair')
            modActivityPolicy=@('ReadPolicy')
            modConfig=@('LoadConfig','CloseTransientConfigAfterLoad')
        }
    }
    foreach($role in @('Operations','Core')){
        $book=$packages['invSys.'+$role+'.xlam']
        if($book.ReadOnly -or -not [string]::Equals($book.FullName,(Join-Path $probeDeploy ('invSys.'+$role+'.xlam')),[StringComparison]::OrdinalIgnoreCase)){
            throw 'View tracing is restricted to writable disposable package copies.'
        }
        $observer=$book.VBProject.VBComponents.Add(1)
        $observer.Name='TestViewTrace'
        $tracePath=(Join-Path $reportRoot ('view-calls-'+$role+'.tsv')).Replace('"','""')
        $source=@'
Option Explicit
#If VBA7 Then
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If
Private mRefreshEntries As Long
Public Function RefreshEntries() As Long
    RefreshEntries = mRefreshEntries
End Function
Public Sub Mark(ByVal label As String)
    Dim handle As Integer, tick As Long
    If label = "frmActionPathView.RefreshView|Enter" Then mRefreshEntries = mRefreshEntries + 1
    tick = GetTickCount()
    handle = FreeFile
    Open "TRACE_PATH" For Append As #handle
    Print #handle, Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & CStr(tick) & vbTab & label
    Close #handle
End Sub
'@
        $observer.CodeModule.AddFromString($source.Replace('TRACE_PATH',$tracePath))
        foreach($component in $targets[$role].Keys){
            $module=$book.VBProject.VBComponents.Item($component).CodeModule
            foreach($procedure in $targets[$role][$component]){
                $start=$module.ProcStartLine($procedure,0)
                $count=$module.ProcCountLines($procedure,0)
                $body=$module.ProcBodyLine($procedure,0)-$start
                $lines=[string[]]($module.Lines($start,$count) -split '\r?\n')
                while($lines[$body].TrimEnd().EndsWith('_')){$body++}
                $label=$component+'.'+$procedure
                $updated=New-Object 'System.Collections.Generic.List[string]'
                for($index=0;$index -lt $lines.Count;$index++){
                    $line=$lines[$index]
                    if($index -gt $body){
                        # Preserve single-line If scope: trace and Exit share the same branch.
                        $line=$line -replace '\bExit (Function|Sub)\b',('TestViewTrace.Mark "'+$label+'|Exit": Exit $1')
                        if($line -match '^\s*End (Function|Sub)\s*$'){$updated.Add('    TestViewTrace.Mark "'+$label+'|End"')}
                    }
                    $updated.Add($line)
                    if($index -eq $body){$updated.Add('    TestViewTrace.Mark "'+$label+'|Enter"')}
                }
                $module.DeleteLines($start,$count)
                $module.InsertLines($start,($updated -join "`r`n"))
            }
        }
    }
    if($CheckGuideLayoutStabilityForTest){
        $form=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('frmActionPathView').CodeModule
        $form.AddFromString(@'
Public Function LayoutStabilityForTest(ByVal resized As Boolean) As String
    Dim before As Long, width As Single
    before = TestViewTrace.RefreshEntries()
    If resized Then
        width = Me.Width
        Me.Width = width + 20
        UserForm_Layout
        Me.Width = width
        UserForm_Layout
    Else
        UserForm_Layout
        UserForm_Layout
    End If
    LayoutStabilityForTest = CStr(TestViewTrace.RefreshEntries() - before)
End Function
'@)
        $module=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('modInventoryViewer').CodeModule
        $module.AddFromString(@'
Public Function GuideLayoutStabilityForTest(ByVal resized As Boolean) As String
    Dim instance As Object, view As frmActionPathView
    GuideLayoutStabilityForTest = "MISSING"
    For Each instance In VBA.UserForms
        If TypeName(instance) = "frmActionPathView" Then
            Set view = instance
            GuideLayoutStabilityForTest = view.LayoutStabilityForTest(resized)
            Exit Function
        End If
    Next instance
End Function
'@)
    }
}
