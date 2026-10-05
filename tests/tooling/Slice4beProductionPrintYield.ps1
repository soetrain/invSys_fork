# Interrupt a real handler at the declared preview seam, after report preparation.
function Install-ProductionPrintYieldProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    $form.InsertLines($form.ProcBodyLine('UserForm_Initialize',0)+1,'    TestProductionDesigner.PrintYieldInitialized')
    $form.AddFromString(@'
Public Function PrintYieldActForTest() As Boolean
    On Error GoTo Failed
    mBtnManagerPrint_Click
    PrintYieldActForTest = True
Failed:
End Function
Public Function PrintYieldGuardsForTest() As Boolean
    PrintYieldGuardsForTest = Not mLoading And Not mDesignerActionInProgress
End Function
'@)
    $adapter=$project.VBComponents.Item('TestProductionDesigner').CodeModule
    $adapter.InsertLines(1,@'
Private mPrintYieldMode As String, mPrintYieldAuth As String, mPrintYieldReached As Boolean
Private mPrintYieldValid As Boolean, mPrintYieldInitializations As Long
Private mPrintYieldBefore As Long, mPrintYieldAfter As Long
'@)
    $start=$adapter.ProcStartLine('PrintPreviewForTest',0)
    $last=$start+$adapter.ProcCountLines('PrintPreviewForTest',0)-1
    while($adapter.Lines($last,1).Trim() -cne 'End Sub'){$last--;if($last -le $start){throw 'Preview end anchor missing; not product RED.'}}
    $adapter.InsertLines($last,'    PrintYieldBoundary')
    $adapter.AddFromString(@'
Public Sub PrintYieldInitialized()
    mPrintYieldInitializations = mPrintYieldInitializations + 1
End Sub
Public Sub PrintYieldArm(ByVal mode As String, ByVal authPath As String)
    mPrintYieldMode = mode: mPrintYieldAuth = authPath
    mPrintYieldReached = False: mPrintYieldValid = False: mPrintYieldInitializations = 0
    mPrintYieldBefore = 0: mPrintYieldAfter = 0
End Sub
Public Sub PrintYieldBoundary()
    Dim mode As String, beforeContext As String, book As Workbook, removed As Boolean
    If mPrintYieldMode = "" Then Exit Sub
    mode = mPrintYieldMode: mPrintYieldMode = "": mPrintYieldReached = True
    beforeContext = modActivity.CaptureContext()
    Select Case mode
        Case "SignedOut"
            modAuth.SignOut
            mPrintYieldValid = Not modAuth.IsSignedIn()
        Case "Target"
            modNasConnection.ClearWarehouseTarget
            mPrintYieldValid = Not modNasConnection.IsTargetResolved()
        Case "Permission"
            PrintYieldRevokePermission
            mPrintYieldValid = Not modRoleUiAccess.CanCurrentUserPerformCapability("PROD_POST") And _
                Not modRoleUiAccess.CanCurrentUserPerformCapability("ADMIN_MAINT") And _
                modAuth.IsSignedIn() And modActivity.CaptureContext() = beforeContext
        Case "ClosedWorkbook"
            mPrintYieldBefore = Application.Workbooks.Count
            mRunLocalCapturedBookForTest.Close False
            mPrintYieldAfter = Application.Workbooks.Count
            removed = True
            For Each book In Application.Workbooks
                If book Is mRunLocalCapturedBookForTest Then removed = False
            Next book
            mPrintYieldValid = removed And mPrintYieldBefore = mPrintYieldAfter + 1
    End Select
End Sub
Private Sub PrintYieldRevokePermission()
    Dim book As Workbook, sheet As Worksheet, table As ListObject, row As ListRow, revoked As Long
    On Error GoTo Failed
    Set book = Application.Workbooks.Open(mPrintYieldAuth, 0, False)
    For Each sheet In book.Worksheets
        For Each table In sheet.ListObjects
            If table.Name = "tblCapabilities" Then
                For Each row In table.ListRows
                    If CStr(row.Range.Cells(1, table.ListColumns("UserId").Index).Value2) = "config-producer" And _
                       CStr(row.Range.Cells(1, table.ListColumns("Capability").Index).Value2) = "PROD_POST" Then
                        row.Range.Cells(1, table.ListColumns("Status").Index).Value2 = "Inactive"
                        revoked = revoked + 1
                    End If
                Next row
            End If
        Next table
    Next sheet
    If revoked <> 1 Then GoTo Failed
    book.Save: book.Close False
    Exit Sub
Failed:
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    On Error GoTo 0
    Err.Raise vbObjectError + 262, , "Print permission interruption fixture unavailable."
End Sub
Public Function PrintYieldAct() As Boolean
    mPrintOwners = 0: mPrintReads = 0: Set mPrintBook = Nothing
    PrintYieldAct = mForm.PrintYieldActForTest()
End Function
Public Function PrintYieldFact(ByVal fact As String) As Boolean
    Select Case fact
        Case "Reached": PrintYieldFact = mPrintYieldReached
        Case "Valid": PrintYieldFact = mPrintYieldValid
        Case "Loaded": PrintYieldFact = modOperationsFormLifetime.IsLoaded(mForm)
        Case "NoReinitialization": PrintYieldFact = (mPrintYieldInitializations = 0)
        Case "GuardsRestored": PrintYieldFact = mForm.PrintYieldGuardsForTest()
    End Select
End Function
'@)
}

function Test-ProductionPrintYield($Fixture,$Other,$Book,$Decoy){
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Hash([string]$Path){$stream=[IO.File]::Open($Path,'Open','Read','ReadWrite');try{(Get-FileHash -InputStream $stream).Hash}finally{$stream.Dispose()}}
    function Fingerprint($Worksheet){ConvertTo-Json -Compress -Depth 10 -InputObject @($Worksheet.UsedRange.Formula)}
    function Pins($Target){$pins=@{};foreach($file in Get-ChildItem -LiteralPath $Target.Root -Recurse -File|Where-Object Name -NotLike '~$*'){$pins[$file.FullName]=Hash $file.FullName};return $pins}
    function Same($Before,$After){if($Before.Count -ne $After.Count){return $false};foreach($path in $Before.Keys){if(-not $After.ContainsKey($path) -or $After[$path] -cne $Before[$path]){return $false}};return $true}
    $auth=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $owned=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    if(-not [IO.Path]::GetFullPath($auth).StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Print auth fixture escaped owned root.'}
    $authBytes=[IO.File]::ReadAllBytes($auth);$authPin=Hash $auth
    $decoyBefore=Fingerprint $Decoy.Worksheets.Item('Production')
    foreach($mode in @('SignedOut','Target','Permission','ClosedWorkbook')){
        $work=$null
        try{
            SelectTarget $Fixture 'config-producer'
            $label='PrintYield.'+$mode;$path=Join-Path $runRoot ($label+'.xlsb')
            if(Test-Path -LiteralPath $path){throw 'Preserve existing Print interruption fixture.'}
            $Book.SaveCopyAs($path);$saved=Hash $path;$work=$excel.Workbooks.Open($path,0,$false)
            $name=$work.Name;$source=Fingerprint $work.Worksheets.Item('Production')
            [void](Probe 'OpenDesigner' @($name));[void](Probe 'RunLocalShowAndCapture' @($name,'PRINT'))
            Initialize-SettingsCapture
            if([InvSysSettingsCapture]::OwnedVisibleForm('Production',[IntPtr]$excel.Hwnd) -eq [IntPtr]::Zero){throw 'Visible Print form prerequisite unavailable.'}
            $before=Pins $Fixture;$otherPins=Pins $Other
            [void](Probe 'ResetPrintPreviewForTest');[void](Probe 'PrintYieldArm' @($mode,$auth));$Decoy.Activate()
            $returned=[bool](Probe 'PrintYieldAct')
            if(-not [bool](Probe 'PrintYieldFact' @('Reached')) -or -not [bool](Probe 'PrintYieldFact' @('Valid'))){throw 'Print preview interruption prerequisite unavailable; not product RED.'}
            Check ($label+'.ActualHandlerReturned') $returned
            Check ($label+'.OneOwnerEntry') ([int](Probe 'PrintOwnerEntries') -eq 1)
            Check ($label+'.OneReportRead') ([int](Probe 'PrintReportReads') -eq 1)
            Check ($label+'.OnePreviewEntry') ([int](Probe 'PrintPreviewCountForTest') -eq 1)
            $visible=([InvSysSettingsCapture]::OwnedVisibleForm('Production',[IntPtr]$excel.Hwnd) -ne [IntPtr]::Zero)
            $loaded=[bool](Probe 'PrintYieldFact' @('Loaded'))
            if($mode -ceq 'ClosedWorkbook'){
                $work=$null
                Check ($label+'.NativeWorkbookRemoved') ($name -cnotin @($excel.Workbooks|ForEach-Object Name))
                Check ($label+'.CapturedBindingRejected') (-not [bool](Probe 'RunLocalClosedBindingCurrent'))
            }else{
                if(-not($visible -and $loaded)){throw 'Surviving Print form unavailable; not product RED.'}
                Check ($label+'.SourceAndExactKeyPreserved') ((Fingerprint $work.Worksheets.Item('Production')) -ceq $source)
            }
            if($visible -and $loaded){
                $expected=if($mode -ceq 'Permission'){'Production permission changed. Reopen Production before continuing.'}else{'Session, warehouse, or captured workbook changed. Reopen Production before editing the draft.'}
                Check ($label+'.VisibleRefusal') ([string](Probe 'PrintStatus') -ceq $expected)
                Check ($label+'.GuardsRestored') ([bool](Probe 'PrintYieldFact' @('GuardsRestored')))
                CaptureOwnedFormByCaptionEvidence 'Production' ($label.ToLowerInvariant()+'.png')
            }else{
                Check ($label+'.DismissedFormNotRecreated') (-not $loaded -and -not $visible)
            }
            Check ($label+'.NoFormReinitialization') ([bool](Probe 'PrintYieldFact' @('NoReinitialization')))
            Check ($label+'.DecoyPreserved') ((Fingerprint $Decoy.Worksheets.Item('Production')) -ceq $decoyBefore)
            [pscustomobject]@{Mode=$mode;NativeWorkbookClosure=($mode -ceq 'ClosedWorkbook');VisibleAfter=$visible;LoadedAfter=$loaded;HandlerReturned=$returned;DismissedControlsQueried=$false}|ConvertTo-Json|Set-Content (Join-Path $reportRoot ($label.ToLowerInvariant()+'.json'))
        }finally{
            [void](Probe 'PrintYieldArm' @('',''));[void](Probe 'RunLocalSafeClose')
            if($null -ne $work){$work.Close($false)}
            [IO.File]::WriteAllBytes($auth,$authBytes)
            SelectTarget $Fixture 'config-producer'
        }
        Check ($label+'.SavedOperatorBytesPreserved') ((Hash $path) -ceq $saved)
        Check ($label+'.AuthorityPreservedAfterAuthRestoration') (Same $before (Pins $Fixture))
        Check ($label+'.OtherWarehousePreserved') (Same $otherPins (Pins $Other))
    }
    Check 'PrintYield.AuthBytesRestored' ((Hash $auth) -ceq $authPin)
}
