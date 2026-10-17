# Actual packaged Tester callback/handlers. Only folder/confirmation input is adapted.
. (Join-Path $PSScriptRoot 'Slice4beActivityAssertions.ps1')
function Install-AdminTesterRecordingProbe {
    $admin=$packages['invSys.Admin.xlam'].VBProject
    $core=$packages['invSys.Core.xlam'].VBProject
    $probe=$core.VBComponents.Add(1);$probe.Name='TestTesterRecording'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mPickerInput As String, mPickerPending As Boolean
Private mSetupCalls As Long, mDeleteCalls As Long
Public Function CurrentCatalogVersion() As Long
    CurrentCatalogVersion = modActivityCatalog.CATALOG_VERSION
End Function
Public Sub ResetOwnerCalls()
    mSetupCalls = 0: mDeleteCalls = 0
End Sub
Public Sub OwnerEntered(ByVal action As String)
    If action = "SETUP" Then mSetupCalls = mSetupCalls + 1
    If action = "DELETE" Then mDeleteCalls = mDeleteCalls + 1
End Sub
Public Function OwnerCalls(ByVal action As String) As Long
    If action = "SETUP" Then OwnerCalls = mSetupCalls
    If action = "DELETE" Then OwnerCalls = mDeleteCalls
End Function
Public Sub ArmPicker(ByVal value As String)
    mPickerInput = value: mPickerPending = True
End Sub
Public Function PickerPending() As Boolean
    PickerPending = mPickerPending
End Function
Public Function PickerResult() As String
    PickerResult = mPickerInput: mPickerPending = False
End Function
Public Function Definition(ByVal id As String, ByVal version As Long) As String
    Dim model As Object
    Set model = modActivityCatalog.Control(id, version)
    If Not model Is Nothing Then Definition = modTrainingJson.EncodeObject(model)
End Function
Public Function PriorDefinitionsPreserved() As Boolean
    Dim id As Variant, before As Object, after As Object
    For Each id In modActivityCatalog.ControlIds(41)
        Set before = modActivityCatalog.Control(CStr(id), 41)
        Set after = modActivityCatalog.Control(CStr(id))
        If before Is Nothing Or after Is Nothing Then Exit Function
        If modTrainingJson.EncodeObject(before) <> modTrainingJson.EncodeObject(after) Then Exit Function
    Next id
    PriorDefinitionsPreserved = True
End Function
Public Function SetPolicy(Optional ByVal mode As String = "Enabled") As Boolean
    Dim context As String, version As Long, wire As String, notice As String
    Dim model As Object, row As Variant, user As Object
    context = modActivity.CaptureContext()
    If Not modTrackingPolicySettings.ReadEditor(context, version, wire, notice) Then Exit Function
    Set model = modTrackingPolicyModel.Defaults(mode <> "Disabled")
    model("ViewerActionPathCaptureEnabled") = (mode = "TesterOnly")
    For Each row In model("Controls")
        If Left$(row("ControlId"), 13) = "ADMIN_TESTER_" Then row("Collect") = (mode <> "Disabled" And mode <> "ControlOff")
        If mode = "TesterOnly" Then
            row("Collect") = (row("ControlId") = "ADMIN_TESTER_SETUP")
            row("SequenceEligible") = (row("ControlId") = "ADMIN_TESTER_SETUP")
        End If
    Next row
    If mode = "UserOff" Then
        Set user = CreateObject("Scripting.Dictionary")
        user.Add "UserId", "config-admin": user.Add "Record", False
        model("Users").Add user
    End If
    SetPolicy = modTrackingPolicySettings.SavePolicy(context, version, modTrainingJson.EncodeObject(model), notice)
End Function
Public Function Concludes(ByVal wire As String) As Boolean
    Concludes = modEvaluationMatches.CommandCompleted(modTrainingJson.DecodeObject(wire))
End Function
'@)
    . (Join-Path $PSScriptRoot 'Slice4beAdminDirectoryBoundaries.ps1')
    Install-AdminDirectoryBoundaries $probe.CodeModule $null -ControlId 'ADMIN_TESTER_SETUP' -Outcome 'COMPLETED' -ProbeName 'TestTesterRecording' -RootAction
    $commands=$core.VBComponents.Item('modTesterCommands').CodeModule
    foreach($entry in @(@('Setup','SETUP'),@('DeleteGenerated','DELETE'))){
        $start=$commands.ProcStartLine($entry[0],0);$count=$commands.ProcCountLines($entry[0],0)
        $body=$commands.Lines($start,$count).Replace('    On Error GoTo Failed',('    TestTesterRecording.OwnerEntered "'+$entry[1]+'"'+"`r`n"+'    On Error GoTo Failed'))
        $commands.DeleteLines($start,$count);$commands.InsertLines($start,$body)
    }
    $form=$admin.VBComponents.Item('frmSetupTesterStation').CodeModule
    $form.AddFromString(@'
Public Sub RecordingGestureForTest(ByVal gesture As String)
    Select Case gesture
        Case "Hub": Call mBtnHubPathHelper_Click
        Case "Share": Call mBtnSharePointHelper_Click
        Case "Open": Call mBtnOpen_Click
        Case "Invalid": mFormBusy = True: Me.txtWarehouseId.Value = "": mFormBusy = False: Call btnOK_Click
    End Select
End Sub
Public Function RecordingCredentialSafeForTest(ByVal wire As String) As Boolean
    Dim secret As String
    secret = CStr(Me.txtWarehouseName.Value)
    RecordingCredentialSafeForTest = (secret = "")
    If secret <> "" Then RecordingCredentialSafeForTest = (InStr(1, wire, secret, vbBinaryCompare) = 0)
    secret = ""
End Function
'@)
    foreach($entry in @(@($core,'modDeploymentPaths','BrowseForFolderPath'),@($admin,'modTesterSetup','BrowseForSharePointRoot'))){
        $module=$entry[0].VBComponents.Item($entry[1]).CodeModule
        $start=$module.ProcStartLine($entry[2],0);$count=$module.ProcCountLines($entry[2],0)
        $body=$module.Lines($start,$count)
        $body=[regex]::Replace($body,'(?m)^(    On Error GoTo )',('    If TestTesterRecording.PickerPending() Then '+$entry[2]+' = TestTesterRecording.PickerResult(): Exit Function'+"`r`n"+'$1'))
        $module.DeleteLines($start,$count);$module.InsertLines($start,$body)
    }
    $module=$core.VBComponents.Item('modTesterProvisioning').CodeModule
    $start=$module.ProcStartLine('DetectSharePointRoot',0);$count=$module.ProcCountLines('DetectSharePointRoot',0)
    $body=$module.Lines($start,$count).Replace('    On Error GoTo FailDetect',"    If TestTesterRecording.PickerPending() Then Exit Function`r`n    On Error GoTo FailDetect")
    $module.DeleteLines($start,$count);$module.InsertLines($start,$body)
    $owner=$admin.VBComponents.Item('TestTesterOwner').CodeModule
    $owner.AddFromString(@'
Public Sub Gesture(ByVal name As String, ByVal selection As String)
    If name = "Hub" Or name = "Share" Then TestTesterRecording.ArmPicker selection
    frmSetupTesterStation.RecordingGestureForTest name
End Sub
Public Function CredentialSafe(ByVal wire As String) As Boolean
    CredentialSafe = frmSetupTesterStation.RecordingCredentialSafeForTest(wire)
End Function
'@)
}

function Test-AdminTesterRecording($Fixture,$Other) {
    function Invoke([string]$Name,[object[]]$Values=@()){Run 'invSys.Admin.xlam' ('TestTesterOwner.'+$Name) $Values}
    function Probe([string]$Name,[object[]]$Values=@()){Run 'invSys.Core.xlam' ('TestTesterRecording.'+$Name) $Values}
    $catalogVersion=[int](Probe 'CurrentCatalogVersion')
    function Observe([string]$Suffix,[string]$Outcome,[scriptblock]$Action,[string]$Label=$Suffix){
        $before=@(Get-Slice4beActivityFiles $Fixture)
        [void](Probe 'ResetOwnerCalls')
        & $Action
        if($Suffix -cin @('SETUP','DELETE')){Check ('Tester.Recording.'+$Label+'.OwnerInvocationCount') ([int](Probe 'OwnerCalls' @($Suffix)) -eq $(if($Outcome -ceq 'COMPLETED'){1}else{0}))}
        $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object{$_ -cnotin $before}|ForEach-Object{Get-Content $_ -Raw|ConvertFrom-Json}|Where-Object ControlId -CEQ ('ADMIN_TESTER_'+$Suffix))
        $attempt=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$result=@($records|Where-Object OutcomeCode -CEQ $Outcome)
        $valid=$records.Count -eq 2 -and $attempt.Count -eq 1 -and $result.Count -eq 1
        Check ('Tester.Recording.'+$Label+'.Pair') $valid
        if($valid){$valid=$attempt[0].ActivityId -ceq $result[0].ActivityId -and $attempt[0].RecordId -cne $result[0].RecordId}
        foreach($r in $records){$valid=$valid -and $r.CatalogVersion -eq $catalogVersion -and $r.OwnerId -ceq 'ADMIN_TESTER_SETUP' -and $r.WarehouseId -ceq $Fixture.Warehouse -and $r.UserId -ceq 'config-admin' -and @($r.SourceEventRefs).Count -eq 0}
        Check ('Tester.Recording.'+$Label+'.ExactSourceContext') $valid
        $safe=$records.Count -eq 2
        foreach($r in $records){$raw=($r|ConvertTo-Json -Depth 12).Replace('\\','\');foreach($value in @($id,$root,$Fixture.Secret,'fixture-user')){if($value -and $raw.Contains($value)){$safe=$false}}}
        Check ('Tester.Recording.'+$Label+'.Redacted') $safe
        if($Suffix -cne 'CLOSE'){
            $credentialSafe=$records.Count -eq 2
            foreach($r in $records){$credentialSafe=$credentialSafe -and [bool](Invoke 'CredentialSafe' @(($r|ConvertTo-Json -Depth 12 -Compress)))}
            Check ('Tester.Recording.'+$Label+'.CredentialExcluded') $credentialSafe
        }
        if($result.Count -eq 1){
            Check ('Tester.Recording.'+$Label+'.Conclusion') ([bool](Probe 'Concludes' @(($result[0]|ConvertTo-Json -Depth 12 -Compress))) -eq ($Outcome -cin @('COMPLETED','OPENED','STAGED','CLOSED')))
            $severity=if($Outcome -ceq 'FAILED'){'Error'}elseif($Outcome -cin @('REJECTED','DENIED','UNAVAILABLE')){'Warning'}else{'Info'}
            Check ('Tester.Recording.'+$Label+'.SeverityAndEffect') ($result[0].Severity -ceq $severity -and $result[0].DataEffect -ceq $(if($Outcome -ceq 'CANCELLED'){'Unchanged'}else{'Unknown'}))
        }
    }
    $book=$excel.Workbooks.Add()
    $id='TEST'+[guid]::NewGuid().ToString('N').Substring(0,8).ToUpperInvariant()
    $root=[IO.Path]::GetFullPath((Join-Path $runRoot 'recording-tester'))
    $operator=Join-Path $root ($id+'.Receiving.Operator.xlsm')
    $keep=Join-Path $root 'unknown-preserved.txt'
    try{
        $book.Windows.Item(1).Visible=$true;$excel.Visible=$true
        SelectTarget $Fixture
        if(-not [bool](Probe 'SetPolicy')){throw 'Tester recording policy prerequisite unavailable.'}
        if($AdminTesterGuideOnly){
            . (Join-Path $PSScriptRoot 'Slice4beAdminTesterGuide.ps1')
            Test-AdminTesterGuide $Fixture $book
            return
        }
        $sourceHash=(Get-FileHash -LiteralPath $Fixture.Config).Hash
        foreach($suffix in @('OPEN','SETUP','DELETE','OPEN_WORKBOOK','BROWSE_HUB','FIND_SHAREPOINT','CLOSE')){
            $wire=[string](Probe 'Definition' @(('ADMIN_TESTER_'+$suffix),42));$d=if($wire){$wire|ConvertFrom-Json}else{$null}
            Check ('Tester.Recording.Definition.'+$suffix) ($null -ne $d -and $d.OwnerId -ceq 'ADMIN_TESTER_SETUP' -and $d.Capability -ceq 'ADMIN_MAINT')
            Check ('Tester.Recording.OldCatalogExcludes.'+$suffix) ([string](Probe 'Definition' @(('ADMIN_TESTER_'+$suffix),41)) -ceq '')
        }
        Check 'Tester.Recording.PriorDefinitionsPreserved' ([bool](Probe 'PriorDefinitionsPreserved'))
        Observe 'OPEN' 'OPENED' {if(-not [bool](Invoke 'OpenForm' @($book.Name))){throw 'Actual Tester form did not open.'}}
        [void](Invoke 'PrepareSetup' @($id,$root,''))
        Observe 'BROWSE_HUB' 'CANCELLED' {[void](Invoke 'Gesture' @('Hub',''))} 'HubCancel'
        Observe 'BROWSE_HUB' 'STAGED' {[void](Invoke 'Gesture' @('Hub',$root))} 'HubSelected'
        Observe 'FIND_SHAREPOINT' 'UNAVAILABLE' {[void](Invoke 'Gesture' @('Share',''))} 'ShareUnavailable'
        Observe 'FIND_SHAREPOINT' 'STAGED' {[void](Invoke 'Gesture' @('Share',(Join-Path $runRoot 'recording-publication')))} 'ShareSelected'
        [void](Invoke 'PrepareSetup' @($id,$root,''))
        Observe 'SETUP' 'COMPLETED' {Check 'Tester.Recording.SetupOwnerSucceeded' ([bool](Invoke 'Setup'))}
        Check 'Tester.Recording.ReceiptCreated' (Test-Path -LiteralPath (Join-Path $root 'invSys.TesterOwnership.json'))
        Observe 'OPEN_WORKBOOK' 'OPENED' {[void](Invoke 'Gesture' @('Open',''))}
        $opened=$null
        foreach($wb in @($excel.Workbooks)){if($wb.FullName -ieq $operator){$opened=$wb;break}}
        Check 'Tester.Recording.ExactGeneratedWorkbookOpened' ($null -ne $opened)
        if($null -ne $opened){$opened.Close($false)}
        Observe 'CLOSE' 'CLOSED' {[void](Invoke 'CloseForm')}
        Observe 'OPEN' 'OPENED' {[void](Invoke 'OpenForm' @($book.Name))} 'Reopen'
        [void](Invoke 'Prepare' @($id,$root,'Cancel'))
        Observe 'DELETE' 'CANCELLED' {[void](Invoke 'DeleteGenerated')} 'DeleteCancelled'
        Check 'Tester.Recording.CancelPreservesGenerated' (Test-Path -LiteralPath $operator)
        [IO.File]::WriteAllText($keep,'Unknown owned-fixture file must survive cleanup.')
        $keepHash=(Get-FileHash -LiteralPath $keep).Hash
        [void](Invoke 'Prepare' @($id,$root,'Admin'))
        Observe 'DELETE' 'COMPLETED' {[void](Invoke 'DeleteGenerated')}
        Check 'Tester.Recording.CleanupConfirmed' (-not (Test-Path -LiteralPath $operator) -and -not (Test-Path -LiteralPath (Join-Path $root 'invSys.TesterOwnership.json')))
        Check 'Tester.Recording.UnknownFilePreserved' ((Get-FileHash -LiteralPath $keep).Hash -ceq $keepHash)
        Observe 'SETUP' 'REJECTED' {[void](Invoke 'Gesture' @('Invalid',''))} 'InvalidSetup'
        Observe 'CLOSE' 'CLOSED' {[void](Invoke 'CloseForm')} 'ReopenedClose'
        Check 'Tester.Recording.SourceConfigPreserved' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $sourceHash)
        . (Join-Path $PSScriptRoot 'Slice4beAdminTesterBoundaries.ps1')
        Test-AdminTesterBoundaries $Fixture $Other $book
        if($AdminTesterGuidePaths){
            . (Join-Path $PSScriptRoot 'Slice4beAdminTesterGuide.ps1')
            Test-AdminTesterGuide $Fixture $book
        }
    }finally{[void](Invoke 'CloseForm');$book.Close($false);SelectTarget $Fixture}
}
