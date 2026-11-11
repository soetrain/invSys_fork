# Actual Create handlers; picker adapters supply operator input, never owner results.
. (Join-Path $PSScriptRoot 'Slice4beActivityAssertions.ps1')
function Install-AdminCreateRecordingProbe {
    $admin=$packages['invSys.Admin.xlam'].VBProject
    $core=$packages['invSys.Core.xlam'].VBProject
    $probe=$core.VBComponents.Add(1);$probe.Name='TestCreateRecording'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Public Function CurrentCatalogVersion() As Long
    CurrentCatalogVersion = modActivityCatalog.CATALOG_VERSION
End Function
Public Function Definition(ByVal id As String, ByVal version As Long) As String
    Dim model As Object
    Set model = modActivityCatalog.Control(id, version)
    If Not model Is Nothing Then Definition = modTrainingJson.EncodeObject(model)
End Function
Public Function PriorDefinitionsPreserved() As Boolean
    Dim id As Variant, before As Object, after As Object
    For Each id In modActivityCatalog.ControlIds(40)
        Set before = modActivityCatalog.Control(CStr(id), 40)
        Set after = modActivityCatalog.Control(CStr(id))
        If before Is Nothing Or after Is Nothing Then Exit Function
        If modTrainingJson.EncodeObject(before) <> modTrainingJson.EncodeObject(after) Then Exit Function
    Next id
    PriorDefinitionsPreserved = True
End Function
Public Function SetPolicy(ByVal mode As String) As Boolean
    Dim context As String, version As Long, wire As String, notice As String
    Dim model As Object, row As Variant, user As Object
    context = modActivity.CaptureContext()
    If Not modTrackingPolicySettings.ReadEditor(context, version, wire, notice) Then Exit Function
    Set model = modTrackingPolicyModel.Defaults(mode <> "Disabled")
    model("ViewerActionPathCaptureEnabled") = (mode = "CreateOnly")
    For Each row In model("Controls")
        If Left$(row("ControlId"), 13) = "ADMIN_CREATE_" Then row("Collect") = (mode <> "Disabled" And mode <> "ControlOff")
        If mode = "CreateOnly" Then
            row("Collect") = (row("ControlId") = "ADMIN_CREATE_SUBMIT")
            row("SequenceEligible") = (row("ControlId") = "ADMIN_CREATE_SUBMIT")
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
    Install-AdminDirectoryBoundaries $probe.CodeModule $null -ControlId 'ADMIN_CREATE_SUBMIT' -Outcome 'CREATED_PUBLISHED' -ProbeName 'TestCreateRecording' -RootAction
    $form=$admin.VBComponents.Item('frmCreateWarehouse').CodeModule
    $text=$form.Lines(1,$form.CountOfLines)
    $purpose=if($text -match 'Private Sub mWarehousePurpose_Change\('){'    Call mWarehousePurpose_Change'}else{"    ' No current selection handler; observations remain behavioral RED."}
    $form.AddFromString(@'
Public Sub RecordingGestureForTest(ByVal gesture As String)
    Select Case gesture
        Case "Hub": Call mBtnHubPathHelper_Click
        Case "Share": Call mBtnSharePointHelper_Click
        Case "Purpose"
            mFormBusy = True: mWarehousePurpose.Value = "Training": mFormBusy = False
PURPOSE_HANDLER
        Case "Option": Me.chkPublishInitial.Value = Not CBool(Me.chkPublishInitial.Value)
        Case "Invalid": mFormBusy = True: Me.txtWarehouseId.Value = "": mFormBusy = False: Call btnOK_Click
    End Select
End Sub
'@.Replace('PURPOSE_HANDLER',$purpose))
    # Folder-dialog selections are supplied at the input boundary; native picker acceptance is separate.
    $text=$form.Lines(1,$form.CountOfLines)
    $text=[regex]::Replace($text,'candidate = modDeploymentPaths.BrowseForFolderPath\([^\r\n]+','candidate = TestCreateOwner.PickFolderForTest()')
    $text=[regex]::Replace($text,'candidate = modTesterSetup.DetectSharePointRoot\([^\r\n]+','candidate = ""')
    $text=[regex]::Replace($text,'candidate = modTesterSetup.BrowseForSharePointRoot\([^\r\n]+','candidate = TestCreateOwner.PickFolderForTest()')
    $form.DeleteLines(1,$form.CountOfLines);$form.AddFromString($text)
    $owner=$admin.VBComponents.Item('TestCreateOwner').CodeModule
    $owner.InsertLines($owner.CountOfDeclarationLines+1,'Private mPickerInput As String')
    $owner.AddFromString(@'
Public Function PickFolderForTest() As String
    PickFolderForTest = mPickerInput
End Function
Public Sub Gesture(ByVal name As String, ByVal selection As String)
    mPickerInput = selection
    frmCreateWarehouse.RecordingGestureForTest name
End Sub
'@)
}

function Test-AdminCreateRecording($Fixture,$Other) {
    function Invoke([string]$Name,[object[]]$Values=@()){Run 'invSys.Admin.xlam' ('TestCreateOwner.'+$Name) $Values}
    function Probe([string]$Name,[object[]]$Values=@()){Run 'invSys.Core.xlam' ('TestCreateRecording.'+$Name) $Values}
    $catalogVersion=[int](Probe 'CurrentCatalogVersion')
    function Observe([string]$Suffix,[string]$Outcome,[scriptblock]$Action,[string]$Label=$Suffix){
        $before=@(Get-Slice4beActivityFiles $Fixture)
        & $Action
        $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object{$_ -cnotin $before}|ForEach-Object{Get-Content $_ -Raw|ConvertFrom-Json}|Where-Object ControlId -CEQ ('ADMIN_CREATE_'+$Suffix))
        $attempt=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$result=@($records|Where-Object OutcomeCode -CEQ $Outcome)
        $valid=$records.Count -eq 2 -and $attempt.Count -eq 1 -and $result.Count -eq 1
        if(-not $valid){Write-Output ('Create observation counts '+$Label+': records='+$records.Count+' requested='+$attempt.Count+' result='+$result.Count+' outcomes='+(@($records|ForEach-Object OutcomeCode)-join ','))}
        Check ('Create.Recording.'+$Label+'.Pair') $valid
        if($valid){$valid=$attempt[0].ActivityId -ceq $result[0].ActivityId -and $attempt[0].RecordId -cne $result[0].RecordId}
        foreach($r in $records){$valid=$valid -and $r.CatalogVersion -eq $catalogVersion -and $r.OwnerId -ceq 'ADMIN_WAREHOUSE_CREATE' -and $r.WarehouseId -ceq $Fixture.Warehouse -and $r.UserId -ceq 'config-admin' -and @($r.SourceEventRefs).Count -eq 0}
        Check ('Create.Recording.'+$Label+'.ExactSourceContext') $valid
        $safe=$records.Count -eq 2
        foreach($r in $records){$raw=($r|ConvertTo-Json -Depth 12).Replace('\\','\');foreach($value in @($script:CreateRecordingId,$script:CreateRecordingRoot,$Fixture.Secret,'Generated Create coverage','create-fixture-admin')){if($value -and $raw.Contains($value)){$safe=$false}}}
        Check ('Create.Recording.'+$Label+'.Redacted') $safe
        if($result.Count -eq 1){
            Check ('Create.Recording.'+$Label+'.Conclusion') ([bool](Probe 'Concludes' @(($result[0]|ConvertTo-Json -Depth 12 -Compress))) -eq ($Outcome -cin @('CREATED_LOCAL','CREATED_PUBLISHED','PUBLISHED','OPENED','STAGED','SELECTED','CLOSED')))
            $severity=if($Outcome -cin @('FAILED','CREATED_PUBLISH_FAILED','CREATED_PUBLISH_DENIED')){'Error'}elseif($Outcome -cin @('REJECTED','DENIED','UNAVAILABLE')){'Warning'}else{'Info'}
            Check ('Create.Recording.'+$Label+'.FixedSeverityAndEffect') ($result[0].Severity -ceq $severity -and $result[0].DataEffect -ceq $(if($Outcome -ceq 'CANCELLED'){'Unchanged'}else{'Unknown'}))
        }
    }
    # Visible test host (VT-EXCEL-01): show Excel only while it is hidden and before the host workbook exists. Setting
    # Application.Visible=True on an already-visible Excel with a workbook open recreates the SDI frame and leaves
    # Workbook.Windows and Application.Windows empty until Excel next runs a macro (PR4-BASE-02 v17.3/v17.4 evidence).
    if(-not [bool]$excel.Visible){$excel.Visible=$true}
    $recordingBook=$excel.Workbooks.Add()
    try{
        $recordingWindow=$recordingBook.Windows.Item(1);if(-not [bool]$recordingWindow.Visible){$recordingWindow.Visible=$true}
        if(-not [bool]$excel.Visible){$excel.Visible=$true}
        SelectTarget $Fixture
        if(-not [bool](Probe 'SetPolicy' @('Enabled'))){throw 'Create recording policy prerequisite unavailable.'}
        foreach($suffix in @('OPEN','SUBMIT','RETRY_PUBLISH','PUBLISH_EXISTING','OPEN_WORKBOOK','BROWSE_HUB','FIND_SHAREPOINT','PURPOSE','PUBLISH_OPTION','CLOSE')){
            $wire=[string](Probe 'Definition' @(('ADMIN_CREATE_'+$suffix),41));$d=if($wire){$wire|ConvertFrom-Json}else{$null}
            Check ('Create.Recording.Definition.'+$suffix) ($null -ne $d -and $d.OwnerId -ceq 'ADMIN_WAREHOUSE_CREATE' -and $d.Capability -ceq 'ADMIN_MAINT')
            Check ('Create.Recording.OldCatalogExcludes.'+$suffix) ([string](Probe 'Definition' @(('ADMIN_CREATE_'+$suffix),40)) -ceq '')
        }
        Check 'Create.Recording.PriorDefinitionsPreserved' ([bool](Probe 'PriorDefinitionsPreserved'))
        $script:CreateRecordingId='CR'+[guid]::NewGuid().ToString('N').Substring(0,8)
        $script:CreateRecordingRoot=Join-Path $runRoot 'recording-created'
        Observe 'OPEN' 'OPENED' {[void](Invoke 'OpenForm' @($recordingBook.Name))}
        if($AdminCreateOpeningOnly){
            Observe 'CLOSE' 'CLOSED' {[void](Invoke 'CloseForm')}
            Observe 'OPEN' 'OPENED' {[void](Invoke 'OpenForm' @($recordingBook.Name))} 'Reopen'
            return
        }
        [void](Invoke 'Prepare' @($script:CreateRecordingId,$script:CreateRecordingRoot,'','Operational'))
        Observe 'BROWSE_HUB' 'CANCELLED' {[void](Invoke 'Gesture' @('Hub',''))} 'HubCancel'
        Observe 'BROWSE_HUB' 'STAGED' {[void](Invoke 'Gesture' @('Hub',$script:CreateRecordingRoot))} 'HubSelected'
        Observe 'FIND_SHAREPOINT' 'UNAVAILABLE' {[void](Invoke 'Gesture' @('Share',''))} 'ShareUnavailable'
        $share=Join-Path $runRoot 'recording-publication';New-Item -ItemType Directory -Path $share|Out-Null
        Observe 'FIND_SHAREPOINT' 'STAGED' {[void](Invoke 'Gesture' @('Share',$share))} 'ShareSelected'
        Observe 'PURPOSE' 'SELECTED' {[void](Invoke 'Gesture' @('Purpose',''))}
        Observe 'PUBLISH_OPTION' 'STAGED' {[void](Invoke 'Gesture' @('Option',''))}
        [void](Invoke 'Prepare' @($script:CreateRecordingId,$script:CreateRecordingRoot,'','Training'))
        Observe 'SUBMIT' 'CREATED_LOCAL' {[void](Invoke 'Create')}
        Check 'Create.Recording.LocalOwnerResult' (([string](Invoke 'Facts')).StartsWith('True|True|'))
        Observe 'OPEN_WORKBOOK' 'OPENED' {[void](Invoke 'OpenReceiving')}
        Observe 'CLOSE' 'CLOSED' {[void](Invoke 'CloseForm')}
        Observe 'OPEN' 'OPENED' {[void](Invoke 'OpenForm' @($recordingBook.Name))} 'Reopen'
        [void](Invoke 'Prepare' @('UNRELATED_ENTRY',$script:CreateRecordingRoot,$share,'Operational'))
        Observe 'PUBLISH_EXISTING' 'PUBLISHED' {[void](Invoke 'PublishExisting')}
        Observe 'CLOSE' 'CLOSED' {[void](Invoke 'CloseForm')} 'ReopenedClose'
        [void](Invoke 'OpenForm' @($recordingBook.Name))
        Observe 'SUBMIT' 'REJECTED' {[void](Invoke 'Gesture' @('Invalid',''))} 'Invalid'
        [void](Invoke 'CloseForm')
        foreach($partial in @($false,$true)){
            [void](Invoke 'OpenForm' @($recordingBook.Name))
            $label=if($partial){'Partial'}else{'Published'}
            $script:CreateRecordingId='CR'+[guid]::NewGuid().ToString('N').Substring(0,8)
            $script:CreateRecordingRoot=Join-Path $runRoot ('recording-'+$label)
            $destination=Join-Path $runRoot ('recording-share-'+$label)
            New-Item -ItemType Directory -Path $destination|Out-Null
            if($partial){[IO.File]::WriteAllText((Join-Path $destination $script:CreateRecordingId),'Generated publication obstruction')}
            [void](Invoke 'Prepare' @($script:CreateRecordingId,$script:CreateRecordingRoot,$destination,'Training'))
            Observe 'SUBMIT' $(if($partial){'CREATED_PUBLISH_FAILED'}else{'CREATED_PUBLISHED'}) {[void](Invoke 'Create')} $label
            Check ('Create.Recording.'+$label+'.LocalCreationRetained') ([string](Invoke 'Facts')).StartsWith('True|')
            if($partial){
                $before=(Get-FileHash -LiteralPath (Join-Path $script:CreateRecordingRoot ($script:CreateRecordingId+'.invSys.Config.xlsb'))).Hash
                $destination=Join-Path $runRoot 'recording-retry-success';New-Item -ItemType Directory -Path $destination|Out-Null
                [void](Invoke 'SetShare' @($destination))
                Observe 'RETRY_PUBLISH' 'PUBLISHED' {[void](Invoke 'Create')} 'Retry'
                Check 'Create.Recording.RetryNoBootstrap' ([string](Invoke 'Facts') -ceq 'True|True|Close|True|1|2')
                Check 'Create.Recording.RetryPreservesConfig' ((Get-FileHash -LiteralPath (Join-Path $script:CreateRecordingRoot ($script:CreateRecordingId+'.invSys.Config.xlsb'))).Hash -ceq $before)
            }
            [void](Invoke 'CloseForm')
        }
        . (Join-Path $PSScriptRoot 'Slice4beAdminCreateBoundaries.ps1')
        if([string](Probe 'Definition' @('ADMIN_CREATE_SUBMIT',41)) -cne ''){Test-AdminCreateBoundaries $Fixture $Other $recordingBook}
        if($AdminCreateGuidePaths){
            . (Join-Path $PSScriptRoot 'Slice4beAdminCreateGuide.ps1')
            Test-AdminCreateGuide $Fixture $recordingBook
        }
    }finally{[void](Invoke 'CloseForm');$recordingBook.Close($false);SelectTarget $Fixture}
}
