# Actual Ribbon/form observations. Only presentation, connection result and
# optional activity-store failure are adapted in disposable compiled packages.
. (Join-Path $PSScriptRoot 'Slice4beActivityAssertions.ps1')
function Install-AggregationRecordingProbe {
    $admin=$packages['invSys.Admin.xlam'].VBProject
    $core=$packages['invSys.Core.xlam'].VBProject
    $probe=$core.VBComponents.Add(1);$probe.Name='TestAggregationActivity'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Public Function Definition(ByVal id As String, ByVal version As Long) As String
    Dim model As Object
    Set model = modActivityCatalog.Control(id, version)
    If Not model Is Nothing Then Definition = modTrainingJson.EncodeObject(model)
End Function
Public Function PriorDefinitionsPreserved() As Boolean
    Dim id As Variant, before As Object, after As Object
    For Each id In modActivityCatalog.ControlIds(42)
        Set before = modActivityCatalog.Control(CStr(id), 42)
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
    model("ViewerActionPathCaptureEnabled") = (mode = "AggregationOnly")
    For Each row In model("Controls")
        If mode = "AggregationOnly" Then
            row("Collect") = (row("ControlId") = "ADMIN_AGGREGATION_AGGREGATE")
            row("SequenceEligible") = (row("ControlId") = "ADMIN_AGGREGATION_AGGREGATE")
        End If
        If Left$(row("ControlId"), 18) = "ADMIN_AGGREGATION_" And mode <> "AggregationOnly" Then row("Collect") = (mode <> "Disabled" And mode <> "ControlOff")
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
    Install-AdminDirectoryBoundaries $probe.CodeModule $null -ControlId 'ADMIN_AGGREGATION_AGGREGATE' -Outcome 'COMPLETED' -ProbeName 'TestAggregationActivity' -RootAction
    $owner=$admin.VBComponents.Item('TestAggregation').CodeModule
    $owner.InsertLines($owner.CountOfDeclarationLines+1,'Private mConnectionResult As Boolean, mTrackingNotices As Long, mAuditFault As Boolean, mAuditReached As Boolean')
    $owner.AddFromString(@'
Public Sub SetAuditFailure(ByVal value As Boolean)
    mAuditFault = value: mAuditReached = False
End Sub
Public Function AuditFailureForTest() As Boolean
    AuditFailureForTest = mAuditFault
    If mAuditFault Then mAuditReached = True
End Function
Public Function AuditFailureReached() As Boolean
    AuditFailureReached = mAuditReached
End Function
Public Function RecordingAcknowledgesPersistence() As Boolean
    If Not mForm Is Nothing Then RecordingAcknowledgesPersistence = mForm.RecordingAcknowledgesPersistenceForTest()
End Function
Public Sub RecordingGesture(ByVal name As String, Optional ByVal root As String = "", Optional ByVal secret As String = "")
    If Not mForm Is Nothing Then mForm.RecordingGestureForTest name, root, secret
End Sub
Public Sub SetConnectionResult(ByVal value As Boolean)
    mConnectionResult = value
End Sub
Public Function ConnectionResult() As Boolean
    ConnectionResult = mConnectionResult
End Function
Public Function TrackingNotice(ByVal text As String, Optional ByVal style As VbMsgBoxStyle = vbOKOnly, Optional ByVal title As String = "") As VbMsgBoxResult
    If Left$(text, 21) = "Tracking unavailable:" Then mTrackingNotices = mTrackingNotices + 1
    TrackingNotice = vbOK
End Function
Public Sub ResetTrackingNotices()
    mTrackingNotices = 0
End Sub
Public Function TrackingNotices() As Long
    TrackingNotices = mTrackingNotices
End Function
Public Function CredentialsCleared() As Boolean
    If Not mForm Is Nothing Then CredentialsCleared = mForm.CredentialsClearedForTest()
End Function
'@)
    $form=$admin.VBComponents.Item('frmAggregationSources').CodeModule
    $form.AddFromString(@'
Public Sub RecordingGestureForTest(ByVal name As String, ByVal root As String, ByVal secret As String)
    Select Case name
        Case "Refresh": mBtnDiscover_Click
        Case "AddServer": mBtnAddServer_Click
        Case "Connect"
            mTxtRoot.Value = root: mTxtUser.Value = "fixture-connection-user": mTxtPassword.Value = secret
            mBtnConnect_Click
    End Select
End Sub
Public Function RecordingAcknowledgesPersistenceForTest() As Boolean
    RecordingAcknowledgesPersistenceForTest = (InStr(1, mLblStatus.Caption, "Complete output was saved", vbTextCompare) > 0)
End Function
Public Function CredentialsClearedForTest() As Boolean
    CredentialsClearedForTest = (CStr(mTxtUser.Value) = "" And CStr(mTxtPassword.Value) = "")
End Function
'@)
    $auditOwner=$admin.VBComponents.Item('modAdminConsole').CodeModule
    $start=$auditOwner.ProcStartLine('RunHQAggregationFromSourceSet',0);$count=$auditOwner.ProcCountLines('RunHQAggregationFromSourceSet',0)
    $body=$auditOwner.Lines($start,$count)
    $marker='        AppendAuditEntry auditWb, "HQ_AGGREGATE",'
    if([regex]::Matches($body,[regex]::Escape($marker)).Count -ne 1){throw 'Post-save workbook-audit boundary unavailable.'}
    $body=$body.Replace($marker,'        If TestAggregation.AuditFailureForTest() Then Err.Raise vbObjectError + 573'+"`r`n"+$marker)
    $auditOwner.DeleteLines($start,$count);$auditOwner.InsertLines($start,$body)
    $connection=$admin.VBComponents.Item('modAdminConsole').CodeModule
    $first=$connection.ProcStartLine('ConnectAggregationServerForAdmin',0);$count=$connection.ProcCountLines('ConnectAggregationServerForAdmin',0)
    $connection.DeleteLines($first,$count)
    $connection.InsertLines($first,@'
Public Function ConnectAggregationServerForAdmin(ByVal serverRoot As String, ByVal userName As String, ByVal windowsPassword As String, Optional ByRef report As String = "") As Boolean
    ConnectAggregationServerForAdmin = TestAggregation.ConnectionResult()
End Function
'@)
    $observer=$admin.VBComponents.Item('cProvisioningActivity').CodeModule
    $body=$observer.Lines(1,$observer.CountOfLines)
    if(-not $body.Contains('MsgBox notice, vbExclamation, "invSys Admin"')){throw 'Optional notice presentation seam changed.'}
    $observer.DeleteLines(1,$observer.CountOfLines)
    $observer.AddFromString($body.Replace('MsgBox notice, vbExclamation, "invSys Admin"','Call TestAggregation.TrackingNotice(notice, vbExclamation, "invSys Admin")'))
}

function Test-AggregationRecording($Fixture,$Other,$HostBook,$Expected) {
    function Activity([string]$Name,[object[]]$Values=@()){Run 'invSys.Core.xlam' ('TestAggregationActivity.'+$Name) $Values}
    function Observe([string]$Suffix,[string]$Outcome,[scriptblock]$Action,[string]$Label=$Suffix,[switch]$NoRetryClaim){
        $before=@(Get-Slice4beActivityFiles $Fixture)
        $ownPins=SourcePins $Fixture;$otherPins=SourcePins $Other
        & $Action
        Check ('Aggregation.Recording.'+$Label+'.SourcesPreserved') ((Preserved $ownPins) -and (Preserved $otherPins))
        $new=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json})
        $records=@($new|Where-Object ControlId -CEQ ('ADMIN_AGGREGATION_'+$Suffix))
        $attempt=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$result=@($records|Where-Object OutcomeCode -CEQ $Outcome)
        $pair=$records.Count -eq 2 -and $attempt.Count -eq 1 -and $result.Count -eq 1
        Check ('Aggregation.Recording.'+$Label+'.Pair') $pair
        if($NoRetryClaim){
            $honest=$result.Count -eq 1
            if($honest){$honest=$result[0].UserMessage -notmatch '(?i)retained (attempt|retry)'}
            Check ('Aggregation.Recording.'+$Label+'.NoUnprovedRetryClaim') $honest
        }
        $valid=$pair
        if($pair){$valid=$attempt[0].ActivityId -ceq $result[0].ActivityId -and $attempt[0].RecordId -cne $result[0].RecordId}
        foreach($record in $records){$valid=$valid -and $record.CatalogVersion -eq 43 -and $record.OwnerId -ceq 'ADMIN_HQ_AGGREGATION' -and $record.WarehouseId -ceq $Fixture.Warehouse -and $record.UserId -ceq 'config-admin' -and $record.StationId -ceq 'S1' -and @($record.SourceEventRefs).Count -eq 0 -and $record.DataEffect -ceq 'Unknown'}
        Check ('Aggregation.Recording.'+$Label+'.ExactSourceContext') $valid
        $safe=$pair
        foreach($record in $records){
            $wire=($record|ConvertTo-Json -Depth 12 -Compress).Replace('\\','\')
            foreach($value in @($Fixture.Root,$Other.Root,$Fixture.Snapshot,$Fixture.Secret,'fixture-connection-user') + @($Expected.Keys|ForEach-Object {($_ -split '\|',2)[1]})){if($value -and $wire.Contains($value)){$safe=$false}}
        }
        Check ('Aggregation.Recording.'+$Label+'.Redacted') $safe
        Check ('Aggregation.Recording.'+$Label+'.NoSpuriousActions') (@($new|Where-Object {$_.ControlId -like 'ADMIN_AGGREGATION_*' -and $_.ControlId -cne ('ADMIN_AGGREGATION_'+$Suffix)}).Count -eq 0)
        $conclusion=$false
        if($result.Count -eq 1){
            $positive=$Outcome -cin @('OPENED','DISCOVERED','STAGED','CONNECTED','SELECTED','COMPLETED','CLOSED')
            $conclusion=([bool](Activity 'Concludes' @(($result[0]|ConvertTo-Json -Depth 12 -Compress))) -eq $positive)
        }
        Check ('Aggregation.Recording.'+$Label+'.Conclusion') $conclusion
    }
    function ExactOutput {
        if(-not(Test-Path -LiteralPath $Fixture.Output)){return $false}
        $actual=ReadEntities $Fixture.Output 'tblGlobalInventorySnapshot'
        if($actual.Count -ne $Expected.Count){return $false}
        foreach($key in $Expected.Keys){if(-not $actual.ContainsKey($key) -or $actual[$key] -ne $Expected[$key]){return $false}}
        return $true
    }
    try{
        SelectTarget $Fixture
        if(-not [bool](Activity 'SetPolicy')){throw 'Aggregation recording policy prerequisite unavailable.'}
        if($AdminAggregationOutcomeOnly){
            . (Join-Path $PSScriptRoot 'Slice4beAdminAggregationOutcomes.ps1')
            Test-AggregationOutcomeRecording $Fixture $Other $HostBook
            return
        }
        if($AdminAggregationGuideOnly){
            . (Join-Path $PSScriptRoot 'Slice4beAdminAggregationGuide.ps1')
            Test-AggregationGuide $Fixture $Other $HostBook $Expected
            return
        }
        foreach($suffix in @('OPEN','REFRESH','ADD_SERVER','CONNECT','SELECT_SOURCE','AGGREGATE','CLOSE')){
            $wire=[string](Activity 'Definition' @(('ADMIN_AGGREGATION_'+$suffix),43));$definition=if($wire){$wire|ConvertFrom-Json}else{$null}
            Check ('Aggregation.Recording.Definition.'+$suffix) ($null -ne $definition -and $definition.OwnerId -ceq 'ADMIN_HQ_AGGREGATION' -and $definition.Capability -ceq 'ADMIN_MAINT')
            Check ('Aggregation.Recording.OldCatalogExcludes.'+$suffix) ([string](Activity 'Definition' @(('ADMIN_AGGREGATION_'+$suffix),42)) -ceq '')
        }
        Check 'Aggregation.Recording.PriorDefinitionsPreserved' ([bool](Activity 'PriorDefinitionsPreserved'))
        Observe 'OPEN' 'OPENED' {if(-not [bool](Invoke 'OpenForm' @($HostBook.Name))){throw 'Actual Aggregation Ribbon/form unavailable.'}}
        Observe 'REFRESH' 'DISCOVERED' {[void](Invoke 'RecordingGesture' @('Refresh'))}
        Observe 'ADD_SERVER' 'STAGED' {[void](Invoke 'RecordingGesture' @('AddServer'))}
        [void](Invoke 'SetConnectionResult' @($false))
        Observe 'CONNECT' 'FAILED' {[void](Invoke 'RecordingGesture' @('Connect',$Fixture.Root,$Fixture.Secret))} 'ConnectionFailed'
        Check 'Aggregation.Recording.ConnectionFailed.CredentialsCleared' ([bool](Invoke 'CredentialsCleared'))
        [void](Invoke 'SetConnectionResult' @($true))
        Observe 'CONNECT' 'CONNECTED' {[void](Invoke 'RecordingGesture' @('Connect',$Fixture.Root,$Fixture.Secret))} 'ConnectionSucceeded'
        Check 'Aggregation.Recording.ConnectionSucceeded.CredentialsCleared' ([bool](Invoke 'CredentialsCleared'))
        Observe 'SELECT_SOURCE' 'SELECTED' {[void](Invoke 'SelectSource' @(''))} 'ClearSelection'
        Observe 'AGGREGATE' 'REJECTED' {[void](Invoke 'Aggregate')} 'NoSelection'
        Check 'Aggregation.Recording.NoSelection.NoOwnerDispatch' ([int](Invoke 'OwnerCalls') -eq 0)
        Observe 'SELECT_SOURCE' 'SELECTED' {[void](Invoke 'SelectSource' @($Fixture.Warehouse))}
        Observe 'AGGREGATE' 'COMPLETED' {[void](Invoke 'Aggregate')}
        Check 'Aggregation.Recording.Completed.ActualOwnerAndExactPersistedOutput' ([int](Invoke 'OwnerCalls') -eq 1 -and [bool](Invoke 'OwnerSuccess') -and (ExactOutput))
        Observe 'CLOSE' 'CLOSED' {[void](Invoke 'CloseForm')}
        foreach($mode in @('Disabled','ControlOff','UserOff','AttemptWriteFailure','OutcomeWriteFailure')){
            SelectTarget $Fixture
            if(-not [bool](Activity 'SetPolicy' @($(if($mode -like '*WriteFailure'){'Enabled'}else{$mode})))){throw 'Aggregation suppression policy unavailable.'}
            if(-not [bool](Invoke 'OpenForm' @($HostBook.Name))){throw 'Suppression control form unavailable.'}
            [void](Invoke 'SelectSource' @($Fixture.Warehouse));[void](Invoke 'ResetTrackingNotices')
            $before=@(Get-Slice4beActivityFiles $Fixture)
            $ownPins=SourcePins $Fixture;$otherPins=SourcePins $Other
            [void](Activity 'ArmBoundary' @($(if($mode -like '*WriteFailure'){$mode}else{''}),'','',''))
            try{
                [void](Invoke 'Aggregate')
                Check ('Aggregation.Recording.'+$mode+'.SourcesPreserved') ((Preserved $ownPins) -and (Preserved $otherPins))
                $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json}|Where-Object ControlId -CEQ 'ADMIN_AGGREGATION_AGGREGATE')
                Check ('Aggregation.Recording.'+$mode+'.OrdinaryOwnerAndExactOutput') ([int](Invoke 'OwnerCalls') -eq 1 -and [bool](Invoke 'OwnerSuccess') -and (ExactOutput))
                $expectedCount=if($mode -ceq 'OutcomeWriteFailure'){1}else{0}
                Check ('Aggregation.Recording.'+$mode+'.HonestEvidence') ($records.Count -eq $expectedCount -and @($records|Where-Object OutcomeCode -CNE 'REQUESTED').Count -eq 0)
                if($mode -like '*WriteFailure'){
                    Check ('Aggregation.Recording.'+$mode+'.ActualFaultReached') ([int](Activity 'WriteFailures') -eq 1)
                    Check ('Aggregation.Recording.'+$mode+'.OneTrackingNotice') ([int](Invoke 'TrackingNotices') -eq 1)
                }
            }finally{[void](Activity 'ArmBoundary' @('','','',''));[void](Invoke 'CloseForm')}
        }
    }finally{[void](Activity 'ArmBoundary' @('','','',''));[void](Invoke 'CloseForm');SelectTarget $Fixture}
}
