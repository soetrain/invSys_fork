# D18 real callback/queue/Domain evidence. No Seed activity is manufactured.
. (Join-Path $PSScriptRoot 'Slice4beActivityAssertions.ps1')
function Install-AdminSeedRecordingProbe($Module) {
    $Module.AddFromString(@'
Public Function ExactSeedEventForTest() As String
    ExactSeedEventForTest = mSeedEvent
End Function
'@)
    $probe=$packages['invSys.Core.xlam'].VBProject.VBComponents.Add(1)
    $probe.Name='TestSeedRecording'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Public Function EnablePolicy(Optional ByVal mode As String = "Enabled") As Boolean
    Dim context As String, version As Long, request As String, report As String, model As Object, row As Variant, user As Object
    context = modActivity.CaptureContext()
    If context = "" Then Exit Function
    If Not modTrackingPolicySettings.ReadEditor(context, version, request, report) Then Exit Function
    Set model = modTrackingPolicyModel.Defaults(True)
    If mode = "Capture" Then model("ViewerActionPathCaptureEnabled") = True
    If mode = "ControlOff" Then
        For Each row In model("Controls")
            If row("ControlId") = "ADMIN_DEMO_SEED" Then row("Collect") = False
        Next row
    ElseIf mode = "UserOff" Then
        Set user = CreateObject("Scripting.Dictionary")
        user.Add "UserId", modAuth.GetCurrentUserId(): user.Add "Record", False
        model("Users").Add user
    End If
    EnablePolicy = modTrackingPolicySettings.SavePolicy(context, version, modTrainingJson.EncodeObject(model), report)
End Function
Public Function CatalogCompatible() As Boolean
    Dim id As Variant, before As Object, after As Object, ids As Variant
    ids = modActivityCatalog.ControlIds(36)
    If UBound(ids) <> 155 Or UBound(modActivityCatalog.ControlIds()) < 159 Then Exit Function
    If Not modActivityCatalog.Control("ADMIN_DEMO_SEED", 36) Is Nothing Then Exit Function
    For Each id In ids
        Set before = modActivityCatalog.Control(CStr(id), 36)
        Set after = modActivityCatalog.Control(CStr(id))
        If before Is Nothing Or after Is Nothing Then Exit Function
        If modTrainingJson.EncodeObject(before) <> modTrainingJson.EncodeObject(after) Then Exit Function
    Next id
    CatalogCompatible = True
End Function
Public Function CurrentCatalogVersion() As Long
    CurrentCatalogVersion = modActivityCatalog.CATALOG_VERSION
End Function
Public Function Definition() As String
    Dim record As Object
    Set record = modActivityCatalog.Control("ADMIN_DEMO_SEED")
    If Not record Is Nothing Then Definition = modTrainingJson.EncodeObject(record)
End Function
Public Function Concludes(ByVal wire As String) As Boolean
    Concludes = modEvaluationMatches.CommandCompleted(modTrainingJson.DecodeObject(wire))
End Function
'@)
}

function Test-AdminSeedRecording($Fixture) {
    SelectTarget $Fixture
    $catalogVersion=[int](Run 'invSys.Core.xlam' 'TestSeedRecording.CurrentCatalogVersion')
    if($catalogVersion -lt 37){throw 'Packaged activity catalog predates Demo Seed; not behavioral RED.'}
    $enabled=[bool](Run 'invSys.Core.xlam' 'TestSeedRecording.EnablePolicy')
    Check 'SeedRecording.Fixture.PolicySavedForSignedInTarget' $enabled
    if(-not $enabled){throw 'Seed recording policy fixture unavailable; not behavioral RED.'}
    if(-not [bool](Run 'invSys.Admin.xlam' 'TestSeedOutcome.Prepare' @($Fixture.Warehouse,$false))){throw 'Seed recording queue fixture unavailable; not behavioral RED.'}
    try {
        $state=([string](Run 'invSys.Admin.xlam' 'TestSeedOutcome.State')).Split('|')
        if($state[0] -cne '0' -or $state[7] -cne '0'){throw 'Seed recording requires no active demo stock or new event; not behavioral RED.'}
        $before=@(Get-Slice4beActivityFiles $Fixture)
        $ack=[bool](Run 'invSys.Admin.xlam' 'TestSeedOutcome.Invoke')
        $state=([string](Run 'invSys.Admin.xlam' 'TestSeedOutcome.State')).Split('|')
        $applied=$ack -and $state[0] -ceq '1' -and $state[2] -ceq 'True' -and $state[6] -ceq 'True' -and $state[7] -ceq '24'
        Check 'SeedRecording.Owner.ExactApplicationAndIdentity' $applied
        if(-not $applied){throw 'Seed callback did not establish the expected application; investigate completion before recording.'}
        $eventId=[string](Run 'invSys.Admin.xlam' 'TestSeedOutcome.ExactSeedEventForTest')
        if($eventId -eq ''){throw 'Exact owner event observation unavailable.'}
        $wire=[string](Run 'invSys.Core.xlam' 'TestSeedRecording.Definition')
        $definition=if($wire){$wire|ConvertFrom-Json}else{$null}
        Check 'SeedRecording.Catalog.RegisteredOwnerAndCapability' ($null -ne $definition -and $definition.OwnerId -ceq 'ADMIN_DEMO_INVENTORY' -and $definition.Class -ceq 'Command' -and $definition.Capability -ceq 'ADMIN_MAINT')
        $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content -LiteralPath $_ -Raw|ConvertFrom-Json}|Where-Object ControlId -CEQ 'ADMIN_DEMO_SEED')
        $attempt=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED')
        $outcome=@($records|Where-Object OutcomeCode -CEQ 'APPLIED')
        $pair=$records.Count -eq 2 -and $attempt.Count -eq 1 -and $outcome.Count -eq 1
        Check 'SeedRecording.Applied.CorrelatedPair' ($pair -and $attempt[0].ActivityId -ceq $outcome[0].ActivityId -and $attempt[0].RecordId -cne $outcome[0].RecordId)
        $bound=$pair;$refs=$pair;$safe=$pair;$concludes=$pair
        foreach($record in $records){
            $terminal=$record.OutcomeCode -ceq 'APPLIED'
            $effect=if($terminal){'Changed'}else{'Unknown'}
            $bound=$bound -and $record.OwnerId -ceq 'ADMIN_DEMO_INVENTORY' -and $record.WarehouseId -ceq $Fixture.Warehouse -and $record.StationId -ceq 'S1' -and $record.UserId -ceq 'config-admin' -and $record.CatalogVersion -eq $catalogVersion -and $record.DataEffect -ceq $effect
            $references=@($record.SourceEventRefs)
            if($terminal){
                $refs=$refs -and $references.Count -eq 1
                if($references.Count -eq 1){$refs=$refs -and $references[0].EventId -ceq $eventId -and $references[0].WarehouseId -ceq $Fixture.Warehouse -and $references[0].SourceKind -ceq 'Inventory' -and $references[0].SubmissionState -ceq 'Submitted'}
            }else{$refs=$refs -and $references.Count -eq 0}
            $json=$record|ConvertTo-Json -Depth 20 -Compress
            foreach($value in @($Fixture.Root,$Fixture.Secret,(CredentialHash $Fixture.Secret),'DEMO-RAW','Seed_DemoInventory','PinHash','Processor=')){$safe=$safe -and -not $json.Contains($value)}
            $concludes=$concludes -and ([bool](Run 'invSys.Core.xlam' 'TestSeedRecording.Concludes' @($json)) -eq $terminal)
        }
        Check 'SeedRecording.Applied.BoundActorAndHonestEffect' $bound
        Check 'SeedRecording.Applied.ExactOwnSourceReference' $refs
        Check 'SeedRecording.Applied.NoBusinessOrCredentialPayload' $safe
        Check 'SeedRecording.Applied.OnlyOwnerSuccessConcludes' $concludes
    }finally{[void](Run 'invSys.Admin.xlam' 'TestSeedOutcome.Release')}
    Test-AdminSeedRecordingEdges $Fixture
}

function Test-AdminSeedRecordingEdges($Fixture) {
    function Prepare([bool]$Earlier=$false){
        if(-not [bool](Run 'invSys.Admin.xlam' 'TestSeedOutcome.Prepare' @($Fixture.Warehouse,$Earlier))){throw 'Recording edge fixture unavailable; not product RED.'}
    }
    function InvokeSeed([string]$Path='',[string]$Action='SEED'){
        [string](Run 'invSys.Admin.xlam' 'TestSeedOutcome.InvokeWorkflow' @($Action,$Path,$true))
    }
    function Fingerprint {[string](Run 'invSys.Admin.xlam' 'TestSeedOutcome.WorkflowFingerprint')}
    function CheckPair([string]$Case,[string[]]$Before,[string]$Outcome,[string]$Effect,[string]$EventId=''){
        $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $Before}|ForEach-Object {Get-Content -LiteralPath $_ -Raw|ConvertFrom-Json}|Where-Object ControlId -CEQ 'ADMIN_DEMO_SEED')
        $requests=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$terminals=@($records|Where-Object OutcomeCode -CEQ $Outcome)
        $pair=$records.Count -eq 2 -and $requests.Count -eq 1 -and $terminals.Count -eq 1
        Check ('SeedRecording.'+$Case+'.CorrelatedPair') ($pair -and $requests[0].ActivityId -ceq $terminals[0].ActivityId -and $requests[0].RecordId -cne $terminals[0].RecordId)
        $bound=$pair;$refs=$pair;$safe=$pair;$concludes=$pair
        foreach($record in $records){
            $terminal=$record.OutcomeCode -ceq $Outcome
            $expectedEffect=if($terminal){$Effect}else{'Unknown'}
            $bound=$bound -and $record.WarehouseId -ceq $Fixture.Warehouse -and $record.StationId -ceq 'S1' -and $record.UserId -ceq 'config-admin' -and $record.OwnerId -ceq 'ADMIN_DEMO_INVENTORY' -and $record.DataEffect -ceq $expectedEffect
            $reference=@($record.SourceEventRefs)
            if($terminal -and $EventId -ne ''){
                $refs=$refs -and $reference.Count -eq 1
                if($reference.Count -eq 1){$refs=$refs -and $reference[0].EventId -ceq $EventId -and $reference[0].WarehouseId -ceq $Fixture.Warehouse -and $reference[0].SourceKind -ceq 'Inventory' -and $reference[0].SubmissionState -ceq 'Submitted'}
            }else{$refs=$refs -and $reference.Count -eq 0}
            $json=$record|ConvertTo-Json -Depth 20 -Compress
            foreach($value in @($Fixture.Root,$Fixture.Secret,(CredentialHash $Fixture.Secret),'DEMO-RECORDING','recording-upload.csv','recording-invalid.csv','Processor=','PinHash')){$safe=$safe -and -not $json.Contains($value)}
            $expectedConclusion=$terminal -and $Outcome -cin @('APPLIED','UNCHANGED')
            $concludes=$concludes -and ([bool](Run 'invSys.Core.xlam' 'TestSeedRecording.Concludes' @($json)) -eq $expectedConclusion)
        }
        Check ('SeedRecording.'+$Case+'.ActorAndEffect') $bound
        Check ('SeedRecording.'+$Case+'.ExactReferencesOnly') $refs
        Check ('SeedRecording.'+$Case+'.Sanitized') $safe
        Check ('SeedRecording.'+$Case+'.HonestConclusion') $concludes
    }
    Check 'SeedRecording.Catalog.Prior156DefinitionsUnchanged' ([bool](Run 'invSys.Core.xlam' 'TestSeedRecording.CatalogCompatible'))
    Prepare
    try {
        foreach($case in @(@('NoOp','','UNCHANGED','Unchanged','OK'),@('MissingUpload',(Join-Path $Fixture.Root 'absent.csv'),'REJECTED','Unchanged','FAIL'),@('InvalidUpload',(Join-Path $Fixture.Root 'recording-invalid.csv'),'FAILED','Unknown','FAIL'))){
            if($case[0] -ceq 'InvalidUpload'){[IO.File]::WriteAllText($case[1],"ITEM_CODE,ITEM,QTY,UOM,LOCATION`r`nREAL-RECORDING,Invalid,1,each,TRAINING`r`n",[Text.UTF8Encoding]::new($false))}
            $before=@(Get-Slice4beActivityFiles $Fixture);$state=Fingerprint
            Check ('SeedRecording.'+$case[0]+'.OwnerResultAndNoMutation') ((InvokeSeed $case[1]) -ceq $case[4] -and (Fingerprint) -ceq $state)
            CheckPair $case[0] $before $case[2] $case[3]
        }
        foreach($mode in @('ControlOff','UserOff')){
            if(-not [bool](Run 'invSys.Core.xlam' 'TestSeedRecording.EnablePolicy' @($mode))){throw 'Suppression policy fixture failed.'}
            $before=@(Get-Slice4beActivityFiles $Fixture);$state=Fingerprint
            Check ('SeedRecording.'+$mode+'.BusinessNoOpUnaffected') ((InvokeSeed) -ceq 'OK' -and (Fingerprint) -ceq $state)
            Check ('SeedRecording.'+$mode+'.NoOptionalRecords') (@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}).Count -eq 0)
        }
        if(-not [bool](Run 'invSys.Core.xlam' 'TestSeedRecording.EnablePolicy')){throw 'Enabled policy restoration failed.'}
    }finally{[void](Run 'invSys.Admin.xlam' 'TestSeedOutcome.Release')}
    $source=Join-Path $Fixture.Root 'recording-upload.csv'
    [IO.File]::WriteAllText($source,"ITEM_CODE,ITEM,QTY,UOM,LOCATION`r`nDEMO-RECORDING,Recording fixture,9,each,TRAINING`r`n",[Text.UTF8Encoding]::new($false))
    Prepare $true
    try {
        $before=@(Get-Slice4beActivityFiles $Fixture)
        $ack=InvokeSeed $source
        $state=([string](Run 'invSys.Admin.xlam' 'TestSeedOutcome.State')).Split('|')
        $eventId=[string](Run 'invSys.Admin.xlam' 'TestSeedOutcome.ExactSeedEventForTest')
        Check 'SeedRecording.Pending.OwnEventPendingOtherApplied' ($ack -ceq 'FAIL' -and $state[0] -ceq '1' -and $state[1] -ceq 'NEW' -and $state[2] -ceq 'False' -and $state[3] -ceq 'True' -and $eventId -ne '')
        CheckPair 'Pending' $before 'PENDING' 'Unknown' $eventId
        $recordPins=@(Get-Slice4beActivityFiles $Fixture|ForEach-Object {[pscustomobject]@{Path=$_;Hash=(Get-FileHash -LiteralPath $_).Hash}})
        $applied=[int](Run 'invSys.Admin.xlam' 'TestSeedOutcome.ApplyPending')
        $state=([string](Run 'invSys.Admin.xlam' 'TestSeedOutcome.State')).Split('|')
        Check 'SeedRecording.Pending.LaterApplicationPreservesObservation' ($applied -eq 1 -and $state[2] -ceq 'True' -and @($recordPins|Where-Object {(Get-FileHash -LiteralPath $_.Path).Hash -cne $_.Hash}).Count -eq 0)
        $before=@(Get-Slice4beActivityFiles $Fixture);$fingerprint=Fingerprint
        Check 'SeedRecording.UploadAlias.NoOpAndNoNewEvent' ((InvokeSeed $source 'UPLOAD') -ceq 'OK' -and (Fingerprint) -ceq $fingerprint)
        CheckPair 'UploadAlias' $before 'UNCHANGED' 'Unchanged'
    }finally{[void](Run 'invSys.Admin.xlam' 'TestSeedOutcome.Release')}
}
