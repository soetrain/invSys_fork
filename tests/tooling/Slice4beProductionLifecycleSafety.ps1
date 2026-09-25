# Unsaved adapters only. Exercise real handlers; never manufacture activity rows.
function Install-ProductionLifecycleSafetyProbe {
    $form=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('frmProduction').CodeModule
    $form.InsertLines(($form.CountOfDeclarationLines+1),@'
Private mLifecycleReenterArmed As Boolean, mLifecycleReenterReached As Boolean
Private mLifecycleReenterDesigner As String, mLifecycleReenterAction As String, mLifecycleReenterResult As String
'@)
    $form.AddFromString(@'
Public Sub LifecycleArmReentryForTest(ByVal designer As String, ByVal action As String)
    mLifecycleReenterDesigner = designer: mLifecycleReenterAction = action
    mLifecycleReenterArmed = True: mLifecycleReenterReached = False: mLifecycleReenterResult = ""
End Sub
Private Sub LifecycleTryReentryForTest()
    If Not mLifecycleReenterArmed Then Exit Sub
    mLifecycleReenterArmed = False: mLifecycleReenterReached = True
    mLifecycleReenterResult = LifecycleActForTest(mLifecycleReenterDesigner, mLifecycleReenterAction)
End Sub
Public Function LifecycleReentryEvidenceForTest() As String
    LifecycleReentryEvidenceForTest = CStr(mLifecycleReenterReached) & "|" & mLifecycleReenterResult
End Function
'@)
    $text=$form.Lines(1,$form.CountOfLines)
    $procedures=if($text.Contains('Private Function SubmitDesignerAction(')){@('SubmitDesignerAction')}else{@('SubmitProcessAction','SubmitRecipeAction')}
    foreach($procedure in $procedures){
        # Called after the real owner returns and before form refresh/completion.
        # The nested action disarms itself before calling the very same handler.
        InsertFault $form $procedure 'RefreshReusableDesignLists' '    LifecycleTryReentryForTest'
    }
    $packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Sub LifecycleArmReentry(ByVal designer As String, ByVal action As String)
    mForm.LifecycleArmReentryForTest designer, action
End Sub
Public Function LifecycleReentryEvidence() As String
    LifecycleReentryEvidence = mForm.LifecycleReentryEvidenceForTest()
End Function
'@)
    $packages['invSys.Core.xlam'].VBProject.VBComponents.Item('TestShippingCatalog').CodeModule.AddFromString(@'
Private mLifecycleOriginalPolicy As String
Public Function LifecycleTrackingOffForTest(ByVal restore As Boolean) As Boolean
    Dim context As String, version As Long, request As String, report As String
    context = modActivity.CaptureContext()
    If Not modTrackingPolicySettings.ReadEditor(context, version, request, report) Then Exit Function
    If restore Then
        If mLifecycleOriginalPolicy = "" Then Exit Function
        request = mLifecycleOriginalPolicy
    Else
        mLifecycleOriginalPolicy = request
        request = modTrainingJson.EncodeObject(modTrackingPolicyModel.Defaults(False))
    End If
    LifecycleTrackingOffForTest = modTrackingPolicySettings.SavePolicy(context, version, request, report)
    If restore And LifecycleTrackingOffForTest Then mLifecycleOriginalPolicy = ""
End Function
'@)
}

function Test-ProductionLifecycleSafety($Fixture,$Book,$Sheet,[string]$Canary) {
    $actions=@(@('Process','Save','DRAFT'),@('Process','Release','RELEASED'),@('Recipe','Save','DRAFT'),@('Recipe','Release','RELEASED'),@('Recipe','Obsolete','OBSOLETE'),@('Process','Obsolete','OBSOLETE'))
    foreach($mode in @('StoreUnavailable','TrackingOff','Reentry')){
        $blocked=$false;$moved=$false;$disabled=$false
        $parent=Join-Path $Fixture.Root 'Training\Activity';$leaf=Join-Path $parent $Fixture.Warehouse
        $held=$leaf+'-lifecycle-held'
        $root=[IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\'
        foreach($path in @($leaf,$held)){
            if(-not [IO.Path]::GetFullPath($path).StartsWith($root,[StringComparison]::OrdinalIgnoreCase)){throw 'Lifecycle activity fixture escaped its disposable root.'}
        }
        if(Test-Path -LiteralPath $held){throw 'Lifecycle held fixture path already exists.'}
        try {
            [void](Probe 'CloseDesigner')
            if($mode -ceq 'TrackingOff'){
                SelectTarget $Fixture
                if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.LifecycleTrackingOffForTest' @($false))){throw 'Explicit tracking-off fixture unavailable; not RED.'}
                $disabled=$true
            }
            SelectTarget $Fixture 'config-producer';[void](Probe 'OpenDesigner' @($Book.Name))
            [void](Probe 'RefreshDesigners')
            $ready=[string](Probe 'LifecyclePrepare' @($Canary))
            if($ready -cnotmatch '^READY\|([^|]+)\|([^|]+)$'){throw 'Lifecycle safety Process prerequisite unavailable; not RED.'}
            $processId=$Matches[1];$processVersion=$Matches[2]
            foreach($action in $actions){
                $designer=$action[0];$verb=$action[1];$id='PRODUCTION_'+$designer.ToUpperInvariant()+'_'+$verb.ToUpperInvariant()
                $case='ProductionLifecycle.Safety.'+$mode+'.'+$designer+'.'+$verb
                if($designer -ceq 'Recipe' -and $verb -ceq 'Save'){
                    if(-not [bool](Probe 'ReleasedRecipe' @($processId,$processVersion,$Canary))){throw 'Lifecycle safety Recipe prerequisite unavailable; not RED.'}
                }
                $before=@(Get-Slice4beActivityFiles $Fixture);$sourceBefore=@()
                if($mode -cne 'Reentry'){$sourceBefore=@(SourceRows)}
                $pins=@{};foreach($file in $before){$pins[$file]=(Get-FileHash -LiteralPath $file).Hash}
                if($mode -ceq 'StoreUnavailable'){
                    New-Item -ItemType Directory -Path $parent -Force|Out-Null
                    if(Test-Path -LiteralPath $leaf){Move-Item -LiteralPath $leaf -Destination $held;$moved=$true}
                    [IO.File]::WriteAllText($leaf,'blocked lifecycle fixture path');$blocked=$true
                }
                if($mode -ceq 'Reentry'){
                    $queueRoot=Join-Path $runRoot ('lifecycle-reentry-'+[guid]::NewGuid().ToString('N'))
                    [void](Run 'invSys.Core.xlam' 'modRoleEventWriter.LifecycleWriterFixtureForTest' @($queueRoot,'Observe',$Fixture.Warehouse))
                    [void](Probe 'LifecycleArmReentry' @($designer,$verb))
                }
                $returned=[string](Probe 'LifecycleAct' @($designer,$verb))
                $notice=[string](Probe 'LifecycleNotice')
                if($mode -ceq 'StoreUnavailable'){
                    Remove-Item -LiteralPath $leaf -Force;$blocked=$false
                    if($moved){Move-Item -LiteralPath $held -Destination $leaf;$moved=$false}
                }
                Check ($case+'.ActualHandlerCompletes') ($returned -ceq 'RETURNED' -and [string](Probe 'LifecycleStatus' @($designer)) -ceq $action[2])
                $oneWrite=$true
                if($mode -ceq 'Reentry'){
                    $writer=([string](Run 'invSys.Core.xlam' 'modRoleEventWriter.LifecycleWriterEvidenceForTest')).Split('|')
                    $oneWrite=$writer.Count -eq 5 -and $writer[1] -ceq '1' -and $writer[2] -ceq 'True'
                    # Duplicate writes are a product failure in their own right.
                    # Do not require a still-readable publication after that failure.
                    $source=@();if($oneWrite){$source=@(SourceRows|Where-Object EventID -CEQ $writer[0])}
                } else {$source=@(SourceRows|Where-Object {$_.EventID -cnotin @($sourceBefore|ForEach-Object EventID)})}
                Check ($case+'.ExactlyOneOwningSubmission') ($oneWrite -and $source.Count -eq 1 -and $source[0].EventType -ceq ($designer.ToUpperInvariant()+'_'+$verb.ToUpperInvariant()) -and $source[0].WarehouseId -ceq $Fixture.Warehouse -and $source[0].AppliedAtUTC -ne '')
                $raw=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {[IO.File]::ReadAllText($_)})
                $safe=$true
                foreach($text in $raw){foreach($value in @($Canary,$Fixture.Secret,(CredentialHash $Fixture.Secret),$Fixture.Root,'mBtn','PayloadJson')){if($text.Contains($value)){$safe=$false}}}
                if($mode -ceq 'Reentry'){
                    Check ($case+'.SecondActualHandlerReached') ([string](Probe 'LifecycleReentryEvidence') -ceq 'True|RETURNED')
                    $rows=@($raw|ForEach-Object {$_|ConvertFrom-Json});$attempt=@($rows|Where-Object OutcomeCode -CEQ 'REQUESTED');$outcome=@($rows|Where-Object OutcomeCode -CEQ 'CONFIRMED')
                    $pair=$rows.Count -eq 2 -and $attempt.Count -eq 1 -and $outcome.Count -eq 1
                    $exact=$pair
                    if($pair){
                        $refs=@($outcome[0].SourceEventRefs)
                        $exact=$source.Count -eq 1 -and $refs.Count -eq 1 -and $refs[0].EventId -ceq $source[0].EventID -and $refs[0].SourceKind -ceq 'Designs' -and $refs[0].SubmissionState -ceq 'Submitted' -and $outcome[0].ControlId -ceq $id -and $outcome[0].DataEffect -ceq 'Unknown' -and $attempt[0].ActivityId -ceq $outcome[0].ActivityId -and $attempt[0].RecordId -cne $outcome[0].RecordId
                    }
                    Check ($case+'.OnlyOneCorrelatedOwnerPair') $exact
                } else {
                    Check ($case+'.NoFabricatedActivity') ($raw.Count -eq 0)
                    if($mode -ceq 'StoreUnavailable'){Check ($case+'.TrackingFailureVisible') $notice.Contains('Tracking unavailable')}
                }
                $same=$true;foreach($file in $pins.Keys){$same=$same -and (Get-FileHash -LiteralPath $file).Hash -ceq $pins[$file]}
                Check ($case+'.PriorActivityBytesPreserved') $same
                Check ($case+'.NoEnteredDataOrSecrets') $safe
                Check ($case+'.UnknownOperatorColumnPreserved') ($Sheet.Cells.Item(1,1).Value2 -ceq 'Operator Extra' -and $Sheet.Cells.Item(2,1).Value2 -ceq $Canary)
            }
        } finally {
            if($mode -ceq 'Reentry'){[void](Run 'invSys.Core.xlam' 'modRoleEventWriter.LifecycleWriterFixtureForTest' @('','',''))}
            if($blocked){Remove-Item -LiteralPath $leaf -Force}
            if($moved){Move-Item -LiteralPath $held -Destination $leaf}
            if($disabled){
                SelectTarget $Fixture
                if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.LifecycleTrackingOffForTest' @($true))){throw 'Original fixture tracking policy did not restore.'}
            }
        }
    }
}

function Test-ProductionLifecycleClosedWorkbook($Fixture,$Other,[string]$Canary,$WindowOwner) {
    $book=$null
    try {
        [void](Probe 'CloseDesigner');SelectTarget $Fixture 'config-producer'
        $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
        $sheet.Cells.Item(1,1).Value2='Operator Extra';$sheet.Cells.Item(2,1).Value2=$Canary
        $path=Join-Path $runRoot 'lifecycle-closed-operator.xlsb';$book.SaveAs($path,50)
        $book.Close($false);$book=$null;$sheet=$null
        $pin=(Get-FileHash -LiteralPath $path).Hash
        $book=$excel.Workbooks.Open($path)
        # Keep the MSForms host window alive while testing loss of the captured
        # workbook. The surviving decoy must never become its replacement binding.
        $WindowOwner.Activate()
        [void](Probe 'OpenDesigner' @($book.Name))
        $states=@{}
        foreach($designer in @('Process','Recipe')){
            [void](Probe 'Stage' @($designer,$Canary));$states[$designer]=[string](Probe 'State' @($designer))
        }
        $before=@(Get-Slice4beActivityFiles $Fixture);$otherBefore=@(Get-Slice4beActivityFiles $Other)
        $sourceBefore=@(SourceRows|ForEach-Object EventID)
        $book.Close($false);$book=$null;$sheet=$null
        # Draft setup precedes closure. There is no attempt to stage a stale form.
        foreach($designer in @('Process','Recipe')){foreach($verb in @('Save','Release','Obsolete')){
            $returned=[string](Probe 'LifecycleAct' @($designer,$verb))
            Check ('ProductionLifecycle.Guard.ClosedWorkbook.'+$designer+'.'+$verb) ($returned -ceq 'RETURNED' -and $states[$designer] -ceq [string](Probe 'State' @($designer)) -and [string](Probe 'LifecycleNotice') -match 'Reopen|reopen')
        }}
        Check 'ProductionLifecycle.Guard.ClosedWorkbook.NoRedirectedActivity' (@(Get-Slice4beActivityFiles $Fixture).Count -eq $before.Count -and @(Get-Slice4beActivityFiles $Other).Count -eq $otherBefore.Count)
        $after=@(SourceRows|ForEach-Object EventID)
        Check 'ProductionLifecycle.Guard.ClosedWorkbook.NoOwningSubmission' (@(Compare-Object $sourceBefore $after -CaseSensitive).Count -eq 0)
        Check 'ProductionLifecycle.Guard.ClosedWorkbook.SavedUnknownColumnBytesPreserved' ((Get-FileHash -LiteralPath $path).Hash -ceq $pin)
    } finally {
        [void](Probe 'CloseDesigner')
        if($null -ne $book){$book.Close($false)}
        SelectTarget $Fixture
    }
}
