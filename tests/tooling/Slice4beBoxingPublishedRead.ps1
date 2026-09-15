# D18: actual Make/Unbox observations and independently verified Inventory lines
# flow through ordinary Admin publication and the operator's Viewer controls.
# Install before the existing five instrumented compiles; never save probe code.
function Install-Slice4beBoxingPublishedReadProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $project.VBComponents.Item('frmEventDetail').CodeModule.AddFromString(@'
Public Function BoxingDetailForTest(ByVal action As String, ByVal caption As String, ByVal expected As String) As String
    Dim index As Long, found As Boolean
    Select Case action
        Case "Lines": BoxingDetailForTest = CStr(mLines.ListCount)
        Case "SelectedLabel"
            If mLines.ListIndex >= 0 Then BoxingDetailForTest = CStr(mLines.List(mLines.ListIndex, 0))
        Case "Prompt": BoxingDetailForTest = CStr(Me.Controls("lblDetailLines").Caption)
        Case "Select"
            index = CLng(expected)
            If index < 0 Or index >= mLines.ListCount Then Exit Function
            mLines.ListIndex = index: mLines_Click
            BoxingDetailForTest = "SELECTED"
        Case "Field", "Absent"
            For index = 0 To mFields.ListCount - 1
                If CStr(mFields.List(index, 0)) = caption Then
                    found = True
                    If action = "Field" Then BoxingDetailForTest = CStr(CStr(mFields.List(index, 1)) = expected)
                    Exit For
                End If
            Next index
            If action = "Absent" Then BoxingDetailForTest = CStr(Not found)
        Case "ReadOnly": BoxingDetailForTest = CStr(mFields.Locked)
        Case "Close": mClose_Click: BoxingDetailForTest = "CLOSED"
    End Select
End Function
'@)
    $project.VBComponents.Item('modInventoryViewer').CodeModule.AddFromString(@'
Public Function BoxingDetailForTest(ByVal action As String, Optional ByVal caption As String = "", Optional ByVal expected As String = "") As String
    Dim form As Object
    For Each form In VBA.UserForms
        If TypeName(form) = "frmEventDetail" Then
            BoxingDetailForTest = form.BoxingDetailForTest(action, caption, expected)
            Exit Function
        End If
    Next form
End Function
'@)
}

function Test-Slice4beBoxingPublishedRead($Fixture,$Other,$Evidence) {
    if($null -eq $Evidence -or @($Evidence.Rows).Count -ne 4 -or @($Evidence.Observations).Count -ne 8){
        throw 'Verified actual Boxing action evidence is unavailable; not behavioral RED.'
    }
    function BoxingRead([string]$Action,[string]$Expected='') {
        $value=Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadActionForTest' @($Action,$Expected)
        return ($value -is [bool] -and $value)
    }
    function BoxingDetail([string]$Action,[string]$Caption='',[string]$Expected='') {
        [string](Run 'invSys.Operations.xlam' 'modInventoryViewer.BoxingDetailForTest' @($Action,$Caption,$Expected))
    }
    CloseRecordingViewer
    SelectTarget $Fixture 'config-admin'
    $published=Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest'
    Check 'Boxing.PublishedRead.ActualAdminPublication' ($published -is [bool] -and $published)
    if($published -isnot [bool] -or -not $published){throw 'Boxing publication command did not complete.'}
    $path=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json')
    $model=[IO.File]::ReadAllText($path)|ConvertFrom-Json
    $valid=Run 'invSys.Core.xlam' 'modInventoryViewerData.PublishedReadFixtureValidForTest' @($path,$Fixture.Warehouse)
    Check 'Boxing.PublishedRead.ValidPersistedArtifact' ($valid -is [bool] -and $valid)
    if($valid -isnot [bool] -or -not $valid){throw 'Boxing published schema/integrity unavailable.'}
    $pins=@{}
    foreach($file in Get-ChildItem -LiteralPath $Fixture.Root -Recurse -File){$pins[$file.FullName]=Get-ShippingActivityHash $file.FullName}
    $otherHash=Get-ShippingActivityHash $Other.FullName
    $authority=[long](Run 'invSys.Operations.xlam' 'modTS_Shipments.PublishedReadAuthorityCallsForTest')
    $publication=[long](Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest')
    SelectTarget $Fixture 'config-reader'
    try {
        $Other.Activate()
        OpenRecordingViewer
        foreach($group in @($Evidence.Rows|Group-Object EventID)){
            $rows=@($group.Group);$id=[string]$rows[0].EventID
            $kind=if($rows[0].EventType -ceq 'BOX_BUILD'){'Make'}else{'Unbox'}
            $label='Boxing.PublishedRead.'+$kind
            $stored=@($model.Groups|Where-Object {$_.Source -ceq 'Inventory' -and $_.SourceId -ceq $id})
            $exact=$stored.Count -eq 1
            if($exact){
                $lines=@($stored[0].Lines);$exact=$lines.Count -eq $rows.Count
                foreach($row in $rows){$exact=$exact -and @($lines|Where-Object {
                    $_.System_Key -ceq $row.System_Key -and $_.EventType -ceq $row.EventType -and
                    [double]$_.QtyDelta -eq [double]$row.QtyDelta -and $_.WarehouseId -ceq $Fixture.Warehouse
                }).Count -eq 1}
            }
            Check ($label+'.AllExactOwnerLinesPublished') $exact
            [void](BoxingRead 'Search' $id)
            Check ($label+'.ActualEventSelection') (BoxingRead 'SelectSource' $id)
            Check ($label+'.EveryContributingLineVisible') ((BoxingDetail 'Lines') -ceq '2')
            Check ($label+'.ReadOnlyFields') ((BoxingDetail 'ReadOnly') -ceq 'True')
            Check ($label+'.SourceNeutralPrompt') ((BoxingDetail 'Prompt') -ceq 'Contributing lines - select a line to inspect its fields')
            if($stored.Count -eq 1){
                $index=0
                foreach($line in $stored[0].Lines){
                    $lineLabel=$label+'.Line'+($index+1)
                    Check ($lineLabel+'.ActualLineSelection') ((BoxingDetail 'Select' '' ([string]$index)) -ceq 'SELECTED')
                    Check ($lineLabel+'.ExactInventoryKeyLabel') ((BoxingDetail 'SelectedLabel') -ceq [string]$line.System_Key)
                    foreach($field in @(@('Source event / activity ID',$id),@('Inventory identity (System_Key)',[string]$line.System_Key),
                        @('Source classification','Business event'),@('Event code',[string]$line.EventType),@('Quantity',[string]$line.QtyDelta))){
                        Check ($lineLabel+'.'+$field[0]) ((BoxingDetail 'Field' $field[0] $field[1]) -ceq 'True')
                    }
                    $index++
                }
            }
            if($CaptureEvidence){Capture-BoxingFormEvidence 'Event Detail' ('boxing-published-'+$kind.ToLowerInvariant()+'.png') ($label+'.VisibleCapture')}
            [void](BoxingDetail 'Close')
        }
        foreach($group in @($Evidence.Observations|Group-Object ActivityId)){
            $records=@($group.Group);$id=[string]$records[0].ActivityId
            $label='Boxing.PublishedRead.Action'+$records[0].Ordinal
            $stored=@($model.Groups|Where-Object {$_.Source -ceq 'Activity' -and $_.SourceId -ceq $id})
            $exact=$stored.Count -eq 1
            if($exact){
                $lines=@($stored[0].Lines);$exact=$lines.Count -eq 2
                foreach($record in $records){
                    $matches=@($lines|Where-Object RecordId -CEQ $record.RecordId)
                    $exact=$exact -and $matches.Count -eq 1
                    if($matches.Count -eq 1){
                        foreach($field in @('ActivityId','ControlId','EventCode','OutcomeCode','OwnerId','DataEffect','SequenceId','Ordinal')){
                            $exact=$exact -and $matches[0].$field -ceq $record.$field
                        }
                        $exact=$exact -and ($matches[0].SourceEventRefs|ConvertTo-Json -Depth 8 -Compress) -ceq ($record.SourceEventRefs|ConvertTo-Json -Depth 8 -Compress)
                    }
                }
            }
            Check ($label+'.ExactObservationPairAndAllReferencesPublished') $exact
            [void](BoxingRead 'Search' $id)
            Check ($label+'.ActualActivitySelection') (BoxingRead 'SelectSource' $id)
            Check ($label+'.BothObservationLinesVisible') ((BoxingDetail 'Lines') -ceq '2')
            Check ($label+'.SourceNeutralPrompt') ((BoxingDetail 'Prompt') -ceq 'Contributing lines - select a line to inspect its fields')
            if($stored.Count -eq 1){
                $index=0
                foreach($line in $stored[0].Lines){
                    Check ($label+'.Line'+$index+'.ActualSelection') ((BoxingDetail 'Select' '' ([string]$index)) -ceq 'SELECTED')
                    Check ($label+'.Line'+$index+'.FixedCaptionAndObservedOutcomeLabel') ((BoxingDetail 'SelectedLabel') -ceq (([string]$line.Caption)+' - '+([string]$line.OutcomeCode)))
                    foreach($field in @(@('Source event / activity ID',$id),@('Source classification','User activity'),
                        @('Owning operation','BOXING_WORKFLOW'),@('Event code',[string]$line.EventCode),
                        @('Outcome',[string]$line.OutcomeCode),@('Data effect',[string]$line.DataEffect))){
                        Check ($label+'.Line'+$index+'.'+$field[0]) ((BoxingDetail 'Field' $field[0] $field[1]) -ceq 'True')
                    }
                    $index++
                }
            }
            if($CaptureEvidence){Capture-BoxingFormEvidence 'Event Detail' ('boxing-published-action'+$records[0].Ordinal+'.png') ($label+'.VisibleCapture')}
            [void](BoxingDetail 'Close')
        }
        Check 'Boxing.PublishedRead.NoShippingAuthorityRead' ($authority -eq [long](Run 'invSys.Operations.xlam' 'modTS_Shipments.PublishedReadAuthorityCallsForTest'))
        Check 'Boxing.PublishedRead.NoImplicitPublication' ($publication -eq [long](Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest'))
    } finally {CloseRecordingViewer}
    $unchanged=$pins.Count -eq @(Get-ChildItem -LiteralPath $Fixture.Root -Recurse -File).Count
    foreach($file in $pins.Keys){$unchanged=$unchanged -and $pins[$file] -ceq (Get-ShippingActivityHash $file)}
    Check 'Boxing.PublishedRead.AllRuntimeAndTrainingBytesPreserved' $unchanged
    Check 'Boxing.PublishedRead.UnrelatedWorkbookPreserved' ($otherHash -ceq (Get-ShippingActivityHash $Other.FullName) -and $Other.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'shipping unrelated sentinel')
}
