# D18 visible evidence from the actual saved result after the packaged Evaluate action.
# Capture only the owned disposable library; no substitute rendering or source edits.
function Capture-EvaluationLibrary([string]$Stage,[string]$View,[string]$Title) {
    try {
        Initialize-SettingsCapture
        $handle=[InvSysSettingsCapture]::OwnedVisibleForm($Title,[IntPtr]$excel.Hwnd).ToInt64()
        if($handle -eq 0){throw 'Requested form window unavailable.'}
        CaptureOwnedFormEvidence $Title ('evaluation-'+$Stage.ToLowerInvariant()+'-'+$View.ToLowerInvariant()+'.png') $handle
        Check ('EvaluationVisual.'+$Stage+'.Capture.'+$View) $true
    } catch {
        $message=$_.Exception.GetBaseException().Message
        if($message -cnotin @('Requested form is not in the foreground.','Requested form window unavailable.')){throw}
        Check ('EvaluationVisual.'+$Stage+'.Capture.'+$View) $false
        Write-Output ('Diagnostic pane capture unavailable: '+$Stage+'/'+$View+'; '+$message)
    }
}

function Test-EvaluationVisualEvidence([string]$Stage) {
    $prefix='EvaluationVisual.'+$Stage+'.'
    $selected=ExpectationControl 'lstActionPaths' 'Value' '' 'frmActionPaths'
    $resultId=Selected-EvaluationId
    $text=ExpectationControl 'txtPathEvaluation' 'Value' '' 'frmActionPaths'
    $status=ExpectationControl 'lblEvaluationStatus' 'Caption' '' 'frmActionPaths'
    $expected=if($Stage -ceq 'Applied'){'Conclusion observed'}else{'Awaiting published result'}
    if($selected -cne $pathId -or $resultId -notmatch '^[0-9a-fA-F-]{36}$' -or
        -not $text.Contains($resultId) -or -not $status.StartsWith($expected,[StringComparison]::Ordinal)){
        throw 'Actual saved-result fixture is unavailable for visible inspection.'
    }
    $title=ExpectationControl '' 'FormCaption' '' 'frmActionPaths'
    Check ('RecordingNotice.'+$Stage+'.AfterEvaluationCaptureOnly') (
        (ExpectationControl 'lblPathStatus' 'Caption' '' 'frmActionPaths') -ceq 'Stopped. Capture frozen.')
    Check ($prefix+'ApprovedTitle') ($title -ceq 'Action Paths')
    if($title -cmatch '^UserForm[0-9]+$'){Write-Output ('Observed generated library title: '+$title)}
    $pins=@{}
    foreach($file in Get-ChildItem -LiteralPath $journalRoot -Recurse -File){$pins[$file.FullName]=(Get-FileHash -LiteralPath $file.FullName).Hash}
    $activity=ActivityPins
    $publicationHash=(Get-FileHash -LiteralPath $eventsPath).Hash
    $publishBefore=[long](Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest')
    $authorityBefore=[long](Run 'invSys.Operations.xlam' 'modTS_Shipments.PublishedReadAuthorityCallsForTest')
    $visibility=$excel.Visible
    if($null -eq $visibility){throw 'Excel visibility is unavailable; not product RED.'}
    try {
        $excel.Visible=$true
        foreach($size in @(@('Minimum',720,520),@('Default',820,640),@('Larger',1000,760),@('Restored',820,640))){
            $layout=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingExpectationLayoutForTest' @('frmActionPaths',[single]$size[1],[single]$size[2]))
            Check ($prefix+'Layout.'+$size[0]) ($layout -ceq 'FITS')
            if($layout -cne 'FITS'){Write-Output ('Diagnostic geometry '+$Stage+'/'+$size[0]+': '+$layout)}
            if((ExpectationControl 'txtPathEvaluation' 'ViewportTop' '' 'frmActionPaths') -cne 'DELIVERED'){throw 'Actual diagnostic text viewport is unavailable.'}
            Capture-EvaluationLibrary $Stage $size[0] $title
        }
        if((ExpectationControl 'txtPathEvaluation' 'ViewportBottom' '' 'frmActionPaths') -cne 'DELIVERED'){throw 'Actual diagnostic source viewport is unavailable.'}
        Capture-EvaluationLibrary $Stage 'Sources' $title
    } finally {
        $excel.Visible=[bool]$visibility
    }
    Check ($prefix+'SelectionAndSavedTextPreserved') (
        (ExpectationControl 'lstActionPaths' 'Value' '' 'frmActionPaths') -ceq $selected -and
        (Selected-EvaluationId) -ceq $resultId -and
        (ExpectationControl 'txtPathEvaluation' 'Value' '' 'frmActionPaths') -ceq $text -and
        (ExpectationControl 'lblEvaluationStatus' 'Caption' '' 'frmActionPaths') -ceq $status)
    $preserved=@(Get-ChildItem -LiteralPath $journalRoot -Recurse -File).Count -eq $pins.Count
    foreach($path in $pins.Keys){$preserved=$preserved -and (Get-FileHash -LiteralPath $path).Hash -ceq $pins[$path]}
    Check ($prefix+'TrainingAndActivityBytesPreserved') ($preserved -and (PinsRetained $activity) -and (ActivityPins).Count -eq $activity.Count)
    Check ($prefix+'NoPublicationOrAuthoritySideEffects') (
        (Get-FileHash -LiteralPath $eventsPath).Hash -ceq $publicationHash -and
        [long](Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest') -eq $publishBefore -and
        [long](Run 'invSys.Operations.xlam' 'modTS_Shipments.PublishedReadAuthorityCallsForTest') -eq $authorityBefore)
}
