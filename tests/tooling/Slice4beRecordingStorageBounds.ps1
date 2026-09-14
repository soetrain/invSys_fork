# Supplemental serialized-store checks after actual operator-handler recording.
# The large metadata is synthetic test data, never a fabricated business outcome.
function Test-Slice4beRecordingStorageBounds($Fixture) {
    SelectTarget $Fixture 'config-admin'
    $core=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modActionRecording').CodeModule
    $core.AddFromString(@'
Public Function RecordingBoundaryForTest(ByVal request As String) As String
    Dim model As Object, target As WarehouseTarget, hash As String, notice As String
    Set target = modNasConnection.GetCurrentTarget()
    Set model = modTrainingJson.DecodeObject(request)
    If model Is Nothing Then RecordingBoundaryForTest = "INVALID": Exit Function
    If Not modRecordingModel.Validate(target, model) Then RecordingBoundaryForTest = "INVALID": Exit Function
    If modRecordingJournal.Append(target, model, hash, notice) Then
        RecordingBoundaryForTest = "VALID|SAVED"
    Else
        RecordingBoundaryForTest = "VALID|REJECTED|" & notice
    End If
End Function
Public Function RecordingBoundaryReadableForTest(ByVal id As String) As Boolean
    Dim model As Object
    Set model = modRecordingJournal.ReadEntry(modNasConnection.GetCurrentTarget(), id, 1)
    RecordingBoundaryReadableForTest = Not model Is Nothing
End Function
'@)
    $root=Join-Path $Fixture.Root ('Training/ActionPaths/'+$Fixture.Warehouse)
    $starts=@(Get-ChildItem -LiteralPath $root -Filter '*.1.json' -File)
    if(-not $starts.Count){throw 'Storage-bound checks require the real recording GREEN fixture.'}
    $source=[IO.File]::ReadAllText($starts[0].FullName)
    foreach($extra in @(0,1)) {
        $model=$source|ConvertFrom-Json
        $model.PSObject.Properties.Remove('ContentSha256')
        $model.RecordId=[guid]::NewGuid().ToString()
        $model.ActionPathId=[guid]::NewGuid().ToString()
        $model.SequenceId=[guid]::NewGuid().ToString()
        $model.BuildIdentity=''
        $base=$model|ConvertTo-Json -Depth 16 -Compress
        $model.BuildIdentity='B'*(1048576-83+$extra-$base.Length)
        $body=$model|ConvertTo-Json -Depth 16 -Compress
        if($body.Length+83 -ne 1048576+$extra){throw 'Serialized byte-bound fixture length is wrong.'}
        $path=Join-Path $root ($model.ActionPathId+'.1.json')
        $result=[string](Run 'invSys.Core.xlam' 'modActionRecording.RecordingBoundaryForTest' @($body))
        if($result -ceq 'INVALID'){throw 'Byte-bound fixture failed schema validation; not a size test.'}
        if($extra -eq 0) {
            Check 'Recording.Storage.Exact1MiBRecordAccepted' ($result -ceq 'VALID|SAVED' -and (Test-Path -LiteralPath $path) -and (Get-Item -LiteralPath $path).Length -eq 1048576)
            Check 'Recording.Storage.Exact1MiBRecordReadsWithIntegrity' ([bool](Run 'invSys.Core.xlam' 'modActionRecording.RecordingBoundaryReadableForTest' @($model.ActionPathId)))
            $pin=(Get-FileHash -LiteralPath $path).Hash
            $retry=[string](Run 'invSys.Core.xlam' 'modActionRecording.RecordingBoundaryForTest' @($body))
            Check 'Recording.Storage.IdenticalAppendPreservesBytes' ($retry -ceq 'VALID|SAVED' -and (Get-FileHash -LiteralPath $path).Hash -ceq $pin)
        } else {
            Check 'Recording.Storage.OneByteOversizeRejectedExplicitly' ($result -match '^VALID\|REJECTED\|.*exceeds 1 MiB' -and -not (Test-Path -LiteralPath $path))
        }
    }
    Check 'Recording.Storage.NoPartialFilesPublished' (@(Get-ChildItem -LiteralPath $root -Filter '*.pending' -File).Count -eq 0)
}
