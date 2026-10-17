# Faults are injected only into disposable packages/owner-created fixtures.
# Results retain booleans and fixed case names, never source values or paths.
function Install-AggregationBoundaryProbe {
    $core=$packages['invSys.Core.xlam'].VBProject
    $probe=$core.VBComponents.Add(1);$probe.Name='TestAggregationBoundary'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mFault As String, mPoint As String, mInjected As Boolean
Private mSource As String, mAlternate As String, mFeed As String
Public Sub Arm(ByVal fault As String, ByVal point As String, ByVal source As String, ByVal alternate As String, ByVal feed As String)
    mFault = fault: mPoint = point: mSource = source: mAlternate = alternate: mFeed = feed: mInjected = False
End Sub
Public Function Injected() As Boolean
    Injected = mInjected
End Function
Public Function IdentityUnavailable() As Boolean
    IdentityUnavailable = (mFault = "IDENTITY_UNAVAILABLE")
    If IdentityUnavailable Then mInjected = True
End Function
Public Sub AtPoint(ByVal point As String)
    Dim fso As Object, target As WarehouseTarget, report As String
    Dim auth As Workbook, sheet As Worksheet, table As ListObject, row As ListRow
    If mInjected Or mFault = "" Or point <> mPoint Then Exit Sub
    Select Case mFault
        Case "SIGN_OUT": modAuth.SignOut: mInjected = True
        Case "SOURCE_CHANGED"
            Set fso = CreateObject("Scripting.FileSystemObject")
            fso.CopyFile mAlternate, mSource, True
            mInjected = True
        Case "SOURCE_MISSING"
            Set fso = CreateObject("Scripting.FileSystemObject")
            fso.DeleteFile mSource, False
            mInjected = True
        Case "FEED_CHANGED"
            Set target = modNasConnection.GetCurrentTarget()
            If Not modConfig.UpdateConfigValue("PathSharePointRoot", mFeed, report, target.WarehouseId, target.StationId) Then Err.Raise vbObjectError + 501
            mInjected = True
        Case "RIGHTS_REVOKED"
            Set target = modNasConnection.GetCurrentTarget()
            Set auth = Application.Workbooks.Open(target.AuthPath, UpdateLinks:=0, ReadOnly:=False)
            For Each sheet In auth.Worksheets
                For Each table In sheet.ListObjects
                    If table.Name = "tblCapabilities" Then
                        For Each row In table.ListRows
                            If CStr(row.Range.Cells(1, table.ListColumns("UserId").Index).Value2) = modAuth.GetCurrentUserId() And _
                               CStr(row.Range.Cells(1, table.ListColumns("Capability").Index).Value2) = "ADMIN_MAINT" Then _
                                row.Range.Cells(1, table.ListColumns("Status").Index).Value2 = "Disabled"
                        Next row
                    End If
                Next table
            Next sheet
            auth.Save: auth.Close False: mInjected = True
        Case "WRITE_FAILURE": mInjected = True: Err.Raise 1004
        Case "PARTIAL_OUTPUT"
            Set fso = CreateObject("Scripting.FileSystemObject")
            With fso.CreateTextFile(mFeed, True, False)
                .Write "Incomplete disposable aggregation output"
                .Close
            End With
            mInjected = True
        Case "CANDIDATE_STATUS_REMOVED", "CANDIDATE_SOURCE_WRONG"
            Set fso = CreateObject("Scripting.FileSystemObject")
            For Each auth In Application.Workbooks
                If StrComp(fso.GetParentFolderName(auth.FullName), fso.GetParentFolderName(mFeed), vbTextCompare) = 0 And _
                        (auth.FullName = mFeed Or LCase$(auth.Name) Like "invsys.aggregation.*.candidate.xlsb") Then
                    If mFault = "CANDIDATE_STATUS_REMOVED" Then
                        auth.Worksheets("GlobalSnapshotStatus").Delete
                    Else
                        Set table = auth.Worksheets("GlobalInventorySnapshot").ListObjects("tblGlobalInventorySnapshot")
                        table.ListColumns("SourceSnapshot").DataBodyRange.Value2 = "Invalid disposable source provenance"
                    End If
                    auth.Save: mInjected = True: Exit For
                End If
            Next auth
            If Not mInjected Then Err.Raise vbObjectError + 502
    End Select
End Sub
Public Function DirectInteractive(ByVal context As String, ByVal sources As String, ByVal outputPath As String) As Boolean
    Dim report As String
    ' Baseline comparison: this is the exact primitive Core writer dispatched by
    ' the current interactive handler. It has no captured-context argument/guard.
    ' Scheduled/fixture authorization is a separate preserved comparison.
    DirectInteractive = modHqAggregator.GenerateGlobalSnapshotFromFiles(sources, outputPath, report)
End Function
Public Function CaptureDirect(ByVal bookName As String, ByVal warehouse As String, ByVal configPath As String, ByVal snapshot As String) As String
    CaptureDirect = modAuthUserCommands.CaptureContext()
End Function
'@)
    $owner=$packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('modAdminConsole').CodeModule
    $first=$owner.ProcStartLine('RunHQAggregationFromSourceSet',0);$count=$owner.ProcCountLines('RunHQAggregationFromSourceSet',0)
    $body=$owner.Lines($first,$count)
    $marker='    TestAggregation.Phase "OutputConfig"'
    $modern=$body.Contains('modAggregationCommands.Aggregate')
    if($modern){
        $files=$core.VBComponents.Item('modAggregationFiles').CodeModule
        $first=$files.ProcStartLine('Identity',0);$count=$files.ProcCountLines('Identity',0)
        $body=$files.Lines($first,$count)
        $marker='    On Error GoTo Finished'
        if(-not $body.Contains($marker)){throw 'File identity fault boundary unavailable.'}
        $body=$body.Replace($marker,$marker+"`r`n"+'    If TestAggregationBoundary.IdentityUnavailable() Then Exit Function')
        $files.DeleteLines($first,$count);$files.InsertLines($first,$body)
        $adapter=$probe.CodeModule
        $first=$adapter.ProcStartLine('DirectInteractive',0);$count=$adapter.ProcCountLines('DirectInteractive',0)
        $adapter.DeleteLines($first,$count)
        $adapter.InsertLines($first,@'
Public Function DirectInteractive(ByVal context As String, ByVal sources As String, ByVal outputPath As String) As Boolean
    Dim report As String
    DirectInteractive = modAggregationCommands.Aggregate(context, sources, report)
End Function
'@)
        $first=$adapter.ProcStartLine('CaptureDirect',0);$count=$adapter.ProcCountLines('CaptureDirect',0)
        $adapter.DeleteLines($first,$count)
        $adapter.InsertLines($first,@'
Public Function CaptureDirect(ByVal bookName As String, ByVal warehouse As String, ByVal configPath As String, ByVal snapshot As String) As String
    Dim context As String
    context = modAggregationCommands.CaptureForm(bookName)
    If modAggregationCommands.RegisterSource(context, warehouse, configPath, snapshot) <> "" Then CaptureDirect = context
End Function
Public Function ResolveKnown(ByVal context As String, ByVal sources As String) As Boolean
    Dim report As String
    ResolveKnown = modAggregationCommands.ResolveRetry(context, sources, True, report)
    modAggregationCommands.ReleaseForm context
End Function
'@)
        $command=$core.VBComponents.Item('modAggregationCommands').CodeModule
        $first=$command.ProcStartLine('Aggregate',0);$count=$command.ProcCountLines('Aggregate',0)
        $body=$command.Lines($first,$count)
        $point='    mWriting = True: mContext = context'
        if(-not $body.Contains($point)){throw 'Core inspection-return boundary unavailable.'}
        $body=$body.Replace($point,$point+"`r`n"+'    TestAggregationBoundary.AtPoint "AFTER_INSPECTION"')
        $promotion=[regex]::Matches($body,'(?im)^\s*promoted\s*=\s*modAggregationFiles\.Promote[^\r\n]*')
        if($promotion.Count -ne 1){throw 'Core promotion-return boundary unavailable.'}
        $body=$body.Replace($promotion[0].Value,$promotion[0].Value+"`r`n"+'    TestAggregationBoundary.AtPoint "AFTER_PROMOTION"')
        $command.DeleteLines($first,$count);$command.InsertLines($first,$body)
    }else{
        if(-not $body.Contains($marker)){throw 'Post-inspection fault boundary unavailable.'}
        $body=$body.Replace($marker,'    TestAggregationBoundary.AtPoint "AFTER_INSPECTION"'+"`r`n"+$marker)
        $owner.DeleteLines($first,$count);$owner.InsertLines($first,$body)
    }
    $writer=$core.VBComponents.Item('modHqAggregator').CodeModule
    $first=$writer.ProcStartLine('WriteGlobalSnapshotWorkbook',0);$count=$writer.ProcCountLines('WriteGlobalSnapshotWorkbook',0)
    $body=$writer.Lines($first,$count)
    $marker='    TestAggregationRoots.Stage "SaveAs"'
    if(-not $body.Contains($marker)){throw 'Writer fault boundary unavailable.'}
    if($modern){
        $guard=[regex]::Matches($body,'(?im)^\s*If Not modAggregationCommands\.CheckPhase[^\r\n]*7512[^\r\n]*')
        if($guard.Count -ne 1){throw 'Core pre-save authority boundary unavailable.'}
        $body=$body.Replace($guard[0].Value,'    TestAggregationBoundary.AtPoint "BEFORE_SAVE"'+"`r`n"+$guard[0].Value)
    }else{$body=$body.Replace($marker,'    TestAggregationBoundary.AtPoint "BEFORE_SAVE"'+"`r`n"+$marker)}
    $close='    TestAggregationRoots.Stage "Close"'
    if(-not $body.Contains($close)){throw 'Post-save fault boundary unavailable.'}
    if($modern){
        $guard=[regex]::Matches($body,'(?im)^\s*If Not modAggregationCommands\.CheckPhase[^\r\n]*7513[^\r\n]*')
        if($guard.Count -ne 1){throw 'Core post-save authority boundary unavailable.'}
        $body=$body.Replace($guard[0].Value,'    TestAggregationBoundary.AtPoint "AFTER_SAVE"'+"`r`n"+$guard[0].Value)
    }else{$body=$body.Replace($close,'    TestAggregationBoundary.AtPoint "AFTER_SAVE"'+"`r`n"+$close)}
    $closed=[regex]::Matches($body,'(?im)^\s*wb\.Close[^\r\n]*')
    if($closed.Count -ne 1){throw 'Completed persistence observation boundary unavailable.'}
    if(-not $modern){$body=$body.Replace($closed[0].Value,$closed[0].Value+"`r`n"+'    TestAggregationBoundary.AtPoint "AFTER_PROMOTION"')}
    $writer.DeleteLines($first,$count);$writer.InsertLines($first,$body)
}

function Test-AggregationBoundaries($Fixture,$Other,$HostBook,$OtherBook,[ref]$Observations) {
    function Call([string]$Name,[object[]]$Values=@()){Run 'invSys.Admin.xlam' ('TestAggregation.'+$Name) $Values}
    function Core([string]$Name,[object[]]$Values=@()){Run 'invSys.Core.xlam' ('TestAggregationBoundary.'+$Name) $Values}
    function FileHash([string]$Path){if(-not(Test-Path -LiteralPath $Path)){return ''};$stream=[IO.File]::Open($Path,'Open','Read',([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete));try{(Get-FileHash -InputStream $stream).Hash}finally{$stream.Dispose()}}
    $owned=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    $sourceBytes=[IO.File]::ReadAllBytes($Fixture.Snapshot)
    $configBytes=[IO.File]::ReadAllBytes($Fixture.Config)
    $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $authBytes=[IO.File]::ReadAllBytes($authPath)
    $feed=Split-Path -Parent (Split-Path -Parent $Fixture.Output)
    $changedFeed=Join-Path $runRoot 'changed-aggregation-feed'
    if(-not [IO.Path]::GetFullPath($changedFeed).StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Changed feed escapes disposable scope.'}
    try{
        $inspection=@('SourceChangedAfterInspection','SourceMissingAfterInspection','FeedChangedAfterInspection','SignedOutAfterInspection','RightsRevokedAfterInspection')
        $write=@('FailureBeforeSave','FailureAfterSave','SourceChangedBeforeSave','FeedChangedBeforeSave','SignedOutBeforeSave','MalformedSelectedSource','DuplicateSelectedEntity','UnknownOutputColumn')
        $ownership=@('BorrowedCleanOutput','BorrowedDirtyOutput','ReadOnlyOutput','DirtyConfig','UnreadableConfig','DirtyAuth')
        $cases=if($AdminAggregationWriteBoundariesOnly){$write}else{
            switch($AdminAggregationBoundaryGroup){
                'Inspection'{$inspection}
                'Write'{$write}
                'Ownership'{$ownership}
                default{$inspection+$write+$ownership}
            }
        }
        if($AdminAggregationBoundaryCase -ne 'All'){$cases=@($cases|Where-Object {$_ -ceq $AdminAggregationBoundaryCase})}
        foreach($case in $cases){
            SelectTarget $Fixture
            [void](Core 'Arm' @('','','','',''))
            [void](Run 'invSys.Core.xlam' 'TestAggregationRoots.PathComparison' @($true,$Fixture.Output))
            if(-not [bool](Run 'invSys.Core.xlam' 'TestAggregationRoots.NormalizedControl' @($Fixture.Snapshot,$Fixture.Output))){throw 'Existing output fixture failed; not behavioral RED.'}
            if($case -ceq 'UnknownOutputColumn'){
                $customBook=$excel.Workbooks.Open($Fixture.Output,0,$false)
                try{
                    $custom=Table $customBook 'tblGlobalInventorySnapshot'
                    $column=$custom.ListColumns.Add();$column.Name='Local Annotation'
                    $column.DataBodyRange.Cells.Item(1,1).Value2='Preserve disposable annotation'
                    $customBook.Save()
                }finally{$customBook.Close($false)}
            }
            $before=FileHash $Fixture.Output
            $borrowed=$null;$authorityBook=$null
            $outputAttributes=[IO.File]::GetAttributes($Fixture.Output)
            if($case -cin @('BorrowedCleanOutput','BorrowedDirtyOutput')){
                $borrowed=$excel.Workbooks.Open($Fixture.Output,0,$false)
                if($case -ceq 'BorrowedDirtyOutput'){$borrowed.Worksheets.Item(1).Cells.Item(1,12).Value2='Preserve disposable output edit'}
            }
            if(-not [bool](Call 'OpenForm' @($HostBook.Name)) -or [int](Call 'SelectSource' @($Fixture.Warehouse)) -ne 1){throw 'Actual aggregation entry unavailable.'}
            if($case -ceq 'ReadOnlyOutput'){[IO.File]::SetAttributes($Fixture.Output,($outputAttributes -bor [IO.FileAttributes]::ReadOnly))}
            if($case -cin @('DirtyConfig','DirtyAuth')){
                $authorityPath=if($case -ceq 'DirtyConfig'){$Fixture.Config}else{$authPath}
                $authorityBook=$excel.Workbooks.Open($authorityPath,0,$false)
                $authorityBook.Worksheets.Item(1).Cells.Item(1,30).Value2='Preserve disposable authority edit'
                Check ('Aggregation.Boundary.'+$case+'.DirtyPrecondition') (-not [bool]$authorityBook.Saved)
            }
            if($case -ceq 'UnreadableConfig'){[IO.File]::WriteAllBytes($Fixture.Config,[Text.Encoding]::ASCII.GetBytes('Invalid disposable configuration'))}
            if($case -cin @('MalformedSelectedSource','DuplicateSelectedEntity')){
                $invalidBook=$excel.Workbooks.Open($Fixture.Snapshot,0,$false)
                try{
                    $invalid=Table $invalidBook 'tblInventorySnapshot'
                    if($case -ceq 'MalformedSelectedSource'){$invalid.ListColumns.Item('System_Key').Name='Missing identity'}else{
                        $duplicate=$invalid.ListRows.Add()
                        [void]$invalid.ListRows.Item(1).Range.Copy($duplicate.Range)
                    }
                    $invalidBook.Save()
                }finally{$invalidBook.Close($false)}
            }
            $fault=switch($case){
                'SourceChangedAfterInspection'{'SOURCE_CHANGED'}
                'SourceMissingAfterInspection'{'SOURCE_MISSING'}
                'FeedChangedAfterInspection'{'FEED_CHANGED'}
                'SignedOutAfterInspection'{'SIGN_OUT'}
                'RightsRevokedAfterInspection'{'RIGHTS_REVOKED'}
                'FailureBeforeSave'{'WRITE_FAILURE'}
                'FailureAfterSave'{'WRITE_FAILURE'}
                'SourceChangedBeforeSave'{'SOURCE_CHANGED'}
                'FeedChangedBeforeSave'{'FEED_CHANGED'}
                'SignedOutBeforeSave'{'SIGN_OUT'}
                default{''}
            }
            $point=if($case -ceq 'FailureAfterSave'){'AFTER_SAVE'}elseif($case -cin @('FailureBeforeSave','SourceChangedBeforeSave','FeedChangedBeforeSave','SignedOutBeforeSave')){'BEFORE_SAVE'}else{'AFTER_INSPECTION'}
            [void](Core 'Arm' @($fault,$point,$Fixture.Snapshot,$Other.Snapshot,$changedFeed))
            [void](Call 'Aggregate')
            $success=[bool](Call 'OwnerSuccess');$calls=[int](Run 'invSys.Core.xlam' 'TestAggregationRoots.Calls')
            $preserved=(FileHash $Fixture.Output) -ceq $before
            if($fault){Check ('Aggregation.Boundary.'+$case+'.FaultInjected') ([bool](Core 'Injected'))}
            if($case -ceq 'UnknownOutputColumn'){
                Check 'Aggregation.Boundary.UnknownOutputColumn.CompletePersistence' $success
                $customPreserved=$false
                if(Test-Path -LiteralPath $Fixture.Output){
                    $customBook=$excel.Workbooks.Open($Fixture.Output,0,$true)
                    try{
                        $custom=Table $customBook 'tblGlobalInventorySnapshot'
                        foreach($column in $custom.ListColumns){
                            if([string]$column.Name -ceq 'Local Annotation'){$customPreserved=[string]$column.DataBodyRange.Cells.Item(1,1).Value2 -ceq 'Preserve disposable annotation'}
                        }
                    }finally{$customBook.Close($false)}
                }
                Check 'Aggregation.Boundary.UnknownOutputColumn.ValuesPreserved' $customPreserved
            }else{
                Check ('Aggregation.Boundary.'+$case+'.NoFalseCompletion') (-not $success)
                if($case -cne 'FailureAfterSave'){Check ('Aggregation.Boundary.'+$case+'.ExistingOutputPreserved') $preserved}
            }
            $redacted=[bool](Call 'FeedbackRedacted' @($Fixture.Root,$feed))
            Check ('Aggregation.Boundary.'+$case+'.FeedbackExcludesEnteredPaths') $redacted
            if($case -cnotin @('FailureBeforeSave','FailureAfterSave','SourceChangedBeforeSave','FeedChangedBeforeSave','SignedOutBeforeSave','UnknownOutputColumn')){
                Check ('Aggregation.Boundary.'+$case+'.RefusedBeforeWriter') ($calls -eq 0)
            }elseif($case -cin @('FailureBeforeSave','FailureAfterSave')){
                $retry=@(Get-ChildItem -LiteralPath (Split-Path -Parent $Fixture.Output) -Recurse -File -Filter '*aggregation*retry*.json' -ErrorAction SilentlyContinue)
                Check ('Aggregation.Boundary.'+$case+'.DurableRetryStateRetained') ($retry.Count -gt 0)
                $boundRetry=$false
                foreach($file in $retry){
                    try{
                        $record=[IO.File]::ReadAllText($file.FullName)|ConvertFrom-Json
                        $sourceHash=FileHash $Fixture.Snapshot
                        $matching=@($record.Sources|Where-Object { $_.Path -ceq $Fixture.Snapshot -and $_.Hash -ieq $sourceHash })
                        if($record.Destination -ceq $Fixture.Output -and $record.Phase -is [string] -and $record.Phase -and $record.Persistence -is [string] -and $record.Persistence -and $matching.Count -eq 1){$boundRetry=$true}
                    }catch{$boundRetry=$false}
                }
                Check ('Aggregation.Boundary.'+$case+'.RetryRecordBindsDestinationSourceAndPhase') $boundRetry
                [void](Core 'Arm' @('','','','',''))
                [void](Call 'Aggregate')
                Check ('Aggregation.Boundary.'+$case+'.SameFormRetryRequiresRecapture') (-not [bool](Call 'OwnerSuccess') -and [int](Run 'invSys.Core.xlam' 'TestAggregationRoots.Calls') -eq 0)
            }
            if($null -ne $borrowed){
                $stillOpen=$false;$statePreserved=$false
                foreach($book in $excel.Workbooks){if([string]$book.FullName -ceq $Fixture.Output){$stillOpen=$true;$statePreserved=([bool]$book.Saved -eq ($case -ceq 'BorrowedCleanOutput'));$book.Close($false);break}}
                Check ('Aggregation.Boundary.'+$case+'.BorrowedWorkbookStillOpen') $stillOpen
                Check ('Aggregation.Boundary.'+$case+'.BorrowedSavedStatePreserved') $statePreserved
                $borrowed=$null
            }
            if($null -ne $authorityBook){
                $stillOpen=$false;$dirty=$false
                foreach($book in $excel.Workbooks){if([string]$book.FullName -ceq $authorityPath){$stillOpen=$true;$dirty=-not [bool]$book.Saved;$book.Close($false);break}}
                Check ('Aggregation.Boundary.'+$case+'.AuthorityStillOpenDirty') ($stillOpen -and $dirty)
                $authorityBook=$null
            }
            $redirected=Test-Path -LiteralPath (Join-Path $changedFeed 'Global\invSys.Global.InventorySnapshot.xlsb')
            Check ('Aggregation.Boundary.'+$case+'.NoRedirectedOutput') (-not $redirected)
            $Observations.Value+=[pscustomobject]@{Case=$case;OwnerSuccess=$success;CoreDispatches=$calls;ExistingOutputPreserved=$preserved;RedirectedOutput=$redirected}
            [void](Call 'CloseForm')
            [void](Core 'Arm' @('','','','',''))
            if(Test-Path -LiteralPath $Fixture.Output){[IO.File]::SetAttributes($Fixture.Output,$outputAttributes)}
            [IO.File]::WriteAllBytes($Fixture.Snapshot,$sourceBytes)
            [IO.File]::WriteAllBytes($Fixture.Config,$configBytes)
            [IO.File]::WriteAllBytes($authPath,$authBytes)
            if($redirected){Remove-Item -LiteralPath (Join-Path $changedFeed 'Global\invSys.Global.InventorySnapshot.xlsb')}
            $pending=@(Get-ChildItem -LiteralPath (Split-Path -Parent $Fixture.Output) -File -Filter 'invSys.Aggregation.*.retry.json' -ErrorAction SilentlyContinue)
            if($pending.Count){
                # Resolve through the real owner after restoring this deliberately
                # mutated fixture and capturing a fresh session. Never erase evidence
                # or overwrite uncertain feed bytes to prepare the next case.
                SelectTarget $Fixture
                $review=[string](Core 'CaptureDirect' @($HostBook.Name,$Fixture.Warehouse,$Fixture.Config,$Fixture.Snapshot))
                $resolved=[bool](Core 'ResolveKnown' @($review,$Fixture.Snapshot))
                Check ('Aggregation.Boundary.'+$case+'.ExplicitFreshReviewResolvedKnownState') $resolved
                if(-not $resolved){throw 'Known-state attempt could not be resolved; independent case preparation stopped.'}
            }
        }
        if($AdminAggregationBoundaryGroup -notin @('All','Ownership')){return}
        foreach($overlap in @('TrustedSource','SelectedAuthority','SelectedSnapshots')){
            SelectTarget $Fixture
            $overlapFeed=switch($overlap){
                'TrustedSource'{$Fixture.Root}
                'SelectedAuthority'{$Other.Root}
                'SelectedSnapshots'{Split-Path -Parent $Other.Snapshot}
            }
            if(-not [bool](Call 'SetFeedForComparison' @($overlapFeed,$Fixture.Warehouse))){throw 'Overlap feed fixture was not established.'}
            $overlapOutput=Join-Path $overlapFeed 'Global\invSys.Global.InventorySnapshot.xlsb'
            if(-not [IO.Path]::GetFullPath($overlapOutput).StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Overlap output escapes disposable scope.'}
            [void](Run 'invSys.Core.xlam' 'TestAggregationRoots.PathComparison' @($true,$overlapOutput))
            if(-not [bool](Run 'invSys.Core.xlam' 'TestAggregationRoots.NormalizedControl' @($Other.Snapshot,$overlapOutput))){throw 'Existing overlap output fixture unavailable.'}
            $overlapHash=FileHash $overlapOutput
            if(-not [bool](Call 'OpenForm' @($HostBook.Name)) -or [int](Call 'SelectSource' @($Other.Warehouse)) -ne 1){throw 'Overlap fixture did not enter the actual source-selection handler.'}
            [void](Call 'Aggregate')
            Check ('Aggregation.Overlap.'+$overlap+'.RefusedBeforeWriter') (-not [bool](Call 'OwnerSuccess') -and [int](Run 'invSys.Core.xlam' 'TestAggregationRoots.Calls') -eq 0)
            Check ('Aggregation.Overlap.'+$overlap+'.ExistingOutputPreserved') ((FileHash $overlapOutput) -ceq $overlapHash)
            [void](Call 'CloseForm')
            [IO.File]::WriteAllBytes($Fixture.Config,$configBytes)
            Remove-Item -LiteralPath $overlapOutput
        }
        SelectTarget $Fixture
        $closedHost=$excel.Workbooks.Add();$closedName=$closedHost.Name
        if(-not [bool](Call 'OpenForm' @($closedName)) -or [int](Call 'SelectSource' @($Fixture.Warehouse)) -ne 1){throw 'Closed-workbook fixture did not enter the actual handler.'}
        $closedHost.Close($false);$OtherBook.Activate()
        $otherSheets=$OtherBook.Worksheets.Count
        [void](Core 'Arm' @('','','','',''))
        [void](Run 'invSys.Core.xlam' 'TestAggregationRoots.PathComparison' @($true,$Fixture.Output))
        [void](Call 'Aggregate')
        Check 'Aggregation.Boundary.ClosedOriginalWorkbook.RefusedBeforeWriter' ([int](Run 'invSys.Core.xlam' 'TestAggregationRoots.Calls') -eq 0)
        Check 'Aggregation.Boundary.ClosedOriginalWorkbook.OtherWorkbookPreserved' ($OtherBook.Worksheets.Count -eq $otherSheets)
        [void](Call 'CloseForm')
        SelectTarget $Fixture
        $context=[string](Core 'CaptureDirect' @($HostBook.Name,$Fixture.Warehouse,$Fixture.Config,$Fixture.Snapshot))
        if(-not $context){throw 'Direct interactive capture unavailable; not behavioral RED.'}
        $positive=Join-Path $runRoot 'direct-interactive-positive.xlsb'
        [void](Core 'Arm' @('','','','',''))
        Check 'Aggregation.DirectCore.BaselinePositiveWriterControl' ([bool](Core 'DirectInteractive' @($context,$Fixture.Snapshot,$positive)))
        $savedFeed=FileHash $Fixture.Output
        [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
        $denied=Join-Path $runRoot 'direct-interactive-signed-out.xlsb'
        $wrote=[bool](Core 'DirectInteractive' @($context,$Fixture.Snapshot,$denied))
        Check 'Aggregation.DirectCore.SignedOutInteractiveDispatchDenied' (-not $wrote -and -not(Test-Path -LiteralPath $denied) -and (FileHash $Fixture.Output) -ceq $savedFeed)
        $Observations.Value+=[pscustomobject]@{Case='DirectCoreSignedOut';OwnerSuccess=$wrote;SourceOutput=(Test-Path -LiteralPath $denied);BaselineAdapter=($packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modHqAggregator').CodeModule.Lines(1,200) -notmatch 'commandContext')}
    }finally{
        [void](Call 'CloseForm');[void](Core 'Arm' @('','','','',''))
        [IO.File]::WriteAllBytes($Fixture.Snapshot,$sourceBytes)
        [IO.File]::WriteAllBytes($Fixture.Config,$configBytes)
        [IO.File]::WriteAllBytes($authPath,$authBytes)
        SelectTarget $Fixture
        [void](Run 'invSys.Core.xlam' 'TestAggregationRoots.PathComparison' @($false,$Fixture.Output))
    }
}
