# D18 owner-query tests. Only generated fixtures are changed by setup/faults.
# Domain applies lifecycle events; the query must not repair or save its source.
function Test-Slice4beDesignsPublicationSource($Fixture,$Other) {
    $domain=$packages['invSys.Designs.Domain.xlam'].VBProject.VBComponents.Item('modDesignsBridgeApi').CodeModule
    $domain.AddFromString(@'
Public Function PrepareDesignsPublicationForTest(ByVal warehouse As String, ByVal populated As Boolean) As Boolean
    Dim wb As Workbook, evt As Object, report As String, status As String, code As String, message As String
    Dim ws As Worksheet, lo As ListObject, events As ListObject, column As ListColumn, i As Long
    Set wb = modDesignsRuntime.ResolveDesignsWorkbook(warehouse, Nothing, report)
    If wb Is Nothing Then Exit Function
    On Error GoTo Done
    If populated Then
        For i = 1 To 2
            Set evt = CreateObject("Scripting.Dictionary")
            evt("EventID") = IIf(i = 1, "Evt-Pub-Design-A", "evt-pub-design-b")
            evt("EventType") = IIf(i = 1, "DESIGN_CREATE", "DESIGN_RELEASE")
            evt("CreatedAtUTC") = Now
            evt("WarehouseId") = warehouse: evt("StationId") = "S1": evt("UserId") = "config-admin"
            evt("DesignId") = "Design-Pub-A": evt("DesignVersion") = "1"
            evt("PayloadJson") = ""
            If i = 1 Then evt("PayloadJson") = "[{""DesignType"":""RECIPE"",""DesignName"":""Publication fixture"",""Description"":""DO-NOT-PUBLISH-PAYLOAD"",""LineNo"":1,""Process"":1,""IOType"":""USED"",""ComponentSKU"":""SKU-GROUP"",""Qty"":1,""UOM"":""EA"",""Percent"":100}]"
            evt("Note") = "Publication note" & vbCrLf & "Detail" & vbTab & "\finish"
            If Not modDesignsApply.ApplyDesignEvent(evt, wb, "PUB-SOURCE-FIXTURE", status, code, message) Then GoTo Done
            If status <> "APPLIED" Then GoTo Done
        Next i
    End If
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = "tblDesignEvents" Then Set events = lo
        Next lo
    Next ws
    If events Is Nothing Then GoTo Done
    Set column = events.ListColumns.Add(1): column.Name = "User publication sentinel"
    If Not events.DataBodyRange Is Nothing Then column.DataBodyRange.Value2 = "DO-NOT-PUBLISH-UNKNOWN"
    events.ListColumns("EventID").Name = "eventID"
    wb.Save
    PrepareDesignsPublicationForTest = (events.ListRows.Count = IIf(populated, 2, 0))
Done:
    wb.Close False
End Function
Public Function SetPublicationHeaderForTest(ByVal workbookName As String, ByVal malformed As Boolean) As Boolean
    Dim wb As Workbook, ws As Worksheet, lo As ListObject, column As ListColumn
    Dim oldName As String, newName As String
    Set wb = Application.Workbooks(workbookName)
    oldName = IIf(malformed, "eventID", "AbsentEventIdentity")
    newName = IIf(malformed, "AbsentEventIdentity", "eventID")
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = "tblDesignEvents" Then
                Set column = lo.ListColumns(oldName)
                column.Name = newName
                wb.Save
                SetPublicationHeaderForTest = (column.Name = newName And wb.Saved)
                Exit Function
            End If
        Next lo
    Next ws
End Function
'@)
    $core=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modDesignsDomainBridge').CodeModule
    $core.AddFromString(@'
Public Function DesignsPublicationQueryForTest(ByVal warehouse As String, ByVal root As String) As String
    DesignsPublicationQueryForTest = CStr(RunDesignsQueryBridge("PUBLICATION_EVENTS", warehouse, root, "", Nothing))
End Function
'@)
    function SourceQuery($Target,[string]$Root='') {
        if($Root -eq ''){$Root=$Target.Root}
        [string](Run 'invSys.Core.xlam' 'modDesignsDomainBridge.DesignsPublicationQueryForTest' @($Target.Warehouse,$Root))
    }
    function DecodeSource([string]$Text) {
        $result=[pscustomobject]@{Valid=$false;Availability='';Warehouse='';Mode='';Reason='';Count=-1;Rows=@()}
        $lines=$Text -split "`r`n";$head=$lines[0] -split "`t"
        if($head.Count -ne 7 -or $head[0] -cne 'EVTSRC1' -or $head[1] -cne 'Designs'){return $result}
        $result.Valid=$true;$result.Availability=$head[2];$result.Warehouse=$head[3];$result.Mode=$head[5];$result.Reason=$head[6]
        [int]$count=0;if(-not [int]::TryParse($head[4],[ref]$count)){$result.Valid=$false;return $result};$result.Count=$count
        if($result.Availability -cne 'Available'){return $result}
        if($lines.Count -ne $count+2){$result.Valid=$false;return $result}
        $fields=$lines[1] -split "`t"
        for($r=0;$r -lt $count;$r++){
            $values=$lines[$r+2] -split "`t";if($values.Count -ne $fields.Count){$result.Valid=$false;return $result}
            $record=[ordered]@{}
            for($c=0;$c -lt $fields.Count;$c++){
                $record[$fields[$c]]=[regex]::Replace($values[$c],'\\([\\trn])',[Text.RegularExpressions.MatchEvaluator]{param($m) switch($m.Groups[1].Value){'t'{"`t"};'r'{"`r"};'n'{"`n"};'\'{'\'}}})
            }
            $result.Rows+=,[pscustomobject]$record
        }
        $result
    }
    function IsUnavailable($Value,[string]$Reason) {$Value.Valid -and $Value.Availability -ceq 'Unavailable' -and $Value.Count -eq 0 -and $Value.Reason -ceq $Reason}
    foreach($entry in @(@($Fixture,$true),@($Other,$false))){
        SelectTarget $entry[0] 'config-admin'
        if(-not [bool](Run 'invSys.Designs.Domain.xlam' 'modDesignsBridgeApi.PrepareDesignsPublicationForTest' @($entry[0].Warehouse,$entry[1]))){throw 'Domain fixture preparation failed.'}
    }
    SelectTarget $Fixture 'config-admin'
    Check 'DesignsPublication.OwnerAppliedLifecycleFixture' $true
    $path=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Designs.xlsb')
    $before=(Get-FileHash -LiteralPath $path).Hash
    $text=SourceQuery $Fixture;$read=DecodeSource $text
    Check 'DesignsPublication.VersionedOwnerEnvelope' ($read.Valid -and $read.Availability -ceq 'Available' -and $read.Warehouse -ceq $Fixture.Warehouse)
    $exact=$read.Rows.Count -eq 2
    if($exact){$exact=$read.Rows[0].EventID -ceq 'Evt-Pub-Design-A' -and $read.Rows[1].EventID -ceq 'evt-pub-design-b' -and $read.Rows[0].DefinitionId -ceq 'Design-Pub-A' -and $read.Rows[1].EventType -ceq 'DESIGN_RELEASE'}
    Check 'DesignsPublication.ExactLifecycleIdentitiesAndAllRows' $exact
    $note=$read.Rows.Count -eq 2
    if($note){$note=$read.Rows[0].Note -ceq "Publication note`r`nDetail`t\finish" -and $read.Rows[0].OccurredAtUTC -match '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$'}
    Check 'DesignsPublication.LosslessTextAndUnverifiedRecordedTime' $note
    Check 'DesignsPublication.PayloadAndUnknownColumnsExcluded' ($read.Valid -and -not $text.Contains('PayloadJson') -and -not $text.Contains('DO-NOT-PUBLISH') -and -not $text.Contains('User publication sentinel'))
    Check 'DesignsPublication.TransientSourceIsReadOnly' ($read.Mode -ceq 'ReadOnly')
    Check 'DesignsPublication.TransientSourceReleased' (@($excel.Workbooks|Where-Object FullName -eq $path).Count -eq 0)
    Check 'DesignsPublication.SourceBytesPreserved' ((Get-FileHash -LiteralPath $path).Hash -ceq $before)
    $book=$excel.Workbooks.Open($path,0,$false)
    try {
        $borrowed=DecodeSource (SourceQuery $Fixture)
        Check 'DesignsPublication.CleanBorrowedSourceRemainsOpen' ($borrowed.Availability -ceq 'Available' -and $borrowed.Mode -ceq 'Borrowed' -and [bool]$book.Saved -and @($excel.Workbooks|Where-Object FullName -eq $path).Count -eq 1)
        $table=Table $book 'tblDesignEvents';$table.ListColumns.Item('User publication sentinel').DataBodyRange.Cells.Item(1,1).Value2='UNSAVED-FIXTURE'
        $dirty=DecodeSource (SourceQuery $Fixture)
        Check 'DesignsPublication.UnsavedSourceUnavailableWithoutSave' ((IsUnavailable $dirty 'UnsavedSource') -and -not [bool]$book.Saved)
    } finally {$book.Close($false)}
    $book=$excel.Workbooks.Open($path,0,$false)
    if(-not [bool](Run 'invSys.Designs.Domain.xlam' 'modDesignsBridgeApi.SetPublicationHeaderForTest' @($book.Name,$true))){throw 'Missing-header fixture rename did not persist.'}
    try {
        $malformedHash=PublicationSourceHash $path;$malformed=DecodeSource (SourceQuery $Fixture)
        Check 'DesignsPublication.MissingHeaderUnavailableWithoutRepair' ((IsUnavailable $malformed 'InvalidSchema') -and (PublicationSourceHash $path) -ceq $malformedHash -and [bool]$book.Saved)
    } finally {
        if(-not [bool](Run 'invSys.Designs.Domain.xlam' 'modDesignsBridgeApi.SetPublicationHeaderForTest' @($book.Name,$false))){throw 'Missing-header fixture restoration failed.'}
        $book.Close($false)
    }
    $withheld=$path+'.publication-withheld'
    if(-not ([IO.Path]::GetFullPath($path)).StartsWith(([IO.Path]::GetFullPath($runRoot)).TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Source fault path escaped its fixture.'}
    Move-Item -LiteralPath $path -Destination $withheld
    try {
        $missing=DecodeSource (SourceQuery $Fixture)
        Check 'DesignsPublication.MissingSourceUnavailableWithoutCreation' ((IsUnavailable $missing 'MissingSource') -and -not (Test-Path -LiteralPath $path))
    } finally {Move-Item -LiteralPath $withheld -Destination $path}
    $wrongRoot=DecodeSource (SourceQuery $Fixture (Join-Path $Fixture.Root 'wrong-root'))
    Check 'DesignsPublication.CapturedRootMismatchDenied' (IsUnavailable $wrongRoot 'ContextMismatch')
    SelectTarget $Other 'config-admin'
    try {
        $stale=DecodeSource (SourceQuery $Fixture);$empty=DecodeSource (SourceQuery $Other)
        Check 'DesignsPublication.ChangedTargetRejectsOldSource' (IsUnavailable $stale 'ContextMismatch')
        Check 'DesignsPublication.ValidEmptySourceIsAvailable' ($empty.Valid -and $empty.Availability -ceq 'Available' -and $empty.Count -eq 0 -and $empty.Rows.Count -eq 0)
    } finally {SelectTarget $Fixture 'config-admin'}
}
