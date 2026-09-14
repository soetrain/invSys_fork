# Preserve historical detail/group fixture semantics while exercising EVENTS1.
# Older package baselines keep their existing published XLSB fixture. A candidate
# with the new reader publishes the same synthetic projection lines through the
# real Core publication class; its reader is never replaced or bypassed.
function Publish-Slice4beProjectionFixture($Fixture,[string]$Snapshot) {
    $reader=$null
    try{$reader=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modPublishedEventsReader')}catch{}
    if($null -eq $reader){return}
    $module=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modInventoryViewerData').CodeModule
    $module.AddFromString(@'
Public Function PublishProjectionFixtureForTest(ByVal workbookName As String) As Boolean
    Dim publisher As cEventsPublication, source As Workbook, report As String, model As Object
    Set source = Application.Workbooks(workbookName)
    If Not source.ReadOnly Then Exit Function
    Set publisher = New cEventsPublication
    publisher.CaptureInventory source.Worksheets("InventoryEvents").ListObjects("tblInventoryEvents")
    If Not publisher.Publish(modNasConnection.GetCurrentTargetWarehouseId(), modNasConnection.GetCurrentTargetRuntimeRoot(), report) Then Exit Function
    Set model = modEventsPublicationStore.Read(modNasConnection.GetCurrentTargetRuntimeRoot() & "\" & _
        modNasConnection.GetCurrentTargetWarehouseId() & ".invSys.Snapshot.Events.json", modNasConnection.GetCurrentTargetWarehouseId())
    PublishProjectionFixtureForTest = Not model Is Nothing
End Function
'@)
    $before=(Get-FileHash -LiteralPath $Snapshot).Hash
    $book=$excel.Workbooks.Open($Snapshot,0,$true)
    try {
        if(-not [bool](Run 'invSys.Core.xlam' 'modInventoryViewerData.PublishProjectionFixtureForTest' @($book.Name))) {throw 'Typed projection publication fixture failed; not behavioral RED.'}
    } finally {$book.Close($false)}
    if((Get-FileHash -LiteralPath $Snapshot).Hash -cne $before){throw 'Publishing the projection fixture changed its saved source.'}
}
