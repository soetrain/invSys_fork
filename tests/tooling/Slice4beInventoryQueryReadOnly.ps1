# Supplemental D3 checks use only Admin-generated/seeded disposable inventory.
# Compare row values in memory; emit booleans only, never values or credentials.
function Install-InventoryQueryReadOnlyProbe {
    $domain=$packages['invSys.Inventory.Domain.xlam'].VBProject.VBComponents.Item('modInventoryBridgeApi').CodeModule
    $domain.InsertLines($domain.CountOfDeclarationLines+1,"Private mQuerySeenForTest As Boolean`r`nPrivate mQueryReadOnlyForTest As Boolean`r`nPrivate mQueryFailForTest As Boolean")
    $domain.AddFromString(@'
Public Sub ResetQueryObservationForTest(ByVal failQuery As Boolean)
    mQuerySeenForTest = False
    mQueryReadOnlyForTest = False
    mQueryFailForTest = failQuery
End Sub
Public Function QueryObservationForTest() As String
    QueryObservationForTest = CStr(mQuerySeenForTest) & "|" & CStr(mQueryReadOnlyForTest)
End Function
Private Sub ObserveQueryForTest(ByVal wb As Workbook)
    mQuerySeenForTest = True
    If Not wb Is Nothing Then mQueryReadOnlyForTest = wb.ReadOnly
End Sub
'@)
    foreach($procedure in @('GetOnHandQtyBridgeResult','GetLocationBalancesBridgeResult','ListInventoryPickerItemsBridgeResult','ListAvailableInventoryEntitiesBridgeResult')){
        $start=$domain.ProcStartLine($procedure,0);$count=$domain.ProcCountLines($procedure,0)
        $lines=$domain.Lines($start,$count) -split '\r?\n'
        $anchors=@(for($i=0;$i -lt $lines.Count;$i++){if($lines[$i] -match ('^\s*'+[regex]::Escape($procedure)+'\s*=')){$start+$i}})
        if($anchors.Count -ne 1){throw 'Unique Domain query observation boundary unavailable.'}
        # The declared Domain query failure envelope is Empty (or zero for
        # quantity), not an unhandled cross-project VBA exception dialog.
        $domain.InsertLines($anchors[0],"    ObserveQueryForTest inventoryWb`r`n    If mQueryFailForTest Then Exit Function")
    }
    $module=$packages['invSys.Core.xlam'].VBProject.VBComponents.Add(1)
    $module.Name='TestInventoryQueryReadOnly'
    $module.CodeModule.AddFromString(@'
Option Explicit
Private mExpected(0 To 3) As Variant
Private mQuerySku As String
Public Function StageDirtyForTest(ByVal workbookName As String) As Boolean
    Dim wb As Workbook, lo As ListObject, added As ListColumn
    Set wb = Application.Workbooks(workbookName)
    Set lo = wb.Worksheets("InventoryEntities").ListObjects("tblInventoryEntities")
    lo.Parent.Unprotect
    Set added = lo.ListColumns.Add
    added.Name = "Operator Annotation"
    added.DataBodyRange.Cells(1, 1).Value2 = "Unsaved query fixture value"
    lo.Parent.Protect
    StageDirtyForTest = Not wb.Saved And DirtyPreservedForTest(workbookName)
End Function
Public Function DirtyPreservedForTest(ByVal workbookName As String) As Boolean
    On Error GoTo CleanFail
    Dim wb As Workbook, lo As ListObject
    Set wb = Application.Workbooks(workbookName)
    Set lo = wb.Worksheets("InventoryEntities").ListObjects("tblInventoryEntities")
    DirtyPreservedForTest = lo.Parent.ProtectContents And _
        CStr(lo.ListColumns("Operator Annotation").DataBodyRange.Cells(1, 1).Value2) = "Unsaved query fixture value"
CleanFail:
End Function
Public Sub CaptureForTest(ByVal referenceName As String)
    Dim reference As Workbook, entities As ListObject
    Set reference = Application.Workbooks(referenceName)
    Set entities = reference.Worksheets("InventoryEntities").ListObjects("tblInventoryEntities")
    mQuerySku = CStr(entities.DataBodyRange.Cells(1, entities.ListColumns("SKU").Index).Value2)
    mExpected(0) = Application.Run("'invSys.Inventory.Domain.xlam'!modInventoryBridgeApi.GetOnHandQtyBridgeResult", mQuerySku, reference)
    mExpected(1) = Application.Run("'invSys.Inventory.Domain.xlam'!modInventoryBridgeApi.GetLocationBalancesBridgeResult", mQuerySku, reference)
    mExpected(2) = Application.Run("'invSys.Inventory.Domain.xlam'!modInventoryBridgeApi.ListInventoryPickerItemsBridgeResult", "", reference)
    mExpected(3) = Application.Run("'invSys.Inventory.Domain.xlam'!modInventoryBridgeApi.ListAvailableInventoryEntitiesBridgeResult", "", reference)
End Sub
Public Function ReadForTest(ByVal query As String, ByVal suppliedName As String, ByVal missing As Boolean) As String
    Dim supplied As Workbook, expected As Variant, actual As Variant, sku As String
    If suppliedName <> "" Then Set supplied = Application.Workbooks(suppliedName)
    sku = mQuerySku
    Select Case query
        Case "Quantity": expected = mExpected(0)
        Case "Locations": expected = mExpected(1)
        Case "Picker": expected = mExpected(2): sku = ""
        Case "Entities": expected = mExpected(3): sku = ""
        Case Else: Err.Raise 5, , "Unknown fixture query."
    End Select
    Select Case query
        Case "Quantity": actual = modInventoryDomainBridge.GetInventoryOnHandQtyBridge(sku, supplied)
        Case "Locations": actual = modInventoryDomainBridge.GetInventoryLocationBalancesBridge(sku, supplied)
        Case "Picker": actual = modInventoryDomainBridge.ListInventoryPickerItemsBridge(sku, supplied)
        Case "Entities": actual = modInventoryDomainBridge.ListAvailableInventoryEntitiesBridge(sku, supplied)
    End Select
    If missing Then
        If query = "Quantity" Then
            ReadForTest = CStr(CDbl(actual) = 0)
        Else
            ReadForTest = CStr(IsEmpty(actual))
        End If
    Else
        ReadForTest = CStr(EquivalentForTest(expected, actual)) & "|" & CStr(HasValuesForTest(expected))
    End If
End Function
Private Function HasValuesForTest(ByVal value As Variant) As Boolean
    If IsArray(value) Then
        HasValuesForTest = (UBound(value, 1) >= LBound(value, 1))
    ElseIf Not IsEmpty(value) Then
        HasValuesForTest = (CDbl(value) <> 0)
    End If
End Function
Private Function EquivalentForTest(ByVal expected As Variant, ByVal actual As Variant) As Boolean
    Dim r As Long, c As Long
    If IsArray(expected) <> IsArray(actual) Then Exit Function
    If IsArray(expected) Then
        If LBound(expected, 1) <> LBound(actual, 1) Or UBound(expected, 1) <> UBound(actual, 1) Then Exit Function
        If LBound(expected, 2) <> LBound(actual, 2) Or UBound(expected, 2) <> UBound(actual, 2) Then Exit Function
        For r = LBound(expected, 1) To UBound(expected, 1)
            For c = LBound(expected, 2) To UBound(expected, 2)
                If StrComp(CStr(expected(r, c)), CStr(actual(r, c)), vbBinaryCompare) <> 0 Then Exit Function
            Next c
        Next r
    Else
        If IsEmpty(expected) <> IsEmpty(actual) Then Exit Function
        If CStr(expected) <> CStr(actual) Then Exit Function
    End If
    EquivalentForTest = True
End Function
'@)
}

function Test-InventoryQueryReadOnly($Fixture) {
    $path=Join-Path $fixture.Root ($fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    $referencePath=Join-Path $fixture.Root 'query-reference.xlsb'
    $held=Join-Path $fixture.Root 'query-source.held'
    $root=[IO.Path]::GetFullPath($fixture.Root).TrimEnd('\')+'\'
    foreach($target in @($path,$referencePath,$held)){
        if(-not [IO.Path]::GetFullPath($target).StartsWith($root,[StringComparison]::OrdinalIgnoreCase)){throw 'Query fixture path escaped owned root.'}
    }
    function SetupStage([string]$Stage){
        $size=0;$entries=0;$valid=$false
        if(Test-Path -LiteralPath $path){
            $size=(Get-Item -LiteralPath $path).Length
            Add-Type -AssemblyName System.IO.Compression
            $stream=[IO.File]::Open($path,'Open','Read','ReadWrite');$zip=$null
            try{$zip=[IO.Compression.ZipArchive]::new($stream,[IO.Compression.ZipArchiveMode]::Read,$false);$entries=$zip.Entries.Count;$valid=$true}finally{if($null -ne $zip){$zip.Dispose()}else{$stream.Dispose()}}
        }
        $opened=@($excel.Workbooks|Where-Object{$_.FullName -ceq $path})
        $format=0;$readOnly=$null;$saved=$null
        if($opened.Count -eq 1){$format=$opened[0].FileFormat;$readOnly=$opened[0].ReadOnly;$saved=$opened[0].Saved}
        [pscustomobject]@{Stage=$Stage;SourceExists=(Test-Path -LiteralPath $path);SourceOpenCount=$opened.Count;Bytes=$size;ZipValid=$valid;Entries=$entries;OpenFormat=$format;OpenReadOnly=$readOnly;OpenSaved=$saved}|ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'inventory-query-setup.jsonl')
    }
    SetupStage 'BeforeSeed'
    SelectTarget $fixture
    $seed=[string](Run 'invSys.Admin.xlam' 'modAdminConsole.SeedDemoInventoryForAutomation' @($fixture.Warehouse,'S1','config-admin'))
    if(-not $seed.StartsWith('OK|')){throw 'Admin Seed query fixture unavailable; not product RED.'}
    SetupStage 'AfterSeed'
    function OpenTargets {@($excel.Workbooks|Where-Object{$_.FullName -ceq $path})}
    function CloseTargets {foreach($target in @(OpenTargets)){$target.Close($false)}}
    function HashTarget {
        $stream=[IO.File]::Open($path,'Open','Read','ReadWrite')
        try{(Get-FileHash -InputStream $stream).Hash}finally{$stream.Dispose()}
    }
    function Query([string]$Kind,[string]$Supplied='', [bool]$Missing=$false,[bool]$FailQuery=$false){
        [void](Run 'invSys.Inventory.Domain.xlam' 'modInventoryBridgeApi.ResetQueryObservationForTest' @($FailQuery))
        [string](Run 'invSys.Core.xlam' 'TestInventoryQueryReadOnly.ReadForTest' @($Kind,$Supplied,$Missing))
    }
    function QueryState {[string](Run 'invSys.Inventory.Domain.xlam' 'modInventoryBridgeApi.QueryObservationForTest')}
    CloseTargets
    $source=$null
    try {
        SetupStage 'BeforeSourceOpen'
        $source=$excel.Workbooks.Open($path,0,$false)
        $entities=Table $source 'tblInventoryEntities'
        if($null -eq $entities.DataBodyRange -or $entities.ListRows.Count -eq 0){throw 'Nonempty Admin-seeded inventory required.'}
        [void](Run 'invSys.Core.xlam' 'TestInventoryQueryReadOnly.CaptureForTest' @($source.Name))
        $source.Close($false);$source=$null
        Copy-Item -LiteralPath $path -Destination $referencePath
        $pin=HashTarget
        SetupStage 'ExpectedQueriesCaptured'
        SelectTarget $fixture 'config-producer'
        foreach($kind in @('Quantity','Locations','Picker','Entities')){
            SetupStage ('BeforeCold'+$kind)
            $result=Query $kind
            Check ('InventoryRead.Cold.'+$kind+'.NonemptyExactResults') ($result -ceq 'True|True')
            Check ('InventoryRead.Cold.'+$kind+'.SourceReadOnly') ((QueryState) -ceq 'True|True')
            Check ('InventoryRead.Cold.'+$kind+'.BytesPreserved') ((HashTarget) -ceq $pin)
            Check ('InventoryRead.Cold.'+$kind+'.TransientClosed') (@(OpenTargets).Count -eq 0)
            CloseTargets
            # Restore only this generated source between independent cases;
            # record preservation before restoration, never repin after a read.
            Copy-Item -LiteralPath $referencePath -Destination $path -Force
        }
        foreach($mode in @('Implicit','Supplied')){
            foreach($kind in @('Quantity','Locations','Picker','Entities')){
                SetupStage ('Before'+$mode+$kind+'Open')
                $source=$excel.Workbooks.Open($path,0,$false)
                SetupStage ('After'+$mode+$kind+'Open')
                if(-not [bool](Run 'invSys.Core.xlam' 'TestInventoryQueryReadOnly.StageDirtyForTest' @($source.Name))){throw 'Typed dirty caller fixture not established.'}
                if($source.Saved){throw 'Dirty caller fixture was not staged.'}
                $supplied='';if($mode -ceq 'Supplied'){$supplied=$source.Name}
                $result=Query $kind $supplied
                SetupStage ('After'+$mode+$kind+'Query')
                $prefix='InventoryRead.'+$mode+'.'+$kind
                Check ($prefix+'.NonemptyExactResults') ($result -ceq 'True|True')
                Check ($prefix+'.CallerStillOpen') (@(OpenTargets).Count -eq 1)
                Check ($prefix+'.DirtyStatePreserved') (-not [bool]$source.Saved)
                Check ($prefix+'.ProtectionAndCustomValuePreserved') ([bool](Run 'invSys.Core.xlam' 'TestInventoryQueryReadOnly.DirtyPreservedForTest' @($source.Name)))
                Check ($prefix+'.SavedBytesPreserved') ((HashTarget) -ceq $pin)
                $source.Close($false);$source=$null
                Copy-Item -LiteralPath $referencePath -Destination $path -Force
            }
        }
        foreach($kind in @('Quantity','Locations','Picker','Entities')){
            Move-Item -LiteralPath $path -Destination $held
            try{
                $result=Query $kind '' $true
                Check ('InventoryRead.Missing.'+$kind+'.ExistingEmptyResult') ($result -ceq 'True')
                Check ('InventoryRead.Missing.'+$kind+'.NoDomainQuery') ((QueryState) -ceq 'False|False')
                Check ('InventoryRead.Missing.'+$kind+'.NoStoreCreated') (-not(Test-Path -LiteralPath $path))
                Check ('InventoryRead.Missing.'+$kind+'.NoWorkbookLeftOpen') (@(OpenTargets).Count -eq 0)
            }finally{
                CloseTargets
                if(Test-Path -LiteralPath $path){Remove-Item -LiteralPath $path}
                Move-Item -LiteralPath $held -Destination $path
            }
        }
        $result=Query 'Picker' '' $true $true
        Check 'InventoryRead.Failure.Cold.EmptyResult' ($result -ceq 'True')
        Check 'InventoryRead.Failure.Cold.SourceReadOnly' ((QueryState) -ceq 'True|True')
        Check 'InventoryRead.Failure.Cold.BytesPreserved' ((HashTarget) -ceq $pin)
        Check 'InventoryRead.Failure.Cold.TransientClosed' (@(OpenTargets).Count -eq 0)
        CloseTargets
        Copy-Item -LiteralPath $referencePath -Destination $path -Force
        $source=$excel.Workbooks.Open($path,0,$false)
        if(-not [bool](Run 'invSys.Core.xlam' 'TestInventoryQueryReadOnly.StageDirtyForTest' @($source.Name))){throw 'Typed failed-query caller fixture not established.'}
        $result=Query 'Picker' $source.Name $true $true
        Check 'InventoryRead.Failure.Supplied.EmptyResult' ($result -ceq 'True')
        Check 'InventoryRead.Failure.Supplied.DirtyProtectedCallerPreserved' (@(OpenTargets).Count -eq 1 -and -not $source.Saved -and [bool](Run 'invSys.Core.xlam' 'TestInventoryQueryReadOnly.DirtyPreservedForTest' @($source.Name)))
        Check 'InventoryRead.Failure.Supplied.SavedBytesPreserved' ((HashTarget) -ceq $pin)
        $source.Close($false);$source=$null
        Check 'InventoryRead.ReferencePreserved' ((Get-FileHash -LiteralPath $referencePath).Hash -ceq $pin)
    }finally{[void](Run 'invSys.Inventory.Domain.xlam' 'modInventoryBridgeApi.ResetQueryObservationForTest' @($false));if($null -ne $source){$source.Close($false)};CloseTargets}
}
