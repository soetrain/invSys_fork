# Supplemental D3 checks use only Admin-generated/seeded disposable inventory.
# Compare row values in memory; emit booleans only, never values or credentials.
function Install-InventoryQueryReadOnlyProbe {
    $module=$packages['invSys.Core.xlam'].VBProject.VBComponents.Add(1)
    $module.Name='TestInventoryQueryReadOnly'
    $module.CodeModule.AddFromString(@'
Option Explicit
Public Function ReadForTest(ByVal query As String, ByVal referenceName As String, ByVal suppliedName As String, ByVal missing As Boolean) As String
    Dim reference As Workbook, supplied As Workbook, entities As ListObject
    Dim expected As Variant, actual As Variant, sku As String, macro As String
    Set reference = Application.Workbooks(referenceName)
    Set entities = reference.Worksheets("InventoryEntities").ListObjects("tblInventoryEntities")
    sku = CStr(entities.DataBodyRange.Cells(1, entities.ListColumns("SKU").Index).Value2)
    If suppliedName <> "" Then Set supplied = Application.Workbooks(suppliedName)
    Select Case query
        Case "Quantity": macro = "GetOnHandQtyBridgeResult"
        Case "Locations": macro = "GetLocationBalancesBridgeResult"
        Case "Picker": macro = "ListInventoryPickerItemsBridgeResult": sku = ""
        Case "Entities": macro = "ListAvailableInventoryEntitiesBridgeResult": sku = ""
        Case Else: Err.Raise 5, , "Unknown fixture query."
    End Select
    expected = Application.Run("'invSys.Inventory.Domain.xlam'!modInventoryBridgeApi." & macro, sku, reference)
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

function Test-InventoryQueryReadOnly {
    $fixture=NewFixture 'inventory-query-read-only';SelectTarget $fixture
    $seed=[string](Run 'invSys.Admin.xlam' 'modAdminConsole.SeedDemoInventoryForAutomation' @($fixture.Warehouse,'S1','config-admin'))
    if(-not $seed.StartsWith('OK|')){throw 'Admin Seed query fixture unavailable; not product RED.'}
    $path=Join-Path $fixture.Root ($fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    $referencePath=Join-Path $fixture.Root 'query-reference.xlsb'
    $held=Join-Path $fixture.Root 'query-source.held'
    $root=[IO.Path]::GetFullPath($fixture.Root).TrimEnd('\')+'\'
    foreach($target in @($path,$referencePath,$held)){
        if(-not [IO.Path]::GetFullPath($target).StartsWith($root,[StringComparison]::OrdinalIgnoreCase)){throw 'Query fixture path escaped owned root.'}
    }
    function OpenTargets {@($excel.Workbooks|Where-Object{$_.FullName -ceq $path})}
    function CloseTargets {foreach($target in @(OpenTargets)){$target.Close($false)}}
    function HashTarget {(Get-FileHash -LiteralPath $path).Hash}
    function Query([string]$Kind,[string]$Supplied='', [bool]$Missing=$false){
        [string](Run 'invSys.Core.xlam' 'TestInventoryQueryReadOnly.ReadForTest' @($Kind,$reference.Name,$Supplied,$Missing))
    }
    CloseTargets
    $reference=$null;$source=$null
    try {
        $source=$excel.Workbooks.Open($path,0,$false)
        $entities=Table $source 'tblInventoryEntities'
        if($null -eq $entities.DataBodyRange -or $entities.ListRows.Count -eq 0){throw 'Nonempty Admin-seeded inventory required.'}
        $custom=$entities.ListColumns.Add();$custom.Name='Operator Annotation';$custom.DataBodyRange.Value2='Query fixture custom value'
        $entities.Parent.Protect()
        $source.Save();$source.Close($false);$source=$null
        Copy-Item -LiteralPath $path -Destination $referencePath
        $pin=HashTarget
        $reference=$excel.Workbooks.Open($referencePath,0,$true)
        SelectTarget $fixture 'config-producer'
        foreach($kind in @('Quantity','Locations','Picker','Entities')){
            $result=Query $kind
            Check ('InventoryRead.Cold.'+$kind+'.NonemptyExactResults') ($result -ceq 'True|True')
            Check ('InventoryRead.Cold.'+$kind+'.BytesPreserved') ((HashTarget) -ceq $pin)
            Check ('InventoryRead.Cold.'+$kind+'.TransientClosed') (@(OpenTargets).Count -eq 0)
            CloseTargets
            # Restore only this generated source between independent cases;
            # record preservation before restoration, never repin after a read.
            Copy-Item -LiteralPath $referencePath -Destination $path -Force
        }
        foreach($mode in @('Implicit','Supplied')){
            foreach($kind in @('Quantity','Locations','Picker','Entities')){
                $source=$excel.Workbooks.Open($path,0,$false)
                $entities=Table $source 'tblInventoryEntities';$entities.Parent.Unprotect()
                $entities.ListColumns.Item('Operator Annotation').DataBodyRange.Cells.Item(1,1).Value2='Unsaved query fixture value'
                $entities.Parent.Protect()
                if($source.Saved){throw 'Dirty caller fixture was not staged.'}
                $supplied='';if($mode -ceq 'Supplied'){$supplied=$source.Name}
                $result=Query $kind $supplied
                $prefix='InventoryRead.'+$mode+'.'+$kind
                Check ($prefix+'.NonemptyExactResults') ($result -ceq 'True|True')
                Check ($prefix+'.CallerStillOpen') (@(OpenTargets).Count -eq 1)
                Check ($prefix+'.DirtyStatePreserved') (-not [bool]$source.Saved)
                Check ($prefix+'.ProtectionAndCustomValuePreserved') ($entities.Parent.ProtectContents -and $entities.ListColumns.Item('Operator Annotation').DataBodyRange.Cells.Item(1,1).Value2 -ceq 'Unsaved query fixture value')
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
                Check ('InventoryRead.Missing.'+$kind+'.NoStoreCreated') (-not(Test-Path -LiteralPath $path))
                Check ('InventoryRead.Missing.'+$kind+'.NoWorkbookLeftOpen') (@(OpenTargets).Count -eq 0)
            }finally{
                CloseTargets
                if(Test-Path -LiteralPath $path){Remove-Item -LiteralPath $path}
                Move-Item -LiteralPath $held -Destination $path
            }
        }
        Check 'InventoryRead.ReferencePreserved' ((Get-FileHash -LiteralPath $referencePath).Hash -ceq $pin -and $reference.Saved -and $reference.ReadOnly)
    }finally{if($null -ne $source){$source.Close($false)};CloseTargets;if($null -ne $reference){$reference.Close($false)}}
}
