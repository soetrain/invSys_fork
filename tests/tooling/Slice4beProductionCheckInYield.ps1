# Unsaved probes interrupt actual Check In reads; original handlers and payloads remain intact.
function Install-ProductionCheckInYieldProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    function After-CheckRead([string]$Component,[string]$Procedure,[string]$Anchor,[string]$Boundary,[string]$Available){
        $module=$project.VBComponents.Item($Component).CodeModule
        $start=$module.ProcStartLine($Procedure,0);$end=$start+$module.ProcCountLines($Procedure,0)
        $hits=@(for($line=$start;$line -lt $end;$line++){if($module.Lines($line,1).Trim().StartsWith($Anchor,[StringComparison]::Ordinal)){$line}})
        if($hits.Count -ne 1){throw ('Check In read-return anchor changed: '+$Procedure+'; not product RED.')}
        $line=$hits[0]
        while($module.Lines($line,1).TrimEnd().EndsWith('_')){$line++;if($line -ge $end){throw 'Incomplete read-return anchor; not product RED.'}}
        $module.InsertLines($line+1,('    TestProductionDesigner.CheckYieldReturned "'+$Boundary+'", '+$Available))
    }
    After-CheckRead 'modProductionRunEntityReads' 'AvailableQuantity' 'entities = modInventoryDomainBridge.ListAvailableInventoryEntitiesBridge(' 'AvailableQuantity' 'IsArray(entities)'
    After-CheckRead 'modProductionRunEntityReads' 'IsNonCounted' 'entities = modInventoryDomainBridge.ListAvailableInventoryEntitiesBridge(' 'EntityKind' 'IsArray(entities)'
    After-CheckRead 'modProductionReusableRun' 'ReusableRunPaletteRows' 'entities = modInventoryDomainBridge.ListAvailableInventoryEntitiesBridge(' 'RunPalette' 'IsArray(entities)'
    After-CheckRead 'modProductionReusableRun' 'ReusableRunManagerCheckRows' 'entity = modInventoryDomainBridge.ListAvailableInventoryEntitiesBridge(' 'ManagerCheck' 'IsArray(entity)'
    After-CheckRead 'modProductionReusableRun' 'AvailableStockForRequirement' 'entities = modInventoryDomainBridge.ListAvailableInventoryEntitiesBridge(' 'AvailableStock' 'IsArray(entities)'
    After-CheckRead 'modProductionReusableRun' 'ExactEntityUom' 'entities = modInventoryDomainBridge.ListAvailableInventoryEntitiesBridge(' 'EntityUom' 'IsArray(entities)'
    After-CheckRead 'modProductionCheckInActions' 'ResolveSelectedKey' 'entities = modInventoryDomainBridge.ListAvailableInventoryEntitiesBridge(' 'ResolveKey' 'IsArray(entities)'
    After-CheckRead 'mProduction' 'LoadProductionRunInventoryPickerItems' 'result = modInventoryDomainBridge.ListInventoryPickerItemsBridge(' 'InventoryPicker' 'IsArray(result)'
    After-CheckRead 'mProduction' 'GetProductionRunDefaultLocation' 'GetProductionRunDefaultLocation = Trim$(modConfig.GetString(' 'DefaultLocation' 'True'
    After-CheckRead 'frmProduction' 'RefreshRunPaletteState' 'choices = mProduction.LoadProductionRunIngredientChoices(' 'IngredientChoices' 'True'
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    $form.AddFromString(@'
Public Sub CheckYieldResetCacheForTest()
    ResetInventoryCache
End Sub
Public Function CheckYieldProjectionForTest() As String
    Dim lists As New Collection, control As MSForms.ListBox, row As Long, column As Long, value As String
    value = RunLocalStateForTest() & "|" & CheckBaselineWorksheetStateForTest() & "|" & _
        CStr(mRunInventoryCacheLoaded) & "|" & CStr(IsEmpty(mRunInventoryRows))
    lists.Add mLstLoaderLines: lists.Add mLstRunPalette: lists.Add mLstManagerCheck
    lists.Add mLstManagerOutput: lists.Add mLstRunInstructions
    For Each control In lists
        value = value & "|" & CStr(control.ListCount) & ":" & CStr(control.ListIndex)
        For row = 0 To control.ListCount - 1
            For column = 0 To control.ColumnCount - 1
                value = value & "|" & NzStr(control.List(row, column))
            Next column
        Next row
    Next control
    CheckYieldProjectionForTest = value & "|" & CStr(mCmbRunLocation.ListCount) & "|" & CStr(mCmbTreeRunLocation.ListCount)
End Function
'@)
    $adapter=$project.VBComponents.Item('TestProductionDesigner').CodeModule
    $adapter.InsertLines(1,@'
Private mCheckYieldTarget As String, mCheckYieldReached As Boolean, mCheckYieldAvailable As Boolean
Private mCheckYieldOrdinal As Long, mCheckYieldSeen As Long, mCheckYieldLater As Long
Private mCheckYieldOwner As String, mCheckYieldProjection As String
Private mCheckYieldInterruption As String, mCheckYieldAuthPath As String, mCheckYieldValid As Boolean
'@)
    $adapter.AddFromString(@'
Public Sub CheckYieldReset()
    mCheckYieldTarget = "": mCheckYieldReached = False: mCheckYieldAvailable = False
    mCheckYieldOrdinal = 1: mCheckYieldSeen = 0: mCheckYieldLater = 0
    mCheckYieldOwner = "": mCheckYieldProjection = ""
    mCheckYieldInterruption = "": mCheckYieldAuthPath = "": mCheckYieldValid = False
End Sub
Public Sub CheckYieldArm(ByVal boundary As String, ByVal ordinal As Long, ByVal interruption As String, ByVal authPath As String)
    CheckYieldReset
    mCheckYieldTarget = boundary: mCheckYieldOrdinal = ordinal
    mCheckYieldInterruption = interruption: mCheckYieldAuthPath = authPath
End Sub
Public Sub CheckYieldReturned(ByVal boundary As String, ByVal available As Boolean)
    Dim beforeContext As String, denied As Boolean
    If mCheckYieldReached Then mCheckYieldLater = mCheckYieldLater + 1
    If mCheckYieldTarget = "" Or boundary <> mCheckYieldTarget Then Exit Sub
    mCheckYieldSeen = mCheckYieldSeen + 1
    If mCheckYieldSeen <> mCheckYieldOrdinal Then Exit Sub
    mCheckYieldTarget = "": mCheckYieldReached = True: mCheckYieldAvailable = available
    mCheckYieldOwner = modProductionReusableRun.RunLocalStateForTest()
    mCheckYieldProjection = mForm.CheckYieldProjectionForTest()
    If mCheckYieldInterruption = "Permission" Then
        beforeContext = modActivity.CaptureContext()
        CheckYieldRevokePermission
        denied = Not modRoleUiAccess.CanCurrentUserPerformCapability("PROD_POST") And _
            Not modRoleUiAccess.CanCurrentUserPerformCapability("ADMIN_MAINT")
        mCheckYieldValid = denied And modAuth.IsSignedIn() And _
            (modActivity.CaptureContext() = beforeContext)
    Else
        modAuth.SignOut
        mCheckYieldValid = Not modAuth.IsSignedIn()
    End If
End Sub
Public Function CheckYieldEvidence() As String
    CheckYieldEvidence = CStr(mCheckYieldReached) & "|" & CStr(mCheckYieldAvailable) & "|" & _
        CStr(mCheckYieldLater) & "|" & CStr(mCheckYieldValid)
End Function
Private Sub CheckYieldRevokePermission()
    Dim wb As Workbook, ws As Worksheet, caps As ListObject, candidate As ListObject
    Dim row As ListRow, revoked As Long
    On Error GoTo Failed
    Set wb = Application.Workbooks.Open(mCheckYieldAuthPath, 0, False)
    For Each ws In wb.Worksheets
        For Each candidate In ws.ListObjects
            If candidate.Name = "tblCapabilities" Then Set caps = candidate
        Next candidate
    Next ws
    If caps Is Nothing Then GoTo Failed
    For Each row In caps.ListRows
        If CStr(row.Range.Cells(1, caps.ListColumns("UserId").Index).Value2) = "config-producer" And _
           CStr(row.Range.Cells(1, caps.ListColumns("Capability").Index).Value2) = "PROD_POST" Then
            row.Range.Cells(1, caps.ListColumns("Status").Index).Value2 = "Inactive"
            revoked = revoked + 1
        End If
    Next row
    If revoked <> 1 Then GoTo Failed
    wb.Save
    wb.Close False
    Exit Sub
Failed:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    On Error GoTo 0
    Err.Raise vbObjectError + 261, , "Check In permission interruption fixture unavailable."
End Sub
Public Function CheckYieldOwnerPreserved() As Boolean
    CheckYieldOwnerPreserved = (modProductionReusableRun.RunLocalStateForTest() = mCheckYieldOwner)
End Function
Public Function CheckYieldProjectionPreserved() As Boolean
    CheckYieldProjectionPreserved = (mForm.CheckYieldProjectionForTest() = mCheckYieldProjection)
End Function
Public Sub CheckYieldResetCache()
    mForm.CheckYieldResetCacheForTest
End Sub
'@)
}

function Test-ProductionCheckInYield($Fixture,$Other,$Book,[string]$SelectedKey,[string]$Canary){
    $cases=@(foreach($boundary in @('AvailableQuantity','EntityKind','RunPalette','ManagerCheck')){@{Mode='Reusable';Boundary=$boundary;Ordinal=1}})
    $cases+=@(@{Mode='Worksheet';Boundary='ResolveKey';Ordinal=1},@{Mode='Worksheet';Boundary='ResolveKey';Ordinal=2})
    foreach($boundary in @('InventoryPicker','DefaultLocation','IngredientChoices')){$cases+=@{Mode='Worksheet';Boundary=$boundary;Ordinal=1}}
    $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $authBytes=[IO.File]::ReadAllBytes($authPath);$authHash=(Get-FileHash -LiteralPath $authPath).Hash
    # Temporarily hide only the disposable local projection so both exact-key reads use the real Domain path.
    $local=$Book.Worksheets.Item('InventoryManagement');$local.Name='CheckYieldLocalSource'
    try{
        foreach($interruption in @('SignedOut','Permission')){foreach($case in $cases){
          try{
            [void](Probe 'CheckYieldReset')
            SelectTarget $Fixture 'config-producer';[void](Probe 'CheckBaselineReopen' @($Book.Name))
            $ready=if($case.Mode -ceq 'Reusable'){[bool](Probe 'CheckBaselineReusableStage' @('Selected'))}else{[bool](Probe 'CheckBaselineWorksheetStage' @($SelectedKey,$Canary))}
            if(-not $ready){throw 'Check In interruption prerequisite unavailable; not product RED.'}
            if($case.Mode -ceq 'Worksheet'){[void](Probe 'CheckYieldResetCache')}
            $before=@(Get-Slice4beActivityFiles $Fixture);$otherBefore=@(Get-Slice4beActivityFiles $Other)
            [void](Probe 'CheckYieldArm' @($case.Boundary,$case.Ordinal,$interruption,$authPath))
            $returned=[bool](Probe 'CheckBaselineAct' @(''))
            $evidence=([string](Probe 'CheckYieldEvidence')).Split('|')
            if($evidence.Count -ne 4 -or $evidence[0] -cne 'True' -or $evidence[1] -cne 'True' -or $evidence[3] -cne 'True'){throw ('Check In real read-return/interruption fixture unavailable: '+$case.Mode+'/'+$case.Boundary+'; not product RED.')}
            $label='CheckInYield.'+$(if($interruption -ceq 'Permission'){'Permission.'}else{''})+$case.Mode+'.'+$case.Boundary+'.'+$case.Ordinal
            Check ($label+'.ActualHandlerReturned') $returned
            Check ($label+'.RealReadReturned') $true
            Check ($label+'.InterruptionApplied') $true
            Check ($label+'.NoLaterReads') ([int]$evidence[2] -eq 0)
            Check ($label+'.OwnerAtBoundaryPreserved') ([bool](Probe 'CheckYieldOwnerPreserved'))
            Check ($label+'.ProjectionAtBoundaryPreserved') ([bool](Probe 'CheckYieldProjectionPreserved'))
            $refused=if($interruption -ceq 'Permission'){[bool](Probe 'CheckBaselinePermissionRefused')}else{[bool](Probe 'CheckBaselineContextRefused')}
            Check ($label+'.VisibleContextRefusal') $refused
            Check ($label+'.GuardsRestored') ([bool](Probe 'CheckBaselineGuards'))
            Check ($label+'.NoActivityOrRedirectedRecords') ((@(Get-Slice4beActivityFiles $Fixture) -join '|') -ceq ($before -join '|') -and (@(Get-Slice4beActivityFiles $Other) -join '|') -ceq ($otherBefore -join '|'))
          }finally{
            if($interruption -ceq 'Permission'){
                [IO.File]::WriteAllBytes($authPath,$authBytes)
                [void](Run 'invSys.Core.xlam' 'modAuth.ReloadAuth' @($Fixture.Warehouse))
            }
          }
        }}
        Check 'CheckInYield.FixtureAuthorizationRestored' ((Get-FileHash -LiteralPath $authPath).Hash -ceq $authHash)
    }finally{
        [void](Probe 'CheckYieldReset');$local.Name='InventoryManagement'
        SelectTarget $Fixture 'config-producer'
    }
}
