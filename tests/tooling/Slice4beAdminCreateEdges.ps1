function Install-CreateEdgeProbe($Core,$Probe){
    $owner=@($Core.VBComponents|Where-Object Name -ceq 'modWarehouseCreateCommands')
    $script:CreateCoreAvailable=$owner.Count -eq 1
    if(-not $script:CreateCoreAvailable){return}
    $command=$owner[0].CodeModule
    $privateArgs=if($command.Lines(1,$command.CountOfLines) -match 'Public Function MatchingInitialCredential\('){', initialPin:=privateEntry, confirmation:=privateEntry'}else{''}
    $Probe.InsertLines($Probe.CountOfDeclarationLines+1,'Private mContext As String, mInterrupt As Long, mAuthorization As Long, mReadStep As Long, mVersionBefore As Long, mTargetBefore As WarehouseTarget')
    $Probe.AddFromString(@'
Public Sub BeforeBootstrap()
    mVersionBefore = modAuthSession.Version(): Set mTargetBefore = modNasConnection.GetCurrentTarget()
End Sub
Public Function ContextFacts() As String
    Dim target As WarehouseTarget
    Set target = modNasConnection.GetCurrentTarget()
    ContextFacts = "SignedIn=" & CStr(modAuth.IsSignedIn()) & "|VersionChanged=" & CStr(mVersionBefore <> modAuthSession.Version()) & "|TargetMissing=" & CStr(target Is Nothing)
    If target Is Nothing Or mTargetBefore Is Nothing Then Exit Function
    ContextFacts = ContextFacts & "|RootChanged=" & CStr(target.RuntimeRoot <> mTargetBefore.RuntimeRoot) & "|HubChanged=" & CStr(target.HubRoot <> mTargetBefore.HubRoot) & "|StationChanged=" & CStr(target.StationId <> mTargetBefore.StationId)
End Function
Public Sub ReadStep(ByVal step As Long)
    mReadStep = step
End Sub
Public Function LastReadStep() As Long
    LastReadStep = mReadStep
End Function
Public Sub CaptureSession()
    mContext = modAuthUserCommands.CaptureContext(): mInterrupt = 0: mAuthorization = 0
End Sub
Public Sub ArmSignOut(ByVal boundary As Long)
    mInterrupt = boundary: mAuthorization = 0
End Sub
Public Sub BeforeAuthorization()
    mAuthorization = mAuthorization + 1
    If mInterrupt > 0 And mAuthorization = mInterrupt Then modAuth.SignOut
End Sub
Public Function DirectCreate(ByVal root As String, ByVal warehouseId As String) As Boolean
    Dim report As String, operatorPath As String, privateEntry As String
    privateEntry = CreateObject("Scripting.FileSystemObject").GetTempName
    DirectCreate = modWarehouseCreateCommands.CreateWarehouse(mContext, warehouseId, "Generated direct fixture", "S1", "destination-admin", root, "", "Training", report, operatorPathPRIVATE_ARGS)
End Function
Public Function DirectPublish(ByVal root As String, ByVal share As String) As Boolean
    Dim report As String
    DirectPublish = modWarehouseCreateCommands.PublishExisting(mContext, root, share, report)
End Function
'@.Replace('PRIVATE_ARGS',$privateArgs))
    $bootstrapLine=$Probe.ProcStartLine('ObserveBootstrap',0)
    $Probe.InsertLines($bootstrapLine+1,'    BeforeBootstrap')
    $module=$owner[0].CodeModule
    $line=@(for($i=1;$i -le $module.CountOfLines;$i++){if($module.Lines($i,1) -match '^Private Function Authorize\('){$i}})
    if($line.Count -ne 1){throw 'Interactive authorization boundary unavailable.'}
    $module.InsertLines($line[0]+1,'    TestCreateCalls.BeforeAuthorization')
    $authMarkers=@('If context =','Set target =','If target Is Nothing','For Each wb In','If Not wb.Saved','If Not modAuth.LoadAuth','actor =','If Not modAuth.CanPerform','If context <>','Authorize = True')
    for($i=$module.CountOfLines;$i -ge $line[0];$i--){
        $text=$module.Lines($i,1).Trim()
        if(@($authMarkers|Where-Object {$text.StartsWith($_)}).Count){$module.InsertLines($i,('    TestCreateCalls.ReadStep '+(1000+$i)))}
    }
    $reader=$Core.VBComponents.Item('modWarehousePublishSource').CodeModule
    # Fixed line numbers only: diagnose reader exits without exposing field values.
    $markers=@('root = modDeploymentPaths','If root =','fileName = Dir','If fileName =','If Dir$()','If Not wb.Saved','Set wh =','If wh Is Nothing','If wh.ListRows','spec.WarehouseId =','If spec.WarehouseId =','If StrComp(fileName','If StrComp(modDeploymentPaths','For row =','If spec.StationId =','If Not fso.FileExists','If Not fso.FolderExists','ReadExisting = True')
    for($i=$reader.CountOfLines;$i -ge 1;$i--){
        $text=$reader.Lines($i,1).Trim()
        if(@($markers|Where-Object {$text.StartsWith($_)}).Count){$reader.InsertLines($i,('    TestCreateCalls.ReadStep '+$i))}
    }
}

function Test-CreateEdges($Fixture,$Other,$HostBook,$Existing){
    Check 'Create.Core.CommandBoundaryAvailable' $script:CreateCoreAvailable
    if(-not $script:CreateCoreAvailable){return}
    function Core([string]$Name,[object[]]$Values=@()){Run 'invSys.Core.xlam' ('TestCreateCalls.'+$Name) $Values}
    function Owner([string]$Name,[object[]]$Values=@()){Run 'invSys.Admin.xlam' ('TestCreateOwner.'+$Name) $Values}
    SelectTarget $Fixture
    [void](Core 'CaptureSession')
    $id='CD'+[guid]::NewGuid().ToString('N').Substring(0,8).ToUpperInvariant()
    $root=Join-Path $runRoot 'direct-authorized-create'
    Check 'Create.Core.NoDestinationRuntimeBeforeCreate' (-not (Test-Path -LiteralPath $root))
    Check 'Create.Core.SourceAdminCreatesNewDestination' ([bool](Core 'DirectCreate' @($root,$id)))
    SelectTarget $Fixture
    Check 'Create.Core.SourceActorHasNoDestinationGrant' (-not [bool](Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ADMIN_MAINT','config-admin',$id,'S1')))
    [void](Core 'CaptureSession')
    $share=Join-Path $runRoot 'direct-authorized-publish'
    Check 'Create.Core.SourceAdminPublishesExistingDestination' ([bool](Core 'DirectPublish' @($root,$share)))

    [void](Core 'CaptureSession');[void](Core 'ArmSignOut' @(1))
    $denied=Join-Path $runRoot 'direct-denied-create'
    Check 'Create.Core.LateSignOutDeniesCreate' (-not [bool](Core 'DirectCreate' @($denied,'DENIED')))
    Check 'Create.Core.LateSignOutNoRuntime' (-not (Test-Path -LiteralPath $denied))
    # An actor hint is not a live invSys session, even at the direct Core boundary.
    [void](Run 'invSys.Core.xlam' 'modRoleEventWriter.SetCurrentUserId' @('config-admin'))
    [void](Core 'CaptureSession')
    Check 'Create.Core.UserHintCannotAuthorize' (-not [bool](Core 'DirectCreate' @($denied,'DENIED')))

    SelectTarget $Fixture;[void](Core 'CaptureSession');[void](Core 'ArmSignOut' @(2))
    $deniedShare=Join-Path $runRoot 'late-denied-publication'
    Check 'Create.Core.ContextLostAfterReadDeniesPublish' (-not [bool](Core 'DirectPublish' @($Existing.Root,$deniedShare)))
    Check 'Create.Core.ContextLostNoPublishedArtifacts' (-not (Test-Path -LiteralPath $deniedShare))
    SelectTarget $Fixture;[void](Core 'CaptureSession')
    $missing=Join-Path $runRoot 'missing-existing-runtime'
    Check 'Create.Core.MissingRuntimeRefused' (-not [bool](Core 'DirectPublish' @($missing,$deniedShare)))
    Check 'Create.Core.MissingRuntimeNotProvisioned' (-not (Test-Path -LiteralPath $missing))
    $original=[IO.File]::ReadAllBytes($Existing.Config)
    try{
        $book=$excel.Workbooks.Open($Existing.Config,0,$false)
        $table=Table $book 'tblWarehouseConfig'
        $table.ListColumns.Item('WarehouseId').Name='Operator_replaced_identity'
        $book.Save();$book.Close($false)
        $invalidHash=(Get-FileHash -LiteralPath $Existing.Config).Hash
        Check 'Create.Core.InvalidRuntimeRefused' (-not [bool](Core 'DirectPublish' @($Existing.Root,$deniedShare)))
        Check 'Create.Core.InvalidRuntimeNotRepaired' ((Get-FileHash -LiteralPath $Existing.Config).Hash -ceq $invalidHash)
    }finally{[IO.File]::WriteAllBytes($Existing.Config,$original)}
    $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $originalAuth=[IO.File]::ReadAllBytes($authPath)
    try{
        SelectTarget $Fixture
        if(-not [bool](Owner 'OpenForm' @($HostBook.Name))){throw 'Lost-rights form fixture unavailable.'}
        [void](Owner 'Prepare' @('LOSTRIGHTS',(Join-Path $runRoot 'lost-rights'),'','Training'))
        $book=$excel.Workbooks.Open($authPath,0,$false);$caps=Table $book 'tblCapabilities'
        $changed=0
        foreach($row in $caps.ListRows){
            if($row.Range.Cells.Item(1,$caps.ListColumns.Item('UserId').Index).Value2 -ceq 'config-admin' -and $row.Range.Cells.Item(1,$caps.ListColumns.Item('Capability').Index).Value2 -ceq 'ADMIN_MAINT'){
                $row.Range.Cells.Item(1,$caps.ListColumns.Item('Status').Index).Value2='INACTIVE';$changed++
            }
        }
        if(-not $changed){throw 'No generated Admin grant was revoked.'}
        $book.Save();$book.Close($false)
        Check 'Create.Core.StaleGrantFixture' ([bool](Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ADMIN_MAINT','config-admin',$Fixture.Warehouse,'S1')))
        [void](Owner 'Create')
        Check 'Create.Core.SavedRevocationDeniesFormCreate' (-not (Test-Path -LiteralPath (Join-Path $runRoot 'lost-rights')))
        [void](Owner 'CloseForm')
    }finally{[IO.File]::WriteAllBytes($authPath,$originalAuth)}
}
