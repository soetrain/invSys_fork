# Packaged callback/form-handler diagnostic. Only generated warehouses are written.
# Test adapters change modal presentation and expose fixed facts, not owner results.
. (Join-Path $PSScriptRoot 'Slice4beAdminCreateAuthority.ps1')
function Install-AdminCreateWarehouseOwnerProbe {
    $project=$packages['invSys.Admin.xlam'].VBProject
    $control=$project.VBComponents.Add(2);$control.Name='TestCreateRibbon'
    $control.CodeModule.AddFromString(@'
Option Explicit
Implements Office.IRibbonControl
Private Property Get IRibbonControl_Id() As String
    IRibbonControl_Id = "btnAdminCreateWarehouse"
End Property
Private Property Get IRibbonControl_Tag() As String
    IRibbonControl_Tag = ""
End Property
Private Property Get IRibbonControl_Context() As Object
    Set IRibbonControl_Context = Application.ActiveWindow
End Property
'@)
    $probe=$project.VBComponents.Add(1);$probe.Name='TestCreateOwner'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Public Function OpenForm(ByVal hostName As String) As Boolean
    Dim control As New TestCreateRibbon, item As Object
    Application.Workbooks(hostName).Activate
    TestCreateCalls.Reset
    RibbonOnActionAdmin control
    DoEvents ' Modeless adapter must deliver real activation before inspecting observations.
    For Each item In VBA.UserForms
        If TypeName(item) = "frmCreateWarehouse" Then OpenForm = item.Visible
    Next item
End Function
Public Sub Prepare(ByVal warehouseId As String, ByVal root As String, ByVal share As String, ByVal purpose As String)
    frmCreateWarehouse.OwnerPrepareForTest warehouseId, root, share, purpose
End Sub
Public Sub Create()
    frmCreateWarehouse.OwnerCreateForTest
End Sub
Public Sub CloseForm()
    Dim item As Object
    For Each item In VBA.UserForms
        If TypeName(item) = "frmCreateWarehouse" Then item.OwnerCancelForTest: Exit Sub
    Next item
End Sub
Public Function Facts() As String
    Facts = frmCreateWarehouse.OwnerFactsForTest & "|" & CStr(TestCreateCalls.Bootstraps()) & "|" & CStr(TestCreateCalls.Publications())
End Function
Public Sub SetShare(ByVal share As String)
    frmCreateWarehouse.OwnerShareForTest share
End Sub
Public Function OpenReceiving() As Boolean
    OpenReceiving = frmCreateWarehouse.OwnerReceivingForTest
End Function
'@)
    $form=$project.VBComponents.Item('frmCreateWarehouse').CodeModule
    $form.AddFromString(@'
Public Sub OwnerPrepareForTest(ByVal warehouseId As String, ByVal root As String, ByVal share As String, ByVal purpose As String)
    mFormBusy = True
    Me.txtWarehouseId.Value = warehouseId
    Me.txtWarehouseName.Value = "Generated Create coverage"
    Me.txtStationId.Value = "S1"
    Me.txtAdminUser.Value = "create-fixture-admin"
    Me.txtPathLocal.Value = root
    Me.txtPathSharePoint.Value = share
    Me.chkPublishInitial.Value = (share <> "")
    mWarehousePurpose.Value = purpose
    mFormBusy = False
End Sub
Public Sub OwnerShareForTest(ByVal share As String)
    Me.txtPathSharePoint.Value = share
    Me.chkPublishInitial.Value = (share <> "")
End Sub
Public Sub OwnerCreateForTest()
    btnOK_Click
End Sub
Public Sub OwnerCancelForTest()
    btnCancel_Click
End Sub
Public Function OwnerFactsForTest() As String
    OwnerFactsForTest = CStr(mLocalBootstrapComplete) & "|" & CStr(Me.Tag = "COMPLETE") & "|" & Me.btnOK.Caption & "|" & CStr(mBtnOpenReceiving.Visible)
End Function
Public Function OwnerReceivingForTest() As Boolean
    mBtnOpenReceiving_Click
    OwnerReceivingForTest = (StrComp(Application.ActiveWorkbook.FullName, mCreatedOperatorWorkbookPath, vbBinaryCompare) = 0)
End Function
'@)
    $entryName='modAdmin';$procedure='Open_CreateWarehouse'
    if(@($project.VBComponents|Where-Object Name -CEQ 'modAdminCreateEntry').Count){$entryName='modAdminCreateEntry';$procedure='OpenFromRibbon'}
    $owner=$project.VBComponents.Item($entryName).CodeModule
    $start=$owner.ProcStartLine($procedure,0);$count=$owner.ProcCountLines($procedure,0)
    $source=$owner.Lines($start,$count)
    if([regex]::Matches($source,'frmCreateWarehouse.Show\b').Count -ne 1){throw 'Create callback presentation boundary changed.'}
    $source=$source.Replace('frmCreateWarehouse.Show','frmCreateWarehouse.Show vbModeless')
    $owner.DeleteLines($start,$count);$owner.InsertLines($start,$source)
    $probe.CodeModule.InsertLines($probe.CodeModule.CountOfDeclarationLines+1,'Private mTrackingNotices As Long')
    $probe.CodeModule.AddFromString(@'
Public Sub ObserveTrackingNotice(ByVal text As String)
    If Left$(text, 21) = "Tracking unavailable:" Then mTrackingNotices = mTrackingNotices + 1
End Sub
Public Function TrackingNotices() As Long
    TrackingNotices = mTrackingNotices
End Function
Public Sub ResetTrackingNotices()
    mTrackingNotices = 0
End Sub
'@)
    $observerName=if(@($project.VBComponents|Where-Object Name -CEQ 'cProvisioningActivity').Count){'cProvisioningActivity'}else{'cCreateWarehouseActivity'}
    if(@($project.VBComponents|Where-Object Name -CEQ $observerName).Count){
        $project.VBComponents.Item($observerName).CodeModule.AddFromString(@'
Private Function MsgBox(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String = "") As VbMsgBoxResult
    TestCreateOwner.ObserveTrackingNotice Prompt
    MsgBox = vbOK
End Function
'@)
    }
    Install-CreateAuthorityProbe $project $form $probe.CodeModule
    if($CheckInitialAdminOnboarding){
        . (Join-Path $PSScriptRoot 'Slice4beInitialAdminOnboarding.ps1')
        Install-InitialAdminFailureProbe
    }
}

function Test-AdminCreateWarehouseOwner($Fixture,$Other) {
    function Invoke-Create([string]$Name,[object[]]$Values=@()){Run 'invSys.Admin.xlam' ('TestCreateOwner.'+$Name) $Values}
    function New-CreateCase([string]$Label,[string]$Purpose='Training',[bool]$Publish=$false){
        SelectTarget $Fixture
        if(-not [bool](Invoke-Create 'OpenForm' @($hostBook.Name))){throw 'Authorized real Create callback did not open; not behavioral RED.'}
        $id='CW'+[guid]::NewGuid().ToString('N').Substring(0,8).ToUpperInvariant()
        $root=Join-Path $runRoot ('create-'+$Label)
        $share=if($Publish){Join-Path $runRoot ('publish-'+$Label)}else{''}
        if($Publish){New-Item -ItemType Directory -Path $share|Out-Null}
        [void](Invoke-Create 'Prepare' @($id,$root,$share,$Purpose))
        [pscustomobject]@{Id=$id;Root=$root;Share=$share;Config=(Join-Path $root ($id+'.invSys.Config.xlsb'))}
    }
    function Assert-CreateArtifacts($Case,[string]$Prefix){
        foreach($suffix in @('.invSys.Config.xlsb','.invSys.Auth.xlsb','.invSys.Data.Inventory.xlsb')){
            Check ($Prefix+'.Artifact'+$suffix) (Test-Path -LiteralPath (Join-Path $Case.Root ($Case.Id+$suffix)))
        }
        Check ($Prefix+'.Inbox') (Test-Path -LiteralPath (Join-Path $Case.Root 'inbox'))
    }
    $hostBook=$excel.Workbooks.Add()
    $sourceBefore=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    $otherBefore=(Get-FileHash -LiteralPath $Other.Config).Hash
    try {
        if($CheckInitialAdminOnboarding){
            . (Join-Path $PSScriptRoot 'Slice4beInitialAdminOnboarding.ps1')
            Test-InitialAdminOnboarding
            return
        }
        foreach($purpose in @('Operational','Training')){
            $case=New-CreateCase $purpose $purpose
            [void](Invoke-Create 'Create')
            if($script:CreateCoreAvailable){
                Check ('Create.'+$purpose+'.SourceSessionPreserved') (([string](Run 'invSys.Core.xlam' 'TestCreateCalls.ContextFacts')) -ceq 'SignedIn=True|VersionChanged=False|TargetMissing=False|RootChanged=False|HubChanged=False|StationChanged=False')
            }
            $facts=[string](Invoke-Create 'Facts')
            Check ('Create.'+$purpose+'.LocalOnlyCompleted') ($facts -ceq 'True|True|Close|True|1|0')
            if($facts -cne 'True|True|Close|True|1|0'){throw 'Real local creation unavailable; no valid owner baseline.'}
            Assert-CreateArtifacts $case ('Create.'+$purpose)
            $config=$excel.Workbooks.Open($case.Config,0,$true)
            try{$saved=Table $config 'tblWarehouseConfig';Check ('Create.'+$purpose+'.SavedPurpose') ([string]$saved.ListColumns.Item('WarehousePurpose').DataBodyRange.Cells.Item(1,1).Value2 -ceq $purpose)}finally{$config.Close($false)}
            Check ('Create.'+$purpose+'.OpenReceivingExactCreatedWorkbook') ([bool](Invoke-Create 'OpenReceiving'))
            $excel.ActiveWorkbook.Close($false)
            [void](Invoke-Create 'CloseForm')
            $before=(Get-FileHash -LiteralPath $case.Config).Hash
            SelectTarget $Fixture
            [void](Invoke-Create 'OpenForm' @($hostBook.Name))
            [void](Invoke-Create 'Prepare' @($case.Id,$case.Root,'',$purpose))
            [void](Invoke-Create 'Create')
            Check ('Create.'+$purpose+'.DuplicateRejectedBeforeOwner') ([string](Invoke-Create 'Facts') -ceq 'False|False|Create|False|0|0')
            [void](Invoke-Create 'CloseForm')
            Check ('Create.'+$purpose+'.DuplicatePreservedBytes') ((Get-FileHash -LiteralPath $case.Config).Hash -ceq $before)
        }
        $case=New-CreateCase 'invalid'
        [void](Invoke-Create 'Prepare' @('invalid/id',$case.Root,'','Training'))
        [void](Invoke-Create 'Create')
        Check 'Create.Invalid.NoOwnerCall' ([string](Invoke-Create 'Facts') -ceq 'False|False|Create|False|0|0')
        Check 'Create.Invalid.NoRuntime' (-not (Test-Path -LiteralPath $case.Root))
        [void](Invoke-Create 'CloseForm')
        $case=New-CreateCase 'cancel'
        [void](Invoke-Create 'CloseForm')
        Check 'Create.Cancel.NoRuntime' (-not (Test-Path -LiteralPath $case.Root))

        foreach($failFirst in @($false,$true)){
            $label=if($failFirst){'retry'}else{'published'}
            $case=New-CreateCase $label 'Training' $true
            # A regular generated file blocks only the first publication folder.
            # No permission changes or production faults are used.
            if($failFirst){[IO.File]::WriteAllText((Join-Path $case.Share $case.Id),'Generated publication obstruction')}
            [void](Invoke-Create 'Create')
            Assert-CreateArtifacts $case ('Create.'+$label)
            if($failFirst){
                Check 'Create.Partial.LocalSuccessPublishFailure' ([string](Invoke-Create 'Facts') -ceq 'True|False|Retry Publish|True|1|1')
                $before=(Get-FileHash -LiteralPath $case.Config).Hash
                $retryShare=Join-Path $runRoot 'publish-retry-success'
                New-Item -ItemType Directory -Path $retryShare|Out-Null
                [void](Invoke-Create 'SetShare' @($retryShare))
                [void](Invoke-Create 'Create')
                Check 'Create.Retry.PublishesWithoutRebootstrap' ([string](Invoke-Create 'Facts') -ceq 'True|True|Close|True|1|2')
                Check 'Create.Retry.LocalConfigBytesPreserved' ((Get-FileHash -LiteralPath $case.Config).Hash -ceq $before)
                $case.Share=$retryShare
            }else{Check 'Create.Publish.BothCompleted' ([string](Invoke-Create 'Facts') -ceq 'True|True|Close|True|1|1')}
            Check ('Create.'+$label+'.PublishedConfig') (Test-Path -LiteralPath (Join-Path $case.Share ($case.Id+'\'+$case.Id+'.invSys.Config.xlsb')))
            Check ('Create.'+$label+'.PublishedDiscovery') (Test-Path -LiteralPath (Join-Path $case.Share ($case.Id+'.config.json')))
            [void](Invoke-Create 'CloseForm')
        }
        # Open through the authorized Ribbon, then invalidate the live actor.
        # The existing signed-in ADMIN_MAINT rule must hold at the write boundary.
        $case=New-CreateCase 'signed-out'
        [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
        Check 'Create.SignedOut.FixtureActuallySignedOut' (-not [bool](Run 'invSys.Core.xlam' 'modAuth.IsSignedIn'))
        [void](Invoke-Create 'Create')
        Check 'Create.SignedOut.DoesNotCreateRuntime' (-not (Test-Path -LiteralPath $case.Config))
        Check 'Create.SignedOut.NoCompletionClaim' (-not ([string](Invoke-Create 'Facts')).StartsWith('True|True|'))
        [void](Invoke-Create 'CloseForm')
        Check 'Create.SourceConfigBytesPreserved' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $sourceBefore)
        Check 'Create.OtherConfigBytesPreserved' ((Get-FileHash -LiteralPath $Other.Config).Hash -ceq $otherBefore)
        Test-CreateAuthority $Fixture $Other $hostBook
    } finally {
        [void](Invoke-Create 'CloseForm')
        $hostBook.Close($false)
    }
}
