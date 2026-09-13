# Supplemental Core checks accompany, and never replace, the actual handler route.
function Install-Slice4beShippingCatalogProbe {
    $module=$packages['invSys.Core.xlam'].VBProject.VBComponents.Add(1)
    $module.Name='TestShippingCatalog'
    $module.CodeModule.AddFromString(@'
Public Function Definition(ByVal id As String, ByVal version As Long) As String
    Dim value As Object
    Set value = modActivityCatalog.Control(id, version)
    If Not value Is Nothing Then Definition = modTrainingJson.EncodeObject(value)
End Function
Public Function Outcome(ByVal id As String, ByVal code As String) As String
    Dim value As Object
    Set value = modActivityCatalog.Outcome(id, code)
    If Not value Is Nothing Then Outcome = modTrainingJson.EncodeObject(value)
End Function
Public Function Ids(ByVal version As Long) As String
    Dim values As Variant
    values = modActivityCatalog.ControlIds(version)
    If IsArray(values) Then Ids = Join(values, vbLf)
End Function
Public Function References(ByVal id As String, ByVal code As String, ByVal json As String) As Boolean
    Dim values As Collection
    Set values = modActivityReferences.Decode("CATALOG_TEST", id, code, json)
    References = Not values Is Nothing
End Function
Public Function Policy(ByVal id As String) As String
    Dim version As Long, collect As Boolean, visible As Boolean, notice As String, valid As Boolean
    valid = modActivityPolicy.ReadPolicy(modNasConnection.GetCurrentTarget(), id, version, collect, visible, notice)
    Policy = CStr(valid) & "|" & CStr(collect) & "|" & CStr(visible) & "|" & CStr(version)
End Function
'@)
}
function Test-Slice4beShippingCatalogPolicy($Fixture) {
    $bytes=[IO.File]::ReadAllBytes($Fixture.Config);$cfg=$null
    try {
        SelectTarget $Fixture
        $cfg=$excel.Workbooks.Open($Fixture.Config,0,$false)
        [void](Add-ActivityFixtureTable $cfg 'tblEventTrackingPolicies' @('PolicyVersion','SchemaVersion','CatalogVersion','CreatedAtUTC','CreatedByUserId','DefaultView','ViewerActionPathCaptureEnabled','AdminViewerEventLoggingEnabled','Operator Extra') @(,@(1.0,1.0,7.0,'2026-09-07T12:00:00.000Z','config-admin','How-To',$false,$true,'preserve')))
        $ids=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(7))).Split("`n"))
        $rows=@(foreach($id in $ids){,@(1.0,$id,$true,$true,$true,'preserve')})
        [void](Add-ActivityFixtureTable $cfg 'tblEventTrackingControls' @('PolicyVersion','ControlId','Collect','Visible','SequenceEligible','Operator Extra') $rows)
        $cfg.Save();$cfg.Close($false);$cfg=$null
        $hash=(Get-FileHash -LiteralPath $Fixture.Config).Hash
        Check 'Shipping.Catalog.Policy.SevenStillValid' ([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Policy' @('ADMIN_SETTINGS_SAVE_VALUE')) -ceq 'True|True|True|1')
        foreach($action in @('ADD','UPDATE','REMOVE','HOLD','RETURN','STAGE','SEND')){
            Check ('Shipping.Catalog.Policy.SevenExcludes.'+$action) ([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Policy' @('SHIPPING_'+$action)) -ceq 'False|False|False|0')
        }
        Check 'Shipping.Catalog.Policy.ReadPreservesBytes' ($hash -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    }finally{
        if($null -ne $cfg){$cfg.Close($false)}
        [IO.File]::WriteAllBytes($Fixture.Config,$bytes)
        [void](Run 'invSys.Core.xlam' 'modConfig.LoadConfig' @($Fixture.Warehouse,'S1'))
    }
}
function Test-Slice4beShippingCatalog {
    $controls=[ordered]@{ADD='Add';UPDATE='Update Row';REMOVE='Remove';HOLD='Send Hold';RETURN='Return';STAGE='To Shipments';SEND='Shipments Sent'}
    $old=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(7))).Split("`n")|Where-Object {$_ -ne ''})
    $new=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(8))).Split("`n")|Where-Object {$_ -ne ''})
    Check 'Shipping.Catalog.EightExtendsSeven' ($old.Count -eq 24 -and $new.Count -eq 31 -and @($old|Where-Object {$_ -cnotin $new}).Count -eq 0 -and @($new|Select-Object -Unique).Count -eq 31)
    foreach($id in $old){
        $prior=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,7))
        $latest=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,8))
        Check ('Shipping.Catalog.Preserve.'+$id) ($prior -ne '' -and $prior -ceq $latest)
    }
    $submitted='[{"WarehouseId":"CATALOG_TEST","SourceKind":"Inventory","EventId":"Source_A","SubmissionState":"Submitted"}]'
    $unknown=$submitted.Replace('Submitted','Unknown')
    $mixed=$submitted.Substring(0,$submitted.Length-1)+','+$unknown.Substring(1).Replace('Source_A','Source_B')
    $invalid=@{
        Duplicate=$submitted.Substring(0,$submitted.Length-1)+','+$submitted.Substring(1)
        CrossWarehouse=$submitted.Replace('CATALOG_TEST','ANOTHER_TEST')
        InvalidIdentity=$submitted.Replace('Source_A','Source A')
        UnknownField=$submitted.Replace('"EventId":','"Extra":"forbidden","EventId":')
        UnknownSource=$submitted.Replace('Inventory','Other')
        UnknownState=$submitted.Replace('Submitted','Applied')
    }
    foreach($action in $controls.Keys){
        $id='SHIPPING_'+$action;$prefix='Shipping.Catalog.'+$action
        $json=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,8))
        $definition=if($json -ne ''){$json|ConvertFrom-Json}else{$null}
        Check ($prefix+'.Definition') ($null -ne $definition -and $definition.ControlId -ceq $id -and $definition.OwnerId -ceq 'SHIPPING_WORKFLOW' -and $definition.Class -ceq 'Command' -and $definition.Role -ceq 'Shipping' -and $definition.Caption -ceq $controls[$action] -and $definition.Surface -ceq 'Operations > Shipping' -and $definition.Capability -ceq 'SHIP_POST' -and $definition.CodePrefix -ceq ($id+'_'))
        $excluded=$true
        foreach($version in 1..7){$excluded=$excluded -and [string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,$version)) -ceq ''}
        Check ($prefix+'.OlderVersionsExclude') $excluded
        $outcomes=[ordered]@{REQUESTED=@('Info','Unknown');DENIED=@('Blocked','Unchanged');REJECTED=@('Warning','Unchanged');FAILED=@('Error','Unknown')}
        if($action -ne 'SEND'){$outcomes.Add('STAGED',@('Info','Changed'))}
        if($action -notin @('HOLD','RETURN')){$outcomes.Add('PENDING',@('Notice','Unknown'))}
        if($action -eq 'SEND'){$outcomes.Add('CONFIRMED',@('Info','Unknown'))}
        foreach($code in @('REQUESTED','DENIED','REJECTED','FAILED','STAGED','PENDING','CONFIRMED','COMPLETED')){
            $json=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Outcome' @($id,$code))
            if($outcomes.Contains($code)){
                $value=if($json -ne ''){$json|ConvertFrom-Json}else{$null}
                Check ($prefix+'.Outcome.'+$code) ($null -ne $value -and $value.OutcomeCode -ceq $code -and $value.EventCode -ceq ($id+'_'+$code) -and $value.Severity -ceq $outcomes[$code][0] -and $value.DataEffect -ceq $outcomes[$code][1] -and $value.UserMessage -ne '' -and $null -ne $value.PSObject.Properties['NextStep'])
            }else{Check ($prefix+'.Outcome.'+$code) ($json -ceq '')}
        }
        foreach($code in @('REQUESTED','DENIED','REJECTED')){
            Check ($prefix+'.References.Empty.'+$code) ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @($id,$code,'[]')))
            Check ($prefix+'.References.RejectSource.'+$code) (-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @($id,$code,$submitted)))
        }
        if($action -ne 'SEND'){
            Check ($prefix+'.References.StagedEmpty') ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @($id,'STAGED','[]')))
            Check ($prefix+'.References.StagedNoSources') (-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @($id,'STAGED',$submitted)))
        }
        if($action -notin @('HOLD','RETURN')){
            Check ($prefix+'.References.PendingSubmitted') ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @($id,'PENDING',$submitted)))
            Check ($prefix+'.References.PendingNotEmpty') (-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @($id,'PENDING','[]')))
            Check ($prefix+'.References.PendingNotUnknown') (-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @($id,'PENDING',$unknown)))
            Check ($prefix+'.References.FailedMixed') ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @($id,'FAILED',$mixed)))
            foreach($kind in $invalid.Keys){Check ($prefix+'.References.Reject'+$kind) (-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @($id,'FAILED',$invalid[$kind])))}
        }else{
            Check ($prefix+'.References.FailedNoSources') (-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @($id,'FAILED',$submitted)))
        }
    }
    Check 'Shipping.Catalog.SEND.References.ConfirmedSubmitted' ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @('SHIPPING_SEND','CONFIRMED',$submitted)))
    Check 'Shipping.Catalog.SEND.References.ConfirmedNotEmpty' (-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @('SHIPPING_SEND','CONFIRMED','[]')))
    Check 'Shipping.Catalog.SEND.References.ConfirmedNotUnknown' (-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @('SHIPPING_SEND','CONFIRMED',$unknown)))
}
