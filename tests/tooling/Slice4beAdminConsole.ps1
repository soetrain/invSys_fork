# Real generated Ribbon and owner observations; no fabricated successful opening.
. (Join-Path $PSScriptRoot 'Slice4beActivityAssertions.ps1')
. (Join-Path $PSScriptRoot 'Slice4beAdminConsoleBoundaries.ps1')
. (Join-Path $PSScriptRoot 'Slice4beAdminConsoleGuide.ps1')
function Install-AdminConsoleProbe {
    $admin=$packages['invSys.Admin.xlam'].VBProject
    $control=$admin.VBComponents.Add(2);$control.Name='TestAdminConsoleRibbon'
    $control.CodeModule.AddFromString(@'
Option Explicit
Implements Office.IRibbonControl
Private Property Get IRibbonControl_Id() As String
    IRibbonControl_Id = "btnAdminOpen"
End Property
Private Property Get IRibbonControl_Tag() As String
    IRibbonControl_Tag = ""
End Property
Private Property Get IRibbonControl_Context() As Object
    Set IRibbonControl_Context = Application.ActiveWindow
End Property
'@)
    $probe=$admin.VBComponents.Add(1);$probe.Name='TestAdminConsole'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mCalls As Long, mOpened As Boolean, mNotices As Long, mUnexpected As Long
Private mHost As Workbook, mPresentedAtOwner As Boolean
Public Sub OwnerEntered()
    mCalls = mCalls + 1
End Sub
Public Sub OwnerOpened()
    mOpened = True
    mPresentedAtOwner = (Application.ActiveWorkbook Is mHost)
    If mPresentedAtOwner Then mPresentedAtOwner = (Application.ActiveSheet.Name = "AdminConsole")
End Sub
Public Function ObserveMessage(ByVal value As String) As VbMsgBoxResult
    If Left$(value, 21) = "Tracking unavailable:" Then
        mNotices = mNotices + 1
    Else
        mUnexpected = mUnexpected + 1
    End If
    ObserveMessage = vbOK
End Function
Public Function Invoke(ByVal hostName As String) As String
    Dim control As New TestAdminConsoleRibbon, wb As Workbook, selected As Boolean, bound As Boolean, originalSurface As Boolean
    Set wb = Application.Workbooks(hostName)
    Set mHost = wb: mPresentedAtOwner = False
    wb.Activate: wb.Worksheets("UserCredentials").Activate
    bound = (modAdmin.ResolveInteractiveAdminWorkbook() Is wb)
    If Not bound Then Err.Raise 5, , "Generated callback host is not bound"
    mCalls = 0: mOpened = False: mNotices = 0: mUnexpected = 0
    RibbonOnActionAdmin control
    selected = (Application.ActiveWorkbook Is wb)
    originalSurface = selected
    If originalSurface Then originalSurface = (Application.ActiveSheet.Name = "UserCredentials")
    If selected Then selected = (Application.ActiveSheet.Name = "AdminConsole")
    Invoke = CStr(mCalls) & "|" & CStr(mOpened) & "|" & CStr(selected) & "|" & CStr(mNotices) & "|" & CStr(mUnexpected) & "|" & CStr(mPresentedAtOwner) & "|" & CStr(originalSurface)
    Set mHost = Nothing
End Function
Public Function Prepare(ByVal name As String) As Boolean
    Dim wb As Workbook, table As ListObject, column As ListColumn, report As String
    Set wb = Application.Workbooks(name)
    If Not modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(wb, report) Then Exit Function
    If Not modAdminConsole.EnsureAdminSchema(wb, report) Then Exit Function
    Set table = wb.Worksheets("UserCredentials").ListObjects("UserCredentials")
    Set column = table.ListColumns.Add: column.Name = "Operator_Custom_Canary"
    If table.ListRows.Count = 0 Then table.ListRows.Add
    column.DataBodyRange.Cells(1, 1).Value2 = "Preserve this custom value"
    Prepare = True
End Function
Public Function HostPreserved(ByVal name As String) As Boolean
    Dim table As ListObject
    Set table = Application.Workbooks(name).Worksheets("UserCredentials").ListObjects("UserCredentials")
    HostPreserved = (table.ListColumns("Operator_Custom_Canary").DataBodyRange.Cells(1, 1).Value2 = "Preserve this custom value")
End Function
Public Sub ProtectConsole(ByVal name As String, ByVal value As Boolean)
    Dim ws As Worksheet
    Set ws = Application.Workbooks(name).Worksheets("AdminConsole")
    If value Then ws.Protect Else ws.Unprotect
End Sub
'@)
    $newOwner=@($admin.VBComponents|Where-Object Name -CEQ 'modAdminConsoleActivity')
    $owner=if($newOwner.Count){$newOwner[0].CodeModule}else{$admin.VBComponents.Item('modAdminConsole').CodeModule}
    $method=if($newOwner.Count){'OpenConsoleView'}else{'OpenAdminConsole'}
    $resolve=if($newOwner.Count){'    Set wb = modAdminWorkbookTarget.ResolveAdminTargetWorkbook(adminWb, ThisWorkbook, True)'}else{'    Set wb = ResolveAdminWorkbook(adminWb)'}
    $start=$owner.ProcStartLine($method,0);$count=$owner.ProcCountLines($method,0)
    $source=$owner.Lines($start,$count)
    $changed=$source.Replace($resolve,"    TestAdminConsole.OwnerEntered`r`n"+$resolve).Replace(('    '+$method+' = True'),('    '+$method+" = True`r`n    TestAdminConsole.OwnerOpened"))
    if($changed -ceq $source){throw 'Console owner observation boundary unavailable; not product RED.'}
    $owner.DeleteLines($start,$count);$owner.InsertLines($start,$changed)
    $ribbon=@($admin.VBComponents|Where-Object {$_.Type -eq 1 -and $_.CodeModule.Lines(1,$_.CodeModule.CountOfLines).Contains('Public Sub RibbonOnActionAdmin(')})
    if($ribbon.Count -ne 1){throw 'One generated Admin Ribbon module required.'}
    $messages=@($admin.VBComponents|Where-Object Name -CEQ 'modAdminConsoleActivity')
    $presentation=if($messages.Count){$messages[0].CodeModule}else{$admin.VBComponents.Item('modAdmin').CodeModule}
    foreach($module in @($presentation,$ribbon[0].CodeModule)){$module.AddFromString(@'
Private Function MsgBox(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String = "") As VbMsgBoxResult
    MsgBox = TestAdminConsole.ObserveMessage(Prompt)
End Function
'@)}
    $core=$packages['invSys.Core.xlam'].VBProject
    $gate=$core.VBComponents.Add(1);$gate.Name='TestConsoleActivity'
    $gate.CodeModule.AddFromString(@'
Option Explicit
Public DenialMessages As Long
Public Function SetPolicy(ByVal mode As String) As Boolean
    Dim context As String, version As Long, request As String, notice As String
    Dim model As Object, row As Variant, users As Collection, user As Object
    context = modActivity.CaptureContext()
    If Not modTrackingPolicySettings.ReadEditor(context, version, request, notice) Then Exit Function
    Set model = modTrackingPolicyModel.Defaults(mode <> "Disabled")
    model("ViewerActionPathCaptureEnabled") = (mode = "Capture")
    For Each row In model("Controls")
        If row("ControlId") = "ADMIN_CONSOLE_OPEN" Then row("Collect") = (mode <> "Disabled" And mode <> "ControlOff")
    Next row
    If mode = "UserOff" Then
        Set user = CreateObject("Scripting.Dictionary")
        user.Add "UserId", "config-admin": user.Add "Record", False
        Set users = model("Users"): users.Add user
    End If
    SetPolicy = modTrackingPolicySettings.SavePolicy(context, version, modTrainingJson.EncodeObject(model), notice)
End Function
Public Function Definition(ByVal version As Long) As String
    Dim value As Object
    Set value = modActivityCatalog.Control("ADMIN_CONSOLE_OPEN", version)
    If Not value Is Nothing Then Definition = modTrainingJson.EncodeObject(value)
End Function
Public Function CurrentCatalogVersion() As Long
    CurrentCatalogVersion = modActivityCatalog.CATALOG_VERSION
End Function
Public Function PriorDefinitionsPreserved() As Boolean
    Dim id As Variant, before As Object, after As Object
    For Each id In modActivityCatalog.ControlIds(35)
        Set before = modActivityCatalog.Control(CStr(id), 35)
        Set after = modActivityCatalog.Control(CStr(id))
        If before Is Nothing Or after Is Nothing Then Exit Function
        If modTrainingJson.EncodeObject(before) <> modTrainingJson.EncodeObject(after) Then Exit Function
    Next id
    PriorDefinitionsPreserved = True
End Function
Public Function Concludes(ByVal wire As String) As Boolean
    Concludes = modEvaluationMatches.CommandCompleted(modTrainingJson.DecodeObject(wire))
End Function
Public Function DenialCount() As Long
    DenialCount = DenialMessages
End Function
'@)
    if($newOwner.Count){Install-AdminConsoleBoundaries $gate.CodeModule $probe.CodeModule $admin}
    if($AdminConsoleGuidePaths){
        . (Join-Path $PSScriptRoot 'Slice4beAdminVerifyGuide.ps1')
        Install-AdminVerifyGuideProbe
    }
    $core.VBComponents.Item('modRoleUiAccess').CodeModule.AddFromString(@'
Private Function MsgBox(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String = "") As VbMsgBoxResult
    TestConsoleActivity.DenialMessages = TestConsoleActivity.DenialMessages + 1
    MsgBox = vbOK
End Function
'@)
}

function Test-AdminConsole($Fixture,$Other) {
    function Probe([string]$Name,[object[]]$Values=@()){Run 'invSys.Core.xlam' ('TestConsoleActivity.'+$Name) $Values}
    $catalogVersion=[int](Probe 'CurrentCatalogVersion')
    if($catalogVersion -lt 36){throw 'Packaged activity catalog predates Admin Console; not behavioral RED.'}
    function FileHash([string]$Path){$stream=[IO.File]::Open($Path,'Open','Read','ReadWrite');try{(Get-FileHash -InputStream $stream).Hash}finally{$stream.Dispose()}}
    function Pins([string]$Root){$values=@{};Get-ChildItem -LiteralPath $Root -Recurse -File|ForEach-Object {$values[$_.FullName]=FileHash $_.FullName};return $values}
    function Retained($Values){foreach($path in $Values.Keys){if(-not(Test-Path -LiteralPath $path) -or (FileHash $path) -cne $Values[$path]){return $false}};return $true}
    SelectTarget $Fixture
    $hostBook=$excel.Workbooks.Add();$otherBook=$excel.Workbooks.Add();$otherCount=$otherBook.Worksheets.Count
    $otherBook.Worksheets.Item(1).Cells.Item(1,1).Value2='Other workbook canary'
    $surfaceResults=@()
    try{
        if(-not [bool](Run 'invSys.Admin.xlam' 'TestAdminConsole.Prepare' @($hostBook.Name))){throw 'Console/custom-column fixture not prepared; not product RED.'}
        Check 'AdminConsole.Fixture.CustomColumnExists' ([bool](Run 'invSys.Admin.xlam' 'TestAdminConsole.HostPreserved' @($hostBook.Name)))
        $wire=[string](Probe 'Definition' @(36));$definition=if($wire){$wire|ConvertFrom-Json}else{$null}
        Check 'AdminConsole.Catalog.RegisteredOwnerAndCapability' ($null -ne $definition -and $definition.OwnerId -ceq 'ADMIN_CONSOLE' -and $definition.Class -ceq 'Navigation' -and $definition.Capability -ceq 'ADMIN_MAINT')
        Check 'AdminConsole.Catalog.AbsentFrom35' ([string](Probe 'Definition' @(35)) -ceq '')
        Check 'AdminConsole.Catalog.PriorDefinitionsPreserved' ([bool](Probe 'PriorDefinitionsPreserved'))
        foreach($case in @('Opened','Protected','Disabled','ControlOff','UserOff','Reader','SignedOut')){
            SelectTarget $Fixture
            $mode=if($case -cin @('Disabled','ControlOff','UserOff')){$case}else{'Enabled'}
            if(-not [bool](Probe 'SetPolicy' @($mode))){throw 'Console policy prerequisite failed.'}
            $authority=Pins $Fixture.Root;$otherAuthority=Pins $Other.Root
            if($case -ceq 'Reader'){SelectTarget $Fixture 'config-reader'}
            if($case -ceq 'SignedOut'){[void](Run 'invSys.Core.xlam' 'modAuth.SignOut')}
            $denied=$case -cin @('Reader','SignedOut');$beforeNotice=[int](Probe 'DenialCount')
            $before=@(Get-Slice4beActivityFiles $Fixture)
            [void](Run 'invSys.Admin.xlam' 'TestAdminConsole.ProtectConsole' @($hostBook.Name,($case -ceq 'Protected')))
            try{$values=([string](Run 'invSys.Admin.xlam' 'TestAdminConsole.Invoke' @($hostBook.Name))) -split '\|'}
            finally{[void](Run 'invSys.Admin.xlam' 'TestAdminConsole.ProtectConsole' @($hostBook.Name,$false))}
            if($values.Count -ne 7){throw 'Console callback observation missing; not product RED.'}
            $prefix='AdminConsole.'+$case+'.';$opened=-not $denied -and $case -cne 'Protected'
            Write-Output ('Console focus '+$case+': owner='+$values[1]+', presented='+$values[5]+', returned='+$values[2])
            $surfaceResults+=[pscustomobject]@{Case=$case;OwnerInvoked=(-not $denied);Passed=if($opened){$values[2] -ceq 'True'}else{$values[6] -ceq 'True'}}
            Check ($prefix+'OwnerInvocationCount') ($values[0] -ceq $(if($denied){'0'}else{'1'}))
            Check ($prefix+'ActualOwnerPresentation') ($values[1] -ceq [string]$opened -and $values[5] -ceq [string]$opened -and $values[2] -ceq [string]$opened)
            Check ($prefix+'NoUnexpectedMessage') ($values[3] -ceq '0' -and $values[4] -ceq '0')
            Check ($prefix+'ActualPermissionGate') (([int](Probe 'DenialCount')-$beforeNotice) -eq $(if($denied){1}else{0}))
            Check ($prefix+'OriginalAndOtherWorkbookPreserved') ([bool](Run 'invSys.Admin.xlam' 'TestAdminConsole.HostPreserved' @($hostBook.Name)) -and $otherBook.Worksheets.Count -eq $otherCount -and [string]$otherBook.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'Other workbook canary')
            Check ($prefix+'SavedAuthorityPreserved') ((Retained $authority) -and (Retained $otherAuthority))
            $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json}|Where-Object ControlId -CEQ 'ADMIN_CONSOLE_OPEN')
            if($denied -or $mode -cne 'Enabled'){Check ($prefix+'NoFabricatedAction') ($records.Count -eq 0);continue}
            $expected=if($opened){'OPENED'}else{'FAILED'}
            $attempt=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$outcome=@($records|Where-Object OutcomeCode -CEQ $expected)
            $pair=$records.Count -eq 2 -and $attempt.Count -eq 1 -and $outcome.Count -eq 1
            Check ($prefix+'ExactRecordedPair') ($pair -and $attempt[0].ActivityId -ceq $outcome[0].ActivityId -and $attempt[0].RecordId -cne $outcome[0].RecordId)
            $safe=$pair;$bound=$pair;$concludes=$pair
            foreach($record in $records){
                $bound=$bound -and $record.OwnerId -ceq 'ADMIN_CONSOLE' -and $record.WarehouseId -ceq $Fixture.Warehouse -and $record.UserId -ceq 'config-admin' -and $record.CatalogVersion -eq $catalogVersion -and $record.DataEffect -ceq 'Unknown'
                $json=$record|ConvertTo-Json -Depth 20 -Compress
                $safe=$safe -and @($record.SourceEventRefs).Count -eq 0
                foreach($value in @($Fixture.Root,$Other.Root,$Fixture.Secret,(CredentialHash $Fixture.Secret),'Operator_Custom_Canary','svc_processor')){$safe=$safe -and -not $json.Contains($value)}
                $concludes=$concludes -and ([bool](Probe 'Concludes' @($json)) -eq ($record.OutcomeCode -ceq 'OPENED'))
            }
            Check ($prefix+'OriginalContextAndHonestEffect') $bound
            Check ($prefix+'NoConsoleValuesOrSourceReferences') $safe
            Check ($prefix+'OnlyOpenedConcludes') $concludes
        }
        foreach($result in $surfaceResults){
            if($result.OwnerInvoked -or $AdminConsoleDenialSurfaceForTest){Check ('AdminConsole.ReturnSurface.'+$result.Case) $result.Passed}
        }
        if([string](Probe 'Definition' @(36)) -cne ''){Test-AdminConsoleBoundaries $Fixture $Other $hostBook $otherBook}
        if($AdminConsoleGuidePaths){Test-AdminConsoleGuide $Fixture $hostBook $catalogVersion}
    }finally{SelectTarget $Fixture;$hostBook.Close($false);$otherBook.Close($false)}
}
