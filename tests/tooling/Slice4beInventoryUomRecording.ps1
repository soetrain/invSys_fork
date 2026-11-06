# Actual Inventory Ribbon and native UOM Change/Click handlers. The existing
# dialog-answer seam supplies input and a real sign-out/session replacement.
function Install-InventoryUomRecordingProbe {
    $core=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('TestInventoryRecordingPolicy').CodeModule
    $core.AddFromString(@'
Public Function Prior198Preserved() As Boolean
    Dim id As Variant, before As Object, after As Object
    For Each id In modActivityCatalog.ControlIds(45)
        Set before = modActivityCatalog.Control(CStr(id), 45)
        Set after = modActivityCatalog.Control(CStr(id))
        If before Is Nothing Or after Is Nothing Then Exit Function
        If modTrainingJson.EncodeObject(before) <> modTrainingJson.EncodeObject(after) Then Exit Function
    Next id
    Prior198Preserved = True
End Function
Public Function UomConclusion(ByVal code As String) As Boolean
    Dim record As Object
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "ControlId", "ADMIN_INVENTORY_UOM_ADD"
    record.Add "CatalogVersion", 46: record.Add "OwnerId", "ADMIN_INVENTORY"
    record.Add "OutcomeCode", code
    UomConclusion = modEvaluationMatches.CommandCompleted(record)
End Function
'@)
    $admin=$packages['invSys.Admin.xlam'].VBProject
    $probe=$admin.VBComponents.Add(1);$probe.Name='TestInventoryUom'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mMode As String, mSecret As String, mCount As Long, mFacts As String, mYield As Boolean
Private mTrace As String
Public Sub Arm(ByVal mode As String, ByVal secret As String)
    mMode = mode: mSecret = secret: mCount = 0: mFacts = "": mYield = False: mTrace = ""
End Sub
Public Function Answer(ByVal prompt As String, ByVal context As String) As String
    Dim status As String
    If prompt <> "New UOM:" Then Err.Raise vbObjectError + 715, , "Unexpected test dialog"
    mCount = mCount + 1
    If mMode = "ExpiredCommandAtPrompt" Then
        modInventoryCommands.Release context
        mYield = True
    End If
    If mMode = "SignedOutAtPrompt" Or mMode = "ReplacementSessionAtPrompt" Then
        modAuth.SignOut
        mYield = Not modAuth.IsSignedIn()
        If mMode = "ReplacementSessionAtPrompt" Then
            status = modAuth.SignInCurrentTargetForAutomation("config-admin", mSecret, "")
            mYield = mYield And (Left$(status, 3) = "OK|")
        End If
    End If
    If mCount <= 2 Then Answer = "UOMSentinel"
    If mCount = 4 Then Answer = "   "
End Function
Public Function Count() As Long
    Count = mCount
End Function
Public Function Yielded() As Boolean
    Yielded = mYield
End Function
Public Sub Facts(ByVal value As Boolean)
    mFacts = CStr(value)
    mSecret = ""
End Sub
Public Function Result() As String
    Result = mFacts
End Function
Public Sub Trace(ByVal turn As Long, ByVal valueMatches As Boolean, ByVal countMatches As Boolean)
    mTrace = mTrace & CStr(turn) & ":" & CStr(mCount) & ":" & CStr(valueMatches) & ":" & CStr(countMatches) & "|"
End Sub
Public Function Diagnostics() As String
    Diagnostics = mTrace
End Function
'@)
    $driver=$admin.VBComponents.Item('TestInventoryRecording').CodeModule
    $first=$driver.ProcStartLine('Drive',0);$count=$driver.ProcCountLines('Drive',0)
    $body=$driver.Lines($first,$count).Replace('form.RecordingNavigationForTest mMode','form.UomRecordingForTest mMode')
    if($body -notmatch 'UomRecordingForTest'){throw 'UOM gesture boundary unavailable.'}
    $driver.DeleteLines($first,$count);$driver.InsertLines($first,$body)
    $form=$admin.VBComponents.Item('frmAddInventoryItem').CodeModule
    $first=$form.ProcStartLine('InputBox',0);$count=$form.ProcCountLines('InputBox',0)
    $body=$form.Lines($first,$count)
    if($body -notmatch 'TestInventoryCommands.DialogAnswer'){throw 'Existing UOM answer seam unavailable.'}
    Write-Output ('UOM original answer seam exact-case match: '+$body.Contains('TestInventoryCommands.DialogAnswer(Prompt)'))
    $body=[regex]::Replace($body,'(?i)TestInventoryCommands\.DialogAnswer\(Prompt\)','TestInventoryUom.Answer(Prompt, mInteractiveContext)')
    if($body -notmatch 'TestInventoryUom\.Answer\('){throw 'UOM answer seam substitution failed; not product RED.'}
    $form.DeleteLines($first,$count);$form.InsertLines($first,$body)
    $form.AddFromString(@'
Public Sub UomRecordingForTest(ByVal mode As String)
    Dim initialCount As Long, valid As Boolean, index As Long, turn As Long
    initialCount = mCmbUom.ListCount
    If mode = "SignedOutAtPrompt" Or mode = "ReplacementSessionAtPrompt" Or mode = "ExpiredCommandAtPrompt" Then
        mCmbUom.Value = ADD_UOM_OPTION
        valid = TestInventoryUom.Yielded And TestInventoryUom.Count = 1
        valid = valid And mCmbUom.ListCount = initialCount And mLblStatus.Caption <> ""
        valid = valid And mCmbUom.Value <> "UOMSENTINEL"
    Else
        valid = True
        For turn = 1 To 4
            mCmbUom.Value = ADD_UOM_OPTION
            TestInventoryUom.Trace turn, mCmbUom.Value = "UOMSENTINEL", mCmbUom.ListCount = initialCount + 1
            valid = valid And mCmbUom.Value = "UOMSENTINEL"
            valid = valid And mCmbUom.ListCount = initialCount + 1
            If mode = "InterruptAfterSecondPrompt" And turn = 2 Then
                TestInventoryUom.Facts valid
                modAuth.SignOut
                mBtnCancel_Click
                Exit Sub
            End If
        Next turn
        valid = valid And TestInventoryUom.Count = 4
        For index = 0 To mCmbUom.ListCount - 1
            If mCmbUom.List(index) = "EA" Then mCmbUom.ListIndex = index: Exit For
        Next index
        valid = valid And mCmbUom.Value = "EA"
        TestInventoryUom.Trace 5, mCmbUom.Value = "EA", mCmbUom.ListCount = initialCount + 1
    End If
    TestInventoryUom.Facts valid
    mBtnCancel_Click
End Sub
'@)
}

function Test-InventoryUomRecording($Fixture) {
    SelectTarget $Fixture
    Check 'InventoryUom.Prior198DefinitionsPreserved' ([bool](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.Prior198Preserved'))
    $wire=[string](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.Definition' @('ADMIN_INVENTORY_UOM_ADD',46))
    $definition=if($wire){$wire|ConvertFrom-Json}else{$null}
    Check 'InventoryUom.Catalog46' ($null -ne $definition -and $definition.OwnerId -ceq 'ADMIN_INVENTORY' -and $definition.Capability -ceq 'ADMIN_MAINT' -and $definition.Class -ceq 'Command')
    Check 'InventoryUom.Version45Excludes' ([string](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.Definition' @('ADMIN_INVENTORY_UOM_ADD',45)) -ceq '')
    foreach($code in @('STAGED','CANCELLED','REFUSED','FAILED','REQUESTED')){
        Check ('InventoryUom.Conclusion.'+$code) ([bool](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.UomConclusion' @($code)) -eq ($code -ceq 'STAGED'))
    }
    $operatorBook=$excel.Workbooks.Add();$operatorBook.SaveAs((Join-Path $runRoot 'inventory-uom-host.xlsx'),51)
    try {
        foreach($mode in @('Enabled','Disabled','ControlOff','UserOff','SignedOutAtPrompt','ReplacementSessionAtPrompt','ExpiredCommandAtPrompt')){
            SelectTarget $Fixture
            $policy=if($mode -cin @('Disabled','ControlOff','UserOff')){$mode}else{'Enabled'}
            if(-not [bool](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.SetPolicy' @($policy))){throw 'UOM policy fixture unavailable; not product RED.'}
            $pins=@{};foreach($store in @('Config','Auth','Data.Inventory')){$path=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.'+$store+'.xlsb');$pins[$path]=Get-AdminInventoryEntryHash $path}
            $before=@(Get-Slice4beActivityFiles $Fixture)
            [void](Run 'invSys.Admin.xlam' 'TestInventoryUom.Arm' @($mode,$Fixture.Secret))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryRecording.Execute' @($operatorBook.Name,$mode))
            $valid=[string](Run 'invSys.Admin.xlam' 'TestInventoryUom.Result') -ceq 'True' -and [bool](Run 'invSys.Admin.xlam' 'TestInventoryRecording.Finished')
            Write-Output ('UOM fixed gesture diagnostics: '+[string](Run 'invSys.Admin.xlam' 'TestInventoryUom.Diagnostics'))
            Check ('InventoryUom.'+$mode+'.ActualOwnerHandlers') $valid
            if(-not $valid){throw 'UOM gesture/owner fixture did not complete; not product RED.'}
            Check ('InventoryUom.'+$mode+'.NoUnexpectedRibbonFailure') ([long](Run 'invSys.Admin.xlam' 'TestInventoryRecording.RibbonFailures') -eq 0)
            $rows=@(Get-Slice4beActivityFiles $Fixture|Where-Object{$_ -cnotin $before}|ForEach-Object{Get-Content $_ -Raw|ConvertFrom-Json}|Where-Object ControlId -Like 'ADMIN_INVENTORY_*')
            $suppressed=$policy -cne 'Enabled';$refused=$mode -ceq 'ExpiredCommandAtPrompt'
            $interrupted=$mode -cin @('SignedOutAtPrompt','ReplacementSessionAtPrompt')
            $expected=if($suppressed){0}elseif($interrupted){3}elseif($refused){6}else{14}
            Check ('InventoryUom.'+$mode+'.ExactRecordCount') ($rows.Count -eq $expected)
            if(-not $suppressed){
                $attempts=@($rows|Where-Object {$_.ControlId -ceq 'ADMIN_INVENTORY_UOM_ADD' -and $_.OutcomeCode -ceq 'REQUESTED'})
                $ends=@($rows|Where-Object {$_.ControlId -ceq 'ADMIN_INVENTORY_UOM_ADD' -and $_.OutcomeCode -cne 'REQUESTED'})
                $wanted=if($interrupted -or $refused){1}else{4}
                Check ('InventoryUom.'+$mode+'.DistinctAttempts') ($attempts.Count -eq $wanted -and @($attempts|ForEach-Object{$_.ActivityId}|Sort-Object -Unique).Count -eq $wanted)
                $valid=if($interrupted){$ends.Count -eq 0}elseif($refused){$ends.Count -eq 1 -and $ends[0].OutcomeCode -ceq 'REFUSED'}else{$ends.Count -eq 4 -and @($ends|Where-Object OutcomeCode -CEQ 'STAGED').Count -eq 2 -and @($ends|Where-Object OutcomeCode -CEQ 'CANCELLED').Count -eq 2}
                foreach($ending in $ends){$valid=$valid -and $ending.DataEffect -ceq 'Unchanged' -and @($attempts|Where-Object ActivityId -CEQ $ending.ActivityId).Count -eq 1}
                Check ('InventoryUom.'+$mode+'.ExactCorrelatedOutcomes') $valid
                $selections=@($rows|Where-Object ControlId -CEQ 'ADMIN_INVENTORY_SELECT_UOM')
                Check ('InventoryUom.'+$mode+'.SubsequentSelectionOnly') ($selections.Count -eq $(if($interrupted -or $refused){0}else{2}))
            }
            $valid=$true;foreach($row in $rows){$valid=$valid -and $row.WarehouseId -ceq $Fixture.Warehouse -and $row.UserId -ceq 'config-admin' -and $row.OwnerId -ceq 'ADMIN_INVENTORY' -and @($row.SourceEventRefs).Count -eq 0}
            Check ('InventoryUom.'+$mode+'.OriginalContextNoCanonicalRefs') $valid
            $json=[string]($rows|ConvertTo-Json -Depth 12 -Compress)
            Check ('InventoryUom.'+$mode+'.Redacted') ($json -notmatch 'UOMSentinel')
            $same=$true;foreach($path in $pins.Keys){$same=$same -and (Get-AdminInventoryEntryHash $path) -ceq $pins[$path]}
            Check ('InventoryUom.'+$mode+'.CanonicalBytesPreserved') $same
        }
    }finally{$operatorBook.Close($false)}
}
