# Extend the existing actual Ribbon/form gate; faults affect only optional
# observation writes and generated workbook/session state, never owner results.
function Install-InventoryRecordingBoundaries($AdminProbe,$CoreProbe) {
    . (Join-Path $PSScriptRoot 'Slice4beAdminDirectoryBoundaries.ps1')
    Install-AdminDirectoryBoundaries $CoreProbe.CodeModule $null -ControlId 'ADMIN_INVENTORY_EDIT_MODE' -Outcome 'SELECTED' -ProbeName 'TestInventoryRecordingPolicy' -RootAction
    $module=$CoreProbe.CodeModule
    $start=$module.ProcStartLine('AtEntry',0);$count=$module.ProcCountLines('AtEntry',0)
    $body=$module.Lines($start,$count)
    $marker='        Case "SignOut"'
    if([regex]::Matches($body,[regex]::Escape($marker)).Count -ne 1){throw 'Session boundary adapter unavailable.'}
    $body=$body.Replace($marker,@'
        Case "ReplacementSession"
            modAuth.SignOut
            status = modAuth.SignInCurrentTargetForAutomation("config-admin", mOtherSecret, "")
            mEntryObserved = (Left$(status, 3) = "OK|")
        Case "SignOut"
'@.TrimEnd())
    $module.DeleteLines($start,$count);$module.InsertLines($start,$body)
    $module=$AdminProbe.CodeModule
    $module.InsertLines($module.CountOfDeclarationLines+1,'Private mBoundary As String, mHost As String, mFinished As Boolean, mBoundaryReached As Boolean, mRibbonFailures As Long')
    $module.AddFromString(@'
Public Sub ArmHostBoundary(ByVal mode As String, ByVal host As String)
    mBoundary = mode: mHost = host: mFinished = False: mBoundaryReached = False
    mRibbonFailures = 0
End Sub
Public Sub AfterWorkbookCapture()
    Dim book As Workbook, path As String
    If mBoundary <> "ClosedWorkbook" And mBoundary <> "ReplacedWorkbook" Then Exit Sub
    Set book = Application.Workbooks(mHost)
    path = book.FullName
    book.Close False
    If mBoundary = "ReplacedWorkbook" Then Set book = Application.Workbooks.Open(path)
    mBoundaryReached = True
End Sub
Public Function BoundaryReached() As Boolean
    BoundaryReached = mBoundaryReached
End Function
Public Function Finished() As Boolean
    Finished = mFinished
End Function
Public Sub RibbonFailure()
    mRibbonFailures = mRibbonFailures + 1
End Sub
Public Function RibbonFailures() As Long
    RibbonFailures = mRibbonFailures
End Function
'@)
    $start=$module.ProcStartLine('Execute',0);$count=$module.ProcCountLines('Execute',0)
    $body=$module.Lines($start,$count)
    $marker='(?im)^[\t ]*mDriving = False[\t ]*\r?$'
    if([regex]::Matches($body,$marker).Count -ne 1){throw 'Actual navigation return boundary unavailable.'}
    $body=[regex]::Replace($body,$marker,'    mDriving = False: mFinished = True')
    $body=[regex]::Replace($body,'(?im)^([\t ]*mDriving = True)',('    On Error GoTo Failed'+"`r`n"+'$1'))
    $body=[regex]::Replace($body,'(?im)^End Sub\s*$',('    Exit Sub'+"`r`n"+'Failed:'+"`r`n"+'    mDriving = False: mFinished = False'+"`r`n"+'End Sub'))
    $module.DeleteLines($start,$count);$module.InsertLines($start,$body)
    $module=$packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('modAdmin').CodeModule
    $start=$module.ProcStartLine('Add_InventoryItem',0);$count=$module.ProcCountLines('Add_InventoryItem',0)
    $body=$module.Lines($start,$count)
    $marker='(?im)^[\t ]*activity\.Initialize targetWb, "ADMIN_INVENTORY_", "Inventory"[\t ]*\r?$'
    if([regex]::Matches($body,$marker).Count -ne 1){throw 'Captured Inventory workbook boundary unavailable.'}
    $body=[regex]::Replace($body,$marker,('$0'+"`r`n"+'    TestInventoryRecording.AfterWorkbookCapture'))
    $module.DeleteLines($start,$count);$module.InsertLines($start,$body)
    $module=$packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('modRibbonGenerated').CodeModule
    $start=$module.ProcStartLine('RibbonOnActionAdmin',0);$count=$module.ProcCountLines('RibbonOnActionAdmin',0)
    $body=$module.Lines($start,$count)
    $marker='(?im)^[\t ]*MsgBox "Ribbon action failed: " & Err.Description, vbExclamation[\t ]*\r?$'
    if([regex]::Matches($body,$marker).Count -ne 1){throw 'Ribbon failure observation boundary unavailable.'}
    $body=[regex]::Replace($body,$marker,'    TestInventoryRecording.RibbonFailure')
    $module.DeleteLines($start,$count);$module.InsertLines($start,$body)
}

function Test-InventoryRecordingBoundaries($Fixture,$Other,$HostBook,[string[]]$Modes=@('SwitchWorkbook','SignOut','ReplacementSession','SwitchTarget','AttemptWriteFailure','OutcomeWriteFailure','ReplacedWorkbook','ClosedWorkbook')) {
    $otherBook=$excel.Workbooks.Add();$otherBook.SaveAs((Join-Path $runRoot 'inventory-recording-other.xlsx'),51)
    $otherPath=$otherBook.FullName;$hostPath=$HostBook.FullName
    function SourcePins($Target){
        $pins=@{}
        foreach($kind in @('Config','Auth','Data.Inventory')){
            $path=Join-Path $Target.Root ($Target.Warehouse+'.invSys.'+$kind+'.xlsb')
            $pins[$kind]=Get-AdminInventoryEntryHash $path
        }
        return $pins
    }
    function SameSource($Before,$After){foreach($key in $Before.Keys){if($Before[$key] -cne $After[$key]){return $false}};return $true}
    try {
        foreach($mode in $Modes){
            SelectTarget $Fixture
            $ready=[bool](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.SetPolicy' @('Enabled'))
            if(-not $ready){throw 'Boundary policy fixture unavailable; not product RED.'}
            $before=@(Get-Slice4beActivityFiles $Fixture);$otherBefore=@(Get-Slice4beActivityFiles $Other)
            $sourcePins=SourcePins $Fixture;$otherPins=SourcePins $Other
            $operatorBook=$excel.Workbooks.Item([IO.Path]::GetFileName($hostPath))
            $secret=if($mode -ceq 'ReplacementSession'){$Fixture.Secret}else{$Other.Secret}
            [void](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.ArmBoundary' @($mode,$Other.Root,$secret,$otherBook.Name))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryRecording.ArmHostBoundary' @($mode,$operatorBook.Name))
            try {
                [void](Run 'invSys.Admin.xlam' 'TestInventoryRecording.Execute' @($operatorBook.Name,'Enabled'))
                $reached=if($mode -cin @('ClosedWorkbook','ReplacedWorkbook')){[bool](Run 'invSys.Admin.xlam' 'TestInventoryRecording.BoundaryReached')}elseif($mode -like '*WriteFailure'){[long](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.WriteFailures') -eq 2}else{[bool](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.EntryObserved')}
                Check ('InventoryBoundary.'+$mode+'.FaultReached') $reached
                if(-not $reached){throw 'Requested recording boundary was not reached; not product RED.'}
                Check ('InventoryBoundary.'+$mode+'.OwnerHandlerReturned') ([bool](Run 'invSys.Admin.xlam' 'TestInventoryRecording.Finished'))
                $expectedPresentations=if($mode -cin @('ClosedWorkbook','ReplacedWorkbook')){0}else{1}
                Check ('InventoryBoundary.'+$mode+'.CapturedPresentationOnly') ([long](Run 'invSys.Admin.xlam' 'TestInventoryRecording.Gestures') -eq $expectedPresentations)
                Check ('InventoryBoundary.'+$mode+'.NoUnexpectedRibbonFailure') ([long](Run 'invSys.Admin.xlam' 'TestInventoryRecording.RibbonFailures') -eq 0)
                $rows=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object{Get-Content $_ -Raw|ConvertFrom-Json}|Where-Object ControlId -Like 'ADMIN_INVENTORY_*')
                $expected=switch($mode){'SwitchWorkbook'{10} 'AttemptWriteFailure'{6} 'OutcomeWriteFailure'{8} 'ClosedWorkbook'{1} 'ReplacedWorkbook'{1} default{2}}
                Check ('InventoryBoundary.'+$mode+'.ExactRecordCount') ($rows.Count -eq $expected)
                $edits=@($rows|Where-Object ControlId -CEQ 'ADMIN_INVENTORY_EDIT_MODE')
                $expectedEdits=switch($mode){'SwitchWorkbook'{4} 'OutcomeWriteFailure'{2} default{0}}
                $exact=$edits.Count -eq $expectedEdits
                if($mode -ceq 'OutcomeWriteFailure'){$exact=$exact -and @($edits|Where-Object OutcomeCode -CNE 'REQUESTED').Count -eq 0}
                Check ('InventoryBoundary.'+$mode+'.NoInventedEditCompletion') $exact
                $bound=$rows.Count -eq $expected
                foreach($row in $rows){$bound=$bound -and $row.WarehouseId -ceq $Fixture.Warehouse -and $row.UserId -ceq 'config-admin' -and $row.OwnerId -ceq 'ADMIN_INVENTORY'}
                Check ('InventoryBoundary.'+$mode+'.OriginalContextOnly') $bound
                Check ('InventoryBoundary.'+$mode+'.OtherWarehouseNotRecorded') (@(Get-Slice4beActivityFiles $Other|Where-Object{$_ -cnotin $otherBefore}).Count -eq 0)
                $notices=[long](Run 'invSys.Admin.xlam' 'TestInventoryRecording.Notices')
                Check ('InventoryBoundary.'+$mode+'.TrackingNotice') $(if($mode -ceq 'SwitchWorkbook'){$notices -eq 0}else{$notices -gt 0})
                Check ('InventoryBoundary.'+$mode+'.CanonicalStoresPreserved') ((SameSource $sourcePins (SourcePins $Fixture)) -and (SameSource $otherPins (SourcePins $Other)))
                Check ('InventoryBoundary.'+$mode+'.OtherWorkbookPreserved') ($otherBook.FullName -ceq $otherPath -and [bool]$otherBook.Saved)
            }finally{
                [void](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.ArmBoundary' @('','','',''))
                [void](Run 'invSys.Admin.xlam' 'TestInventoryRecording.ArmHostBoundary' @('',''))
            }
        }
    }finally{$otherBook.Close($false);SelectTarget $Fixture}
}
