# Instrument disposable Operations memory, then use the existing packaged
# launcher and typed form-layout adapter. No business handler is replaced.
function Install-ProductionPaletteProbe {
    param($Excel,[hashtable]$Packages,[string]$PackageRoot)
    $form=$Packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('frmProduction').CodeModule
    $form.AddFromString(@'
Public Sub TestPaletteCloseForTest()
    mBtnClose_Click
End Sub
Public Function TestPaletteVisibleForTest(ByVal size As String) As String
    Dim previousLoading As Boolean, row As Long, geometry As String
    previousLoading = mLoading
    On Error GoTo Failed
    mLoading = True
    Select Case size
        Case "Minimum", "Default", "Restored"
            geometry = TestLayoutGeometryReportForSize(PRODUCTION_DEFAULT_WIDTH, PRODUCTION_DEFAULT_HEIGHT, 3)
        Case "NativeMaximize"
            geometry = TestCurrentLayoutGeometryReport(3)
        Case Else: Err.Raise 5, , "Unknown palette capture size."
    End Select
    mPages.Value = 3: mPages.Pages(3).ScrollTop = 0
    ' Display-only calibration rows, never inventory entities or workflow input.
    mLstRunPalette.Clear
    For row = 1 To 8
        mLstRunPalette.AddItem ""
        mLstRunPalette.List(row - 1, 2) = "Visible palette row " & CStr(row)
        mLstRunPalette.List(row - 1, 4) = "Display-only fixture"
    Next row
    mLstRunPalette.ListIndex = -1: mLstRunPalette.TopIndex = 0
    TestPaletteVisibleForTest = "Rows=" & CStr(mLstRunPalette.ListCount) & _
        "|TopIndex=" & CStr(mLstRunPalette.TopIndex) & "|Geometry=" & CStr(Left$(geometry, 3) = "OK|")
    mLoading = previousLoading
    Exit Function
Failed:
    mLoading = previousLoading
    Err.Raise Err.Number, , Err.Description
End Function
Private Function TestPaletteMeasurements() As String
    Dim originalWidth As Double, originalHeight As Double
    Dim dimensions As Variant, names As Variant, i As Long
    Dim geometry As String, result As String
    Dim probe As MSForms.ListBox, requestedHeight As Double, returnedHeight As Double
    originalWidth = Me.Width: originalHeight = Me.Height
    dimensions = Array(Array(PRODUCTION_MIN_WIDTH, PRODUCTION_MIN_HEIGHT), _
                       Array(PRODUCTION_DEFAULT_WIDTH, PRODUCTION_DEFAULT_HEIGHT), _
                       Array(PRODUCTION_LAYOUT_TEST_MAX_WIDTH, PRODUCTION_LAYOUT_TEST_MAX_HEIGHT), _
                       Array(originalWidth, originalHeight))
    names = Array("Minimum", "Default", "Expanded", "Restored")
    For i = 0 To 3
        geometry = TestLayoutGeometryReportForSize(CDbl(dimensions(i)(0)), CDbl(dimensions(i)(1)), 3)
        result = result & "|Palette" & names(i) & "Height=" & Format$(mLstRunPalette.Height, "0.000") & _
                 "|Palette" & names(i) & "Geometry=" & CStr(Left$(geometry, 3) = "OK|")
    Next i
    ' Supplemental cause measurement: call the real factory on a temporary
    ' control, then assign the same height after IntegralHeight is already off.
    requestedHeight = 96
    Set probe = AddList(mPages.Pages(3), "paletteHeightProbe", 0, 0, 1018, requestedHeight, 10, RUN_PALETTE_WIDTHS)
    returnedHeight = probe.Height
    probe.Height = requestedHeight
    result = result & "|FactoryHeight=" & Format$(returnedHeight, "0.000") & _
             "|ReassignedHeight=" & Format$(probe.Height, "0.000")
    mPages.Pages(3).Controls.Remove "paletteHeightProbe"
    Set probe = Nothing
    TestPaletteMeasurements = result
End Function
'@)
    $module=$Packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('mProduction').CodeModule
    $module.AddFromString(@'
Public Function PaletteVisibleForTest(ByVal size As String) As String
    If Not frmProduction.Visible Then Err.Raise 5, , "The launcher-owned Production form is not visible."
    PaletteVisibleForTest = frmProduction.TestPaletteVisibleForTest(size)
End Function
Public Function PaletteCaptionForTest() As String
    If Not frmProduction.Visible Then Err.Raise 5, , "The launcher-owned Production form is not visible."
    PaletteCaptionForTest = frmProduction.Caption
End Function
Public Sub ClosePaletteForTest()
    Dim instance As Object
    For Each instance In VBA.UserForms
        If TypeName(instance) = "frmProduction" Then
            instance.TestPaletteCloseForTest
            Exit Sub
        End If
    Next instance
End Sub
'@)
    $name='TestRunListResponsiveLayoutReportForSize'
    $start=$form.ProcStartLine($name,0);$count=$form.ProcCountLines($name,0)
    $body=$form.Lines($start,$count)
    $pattern='(?im)^End Function[ \t]*\r?$'
    if(([regex]::Matches($body,$pattern)).Count -ne 1){throw 'Layout adapter anchor ambiguous; not behavioral RED.'}
    $body=[regex]::Replace($body,$pattern,('    '+$name+' = '+$name+' & TestPaletteMeasurements()'+"`r`nEnd Function"))
    $form.DeleteLines($start,$count);$form.InsertLines($start,$body)
    $visible=$Excel.VBE.MainWindow.Visible
    try {
        foreach($name in @('invSys.Core.xlam','invSys.Inventory.Domain.xlam','invSys.Designs.Domain.xlam','invSys.Operations.xlam')){
            $project=$Packages[$name].VBProject
            foreach($reference in $project.References){
                if($reference.IsBroken){throw 'Broken instrumented reference.'}
                if($reference.Name -like 'invSys_*' -and -not [string]::Equals((Split-Path -Parent $reference.FullPath),$PackageRoot,[StringComparison]::OrdinalIgnoreCase)){throw 'Probe dependency outside disposable packages.'}
            }
            $Excel.VBE.ActiveVBProject=$project
            foreach($component in $project.VBComponents){if($component.Type -eq 1){$component.CodeModule.CodePane.Show();break}}
            $command=$Excel.VBE.CommandBars.FindControl(1,578)
            if($null -eq $command){throw 'Compile command unavailable.'}
            if($command.Enabled){$command.Execute()}
            if($Excel.VBE.CommandBars.FindControl(1,578).Enabled){throw 'Instrumented compile incomplete; not behavioral RED.'}
            Write-Output ('PALETTE_PROBE_COMPILE_PASS '+$name)
        }
    } finally {$Excel.VBE.MainWindow.Visible=$visible}
}

function Close-ProductionPaletteFixture {
    param($Excel,[string]$RuntimeRoot,[string]$PackageRoot)
    [void](Run-WorkbookMacro -Excel $Excel -WorkbookName 'invSys.Operations.xlam' -MacroName 'mProduction.ClosePaletteForTest')
    $books=@($Excel.Workbooks)
    foreach($book in $books){
        $path=[IO.Path]::GetFullPath([string]$book.FullName)
        $fixture=$path.StartsWith([IO.Path]::GetFullPath($RuntimeRoot).TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)
        $package=[bool]$book.IsAddin -and $path.StartsWith([IO.Path]::GetFullPath($PackageRoot).TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)
        if(-not ($fixture -or $package)){throw 'Palette cleanup found a workbook outside its isolated roots.'}
    }
    $Excel.EnableEvents=$false;$Excel.DisplayAlerts=$false
    foreach($book in @($books|Sort-Object {[bool]$_.IsAddin})){$book.Close($false)}
    $empty=[int]$Excel.Workbooks.Count -eq 0
    foreach($book in $books){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}
    return $empty
}

function Invoke-ProductionPaletteCapture {
    param($Excel,$Rows,[string]$OutputPath,[string]$RepoRoot)
    # Reuse the calibrated ownership, physical-pixel and foreground checks.
    $reportRoot=$OutputPath
    $GuideCaptureVisibleExcelForTest=$false;$GuideCaptureSavedWorkbookForTest=$false;$TraceGuideResourcesForTest=$false
    $tokens=$null;$parseErrors=$null
    $tree=[Management.Automation.Language.Parser]::ParseFile((Join-Path $RepoRoot 'tests/tooling/Test-Slice4beConfigCommands.ps1'),[ref]$tokens,[ref]$parseErrors)
    if($parseErrors.Count){throw 'Capture helper source does not parse.'}
    foreach($name in @('Initialize-SettingsCapture','CaptureFormEvidence','CaptureOwnedFormEvidence')){
        $definition=$tree.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -ceq $name},$true)
        if($null -eq $definition){throw 'Calibrated capture helper is missing.'}
        . ([scriptblock]::Create($definition.Extent.Text))
    }
    Initialize-SettingsCapture
    if(-not ('ProductionPaletteNative' -as [type])){
        Add-Type @'
using System;using System.Runtime.InteropServices;
public static class ProductionPaletteNative {
 [DllImport("user32.dll")]public static extern bool ShowWindow(IntPtr window,int command);
 [DllImport("user32.dll")]public static extern bool IsZoomed(IntPtr window);
}
'@
    }
    $title=[string](Run-WorkbookMacro -Excel $Excel -WorkbookName 'invSys.Operations.xlam' -MacroName 'mProduction.PaletteCaptionForTest')
    $window=[InvSysSettingsCapture]::OwnedVisibleForm($title,[IntPtr]$Excel.Hwnd)
    if($window -eq [IntPtr]::Zero){throw 'Unique launcher-owned palette form is unavailable.'}
    foreach($size in @('Minimum','Default','NativeMaximize','Restored')){
        if($size -eq 'NativeMaximize'){[void][ProductionPaletteNative]::ShowWindow($window,3)}
        elseif($size -eq 'Restored'){[void][ProductionPaletteNative]::ShowWindow($window,9)}
        $report=[string](Run-WorkbookMacro -Excel $Excel -WorkbookName 'invSys.Operations.xlam' -MacroName 'mProduction.PaletteVisibleForTest' -Arguments @($size))
        $passed=$report -ceq 'Rows=8|TopIndex=0|Geometry=True'
        if($size -eq 'NativeMaximize'){$passed=$passed -and [ProductionPaletteNative]::IsZoomed($window)}
        if($size -eq 'Restored'){$passed=$passed -and -not [ProductionPaletteNative]::IsZoomed($window)}
        CaptureOwnedFormEvidence $title ('production-palette-'+$size.ToLowerInvariant()+'.png') $window.ToInt64()
        Add-Evidence -Rows $Rows -Callback ('Production.Palette.Visible.'+$size) -Expected 'The real launcher form contains eight display-only rows with valid geometry; complete row visibility requires image review.' -Passed $passed -Observed $report
    }
}

function Add-ProductionPaletteEvidence {
    param([string]$Report,$Rows,[string]$OutputPath)
    $values=@{}
    foreach($part in $Report.Split('|')){
        if($part -match '^(Palette(?:Minimum|Default|Expanded|Restored)(?:Height|Geometry)|FactoryHeight|ReassignedHeight)=(.+)$'){$values[$Matches[1]]=$Matches[2]}
    }
    if($values.Count -ne 10){throw 'Palette measurement fields missing; not behavioral RED.'}
    $values|ConvertTo-Json|Set-Content (Join-Path $OutputPath 'palette-geometry.json')
    foreach($size in @('Minimum','Default','Expanded','Restored')){
        $height=[double]::Parse($values['Palette'+$size+'Height'],[Globalization.CultureInfo]::InvariantCulture)
        Add-Evidence -Rows $Rows -Callback ('Production.Palette.'+$size+'.EightRows') -Expected 'Preserve the accepted palette height of at least 90 points.' -Passed ($height -ge 90) -Observed ('Height='+$height)
        Add-Evidence -Rows $Rows -Callback ('Production.Palette.'+$size+'.Geometry') -Expected 'Run List controls fit without interactive overlap.' -Passed ($values['Palette'+$size+'Geometry'] -ceq 'True') -Observed $values['Palette'+$size+'Geometry']
    }
    foreach($key in @('FactoryHeight','ReassignedHeight')){
        $height=[double]::Parse($values[$key],[Globalization.CultureInfo]::InvariantCulture)
        Add-Evidence -Rows $Rows -Callback ('Production.ListFactory.'+$key) -Expected 'The requested 96-point height is preserved with IntegralHeight disabled.' -Passed ([Math]::Abs($height-96) -lt 0.1) -Observed ('Height='+$height)
    }
}
