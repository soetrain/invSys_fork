# Instrument disposable Operations memory, then use the existing packaged
# launcher and typed form-layout adapter. No business handler is replaced.
function Install-ProductionPaletteProbe {
    param($Excel,[hashtable]$Packages,[string]$PackageRoot)
    $form=$Packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('frmProduction').CodeModule
    $form.AddFromString(@'
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
