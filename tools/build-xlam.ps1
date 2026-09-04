[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$RepoRoot = ".",

    [Parameter(Mandatory = $false)]
    [string]$OutputRoot = "deploy/current",

    [Parameter(Mandatory = $false)]
    [string[]]$Projects = @(),

    [Parameter(Mandatory = $false)]
    [switch]$IncludeOperationsShadow,

    [Parameter(Mandatory = $false)]
    [switch]$Apply
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-OpenXmlAssemblyPath {
    $candidates = @(
        "C:\Program Files\Microsoft Office\root\Office16\ADDINS\Microsoft Power Query for Excel Integrated\bin\DocumentFormat.OpenXml.dll",
        "C:\Program Files\Microsoft Office\root\vfs\ProgramFilesCommonX64\Microsoft Shared\Filters\Documentformat.OpenXml.dll",
        "C:\Program Files\Microsoft Office\root\vfs\ProgramFilesX86\Microsoft Office\Office16\DCF\DocumentFormat.OpenXml.dll"
    )

    foreach ($path in $candidates) {
        if (Test-Path -LiteralPath $path) {
            return $path
        }
    }

    throw "DocumentFormat.OpenXml.dll not found in known Office locations."
}

$openXmlAssemblyPath = Get-OpenXmlAssemblyPath
[void][System.Reflection.Assembly]::LoadFrom($openXmlAssemblyPath)

function Release-ComObject {
    param([object]$Obj)
    if ($null -ne $Obj) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Obj) } catch {}
    }
}

function Get-DeployedOutputPath {
    param(
        [object]$Project,
        [string]$OutputDir
    )

    Join-Path $OutputDir $Project.OutputFile
}

function Get-CodeFiles {
    param(
        [string[]]$SourceDirs
    )

    $files = foreach ($dir in $SourceDirs) {
        if (-not (Test-Path -LiteralPath $dir)) {
            throw "Source directory not found: $dir"
        }

        Get-ChildItem -Path $dir -Recurse -File |
            Where-Object {
                $_.Extension -in @(".bas", ".cls", ".frm") -and
                $_.Name -notlike "*.bak"
            }
    }

    $files | Sort-Object FullName -Unique
}

function Get-SheetModuleFiles {
    param(
        [System.IO.FileInfo[]]$CodeFiles
    )

    $CodeFiles | Where-Object {
        $_.Extension -eq ".cls" -and $_.FullName -match "\\ClassModules\\Sheets\\"
    }
}

function Get-ImportFiles {
    param(
        [System.IO.FileInfo[]]$CodeFiles
    )

    $CodeFiles | Where-Object {
        -not ($_.Extension -eq ".cls" -and $_.FullName -match "\\ClassModules\\Sheets\\") -and
        $_.Extension -ne ".frm"
    }
}

function Get-FormFiles {
    param(
        [System.IO.FileInfo[]]$CodeFiles
    )

    $CodeFiles | Where-Object { $_.Extension -eq ".frm" }
}

function Remove-VbaTestOnlyRegions {
    param(
        [string]$SourceText,
        [string]$SourcePath
    )

    $beginMarker = "'@TestOnlyBegin"
    $endMarker = "'@TestOnlyEnd"
    $beginCount = ([regex]::Matches($SourceText, "(?im)^[ \t]*" + [regex]::Escape($beginMarker) + "[ \t]*$")).Count
    $endCount = ([regex]::Matches($SourceText, "(?im)^[ \t]*" + [regex]::Escape($endMarker) + "[ \t]*$")).Count
    if ($beginCount -ne $endCount) {
        throw "Unbalanced VBA test-only markers in ${SourcePath}: begin=$beginCount end=$endCount"
    }
    if ($beginCount -eq 0) {
        return $SourceText
    }

    $regionPattern = "(?ms)^[ \t]*" + [regex]::Escape($beginMarker) +
        "[ \t]*\r?\n.*?^[ \t]*" + [regex]::Escape($endMarker) +
        "[ \t]*(?:\r?\n|$)"
    $stripped = [regex]::Replace($SourceText, $regionPattern, "")
    if ($stripped -match "(?im)^[ \t]*(')?@TestOnly(Begin|End)[ \t]*$") {
        throw "VBA test-only marker remained after stripping ${SourcePath}."
    }
    return $stripped
}

function New-NormalizedImportFile {
    param(
        [System.IO.FileInfo]$SourceFile
    )

    $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("invsys-build-" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

    $tempPath = Join-Path $tempDir $SourceFile.Name
    $raw = Get-Content -LiteralPath $SourceFile.FullName -Raw
    $raw = Remove-VbaTestOnlyRegions -SourceText $raw -SourcePath $SourceFile.FullName
    $normalized = $raw -replace "`r?`n", "`r`n"
    [System.IO.File]::WriteAllText($tempPath, $normalized, [System.Text.Encoding]::ASCII)

    return $tempPath
}

function New-NormalizedFormImportFile {
    param(
        [System.IO.FileInfo]$FormFile
    )

    $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("invsys-form-" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

    $tempFrmPath = Join-Path $tempDir $FormFile.Name
    $raw = Get-Content -LiteralPath $FormFile.FullName -Raw
    $raw = Remove-VbaTestOnlyRegions -SourceText $raw -SourcePath $FormFile.FullName
    $normalized = $raw -replace "`r?`n", "`r`n"
    [System.IO.File]::WriteAllText($tempFrmPath, $normalized, [System.Text.Encoding]::ASCII)

    $sourceFrxPath = [System.IO.Path]::ChangeExtension($FormFile.FullName, ".frx")
    if (Test-Path -LiteralPath $sourceFrxPath) {
        $tempFrxPath = [System.IO.Path]::ChangeExtension($tempFrmPath, ".frx")
        Copy-Item -LiteralPath $sourceFrxPath -Destination $tempFrxPath -Force
    }

    return $tempFrmPath
}

function Get-VbComponentNameFromFile {
    param(
        [System.IO.FileInfo]$SourceFile
    )

    return [System.IO.Path]::GetFileNameWithoutExtension($SourceFile.Name)
}

function Remove-ExistingVBComponent {
    param(
        [object]$VBProject,
        [string]$ComponentName
    )

    if ([string]::IsNullOrWhiteSpace($ComponentName)) {
        return
    }

    try {
        $existing = $VBProject.VBComponents.Item($ComponentName)
    }
    catch {
        $existing = $null
    }

    if ($null -eq $existing) {
        return
    }

    if ($existing.Type -eq 100) {
        throw "Refusing to remove document component '$ComponentName'."
    }

    [void]$VBProject.VBComponents.Remove($existing)
}

function Assert-VBComponentType {
    param(
        [object]$VBProject,
        [string]$ComponentName,
        [int]$ExpectedType,
        [string]$Context
    )

    try {
        $component = $VBProject.VBComponents.Item($ComponentName)
    }
    catch {
        throw "$Context failed: component '$ComponentName' was not present after import."
    }

    if ($component.Type -ne $ExpectedType) {
        throw "$Context failed: component '$ComponentName' imported with type $($component.Type), expected $ExpectedType."
    }
}

function Ensure-WorksheetNames {
    param(
        [object]$Workbook,
        [string[]]$SheetNames
    )

    if (-not $SheetNames -or $SheetNames.Count -eq 0) {
        return
    }

    while ($Workbook.Worksheets.Count -lt $SheetNames.Count) {
        [void]$Workbook.Worksheets.Add()
    }

    for ($i = 0; $i -lt $SheetNames.Count; $i++) {
        $Workbook.Worksheets.Item($i + 1).Name = $SheetNames[$i]
    }
}

function Import-Components {
    param(
        [object]$VBProject,
        [System.IO.FileInfo[]]$Files
    )

    foreach ($file in $Files) {
        if ($file.Extension -eq ".cls") {
            $firstLine = Get-Content -LiteralPath $file.FullName -TotalCount 1
            $componentName = Get-VbComponentNameFromFile -SourceFile $file
            Remove-ExistingVBComponent -VBProject $VBProject -ComponentName $componentName
            if ($firstLine -match '^VERSION 1\.0 CLASS') {
                Write-Host ("  Importing " + $file.FullName)
                $normalizedPath = New-NormalizedImportFile -SourceFile $file
                try {
                    [void]$VBProject.VBComponents.Import($normalizedPath)
                    Assert-VBComponentType -VBProject $VBProject -ComponentName $componentName -ExpectedType 2 -Context $file.FullName
                }
                finally {
                    Remove-Item -LiteralPath (Split-Path $normalizedPath -Parent) -Recurse -Force -ErrorAction SilentlyContinue
                }
                continue
            }

            Write-Host ("  Creating class module " + $componentName)
            $rawLines = Get-Content -LiteralPath $file.FullName
            $codeLines = New-Object System.Collections.Generic.List[string]
            $inHeader = $true

            foreach ($line in $rawLines) {
                if ($inHeader) {
                    if (
                        $line -match '^VERSION ' -or
                        $line -match '^BEGIN$' -or
                        $line -match '^End$' -or
                        $line -match '^\s+\w+\s*=' -or
                        $line -match '^Attribute VB_'
                    ) {
                        continue
                    }

                    if ([string]::IsNullOrWhiteSpace($line)) {
                        continue
                    }

                    $inHeader = $false
                }

                if ($line -match '^Attribute ') {
                    continue
                }

                [void]$codeLines.Add($line)
            }

            $component = $VBProject.VBComponents.Add(2)
            $component.Name = $componentName
            $module = $component.CodeModule
            if ($module.CountOfLines -gt 0) {
                $module.DeleteLines(1, $module.CountOfLines)
            }
            $module.AddFromString(([string]::Join([Environment]::NewLine, $codeLines)))
            Assert-VBComponentType -VBProject $VBProject -ComponentName $componentName -ExpectedType 2 -Context $file.FullName
            continue
        }

        Remove-ExistingVBComponent -VBProject $VBProject -ComponentName (Get-VbComponentNameFromFile -SourceFile $file)
        Write-Host ("  Importing " + $file.FullName)
        $normalizedPath = New-NormalizedImportFile -SourceFile $file
        try {
            [void]$VBProject.VBComponents.Import($normalizedPath)
        }
        finally {
            Remove-Item -LiteralPath (Split-Path $normalizedPath -Parent) -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Test-FormRequiresStub {
    param(
        [System.IO.FileInfo]$FormFile
    )

    $frxPath = [System.IO.Path]::ChangeExtension($FormFile.FullName, ".frx")
    return -not (Test-Path -LiteralPath $frxPath)
}

function Add-StubUserForm {
    param(
        [object]$VBProject,
        [System.IO.FileInfo]$FormFile
    )

    $formName = [System.IO.Path]::GetFileNameWithoutExtension($FormFile.Name)
    Write-Host ("  Stubbing userform " + $formName + " (missing FRX designer)")
    Remove-ExistingVBComponent -VBProject $VBProject -ComponentName $formName
    $component = $VBProject.VBComponents.Add(3)
    $component.Name = $formName
    $captionLine = Get-Content -LiteralPath $FormFile.FullName | Where-Object { $_ -match '^\s*Caption\s*=\s*"' } | Select-Object -First 1
    if ($null -ne $captionLine) {
        $caption = [regex]::Match($captionLine, '"([^"]*)"').Groups[1].Value
        if ($caption -ne "") {
            try { $component.Designer.Caption = $caption } catch {}
        }
    }
    $module = $component.CodeModule
    if ($module.CountOfLines -gt 0) {
        $module.DeleteLines(1, $module.CountOfLines)
    }
    $stubCode = Get-StubUserFormCode -FormFile $FormFile
    $module.AddFromString($stubCode)
    Assert-VBComponentType -VBProject $VBProject -ComponentName $formName -ExpectedType 3 -Context $FormFile.FullName
}

function Get-StubUserFormCode {
    param(
        [System.IO.FileInfo]$FormFile
    )

    $rawLines = Get-Content -LiteralPath $FormFile.FullName
    $hasRuntimeMarker = $false
    foreach ($line in $rawLines) {
        if ($line -match "'@RuntimeStubUserFormCode") {
            $hasRuntimeMarker = $true
            break
        }
    }

    if (-not $hasRuntimeMarker) {
        return "Option Explicit"
    }

    $codeLines = New-Object System.Collections.Generic.List[string]
    $inCode = $false
    foreach ($line in $rawLines) {
        if (-not $inCode) {
            if ($line -match "'@RuntimeStubUserFormCode") {
                $inCode = $true
            }
            continue
        }

        if ($line -match '^Attribute ') {
            continue
        }

        [void]$codeLines.Add($line)
    }

    if ($codeLines.Count -eq 0) {
        return "Option Explicit"
    }

    return [string]::Join([Environment]::NewLine, $codeLines)
}

function Import-Forms {
    param(
        [object]$VBProject,
        [System.IO.FileInfo[]]$FormFiles
    )

    foreach ($formFile in $FormFiles) {
        $formName = [System.IO.Path]::GetFileNameWithoutExtension($formFile.Name)
        $formSource = Get-Content -LiteralPath $formFile.FullName -Raw
        if ($formSource -match "'@RuntimeStubUserFormCode") {
            if (
                $formSource -notmatch "(EnableResizableUserForm|modProductionFormWindow\.EnableResizable|modReceivingFormWindow\.EnableReceivingResizable)" -or
                $formSource -notmatch "True\s*,\s*True"
            ) {
                throw "$($formFile.FullName) violates the runtime UserForm window standard: Andy Pope/Windows API resize with minimize and maximize must be enabled."
            }
        }
        if (Test-FormRequiresStub -FormFile $formFile) {
            Add-StubUserForm -VBProject $VBProject -FormFile $formFile
        }
        else {
            Write-Host ("  Importing " + $formFile.FullName)
            Remove-ExistingVBComponent -VBProject $VBProject -ComponentName $formName
            $normalizedPath = New-NormalizedFormImportFile -FormFile $formFile
            try {
                [void]$VBProject.VBComponents.Import($normalizedPath)
                Assert-VBComponentType -VBProject $VBProject -ComponentName $formName -ExpectedType 3 -Context $formFile.FullName
            }
            finally {
                Remove-Item -LiteralPath (Split-Path $normalizedPath -Parent) -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

function Add-RibbonCallbacksModule {
    param(
        [object]$VBProject,
        [hashtable]$RibbonConfig
    )

    if ($null -eq $RibbonConfig) {
        return
    }

    $enabledCallbackName = "RibbonRequiredCapabilityGetEnabled"
    if ($RibbonConfig.ContainsKey("EnabledCallbackName") -and -not [string]::IsNullOrWhiteSpace($RibbonConfig.EnabledCallbackName)) {
        $enabledCallbackName = [string]$RibbonConfig.EnabledCallbackName
    }

    $lines = New-Object System.Collections.Generic.List[string]
    [void]$lines.Add("Option Explicit")
    [void]$lines.Add("")
    [void]$lines.Add("Public Sub RibbonOnLoad(ribbon As IRibbonUI)")
    [void]$lines.Add("    On Error Resume Next")
    [void]$lines.Add("    modRibbonRuntimeStatus.RegisterRibbonUi ribbon")
    [void]$lines.Add("    ribbon.Invalidate")
    [void]$lines.Add("    On Error GoTo 0")
    [void]$lines.Add("End Sub")
    [void]$lines.Add("")
    [void]$lines.Add("Public Sub " + $RibbonConfig.CallbackName + "(control As IRibbonControl)")
    [void]$lines.Add("    On Error GoTo ErrHandler")
    [void]$lines.Add("    Select Case control.ID")

    foreach ($group in $RibbonConfig.Groups) {
        $callbackButtons = @()
        if ($group.ContainsKey("Buttons")) { $callbackButtons += @($group.Buttons) }
        if ($group.ContainsKey("PostStatusMenuButtons")) { $callbackButtons += @($group.PostStatusMenuButtons) }
        foreach ($button in $callbackButtons) {
            [void]$lines.Add(("        Case ""{0}""" -f $button.Id))
            if ($button.ContainsKey("RequiredCapability") -and -not [string]::IsNullOrWhiteSpace($button.RequiredCapability)) {
                [void]$lines.Add(("            If Not modRoleUiAccess.RequireCurrentUserCapabilityCached(""{0}"", ""Current user does not have {0} for this warehouse/station."") Then Exit Sub" -f $button.RequiredCapability))
            }
            if ($button.ContainsKey("DirectAction") -and -not [string]::IsNullOrWhiteSpace($button.DirectAction)) {
                [void]$lines.Add(("            {0}" -f $button.DirectAction))
            } else {
                [void]$lines.Add(("            {0}" -f $button.Macro))
            }
        }
    }

    [void]$lines.Add("    End Select")
    [void]$lines.Add("    Exit Sub")
    [void]$lines.Add("ErrHandler:")
    [void]$lines.Add('    MsgBox "Ribbon action failed: " & Err.Description, vbExclamation')
    [void]$lines.Add("End Sub")
    [void]$lines.Add("")
    [void]$lines.Add("Public Sub RibbonRuntimeStatusGetLabel(control As IRibbonControl, ByRef returnedVal)")
    [void]$lines.Add("    returnedVal = modRibbonRuntimeStatus.GetStatusLabel(control.ID)")
    [void]$lines.Add("End Sub")
    [void]$lines.Add("")
    [void]$lines.Add("Public Sub RibbonServerStatusGetLabel(control As IRibbonControl, ByRef returnedVal)")
    [void]$lines.Add("    returnedVal = modRibbonRuntimeStatus.GetServerStatusLabel(control.ID)")
    [void]$lines.Add("End Sub")
    [void]$lines.Add("")
    [void]$lines.Add("Public Sub RibbonAccessStatusGetLabel(control As IRibbonControl, ByRef returnedVal)")
    [void]$lines.Add("    returnedVal = modRibbonRuntimeStatus.GetAccessStatusLabel(control.ID)")
    [void]$lines.Add("End Sub")
    [void]$lines.Add("")
    [void]$lines.Add("Public Sub RibbonRuntimeStatusRefresh(control As IRibbonControl)")
    [void]$lines.Add("    modRibbonRuntimeStatus.RefreshRuntimeContext")
    [void]$lines.Add("End Sub")
    [void]$lines.Add("")
    [void]$lines.Add("Public Sub RibbonCurrentUserGetLabel(control As IRibbonControl, ByRef returnedVal)")
    [void]$lines.Add("    returnedVal = modRibbonRuntimeStatus.GetCurrentUserActionLabel()")
    [void]$lines.Add("End Sub")
    [void]$lines.Add("")
    [void]$lines.Add("Public Sub RibbonServerSessionGetLabel(control As IRibbonControl, ByRef returnedVal)")
    [void]$lines.Add("    returnedVal = modRibbonRuntimeStatus.GetServerSessionActionLabel()")
    [void]$lines.Add("End Sub")
    [void]$lines.Add("")
    [void]$lines.Add("Public Sub " + $enabledCallbackName + "(control As IRibbonControl, ByRef returnedVal As Variant)")
    [void]$lines.Add("    On Error GoTo Disabled")
    [void]$lines.Add("    returnedVal = CBool(RibbonRequiredCapabilityIsEnabledById(control.ID))")
    [void]$lines.Add("    Exit Sub")
    [void]$lines.Add("Disabled:")
    [void]$lines.Add("    returnedVal = False")
    [void]$lines.Add("End Sub")
    [void]$lines.Add("")
    [void]$lines.Add("Public Function RibbonRequiredCapabilityIsEnabledById(ByVal controlId As String) As Boolean")
    [void]$lines.Add("    On Error GoTo Disabled")
    [void]$lines.Add("    RibbonRequiredCapabilityIsEnabledById = False")
    [void]$lines.Add("    Select Case controlId")
    foreach ($group in $RibbonConfig.Groups) {
        foreach ($button in $group.Buttons) {
            if ($button.ContainsKey("RequiredCapability") -and -not [string]::IsNullOrWhiteSpace($button.RequiredCapability)) {
                [void]$lines.Add(("        Case ""{0}""" -f $button.Id))
                [void]$lines.Add(("            RibbonRequiredCapabilityIsEnabledById = modRoleUiAccess.CanCurrentUserPerformCapabilityCached(""{0}"")" -f $button.RequiredCapability))
            }
        }
    }
    [void]$lines.Add("        Case Else")
    [void]$lines.Add("            RibbonRequiredCapabilityIsEnabledById = True")
    [void]$lines.Add("    End Select")
    [void]$lines.Add("    Exit Function")
    [void]$lines.Add("Disabled:")
    [void]$lines.Add("    RibbonRequiredCapabilityIsEnabledById = False")
    [void]$lines.Add("End Function")
    [void]$lines.Add("")
    [void]$lines.Add("Public Sub RibbonWarehouseGetItemCount(control As IRibbonControl, ByRef returnedVal)")
    [void]$lines.Add("    returnedVal = modRibbonRuntimeStatus.GetWarehouseTargetCount()")
    [void]$lines.Add("End Sub")
    [void]$lines.Add("")
    [void]$lines.Add("Public Sub RibbonWarehouseGetItemLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)")
    [void]$lines.Add("    returnedVal = modRibbonRuntimeStatus.GetWarehouseTargetLabel(index)")
    [void]$lines.Add("End Sub")
    [void]$lines.Add("")
    [void]$lines.Add("Public Sub RibbonWarehouseGetSelectedItemIndex(control As IRibbonControl, ByRef returnedVal)")
    [void]$lines.Add("    returnedVal = modRibbonRuntimeStatus.GetSelectedWarehouseTargetIndex()")
    [void]$lines.Add("End Sub")
    [void]$lines.Add("")
    [void]$lines.Add("Public Sub RibbonWarehouseOnAction(control As IRibbonControl, selectedId As String, selectedIndex As Integer)")
    [void]$lines.Add("    modRibbonRuntimeStatus.SelectWarehouseTarget selectedIndex")
    [void]$lines.Add("End Sub")

    $component = $VBProject.VBComponents.Add(1)
    $component.Name = "modRibbonGenerated"
    $component.CodeModule.AddFromString([string]::Join([Environment]::NewLine, $lines))
}

function Get-RibbonXml {
    param(
        [hashtable]$RibbonConfig
    )

    if ($null -eq $RibbonConfig) {
        return $null
    }

    $xml = New-Object System.Text.StringBuilder
    $enabledCallbackName = "RibbonRequiredCapabilityGetEnabled"
    if ($RibbonConfig.ContainsKey("EnabledCallbackName") -and -not [string]::IsNullOrWhiteSpace($RibbonConfig.EnabledCallbackName)) {
        $enabledCallbackName = [string]$RibbonConfig.EnabledCallbackName
    }

    [void]$xml.AppendLine("<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>")
    [void]$xml.AppendLine("<customUI xmlns=""http://schemas.microsoft.com/office/2006/01/customui"" onLoad=""RibbonOnLoad"">")
    [void]$xml.AppendLine("  <ribbon startFromScratch=""false"">")
    [void]$xml.AppendLine("    <tabs>")
    [void]$xml.AppendLine(("      <tab id=""{0}"" label=""{1}"">" -f $RibbonConfig.TabId, $RibbonConfig.Label))

    foreach ($group in $RibbonConfig.Groups) {
        [void]$xml.AppendLine(("        <group id=""{0}"" label=""{1}"">" -f $group.Id, $group.Label))
        foreach ($button in $group.Buttons) {
            if ($null -eq $button) {
                continue
            }
            $imageXml = ""
            $showImage = "false"
            $screentipXml = ""
            $labelXml = (' label="{0}"' -f $button.Label)
            if ($button.ContainsKey("ImageMso") -and -not [string]::IsNullOrWhiteSpace($button.ImageMso)) {
                $imageXml = (' imageMso="{0}"' -f $button.ImageMso)
                $showImage = "true"
            }
            if ($button.ContainsKey("Screentip") -and -not [string]::IsNullOrWhiteSpace($button.Screentip)) {
                $screentipXml = (' screentip="{0}"' -f $button.Screentip)
            }
            if ($button.ContainsKey("GetLabel") -and -not [string]::IsNullOrWhiteSpace($button.GetLabel)) {
                $labelXml = (' getLabel="{0}"' -f $button.GetLabel)
            }
            $enabledXml = ""
            if ($button.ContainsKey("RequiredCapability") -and -not [string]::IsNullOrWhiteSpace($button.RequiredCapability)) {
                $enabledXml = (' getEnabled="{0}"' -f $enabledCallbackName)
            }
            [void]$xml.AppendLine(("          <button id=""{0}""{1} size=""large"" showImage=""{2}""{3}{4}{5} onAction=""{6}""/>" -f $button.Id, $labelXml, $showImage, $imageXml, $screentipXml, $enabledXml, $RibbonConfig.CallbackName))
        }
        if ($group.ContainsKey("StatusLabels")) {
            foreach ($statusLabel in $group.StatusLabels) {
                if ($null -eq $statusLabel) {
                    continue
                }
                [void]$xml.AppendLine(("          <labelControl id=""{0}"" getLabel=""{1}""/>" -f $statusLabel.Id, $statusLabel.GetLabel))
            }
        }
        if ($group.ContainsKey("WarehouseSelector")) {
            $selector = $group.WarehouseSelector
            [void]$xml.AppendLine(("          <dropDown id=""{0}"" label=""{1}"" getItemCount=""RibbonWarehouseGetItemCount"" getItemLabel=""RibbonWarehouseGetItemLabel"" getSelectedItemIndex=""RibbonWarehouseGetSelectedItemIndex"" onAction=""RibbonWarehouseOnAction""/>" -f $selector.Id, $selector.Label))
        }
        if ($group.ContainsKey("StatusMenus")) {
            foreach ($menu in $group.StatusMenus) {
                if ($null -eq $menu) {
                    continue
                }
                $imageXml = ""
                $showImage = "false"
                if ($menu.ContainsKey("ImageMso") -and -not [string]::IsNullOrWhiteSpace($menu.ImageMso)) {
                    $imageXml = (' imageMso="{0}"' -f $menu.ImageMso)
                    $showImage = "true"
                }
                [void]$xml.AppendLine(("          <menu id=""{0}"" label=""{1}"" size=""large"" showImage=""{2}""{3}>" -f $menu.Id, $menu.Label, $showImage, $imageXml))
                foreach ($statusButton in $menu.StatusButtons) {
                    if ($null -eq $statusButton) {
                        continue
                    }
                    [void]$xml.AppendLine(("            <button id=""{0}"" getLabel=""RibbonRuntimeStatusGetLabel"" enabled=""false""/>" -f $statusButton.Id))
                }
                [void]$xml.AppendLine("            <menuSeparator id=""sepRuntimeContextRefresh""/>")
                [void]$xml.AppendLine(("            <button id=""{0}"" label=""Refresh / Details"" imageMso=""Refresh"" onAction=""RibbonRuntimeStatusRefresh""/>" -f $menu.RefreshButtonId))
                [void]$xml.AppendLine("          </menu>")
            }
        }
        if ($group.ContainsKey("PostStatusMenuButtons")) {
            foreach ($button in $group.PostStatusMenuButtons) {
                if ($null -eq $button) { continue }
                $imageXml = ""
                $showImage = "false"
                $screentipXml = ""
                $labelXml = (' label="{0}"' -f $button.Label)
                if ($button.ContainsKey("ImageMso") -and -not [string]::IsNullOrWhiteSpace($button.ImageMso)) {
                    $imageXml = (' imageMso="{0}"' -f $button.ImageMso)
                    $showImage = "true"
                }
                if ($button.ContainsKey("Screentip") -and -not [string]::IsNullOrWhiteSpace($button.Screentip)) {
                    $screentipXml = (' screentip="{0}"' -f $button.Screentip)
                }
                if ($button.ContainsKey("GetLabel") -and -not [string]::IsNullOrWhiteSpace($button.GetLabel)) {
                    $labelXml = (' getLabel="{0}"' -f $button.GetLabel)
                }
                $enabledXml = ""
                if ($button.ContainsKey("RequiredCapability") -and -not [string]::IsNullOrWhiteSpace($button.RequiredCapability)) {
                    $enabledXml = (' getEnabled="{0}"' -f $enabledCallbackName)
                }
                [void]$xml.AppendLine(("          <button id=""{0}""{1} size=""large"" showImage=""{2}""{3}{4}{5} onAction=""{6}""/>" -f $button.Id, $labelXml, $showImage, $imageXml, $screentipXml, $enabledXml, $RibbonConfig.CallbackName))
            }
        }
        [void]$xml.AppendLine("        </group>")
    }

    [void]$xml.AppendLine("      </tab>")
    [void]$xml.AppendLine("    </tabs>")
    [void]$xml.AppendLine("  </ribbon>")
    [void]$xml.AppendLine("</customUI>")
    $xml.ToString()
}

function Install-RibbonCustomUi {
    param(
        [string]$WorkbookPath,
        [hashtable]$RibbonConfig
    )

    if ($null -eq $RibbonConfig) {
        return
    }

    $ribbonXml = Get-RibbonXml -RibbonConfig $RibbonConfig
    $document = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Open($WorkbookPath, $true)
    try {
        $existingPart = $document.RibbonExtensibilityPart
        if ($null -ne $existingPart) {
            $document.DeletePart($existingPart)
        }

        $part = $document.AddRibbonExtensibilityPart()
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($ribbonXml)
        $stream = New-Object System.IO.MemoryStream(,$bytes)
        try {
            $part.FeedData($stream)
        }
        finally {
            $stream.Dispose()
        }
    }
    finally {
        $document.Dispose()
    }
}

function Add-ReferenceByPath {
    param(
        [object]$VBProject,
        [string]$ReferencePath
    )

    foreach ($ref in $VBProject.References) {
        if ($ref.FullPath -and ([string]::Equals($ref.FullPath, $ReferencePath, [System.StringComparison]::OrdinalIgnoreCase))) {
            return
        }
    }

    [void]$VBProject.References.AddFromFile($ReferencePath)
}

function Add-ReferenceByGuidSafe {
    param(
        [object]$VBProject,
        [string]$Guid,
        [int]$Major,
        [int]$Minor
    )

    foreach ($ref in $VBProject.References) {
        if ($ref.Guid -eq $Guid) {
            return
        }
    }

    try {
        [void]$VBProject.References.AddFromGuid($Guid, $Major, $Minor)
    }
    catch {
        Write-Warning "Unable to add reference $Guid ($Major.$Minor): $($_.Exception.Message)"
    }
}

function Remove-ExistingFile {
    param(
        [string]$Path
    )

    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Force
    }
}

function Write-FivePackageManifest {
    param([string]$OutputDir)

    $expectedNames = @(
        "invSys.Core.xlam",
        "invSys.Inventory.Domain.xlam",
        "invSys.Designs.Domain.xlam",
        "invSys.Operations.xlam",
        "invSys.Admin.xlam"
    )
    $actualNames = @(
        Get-ChildItem -LiteralPath $OutputDir -Filter "*.xlam" -File |
            Select-Object -ExpandProperty Name |
            Sort-Object
    )
    $differences = @(Compare-Object ($expectedNames | Sort-Object) $actualNames)
    if ($actualNames.Count -ne $expectedNames.Count -or $differences.Count -gt 0) {
        throw "Cannot publish the R1-5 package manifest because the XLAM set is not exactly the normative five."
    }

    $packages = @()
    foreach ($name in $expectedNames) {
        $path = Join-Path $OutputDir $name
        $item = Get-Item -LiteralPath $path
        $packages += [ordered]@{
            name = $name
            packageSetVersion = "R1-5"
            sizeBytes = [long]$item.Length
            sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
        }
    }

    $manifest = [ordered]@{
        schemaVersion = "1.0.0"
        packageSetVersion = "R1-5"
        generatedUtc = [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")
        packages = $packages
    }
    $manifestPath = Join-Path $OutputDir "addins-manifest.json"
    [IO.File]::WriteAllText(
        $manifestPath,
        (($manifest | ConvertTo-Json -Depth 5) + "`n"),
        (New-Object Text.UTF8Encoding($false))
    )
    Write-Host ("Published " + $manifestPath)
}

$repo = (Resolve-Path $RepoRoot).Path
if ([IO.Path]::IsPathRooted($OutputRoot)) {
    $outputDir = [IO.Path]::GetFullPath($OutputRoot)
}
else {
    $outputDir = Join-Path $repo $OutputRoot
}

$projectMap = @(
    @{
        Key        = "Core"
        Project    = "invSys_Core"
        OutputFile = "invSys.Core.xlam"
        LegacyOutputFiles = @()
        SourceDirs = @((Join-Path $repo "src/Core"))
        References = @()
        Sheets     = @("INVENTORY MANAGEMENT", "ErrorLog", "Notes", "TestSummary")
        AddVbideReference = $true
        Ribbon     = $null
    }
    @{
        Key        = "InventoryDomain"
        Project    = "invSys_Inventory_Domain"
        OutputFile = "invSys.Inventory.Domain.xlam"
        LegacyOutputFiles = @()
        SourceDirs = @((Join-Path $repo "src/InventoryDomain"))
        ExcludeFiles = @("modInvMan.bas", "cInventoryAppEvents.cls")
        References = @("Core")
        Sheets     = @("INVENTORY MANAGEMENT", "InventoryLog", "AppliedEvents", "Locks")
        AddVbideReference = $false
        Ribbon     = $null
    }
    @{
        Key        = "DesignsDomain"
        Project    = "invSys_Designs_Domain"
        OutputFile = "invSys.Designs.Domain.xlam"
        LegacyOutputFiles = @()
        SourceDirs = @((Join-Path $repo "src/DesignsDomain"))
        References = @("Core")
        Sheets     = @()
        AddVbideReference = $false
        Ribbon     = $null
    }
    @{
        Key        = "Operations"
        Project    = "invSys_Operations"
        OutputFile = "invSys.Operations.xlam"
        LegacyOutputFiles = @(
            "invSysReceiving.xlam",
            "invSys.Receiving.xlam",
            "invSys.Production.xlam",
            "invSys.Shipping.xlam"
        )
        SourceDirs = @(
            (Join-Path $repo "src/Operations"),
            (Join-Path $repo "src/Receiving"),
            (Join-Path $repo "src/Production"),
            (Join-Path $repo "src/Shipping")
        )
        ExcludeFiles = @(
            "modReceivingAutoOpen.bas",
            "modProductionAutoOpen.bas",
            "modShippingAutoOpen.bas"
        )
        References = @("Core")
        Sheets     = @(
            "ReceivedTally",
            "InventoryManagement",
            "ReceivedLog",
            "ShipmentsTally",
            "Production",
            "Recipes"
        )
        AddVbideReference = $false
        Ribbon     = @{
            TabId  = "tabInvSysOperations"
            Label  = "Operations"
            CallbackName = "RibbonOnActionOperations"
            EnabledCallbackName = "RibbonRequiredCapabilityGetEnabledOperations"
            Groups = @(
                @{
                    Id      = "grpOperationsSession"
                    Label   = "Session"
                    WarehouseSelector = @{
                        Id = "ddOperationsWarehouseTarget"
                        Label = "Send To"
                    }
                    StatusMenus = @(
                        @{
                            Id = "mnuOperationsRuntimeContext"
                            Label = "Runtime Context"
                            ImageMso = "Info"
                            RefreshButtonId = "btnOperationsRuntimeRefresh"
                            StatusButtons = @(
                                @{ Id = "btnRuntimeWarehouse" },
                                @{ Id = "btnRuntimeDataRoot" },
                                @{ Id = "btnRuntimeInboxRoot" },
                                @{ Id = "btnRuntimeUser" },
                                @{ Id = "btnRuntimeProcessor" },
                                @{ Id = "btnRuntimeHqAggregator" }
                            )
                        }
                    )
                    Buttons = @(
                        @{ Id = "btnOperationsServerSession"; Label = "Server Sign In"; GetLabel = "RibbonServerSessionGetLabel"; DirectAction = "modRoleEventWriter.ToggleServerSessionForCapability"; ImageMso = "FileOpen"; Screentip = "Sign in to or sign out of warehouse server storage" }
                    )
                    PostStatusMenuButtons = @(
                        @{ Id = "btnOperationsCurrentUser"; Label = "invSys Sign In"; GetLabel = "RibbonCurrentUserGetLabel"; DirectAction = "modRoleEventWriter.ToggleCurrentInvSysUserForCapability"; ImageMso = "AddressBook"; Screentip = "Sign in to or sign out of invSys" }
                    )
                    StatusLabels = @(
                        @{ Id = "lblOperationsServerStatus"; GetLabel = "RibbonServerStatusGetLabel" },
                        @{ Id = "lblOperationsAccessStatus"; GetLabel = "RibbonAccessStatusGetLabel" }
                    )
                },
                @{
                    Id      = "grpOperationsOverview"
                    Label   = "Overview"
                    Buttons = @(
                        @{ Id = "btnOperationsInventoryViewer"; Label = "Viewer"; Macro = "modInventoryViewer.OpenInventoryViewer"; ImageMso = "PivotTableInsert"; Screentip = "View current inventory levels" }
                    )
                },
                @{
                    Id      = "grpOperationsReceiving"
                    Label   = "Receiving"
                    Buttons = @(
                        @{ Id = "btnOperationsReceivingForm"; Label = "Receiving"; Macro = "modTS_Received.ShowReceivingForm"; ImageMso = "FormControlButton"; RequiredCapability = "RECEIVE_POST" }
                    )
                },
                @{
                    Id      = "grpOperationsProduction"
                    Label   = "Production"
                    Buttons = @(
                        @{ Id = "btnOperationsProductionForm"; Label = "Production"; Macro = "mProduction.BtnOpenProductionForm"; ImageMso = "CreateForm"; RequiredCapability = "PROD_POST" }
                    )
                },
                @{
                    Id      = "grpOperationsShipping"
                    Label   = "Shipping"
                    Buttons = @(
                        @{ Id = "btnOperationsShippingForm"; Label = "Shipping"; Macro = "modTS_Shipments.BtnOpenShipmentsForm"; ImageMso = "FileSendAsAttachment"; RequiredCapability = "SHIP_POST" }
                    )
                }
            )
        }
    }
    @{
        Key        = "OperationsShadow"
        Project    = "invSys_Operations_Shadow"
        OutputFile = "invSys.Operations.xlam"
        LegacyOutputFiles = @()
        Deployable = $false
        SourceDirs = @(
            (Join-Path $repo "src/Operations"),
            (Join-Path $repo "src/Receiving"),
            (Join-Path $repo "src/Production"),
            (Join-Path $repo "src/Shipping")
        )
        ExcludeFiles = @(
            "modReceivingAutoOpen.bas",
            "modProductionAutoOpen.bas",
            "modShippingAutoOpen.bas"
        )
        References = @("Core")
        Sheets     = @(
            "ReceivedTally",
            "InventoryManagement",
            "ReceivedLog",
            "ShipmentsTally",
            "Production",
            "Recipes"
        )
        AddVbideReference = $false
        Ribbon     = $null
    }
    @{
        Key        = "Admin"
        Project    = "invSys_Admin"
        OutputFile = "invSys.Admin.xlam"
        LegacyOutputFiles = @()
        SourceDirs = @((Join-Path $repo "src/Admin"))
        References = @("Core")
        Sheets     = @("UserCredentials", "Emails")
        AddVbideReference = $false
        Ribbon     = @{
            TabId  = "tabInvSysAdmin"
            Label  = "invSys Admin"
            CallbackName = "RibbonOnActionAdmin"
            EnabledCallbackName = "RibbonRequiredCapabilityGetEnabledAdmin"
            Groups = @(
                @{
                    Id      = "grpAdminSession"
                    Label   = "Session"
                    WarehouseSelector = @{
                        Id = "ddAdminWarehouseTarget"
                        Label = "Send To"
                    }
                    Buttons = @(
                        @{ Id = "btnAdminServerSession"; Label = "Server Sign In"; GetLabel = "RibbonServerSessionGetLabel"; DirectAction = "modRoleEventWriter.ToggleServerSessionForCapability ""ADMIN_MAINT"""; ImageMso = "FileOpen"; Screentip = "Sign in to or sign out of warehouse server storage" },
                        @{ Id = "btnAdminCurrentUser"; Label = "invSys Sign In"; GetLabel = "RibbonCurrentUserGetLabel"; DirectAction = "modRoleEventWriter.ToggleCurrentInvSysUserForCapability ""ADMIN_MAINT"""; ImageMso = "AddressBook"; Screentip = "Sign in to or sign out of invSys" }
                    )
                    StatusLabels = @(
                        @{ Id = "lblAdminServerStatus"; GetLabel = "RibbonServerStatusGetLabel" },
                        @{ Id = "lblAdminAccessStatus"; GetLabel = "RibbonAccessStatusGetLabel" }
                    )
                }
                @{
                    Id      = "grpAdminActions"
                    Label   = "Actions"
                    Buttons = @(
                        @{ Id = "btnAdminOpen"; Label = "Admin Console"; Macro = "modAdmin.Admin_Click"; ImageMso = "FileOpen"; RequiredCapability = "ADMIN_MAINT" },
                        @{ Id = "btnAdminUsers"; Label = "Users and Roles"; Macro = "modAdmin.Open_CreateDeleteUser"; ImageMso = "FileOpen"; RequiredCapability = "ADMIN_MAINT" },
                        @{ Id = "btnAdminSettings"; Label = "Settings"; Macro = "modAdmin.Open_Settings"; ImageMso = "FileProperties"; RequiredCapability = "ADMIN_MAINT" },
                        @{ Id = "btnAdminWarehouses"; Label = "View Warehouses"; Macro = "modAdmin.Open_WarehouseDirectory"; ImageMso = "TablePropertiesDialog"; RequiredCapability = "ADMIN_MAINT" },
                        @{ Id = "btnAdminWarehouseRoot"; Label = "Add Warehouse Root"; Macro = "modAdmin.Add_WarehouseDirectoryRoot"; ImageMso = "Folder"; RequiredCapability = "ADMIN_MAINT" },
                        @{ Id = "btnAdminCreateWarehouse"; Label = "Create New Warehouse"; Macro = "modAdmin.Open_CreateWarehouse"; ImageMso = "FileNew"; RequiredCapability = "ADMIN_MAINT" },
                        @{ Id = "btnAdminSetupTesterStation"; Label = "Test Environment Setup"; Macro = "modAdmin.Admin_SetupTesterStation_Click"; ImageMso = "CreateForm"; Screentip = "Provision an isolated warehouse and operator workbook for diagnostics and regression testing"; RequiredCapability = "ADMIN_MAINT" },
                        @{ Id = "btnAdminAddInventoryItem"; Label = "Add/Edit Inventory Items"; Macro = "modAdmin.Add_InventoryItem"; ImageMso = "TableInsertRowsAbove"; RequiredCapability = "ADMIN_MAINT" },
                        @{ Id = "btnAdminSeedInventory"; Label = "Demo Inventory"; Macro = "modAdmin.Seed_DemoInventory"; ImageMso = "TableInsertRowsAbove"; Screentip = "Seed, delete, or upload guarded demo inventory"; RequiredCapability = "ADMIN_MAINT" },
                        @{ Id = "btnAdminShipmentReconcile"; Label = "Shipment Reconcile"; Macro = "modAdminShipmentReconcile.OpenShipmentReconcileTool"; ImageMso = "RefreshAll"; Screentip = "Queue an audited admin correction linked to a Shipments Sent event"; RequiredCapability = "ADMIN_MAINT" },
                        @{ Id = "btnAdminHqAggregator"; Label = "Aggregate Global Snapshot"; Macro = "modAdmin.Admin_AggregateGlobalSnapshot_Click"; ImageMso = "RefreshAll"; Screentip = "Rebuild the advisory-only global inventory snapshot from published warehouse snapshots"; RequiredCapability = "ADMIN_MAINT" },
                        @{ Id = "btnAdminDesignLifecycle"; Label = "Design Lifecycle"; Macro = "modAdminDesignLifecycle.Admin_DesignLifecycle_Click"; ImageMso = "AcceptInvitation"; Screentip = "Release or obsolete immutable Designs Domain recipes"; RequiredCapability = "ADMIN_MAINT" },
                        @{ Id = "btnAdminVerifyAddinsPublished"; Label = "Verify Add-ins Published"; Macro = "modAdmin.Verify_AddinsPublished"; ImageMso = "FileDocumentInspect"; RequiredCapability = "ADMIN_MAINT" },
                        @{ Id = "btnAdminRetireMigrateWarehouse"; Label = "Retire / Migrate Warehouse"; Macro = "modAdmin.Admin_RetireMigrateWarehouse_Click"; ImageMso = "DeleteSite"; Screentip = "Archive, migrate, retire, or delete a warehouse runtime"; RequiredCapability = "ADMIN_MAINT" }
                    )
                }
            )
        }
    }
)

$availableProjects = @($projectMap)
if (-not $IncludeOperationsShadow) {
    $availableProjects = @(
        $availableProjects |
            Where-Object { $_.Key -ne "OperationsShadow" }
    )
}

$requestedKeys = @($Projects | Where-Object {
    -not [string]::IsNullOrWhiteSpace([string]$_)
})
if ($requestedKeys.Count -eq 0) {
    $requestedKeys = @($availableProjects | ForEach-Object { $_.Key })
}

$availableByKey = @{}
foreach ($availableProject in $availableProjects) {
    $availableByKey[[string]$availableProject.Key] = $availableProject
}
foreach ($requestedKey in $requestedKeys) {
    if (-not $availableByKey.ContainsKey([string]$requestedKey)) {
        throw "Unknown or unavailable project '$requestedKey'."
    }
}

$selectedKeys = @{}
foreach ($requestedKey in $requestedKeys) {
    $selectedKeys[[string]$requestedKey] = $true
}
$selectionChanged = $true
while ($selectionChanged) {
    $selectionChanged = $false
    foreach ($selectedKey in @($selectedKeys.Keys)) {
        $selectedProject = $availableByKey[$selectedKey]
        foreach ($referenceKey in @($selectedProject.References)) {
            if (-not $selectedKeys.ContainsKey([string]$referenceKey)) {
                if (-not $availableByKey.ContainsKey([string]$referenceKey)) {
                    throw "Dependency '$referenceKey' for '$selectedKey' is unavailable."
                }
                $selectedKeys[[string]$referenceKey] = $true
                $selectionChanged = $true
            }
        }
    }
}
$projectMap = @(
    $availableProjects |
        Where-Object { $selectedKeys.ContainsKey([string]$_.Key) }
)

$currentDeployRoot = [IO.Path]::GetFullPath(
    (Join-Path $repo "deploy\current")
).TrimEnd("\")
$resolvedOutputRoot = [IO.Path]::GetFullPath($outputDir).TrimEnd("\")
foreach ($project in $projectMap) {
    if ($project.ContainsKey("Deployable") -and
        -not [bool]$project.Deployable -and
        [string]::Equals(
            $currentDeployRoot,
            $resolvedOutputRoot,
            [StringComparison]::OrdinalIgnoreCase
        )) {
        throw "Non-deployable project '$($project.Key)' cannot target deploy/current."
    }
}

Write-Host "invSys build-xlam.ps1"
Write-Host "RepoRoot: $repo"
Write-Host "OutputRoot: $outputDir"

if (-not (Test-Path -LiteralPath $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

Write-Host "Planned outputs:"
foreach ($project in $projectMap) {
    Write-Host ("- " + (Join-Path $outputDir $project.OutputFile))
}

$legacyArtifacts = @()
foreach ($project in $projectMap) {
    foreach ($legacyName in $project.LegacyOutputFiles) {
        $legacyPath = Join-Path $outputDir $legacyName
        if (Test-Path -LiteralPath $legacyPath) {
            $legacyArtifacts += [pscustomobject]@{
                Project = $project.Key
                Path    = $legacyPath
                Name    = $legacyName
            }
        }
    }
}

if ($legacyArtifacts.Count -gt 0) {
    Write-Host "Legacy outputs queued for archive:"
    foreach ($artifact in $legacyArtifacts) {
        Write-Host ("- " + $artifact.Path)
    }
}

if (-not $Apply) {
    Write-Host "Dry run only. Re-run with -Apply to build the XLAMs."
    exit 0
}

$archiveDir = Join-Path (Split-Path $outputDir -Parent) "archive"
$stagingDir = Join-Path $outputDir (".build-staging-" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $stagingDir -Force | Out-Null

$referenceDir = Join-Path $outputDir ".refs"
if (Test-Path -LiteralPath $referenceDir) {
    Write-Host ("Removing legacy reference copy directory " + $referenceDir)
    Remove-Item -LiteralPath $referenceDir -Recurse -Force
}

$builtOutputs = @{}
$excel = $null
try {
    if (($legacyArtifacts.Count -gt 0) -and (-not (Test-Path -LiteralPath $archiveDir))) {
        New-Item -ItemType Directory -Path $archiveDir -Force | Out-Null
    }

    foreach ($artifact in $legacyArtifacts) {
        $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $archivePath = Join-Path $archiveDir (($artifact.Project + "." + $timestamp + "." + $artifact.Name))
        Write-Host ("Archiving legacy output " + $artifact.Path + " -> " + $archivePath)
        Move-Item -LiteralPath $artifact.Path -Destination $archivePath -Force
    }

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1

    foreach ($project in $projectMap) {
        Write-Host ("Building " + $project.OutputFile + " ...")
        $wb = $null
        try {
            $codeFiles = @(Get-CodeFiles -SourceDirs $project.SourceDirs)
            if ($project.ContainsKey("ExcludeFiles")) {
                $excludedNames = @($project.ExcludeFiles)
                $codeFiles = @($codeFiles | Where-Object { $_.Name -notin $excludedNames })
            }
            $sheetFiles = @(Get-SheetModuleFiles -CodeFiles $codeFiles)
            $importFiles = @(Get-ImportFiles -CodeFiles $codeFiles)
            $formFiles = @(Get-FormFiles -CodeFiles $codeFiles)
            $wb = $excel.Workbooks.Add()
            $vbProject = $wb.VBProject
            $vbProject.Name = $project.Project

            if ($project.AddVbideReference) {
                Write-Host "  Adding VBIDE reference"
                Add-ReferenceByGuidSafe -VBProject $vbProject -Guid "{0002E157-0000-0000-C000-000000000046}" -Major 5 -Minor 3
            }

            foreach ($referenceKey in $project.References) {
                if (-not $builtOutputs.ContainsKey($referenceKey)) {
                    throw "Referenced project '$referenceKey' has not been built yet."
                }

                $referenceProject = $projectMap | Where-Object { $_.Key -eq $referenceKey } | Select-Object -First 1
                if ($null -eq $referenceProject) {
                    throw "Referenced project '$referenceKey' is not defined in projectMap."
                }

                $referencePath = Get-DeployedOutputPath -Project $referenceProject -OutputDir $outputDir
                if (-not (Test-Path -LiteralPath $referencePath)) {
                    throw "Referenced project output is not published yet: $referencePath"
                }

                Write-Host ("  Adding project reference " + $referenceKey + " -> " + $referencePath)
                Add-ReferenceByPath -VBProject $vbProject -ReferencePath $referencePath
            }

            if ($project.Sheets.Count -gt 0) {
                Write-Host "  Preparing placeholder worksheets"
                Ensure-WorksheetNames -Workbook $wb -SheetNames $project.Sheets
            }

            Write-Host "  Importing standard/class/form components"
            Import-Components -VBProject $vbProject -Files $importFiles
            Import-Forms -VBProject $vbProject -FormFiles $formFiles
            Add-RibbonCallbacksModule -VBProject $vbProject -RibbonConfig $project.Ribbon

            $wb.IsAddin = $true
            $outputPath = Join-Path $stagingDir $project.OutputFile
            Remove-ExistingFile -Path $outputPath
            Write-Host ("  Saving " + $outputPath)
            $wb.SaveAs($outputPath, 55)
            $builtOutputs[$project.Key] = $outputPath
            if ($null -ne $project.Ribbon) {
                Write-Host "  Installing RibbonX package"
                Install-RibbonCustomUi -WorkbookPath $outputPath -RibbonConfig $project.Ribbon
            }
            Write-Host ("Built " + $outputPath)
        }
        finally {
            if ($null -ne $wb) {
                try { $wb.Close($false) } catch {}
                Release-ComObject $wb
            }
        }

        $stagedPath = $builtOutputs[$project.Key]
        $finalPath = Join-Path $outputDir $project.OutputFile
        Remove-ExistingFile -Path $finalPath
        Copy-Item -LiteralPath $stagedPath -Destination $finalPath -Force
        Write-Host ("Published " + $finalPath)
    }
}
finally {
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        Release-ComObject $excel
        $excel = $null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    if (Test-Path -LiteralPath $stagingDir) {
        try {
            Remove-Item -LiteralPath $stagingDir -Recurse -Force
        }
        catch {
            Write-Warning ("Could not remove staging directory " + $stagingDir + ": " + $_.Exception.Message)
        }
    }
}

if ([string]::Equals(
    $currentDeployRoot,
    $resolvedOutputRoot,
    [StringComparison]::OrdinalIgnoreCase
)) {
    Write-FivePackageManifest -OutputDir $outputDir
}

Write-Host "Build complete."
