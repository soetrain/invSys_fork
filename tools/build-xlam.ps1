[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$RepoRoot = ".",

    [Parameter(Mandatory = $false)]
    [string]$OutputRoot = "deploy/current",

    [Parameter(Mandatory = $false)]
    [switch]$Apply
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Release-ComObject {
    param([object]$Obj)
    if ($null -ne $Obj) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Obj) } catch {}
    }
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
            $componentName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
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

                [void]$codeLines.Add($line)
            }

            $component = $VBProject.VBComponents.Add(2)
            $component.Name = $componentName
            $module = $component.CodeModule
            if ($module.CountOfLines -gt 0) {
                $module.DeleteLines(1, $module.CountOfLines)
            }
            $module.AddFromString(([string]::Join([Environment]::NewLine, $codeLines)))
            continue
        }

        Write-Host ("  Importing " + $file.FullName)
        [void]$VBProject.VBComponents.Import($file.FullName)
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
    $module.AddFromString("Option Explicit")
}

function Import-Forms {
    param(
        [object]$VBProject,
        [System.IO.FileInfo[]]$FormFiles
    )

    foreach ($formFile in $FormFiles) {
        if (Test-FormRequiresStub -FormFile $formFile) {
            Add-StubUserForm -VBProject $VBProject -FormFile $formFile
        }
        else {
            Write-Host ("  Importing " + $formFile.FullName)
            [void]$VBProject.VBComponents.Import($formFile.FullName)
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

    $lines = New-Object System.Collections.Generic.List[string]
    [void]$lines.Add("Option Explicit")
    [void]$lines.Add("")
    [void]$lines.Add("Private mRibbon As Object")
    [void]$lines.Add("")
    [void]$lines.Add("Public Sub RibbonOnLoad(ribbon As Object)")
    [void]$lines.Add("    Set mRibbon = ribbon")
    [void]$lines.Add("End Sub")
    [void]$lines.Add("")
    [void]$lines.Add("Public Sub RibbonOnAction(control As Object)")
    [void]$lines.Add("    On Error GoTo ErrHandler")
    [void]$lines.Add("    Select Case control.ID")

    foreach ($group in $RibbonConfig.Groups) {
        foreach ($button in $group.Buttons) {
            [void]$lines.Add(("        Case ""{0}""" -f $button.Id))
            [void]$lines.Add(("            Application.Run ""'"" & ThisWorkbook.Name & ""'!{0}""" -f $button.Macro))
        }
    }

    [void]$lines.Add("    End Select")
    [void]$lines.Add("    Exit Sub")
    [void]$lines.Add("ErrHandler:")
    [void]$lines.Add('    MsgBox "Ribbon action failed: " & Err.Description, vbExclamation')
    [void]$lines.Add("End Sub")
    [void]$lines.Add("")
    [void]$lines.Add("Public Sub RibbonInvalidate()")
    [void]$lines.Add("    If Not mRibbon Is Nothing Then mRibbon.Invalidate")
    [void]$lines.Add("End Sub")

    $component = $VBProject.VBComponents.Item("ThisWorkbook")
    $module = $component.CodeModule
    $insertAt = $module.CountOfLines + 1
    $module.InsertLines($insertAt, [string]::Join([Environment]::NewLine, $lines))
}

function Get-RibbonXml {
    param(
        [hashtable]$RibbonConfig
    )

    if ($null -eq $RibbonConfig) {
        return $null
    }

    $xml = New-Object System.Text.StringBuilder
    [void]$xml.AppendLine("<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>")
    [void]$xml.AppendLine("<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""ThisWorkbook.RibbonOnLoad"">")
    [void]$xml.AppendLine("  <ribbon startFromScratch=""false"">")
    [void]$xml.AppendLine("    <tabs>")
    [void]$xml.AppendLine(("      <tab id=""{0}"" label=""{1}"">" -f $RibbonConfig.TabId, $RibbonConfig.Label))

    foreach ($group in $RibbonConfig.Groups) {
        [void]$xml.AppendLine(("        <group id=""{0}"" label=""{1}"">" -f $group.Id, $group.Label))
        foreach ($button in $group.Buttons) {
            [void]$xml.AppendLine(("          <button id=""{0}"" label=""{1}"" size=""large"" onAction=""ThisWorkbook.RibbonOnAction""/>" -f $button.Id, $button.Label))
        }
        [void]$xml.AppendLine("        </group>")
    }

    [void]$xml.AppendLine("      </tab>")
    [void]$xml.AppendLine("    </tabs>")
    [void]$xml.AppendLine("  </ribbon>")
    [void]$xml.AppendLine("</customUI>")
    $xml.ToString()
}

function Set-ZipEntryText {
    param(
        [System.IO.Compression.ZipArchive]$Zip,
        [string]$EntryName,
        [string]$Text
    )

    $existing = $Zip.GetEntry($EntryName)
    if ($null -ne $existing) {
        $existing.Delete()
    }

    $entry = $Zip.CreateEntry($EntryName)
    $stream = $entry.Open()
    $writer = New-Object System.IO.StreamWriter($stream)
    try {
        $writer.Write($Text)
    }
    finally {
        $writer.Dispose()
    }
}

function Get-ZipEntryText {
    param(
        [System.IO.Compression.ZipArchive]$Zip,
        [string]$EntryName
    )

    $entry = $Zip.GetEntry($EntryName)
    if ($null -eq $entry) {
        throw "Zip entry not found: $EntryName"
    }

    $reader = New-Object System.IO.StreamReader($entry.Open())
    try {
        return $reader.ReadToEnd()
    }
    finally {
        $reader.Dispose()
    }
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
    $zip = [System.IO.Compression.ZipFile]::Open($WorkbookPath, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        Set-ZipEntryText -Zip $zip -EntryName 'customUI/customUI14.xml' -Text $ribbonXml

        [xml]$relsXml = Get-ZipEntryText -Zip $zip -EntryName '_rels/.rels'
        $relsNs = 'http://schemas.openxmlformats.org/package/2006/relationships'
        $relsMgr = New-Object System.Xml.XmlNamespaceManager($relsXml.NameTable)
        $relsMgr.AddNamespace('r', $relsNs)
        $uiType = 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility'
        $existingRel = $relsXml.SelectSingleNode("//r:Relationship[@Type='$uiType']", $relsMgr)
        if ($null -eq $existingRel) {
            $newRel = $relsXml.CreateElement('Relationship', $relsNs)
            $newRel.SetAttribute('Id', 'rIdInvSysCustomUi')
            $newRel.SetAttribute('Type', $uiType)
            $newRel.SetAttribute('Target', 'customUI/customUI14.xml')
            [void]$relsXml.DocumentElement.AppendChild($newRel)
        }
        else {
            $existingRel.SetAttribute('Target', 'customUI/customUI14.xml')
        }
        Set-ZipEntryText -Zip $zip -EntryName '_rels/.rels' -Text $relsXml.OuterXml

        [xml]$ctXml = Get-ZipEntryText -Zip $zip -EntryName '[Content_Types].xml'
        $ctNs = 'http://schemas.openxmlformats.org/package/2006/content-types'
        $ctMgr = New-Object System.Xml.XmlNamespaceManager($ctXml.NameTable)
        $ctMgr.AddNamespace('ct', $ctNs)
        $existingOverride = $ctXml.SelectSingleNode("//ct:Override[@PartName='/customUI/customUI14.xml']", $ctMgr)
        if ($null -eq $existingOverride) {
            $override = $ctXml.CreateElement('Override', $ctNs)
            $override.SetAttribute('PartName', '/customUI/customUI14.xml')
            $override.SetAttribute('ContentType', 'application/vnd.ms-office.customUI+xml')
            [void]$ctXml.DocumentElement.AppendChild($override)
        }
        Set-ZipEntryText -Zip $zip -EntryName '[Content_Types].xml' -Text $ctXml.OuterXml
    }
    finally {
        $zip.Dispose()
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

$repo = (Resolve-Path $RepoRoot).Path
$outputDir = Join-Path $repo $OutputRoot

$projectMap = @(
    @{
        Key        = "Core"
        Project    = "invSys_Core"
        OutputFile = "invSys.Core.xlam"
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
        SourceDirs = @((Join-Path $repo "src/InventoryDomain"))
        References = @("Core")
        Sheets     = @("INVENTORY MANAGEMENT", "InventoryLog", "AppliedEvents", "Locks")
        AddVbideReference = $false
        Ribbon     = $null
    }
    @{
        Key        = "DesignsDomain"
        Project    = "invSys_Designs_Domain"
        OutputFile = "invSys.Designs.Domain.xlam"
        SourceDirs = @((Join-Path $repo "src/DesignsDomain"))
        References = @("Core")
        Sheets     = @()
        AddVbideReference = $false
        Ribbon     = $null
    }
    @{
        Key        = "Receiving"
        Project    = "invSys_Receiving"
        OutputFile = "invSys.Receiving.xlam"
        SourceDirs = @((Join-Path $repo "src/Receiving"))
        References = @("Core")
        Sheets     = @("ReceivedTally", "InventoryManagement", "ReceivedLog")
        AddVbideReference = $false
        Ribbon     = @{
            TabId  = "tabInvSysReceiving"
            Label  = "invSys Receiving"
            Groups = @(
                @{
                    Id      = "grpReceivingActions"
                    Label   = "Actions"
                    Buttons = @(
                        @{ Id = "btnReceivingSetup"; Label = "Setup UI"; Macro = "modTS_Received.EnsureGeneratedButtons" },
                        @{ Id = "btnReceivingConfirm"; Label = "Confirm Writes"; Macro = "modTS_Received.ConfirmWrites" },
                        @{ Id = "btnReceivingUndo"; Label = "Undo"; Macro = "modTS_Received.MacroUndo" },
                        @{ Id = "btnReceivingRedo"; Label = "Redo"; Macro = "modTS_Received.MacroRedo" }
                    )
                }
            )
        }
    }
    @{
        Key        = "Shipping"
        Project    = "invSys_Shipping"
        OutputFile = "invSys.Shipping.xlam"
        SourceDirs = @((Join-Path $repo "src/Shipping"))
        References = @("Core", "InventoryDomain")
        Sheets     = @("ShipmentsTally", "InventoryManagement", "ShippingBOM")
        AddVbideReference = $false
        Ribbon     = @{
            TabId  = "tabInvSysShipping"
            Label  = "invSys Shipping"
            Groups = @(
                @{
                    Id      = "grpShippingActions"
                    Label   = "Actions"
                    Buttons = @(
                        @{ Id = "btnShippingSetup"; Label = "Setup UI"; Macro = "modTS_Shipments.InitializeShipmentsUI" },
                        @{ Id = "btnShippingConfirm"; Label = "Confirm Inventory"; Macro = "modTS_Shipments.BtnConfirmInventory" },
                        @{ Id = "btnShippingStage"; Label = "To Shipments"; Macro = "modTS_Shipments.BtnToShipments" },
                        @{ Id = "btnShippingSend"; Label = "Shipments Sent"; Macro = "modTS_Shipments.BtnShipmentsSent" }
                    )
                }
            )
        }
    }
    @{
        Key        = "Production"
        Project    = "invSys_Production"
        OutputFile = "invSys.Production.xlam"
        SourceDirs = @((Join-Path $repo "src/Production"))
        References = @("Core", "InventoryDomain", "DesignsDomain")
        Sheets     = @("Production", "InventoryManagement", "Recipes")
        AddVbideReference = $false
        Ribbon     = @{
            TabId  = "tabInvSysProduction"
            Label  = "invSys Production"
            Groups = @(
                @{
                    Id      = "grpProductionActions"
                    Label   = "Actions"
                    Buttons = @(
                        @{ Id = "btnProductionSetup"; Label = "Setup UI"; Macro = "mProduction.InitializeProductionUI" },
                        @{ Id = "btnProductionLoad"; Label = "Load Recipe"; Macro = "mProduction.BtnLoadRecipe" },
                        @{ Id = "btnProductionUsed"; Label = "To Used"; Macro = "mProduction.BtnToUsed" },
                        @{ Id = "btnProductionMade"; Label = "To Made"; Macro = "mProduction.BtnToMade" },
                        @{ Id = "btnProductionTotal"; Label = "To Total Inv"; Macro = "mProduction.BtnToTotalInv" }
                    )
                }
            )
        }
    }
    @{
        Key        = "Admin"
        Project    = "invSys_Admin"
        OutputFile = "invSys.Admin.xlam"
        SourceDirs = @((Join-Path $repo "src/Admin"))
        References = @("Core", "InventoryDomain", "DesignsDomain")
        Sheets     = @("UserCredentials", "Emails")
        AddVbideReference = $false
        Ribbon     = @{
            TabId  = "tabInvSysAdmin"
            Label  = "invSys Admin"
            Groups = @(
                @{
                    Id      = "grpAdminActions"
                    Label   = "Actions"
                    Buttons = @(
                        @{ Id = "btnAdminOpen"; Label = "Admin Console"; Macro = "modAdmin.Admin_Click" },
                        @{ Id = "btnAdminUsers"; Label = "Users and Roles"; Macro = "modAdmin.Open_CreateDeleteUser" }
                    )
                }
            )
        }
    }
)

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

if (-not $Apply) {
    Write-Host "Dry run only. Re-run with -Apply to build the XLAMs."
    exit 0
}

$builtOutputs = @{}
$excel = $null
try {
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
                Write-Host ("  Adding project reference " + $referenceKey + " -> " + $builtOutputs[$referenceKey])
                Add-ReferenceByPath -VBProject $vbProject -ReferencePath $builtOutputs[$referenceKey]
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
            $outputPath = Join-Path $outputDir $project.OutputFile
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
    }
}
finally {
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        Release-ComObject $excel
    }
}

Write-Host "Build complete."
