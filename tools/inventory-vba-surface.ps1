[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourceRoot,

    [Parameter(Mandatory = $true)]
    [string]$BuildMapPath,

    [Parameter(Mandatory = $true)]
    [string]$RibbonRoot,

    [Parameter(Mandatory = $true)]
    [string]$TestRoot,

    [Parameter(Mandatory = $true)]
    [string]$RootRegistryPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputDirectory,

    [string]$ReportTimestampUtc = "",

    [string]$ImplementationSchemaPath = "",

    [string]$MaintenanceSchemaPath = ""
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = (Resolve-Path (Join-Path $scriptRoot "..")).Path

if ([string]::IsNullOrWhiteSpace($ImplementationSchemaPath)) {
    $ImplementationSchemaPath = Join-Path $scriptRoot "contracts\implementation-manifest.schema.json"
}
if ([string]::IsNullOrWhiteSpace($MaintenanceSchemaPath)) {
    $MaintenanceSchemaPath = Join-Path $scriptRoot "contracts\maintenance-candidates.schema.json"
}
if ([string]::IsNullOrWhiteSpace($ReportTimestampUtc)) {
    $ReportTimestampUtc = [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")
}

function Resolve-RequiredPath {
    param(
        [string]$Path,
        [string]$Description
    )
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "$Description not found: $Path"
    }
    return (Resolve-Path -LiteralPath $Path).Path
}

function ConvertTo-NormalizedPath {
    param([string]$Path)
    return ($Path -replace "\\", "/")
}

function ConvertTo-IdToken {
    param([string]$Value)
    return ((ConvertTo-NormalizedPath $Value) -replace '[^A-Za-z0-9_.-]', '_')
}

function Get-RelativePath {
    param(
        [string]$BasePath,
        [string]$TargetPath
    )
    $baseFull = [IO.Path]::GetFullPath($BasePath)
    if (-not $baseFull.EndsWith([IO.Path]::DirectorySeparatorChar)) {
        $baseFull += [IO.Path]::DirectorySeparatorChar
    }
    $targetFull = [IO.Path]::GetFullPath($TargetPath)
    $baseUri = New-Object Uri($baseFull)
    $targetUri = New-Object Uri($targetFull)
    $relative = [Uri]::UnescapeDataString($baseUri.MakeRelativeUri($targetUri).ToString())
    return (ConvertTo-NormalizedPath $relative)
}

function Get-DisplayPath {
    param([string]$Path)
    $full = [IO.Path]::GetFullPath($Path)
    $repoPrefix = $repoRoot.TrimEnd("\", "/") + [IO.Path]::DirectorySeparatorChar
    if ($full.StartsWith($repoPrefix, [StringComparison]::OrdinalIgnoreCase) -or
        $full.Equals($repoRoot, [StringComparison]::OrdinalIgnoreCase)) {
        return (Get-RelativePath -BasePath $repoRoot -TargetPath $full)
    }
    return (ConvertTo-NormalizedPath ([IO.Path]::GetFileName($full)))
}

function Write-Utf8NoBom {
    param(
        [string]$Path,
        [string]$Content
    )
    $normalized = $Content -replace "`r`n", "`n"
    if (-not $normalized.EndsWith("`n")) {
        $normalized += "`n"
    }
    $encoding = New-Object System.Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($Path, $normalized, $encoding)
}

function Get-StringSha256 {
    param([string]$Value)
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes($Value)
        return ([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace("-", "").ToLowerInvariant()
    }
    finally {
        $sha.Dispose()
    }
}

function Read-Json {
    param([string]$Path)
    # PowerShell 7.5+ otherwise materializes ISO timestamps as DateTime values,
    # although the JSON contract declares them as strings. Preserve their text;
    # Windows PowerShell 5.1 has no DateKind parameter and retains its reader.
    if ((Get-Command ConvertFrom-Json).Parameters.ContainsKey('DateKind')) {
        return (Get-Content -Raw -LiteralPath $Path | ConvertFrom-Json -DateKind String)
    }
    return (Get-Content -Raw -LiteralPath $Path | ConvertFrom-Json)
}

function Test-HasProperty {
    param(
        $Object,
        [string]$Name
    )
    if ($null -eq $Object) {
        return $false
    }
    return ($null -ne $Object.PSObject.Properties[$Name])
}

function Get-JsonType {
    param($Value)
    if ($null -eq $Value) { return "null" }
    if ($Value -is [string]) { return "string" }
    if ($Value -is [bool]) { return "boolean" }
    if ($Value -is [byte] -or $Value -is [int16] -or $Value -is [int32] -or
        $Value -is [int64] -or $Value -is [uint16] -or $Value -is [uint32] -or
        $Value -is [uint64]) { return "integer" }
    if ($Value -is [single] -or $Value -is [double] -or $Value -is [decimal]) {
        return "number"
    }
    if ($Value -is [System.Collections.IEnumerable] -and
        -not ($Value -is [string]) -and
        -not ($Value -is [System.Management.Automation.PSCustomObject])) {
        return "array"
    }
    return "object"
}

function Resolve-SchemaReference {
    param(
        $RootSchema,
        [string]$Reference
    )
    if (-not $Reference.StartsWith("#/")) {
        throw "Only local JSON Schema references are supported: $Reference"
    }
    $node = $RootSchema
    foreach ($segment in $Reference.Substring(2).Split("/")) {
        $decoded = $segment.Replace("~1", "/").Replace("~0", "~")
        if (-not (Test-HasProperty -Object $node -Name $decoded)) {
            throw "JSON Schema reference not found: $Reference"
        }
        $node = $node.$decoded
    }
    return $node
}

function Test-JsonSchemaNode {
    param(
        $Value,
        $Schema,
        $RootSchema,
        [string]$Path,
        [System.Collections.Generic.List[string]]$Errors
    )

    if (Test-HasProperty -Object $Schema -Name '$ref') {
        $resolved = Resolve-SchemaReference -RootSchema $RootSchema -Reference $Schema.'$ref'
        Test-JsonSchemaNode -Value $Value -Schema $resolved -RootSchema $RootSchema `
            -Path $Path -Errors $Errors
        return
    }

    if (Test-HasProperty -Object $Schema -Name "const") {
        if ([string]$Value -cne [string]$Schema.const) {
            $Errors.Add("$Path must equal '$($Schema.const)'.")
            return
        }
    }

    if (Test-HasProperty -Object $Schema -Name "enum") {
        $enumMatch = $false
        foreach ($candidate in @($Schema.enum)) {
            if ([string]$Value -ceq [string]$candidate) {
                $enumMatch = $true
                break
            }
        }
        if (-not $enumMatch) {
            $Errors.Add("$Path is not one of the allowed enum values.")
            return
        }
    }

    if (Test-HasProperty -Object $Schema -Name "type") {
        $actualType = Get-JsonType $Value
        $allowedTypes = @($Schema.type)
        if ($actualType -eq "integer" -and "number" -in $allowedTypes) {
            $actualType = "number"
        }
        if ($actualType -notin $allowedTypes) {
            $Errors.Add("$Path has type '$actualType'; expected $($allowedTypes -join '|').")
            return
        }
    }

    if ($null -eq $Value) {
        return
    }

    $valueType = Get-JsonType $Value
    if ($valueType -eq "object") {
        if (Test-HasProperty -Object $Schema -Name "required") {
            foreach ($requiredName in @($Schema.required)) {
                if (-not (Test-HasProperty -Object $Value -Name ([string]$requiredName))) {
                    $Errors.Add("$Path is missing required property '$requiredName'.")
                }
            }
        }

        $allowedPropertyNames = @()
        if (Test-HasProperty -Object $Schema -Name "properties") {
            $allowedPropertyNames = @($Schema.properties.PSObject.Properties.Name)
            foreach ($propertyName in $allowedPropertyNames) {
                if (Test-HasProperty -Object $Value -Name $propertyName) {
                    Test-JsonSchemaNode -Value $Value.$propertyName `
                        -Schema $Schema.properties.$propertyName `
                        -RootSchema $RootSchema `
                        -Path ($Path + "." + $propertyName) `
                        -Errors $Errors
                }
            }
        }

        if ((Test-HasProperty -Object $Schema -Name "additionalProperties") -and
            $Schema.additionalProperties -eq $false) {
            foreach ($actualName in @($Value.PSObject.Properties.Name)) {
                if ($actualName -notin $allowedPropertyNames) {
                    $Errors.Add("$Path contains undeclared property '$actualName'.")
                }
            }
        }
    }
    elseif ($valueType -eq "array" -and (Test-HasProperty -Object $Schema -Name "items")) {
        $index = 0
        foreach ($item in @($Value)) {
            Test-JsonSchemaNode -Value $item -Schema $Schema.items `
                -RootSchema $RootSchema -Path ($Path + "[$index]") -Errors $Errors
            $index += 1
        }
    }

    if ((Test-HasProperty -Object $Schema -Name "minimum") -and
        ([double]$Value -lt [double]$Schema.minimum)) {
        $Errors.Add("$Path is below minimum $($Schema.minimum).")
    }
    if ((Test-HasProperty -Object $Schema -Name "minLength") -and
        ([string]$Value).Length -lt [int]$Schema.minLength) {
        $Errors.Add("$Path is shorter than minLength $($Schema.minLength).")
    }
    if ((Test-HasProperty -Object $Schema -Name "format") -and
        $Schema.format -eq "date-time") {
        $parsed = [DateTimeOffset]::MinValue
        if (-not [DateTimeOffset]::TryParse([string]$Value, [ref]$parsed)) {
            $Errors.Add("$Path is not a valid date-time.")
        }
    }
}

function Assert-JsonAgainstSchema {
    param(
        [string]$JsonPath,
        [string]$SchemaPath
    )
    $value = Read-Json $JsonPath
    $schema = Read-Json $SchemaPath
    $errors = New-Object System.Collections.Generic.List[string]
    Test-JsonSchemaNode -Value $value -Schema $schema -RootSchema $schema `
        -Path '$' -Errors $errors
    if ($errors.Count -gt 0) {
        throw ("JSON schema validation failed for $JsonPath`n" + ($errors -join "`n"))
    }
}

function Get-StringLiterals {
    param([string]$Text)
    $results = New-Object System.Collections.Generic.List[string]
    foreach ($match in [regex]::Matches($Text, '"(?<value>(?:""|[^"])*)"')) {
        $results.Add($match.Groups["value"].Value.Replace('""', '"'))
    }
    return $results.ToArray()
}

function Remove-VbaStringsAndComments {
    param([string]$Text)
    $output = New-Object System.Collections.Generic.List[string]
    foreach ($line in ($Text -split "`r?`n")) {
        $builder = New-Object Text.StringBuilder
        $inString = $false
        for ($i = 0; $i -lt $line.Length; $i++) {
            $char = $line[$i]
            if ($char -eq '"') {
                if ($inString -and $i + 1 -lt $line.Length -and $line[$i + 1] -eq '"') {
                    [void]$builder.Append("  ")
                    $i += 1
                    continue
                }
                $inString = -not $inString
                [void]$builder.Append(" ")
                continue
            }
            if (-not $inString -and $char -eq "'") {
                break
            }
            if ($inString) {
                [void]$builder.Append(" ")
            }
            else {
                [void]$builder.Append($char)
            }
        }
        $clean = $builder.ToString()
        if ($clean -match '^\s*Rem(\s|$)') {
            $clean = ""
        }
        $output.Add($clean)
    }
    return ($output -join "`n")
}

function Get-ProcedureRecords {
    param(
        [string[]]$Lines,
        [string]$ComponentName,
        [string]$SourcePath
    )

    $records = New-Object System.Collections.Generic.List[object]
    $startPattern = '^\s*(?:(Public|Private|Friend)\s+)?(?:Static\s+)?' +
        '(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_][A-Za-z0-9_]*)\b'

    for ($index = 0; $index -lt $Lines.Count; $index++) {
        $match = [regex]::Match($Lines[$index], $startPattern, "IgnoreCase")
        if (-not $match.Success) {
            continue
        }

        $visibility = $match.Groups[1].Value.ToUpperInvariant()
        if ([string]::IsNullOrWhiteSpace($visibility)) {
            $visibility = "DEFAULT"
        }
        $rawKind = $match.Groups[2].Value.ToUpperInvariant()
        $kind = $rawKind.Replace(" ", "_")
        $name = $match.Groups[3].Value
        $endToken = "End " + $(if ($rawKind.StartsWith("PROPERTY")) { "Property" } else { $rawKind })
        $endIndex = $index
        for ($cursor = $index + 1; $cursor -lt $Lines.Count; $cursor++) {
            if ($Lines[$cursor] -match ('^\s*' + [regex]::Escape($endToken) + '\s*$')) {
                $endIndex = $cursor
                break
            }
        }
        if ($endIndex -eq $index) {
            $endIndex = $Lines.Count - 1
        }

        $bodyText = ($Lines[$index..$endIndex] -join "`n")
        $records.Add([ordered]@{
            componentName = $ComponentName
            name = $name
            visibility = $visibility
            kind = $kind
            startLine = $index + 1
            endLine = $endIndex + 1
            lineCount = $endIndex - $index + 1
            directCalls = @()
            literalApplicationRunTargets = @()
            unresolvedApplicationRunExpressions = @()
            rootIds = @()
            sourcePath = $SourcePath
            bodyText = $bodyText
        })
        $index = $endIndex
    }
    return $records.ToArray()
}

function Get-NormalizedProcedureBody {
    param($Procedure)
    $lines = @($Procedure.bodyText -split "`r?`n")
    if ($lines.Count -le 2) {
        return ""
    }
    $bodyLines = $lines[1..($lines.Count - 2)]
    $clean = Remove-VbaStringsAndComments ($bodyLines -join "`n")
    $clean = [regex]::Replace(
        $clean,
        ('\b' + [regex]::Escape([string]$Procedure.name) + '\b'),
        "<PROC>",
        "IgnoreCase"
    )
    $clean = [regex]::Replace($clean, '\s+', ' ').Trim().ToLowerInvariant()
    return $clean
}

function Get-BuildDefinition {
    param([string]$Path)
    $extension = [IO.Path]::GetExtension($Path).ToLowerInvariant()
    $packages = New-Object System.Collections.Generic.List[object]
    $excludes = @{}
    $blocks = @{}

    if ($extension -eq ".json") {
        $definition = Read-Json $Path
        foreach ($project in @($definition.projects)) {
            $packages.Add([ordered]@{
                key = [string]$project.key
                projectName = [string]$project.projectName
                outputFile = [string]$project.outputFile
                sourcePaths = @($project.sourcePaths | ForEach-Object {
                    ConvertTo-NormalizedPath ([string]$_)
                })
                componentNames = @()
            })
            $excludes[[string]$project.key] = @()
        }
    }
    elseif ($extension -eq ".ps1") {
        $text = Get-Content -Raw -LiteralPath $Path
        $pattern = '(?ms)@\{\s*Key\s*=\s*"(?<key>[^"]+)"(?<body>.*?)' +
            '(?=\r?\n\s*@\{\s*Key\s*=|\r?\n\)\s*\r?\n\s*(?:' +
            '\$availableProjects|Write-Host))'
        foreach ($match in [regex]::Matches($text, $pattern)) {
            $key = $match.Groups["key"].Value
            $body = $match.Groups["body"].Value
            $projectMatch = [regex]::Match($body, 'Project\s*=\s*"([^"]+)"')
            $outputMatch = [regex]::Match($body, 'OutputFile\s*=\s*"([^"]+)"')
            $sourceMatches = [regex]::Matches(
                $body,
                'Join-Path\s+\$repo\s+"(src[/\\][^"]+)"'
            )
            if (-not ($projectMatch.Success -and
                      $outputMatch.Success -and
                      $sourceMatches.Count -gt 0)) {
                throw "Unable to parse build project '$key' from $Path."
            }
            $packages.Add([ordered]@{
                key = $key
                projectName = $projectMatch.Groups[1].Value
                outputFile = $outputMatch.Groups[1].Value
                sourcePaths = @(
                    $sourceMatches |
                        ForEach-Object {
                            ConvertTo-NormalizedPath $_.Groups[1].Value
                        } |
                        Sort-Object -Unique
                )
                componentNames = @()
            })
            $excludeMatch = [regex]::Match($body, 'ExcludeFiles\s*=\s*@\((?<values>[^)]*)\)')
            if ($excludeMatch.Success) {
                $excludes[$key] = @(
                    [regex]::Matches($excludeMatch.Groups["values"].Value, '"([^"]+)"') |
                        ForEach-Object { $_.Groups[1].Value }
                )
            }
            else {
                $excludes[$key] = @()
            }
            $blocks[$key] = $body
        }
    }
    else {
        throw "Unsupported build-map format: $Path"
    }

    return [ordered]@{
        packages = @($packages.ToArray() | Sort-Object { $_.key })
        excludes = $excludes
        blocks = $blocks
    }
}

function Get-PackageKeyForPath {
    param(
        [string]$SourcePath,
        [object[]]$Packages
    )
    $normalized = "/" + (ConvertTo-NormalizedPath $SourcePath).Trim("/") + "/"
    foreach ($package in $Packages) {
        foreach ($sourcePath in @($package.sourcePaths)) {
            $leaf = ([string]$sourcePath).Trim("/").Split("/")[-1]
            if ($normalized -match ('/' + [regex]::Escape($leaf) + '/')) {
                return [string]$package.key
            }
        }
    }
    return ""
}

function Convert-RibbonXmlNode {
    param($Node)
    $callbacks = New-Object System.Collections.Generic.List[string]
    foreach ($attributeName in @(
        "onAction", "getEnabled", "getLabel", "getVisible", "getImage",
        "onChange", "getItemCount", "getItemLabel", "getSelectedItemIndex"
    )) {
        if ($null -ne $Node.Attributes[$attributeName]) {
            $callbacks.Add([string]$Node.Attributes[$attributeName].Value)
        }
    }
    $children = New-Object System.Collections.Generic.List[object]
    foreach ($child in $Node.ChildNodes) {
        if ($child.NodeType -eq [System.Xml.XmlNodeType]::Element) {
            $children.Add((Convert-RibbonXmlNode $child))
        }
    }
    $id = ""
    $label = ""
    $capability = $null
    if ($null -ne $Node.Attributes["id"]) { $id = [string]$Node.Attributes["id"].Value }
    if ($null -ne $Node.Attributes["label"]) { $label = [string]$Node.Attributes["label"].Value }
    if ($null -ne $Node.Attributes["tag"]) { $capability = [string]$Node.Attributes["tag"].Value }
    return [ordered]@{
        id = $id
        nodeType = [string]$Node.LocalName
        label = $label
        callbacks = @($callbacks.ToArray() | Sort-Object -Unique)
        capabilityGate = $capability
        children = $children.ToArray()
    }
}

function Get-RibbonDefinitions {
    param(
        [string]$Path,
        $BuildDefinition
    )
    $definitions = New-Object System.Collections.Generic.List[object]

    if (Test-Path -LiteralPath $Path -PathType Container) {
        foreach ($file in @(Get-ChildItem -LiteralPath $Path -Recurse -File -Filter "*.xml" |
            Sort-Object FullName)) {
            try {
                [xml]$xml = Get-Content -Raw -LiteralPath $file.FullName
                $tabs = New-Object System.Collections.Generic.List[object]
                foreach ($tab in $xml.SelectNodes("//*[local-name()='tab']")) {
                    $tabs.Add((Convert-RibbonXmlNode $tab))
                }
                if ($tabs.Count -gt 0) {
                    $definitions.Add([ordered]@{
                        sourcePath = Get-DisplayPath $file.FullName
                        tabs = $tabs.ToArray()
                    })
                }
            }
            catch {
                throw (
                    "Unable to parse RibbonX file $($file.FullName): " +
                    "$($_.Exception.Message) $($_.ScriptStackTrace)"
                )
            }
        }
    }
    elseif ([IO.Path]::GetExtension($Path).ToLowerInvariant() -eq ".ps1") {
        foreach ($package in @($BuildDefinition.packages)) {
            if (-not $BuildDefinition.blocks.ContainsKey([string]$package.key)) {
                continue
            }
            $block = [string]$BuildDefinition.blocks[[string]$package.key]
            $tabMatch = [regex]::Match($block, 'TabId\s*=\s*"([^"]+)"')
            if (-not $tabMatch.Success) {
                continue
            }
            $labelMatch = [regex]::Match($block, 'Label\s*=\s*"([^"]+)"')
            $tabCallbacks = New-Object System.Collections.Generic.List[string]
            foreach ($callbackField in @("CallbackName", "EnabledCallbackName")) {
                $callbackMatch = [regex]::Match(
                    $block,
                    ([regex]::Escape($callbackField) + '\s*=\s*"([^"]+)"')
                )
                if ($callbackMatch.Success) {
                    $tabCallbacks.Add($callbackMatch.Groups[1].Value)
                }
            }

            $controls = New-Object System.Collections.Generic.List[object]
            $controlPattern = '(?ms)@\{\s*Id\s*=\s*"(?<id>[^"]+)"(?<attrs>[^}]*)\}'
            foreach ($controlMatch in [regex]::Matches($block, $controlPattern)) {
                $controlId = $controlMatch.Groups["id"].Value
                if ($controlId.StartsWith("grp")) {
                    continue
                }
                $attrs = $controlMatch.Groups["attrs"].Value
                $controlLabelMatch = [regex]::Match($attrs, 'Label\s*=\s*"([^"]+)"')
                $controlCallbacks = New-Object System.Collections.Generic.List[string]
                foreach ($field in @("Macro", "DirectAction", "GetLabel")) {
                    $fieldMatch = [regex]::Match(
                        $attrs,
                        ([regex]::Escape($field) + '\s*=\s*"([^"]+)"')
                    )
                    if ($fieldMatch.Success) {
                        $controlCallbacks.Add($fieldMatch.Groups[1].Value.Replace('""', '"'))
                    }
                }
                $capabilityMatch = [regex]::Match(
                    $attrs,
                    'RequiredCapability\s*=\s*"([^"]+)"'
                )
                $controlType = "control"
                if ($controlId.StartsWith("btn")) { $controlType = "button" }
                elseif ($controlId.StartsWith("dd")) { $controlType = "dropDown" }
                elseif ($controlId.StartsWith("mnu")) { $controlType = "menu" }
                elseif ($controlId.StartsWith("lbl")) { $controlType = "label" }
                $controls.Add([ordered]@{
                    id = $controlId
                    nodeType = $controlType
                    label = $(if ($controlLabelMatch.Success) {
                        $controlLabelMatch.Groups[1].Value
                    } else { "" })
                    callbacks = @($controlCallbacks | Sort-Object -Unique)
                    capabilityGate = $(if ($capabilityMatch.Success) {
                        $capabilityMatch.Groups[1].Value
                    } else { $null })
                    children = @()
                })
            }

            $groupIdMatch = [regex]::Match($block, 'Id\s*=\s*"(grp[^"]+)"')
            $group = [ordered]@{
                id = $(if ($groupIdMatch.Success) { $groupIdMatch.Groups[1].Value } else {
                    "grp" + [string]$package.key
                })
                nodeType = "group"
                label = ""
                callbacks = @()
                capabilityGate = $null
                children = @($controls.ToArray() | Sort-Object { $_.id })
            }
            $tab = [ordered]@{
                id = $tabMatch.Groups[1].Value
                nodeType = "tab"
                label = $(if ($labelMatch.Success) { $labelMatch.Groups[1].Value } else { "" })
                callbacks = @($tabCallbacks | Sort-Object -Unique)
                capabilityGate = $null
                children = @($group)
            }
            $definitions.Add([ordered]@{
                sourcePath = Get-DisplayPath $Path
                tabs = @($tab)
            })
        }
    }
    return @(
        $definitions.ToArray() |
            Sort-Object { ([string]$_.sourcePath) + "|" + ([string]$_.tabs[0].id) }
    )
}

function Get-RibbonCallbackRecords {
    param([object[]]$Ribbons)
    $records = New-Object System.Collections.Generic.List[object]

    function Visit-RibbonNode {
        param(
            $Node,
            [string]$SourcePath
        )
        foreach ($callback in @($Node.callbacks)) {
            $raw = [string]$callback
            $procedure = $raw
            if ($procedure.Contains("!")) {
                $procedure = $procedure.Split("!")[-1]
            }
            if ($procedure.Contains(".")) {
                $procedure = $procedure.Split(".")[-1]
            }
            $procedure = ([regex]::Match($procedure, '^[A-Za-z_][A-Za-z0-9_]*')).Value
            if (-not [string]::IsNullOrWhiteSpace($procedure)) {
                $records.Add([ordered]@{
                    id = "ribbon:" + $Node.id + ":" + $procedure
                    procedure = $procedure
                    rootKind = "RIBBON_CALLBACK"
                    reason = "Referenced by Ribbon definition node " + $Node.id + "."
                    source = $SourcePath
                })
            }
        }
        foreach ($child in @($Node.children)) {
            Visit-RibbonNode -Node $child -SourcePath $SourcePath
        }
    }

    foreach ($ribbon in $Ribbons) {
        foreach ($tab in @($ribbon.tabs)) {
            Visit-RibbonNode -Node $tab -SourcePath ([string]$ribbon.sourcePath)
        }
    }
    return $records.ToArray()
}

function Test-RegistryMatch {
    param(
        [string]$Value,
        [string]$Pattern,
        [string]$MatchType
    )
    if ($MatchType -eq "EXACT") {
        return $Value.Equals($Pattern, [StringComparison]::OrdinalIgnoreCase)
    }
    return [regex]::IsMatch($Value, $Pattern, "IgnoreCase")
}

function Get-RegistryRootsForProcedure {
    param(
        $Procedure,
        $Registry
    )
    $roots = New-Object System.Collections.Generic.List[object]
    foreach ($entry in @($Registry.roots)) {
        if (-not (Test-RegistryMatch -Value ([string]$Procedure.componentName) `
            -Pattern ([string]$entry.componentPattern) -MatchType ([string]$entry.matchType))) {
            continue
        }
        if (-not (Test-RegistryMatch -Value ([string]$Procedure.name) `
            -Pattern ([string]$entry.procedurePattern) -MatchType ([string]$entry.matchType))) {
            continue
        }
        if ((Test-HasProperty -Object $entry -Name "sourcePathPattern") -and
            $null -ne $entry.sourcePathPattern -and
            -not [regex]::IsMatch(
                [string]$Procedure.sourcePath,
                [string]$entry.sourcePathPattern,
                "IgnoreCase"
            )) {
            continue
        }
        $roots.Add([ordered]@{
            id = [string]$entry.id + ":" +
                (ConvertTo-IdToken ([string]$Procedure.sourcePath)) + ":" +
                [string]$Procedure.componentName + "." + [string]$Procedure.name
            procedure = [string]$Procedure.name
            rootKind = [string]$entry.rootKind
            reason = [string]$entry.reason
            source = Get-DisplayPath $RootRegistryPath
        })
    }
    return $roots.ToArray()
}

function Get-FormControls {
    param(
        [string[]]$Lines,
        [string]$FormName,
        [string]$SourcePath
    )
    $controls = New-Object System.Collections.Generic.List[object]
    $stack = New-Object System.Collections.Generic.List[object]
    foreach ($line in $Lines) {
        $beginMatch = [regex]::Match(
            $line,
            '^\s*Begin\s+([A-Za-z0-9_.]+)\s+([A-Za-z_][A-Za-z0-9_]*)\s*$'
        )
        if ($beginMatch.Success) {
            $owner = $FormName
            if ($stack.Count -gt 0) {
                $owner = [string]$stack[$stack.Count - 1].name
            }
            $entry = [ordered]@{
                name = $beginMatch.Groups[2].Value
                controlType = $beginMatch.Groups[1].Value
                owner = $owner
                left = 0.0
                top = 0.0
                width = 0.0
                height = 0.0
                isRoot = ($beginMatch.Groups[1].Value -eq "VB.UserForm")
            }
            $stack.Add($entry)
            continue
        }
        if ($line -match '^\s*End\s*$' -and $stack.Count -gt 0) {
            $entry = $stack[$stack.Count - 1]
            $stack.RemoveAt($stack.Count - 1)
            if (-not $entry.isRoot) {
                $controls.Add([ordered]@{
                    name = $entry.name
                    controlType = $entry.controlType
                    owner = $entry.owner
                    left = [double]$entry.left
                    top = [double]$entry.top
                    width = [double]$entry.width
                    height = [double]$entry.height
                })
            }
            continue
        }
        if ($stack.Count -eq 0) {
            continue
        }
        $propertyMatch = [regex]::Match(
            $line,
            '^\s*(Left|Top|Width|Height|ClientWidth|ClientHeight)\s*=\s*([-0-9.]+)'
        )
        if ($propertyMatch.Success) {
            $propertyName = $propertyMatch.Groups[1].Value
            if ($propertyName -eq "ClientWidth") { $propertyName = "Width" }
            if ($propertyName -eq "ClientHeight") { $propertyName = "Height" }
            $stack[$stack.Count - 1][$propertyName.ToLowerInvariant()] =
                [double]$propertyMatch.Groups[2].Value
        }
    }
    $form = [ordered]@{
        name = $FormName
        sourcePath = $SourcePath
        controls = @(
            $controls.ToArray() |
                Sort-Object { ([string]$_.owner) + "|" + ([string]$_.name) }
        )
    }
    $layoutMatch = [regex]::Match(
        ($Lines -join "`n"),
        "(?im)^\s*'@FormLayout\s+" +
        "Strategy=(?<strategy>[A-Za-z0-9_]+)\s+" +
        "MinWidth=(?<minWidth>[0-9.]+)\s+MinHeight=(?<minHeight>[0-9.]+)\s+" +
        "DefaultWidth=(?<defaultWidth>[0-9.]+)\s+DefaultHeight=(?<defaultHeight>[0-9.]+)\s+" +
        "ExpandedWidth=(?<expandedWidth>[0-9.]+)\s+ExpandedHeight=(?<expandedHeight>[0-9.]+)\s*$"
    )
    if ($layoutMatch.Success) {
        $form["layout"] = [ordered]@{
            strategy = $layoutMatch.Groups["strategy"].Value
            minimum = [ordered]@{
                width = [double]$layoutMatch.Groups["minWidth"].Value
                height = [double]$layoutMatch.Groups["minHeight"].Value
            }
            default = [ordered]@{
                width = [double]$layoutMatch.Groups["defaultWidth"].Value
                height = [double]$layoutMatch.Groups["defaultHeight"].Value
            }
            expandedTest = [ordered]@{
                width = [double]$layoutMatch.Groups["expandedWidth"].Value
                height = [double]$layoutMatch.Groups["expandedHeight"].Value
            }
        }
    }
    return $form
}

$resolvedSourceRoot = Resolve-RequiredPath -Path $SourceRoot -Description "Source root"
$resolvedBuildMap = Resolve-RequiredPath -Path $BuildMapPath -Description "Build map"
$resolvedRibbonRoot = Resolve-RequiredPath -Path $RibbonRoot -Description "Ribbon root"
$resolvedTestRoot = Resolve-RequiredPath -Path $TestRoot -Description "Test root"
$resolvedRegistry = Resolve-RequiredPath -Path $RootRegistryPath -Description "Root registry"
$resolvedImplementationSchema = Resolve-RequiredPath `
    -Path $ImplementationSchemaPath -Description "Implementation schema"
$resolvedMaintenanceSchema = Resolve-RequiredPath `
    -Path $MaintenanceSchemaPath -Description "Maintenance schema"

if (-not (Test-Path -LiteralPath $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
}
$resolvedOutput = (Resolve-Path -LiteralPath $OutputDirectory).Path

$buildDefinition = Get-BuildDefinition $resolvedBuildMap
$packages = @($buildDefinition.packages)
$registry = Read-Json $resolvedRegistry
$componentInternals = New-Object System.Collections.Generic.List[object]
$procedures = New-Object System.Collections.Generic.List[object]
$forms = New-Object System.Collections.Generic.List[object]
$allStringsByPath = @{}

$sourceFiles = @(
    Get-ChildItem -LiteralPath $resolvedSourceRoot -Recurse -File |
        Where-Object { $_.Extension.ToLowerInvariant() -in @(".bas", ".cls", ".frm", ".frx") } |
        Sort-Object FullName
)

foreach ($file in $sourceFiles) {
    $sourcePath = Get-DisplayPath $file.FullName
    $extension = $file.Extension.ToLowerInvariant()
    $componentName = [IO.Path]::GetFileNameWithoutExtension($file.Name)
    $kind = switch ($extension) {
        ".bas" { "STANDARD_MODULE" }
        ".cls" { "CLASS_MODULE" }
        ".frm" { "FORM" }
        ".frx" { "FORM_BINARY" }
    }

    if ($extension -eq ".frx") {
        $componentInternals.Add([ordered]@{
            name = $componentName
            kind = $kind
            sourcePath = $sourcePath
            lineCount = 0
            procedureNames = @()
            metadata = [ordered]@{
                byteLength = [int64]$file.Length
                sha256 = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
            }
            packageKey = Get-PackageKeyForPath -SourcePath $sourcePath -Packages $packages
        })
        continue
    }

    $lines = @(Get-Content -LiteralPath $file.FullName)
    $text = $lines -join "`n"
    $nameMatch = [regex]::Match($text, 'Attribute\s+VB_Name\s*=\s*"([^"]+)"')
    if ($nameMatch.Success) {
        $componentName = $nameMatch.Groups[1].Value
    }
    $componentProcedures = @(
        Get-ProcedureRecords -Lines $lines -ComponentName $componentName -SourcePath $sourcePath
    )
    foreach ($procedure in $componentProcedures) {
        $procedures.Add($procedure)
    }

    $metadata = [ordered]@{}
    if ($extension -eq ".frm") {
        $form = Get-FormControls -Lines $lines -FormName $componentName -SourcePath $sourcePath
        $forms.Add($form)
        $metadata["controlCount"] = @($form.controls).Count
    }
    $componentInternals.Add([ordered]@{
        name = $componentName
        kind = $kind
        sourcePath = $sourcePath
        lineCount = $lines.Count
        procedureNames = @($componentProcedures | ForEach-Object { $_.name } | Sort-Object)
        metadata = $metadata
        packageKey = Get-PackageKeyForPath -SourcePath $sourcePath -Packages $packages
    })
    $allStringsByPath[$sourcePath] = @(Get-StringLiterals $text)
}

$procedureNames = @(
    $procedures.ToArray() |
        ForEach-Object { [string]$_.name } |
        Sort-Object -Unique
)
$procedureNameLookup = New-Object 'System.Collections.Generic.Dictionary[string,string]' `
    ([StringComparer]::OrdinalIgnoreCase)
foreach ($procedureName in $procedureNames) {
    if (-not $procedureNameLookup.ContainsKey($procedureName)) {
        $procedureNameLookup.Add($procedureName, $procedureName)
    }
}
foreach ($procedure in $procedures) {
    $cleanBody = Remove-VbaStringsAndComments ([string]$procedure.bodyText)
    $directCalls = New-Object System.Collections.Generic.List[string]
    foreach ($tokenMatch in [regex]::Matches($cleanBody, '\b[A-Za-z_][A-Za-z0-9_]*\b')) {
        $token = $tokenMatch.Value
        if ($procedureNameLookup.ContainsKey($token)) {
            $candidateName = $procedureNameLookup[$token]
            if (-not $candidateName.Equals(
                [string]$procedure.name,
                [StringComparison]::OrdinalIgnoreCase
            )) {
                $directCalls.Add($candidateName)
            }
        }
    }
    $procedure.directCalls = @($directCalls.ToArray() | Sort-Object -Unique)

    $literalTargets = New-Object System.Collections.Generic.List[string]
    $dynamicExpressions = New-Object System.Collections.Generic.List[string]
    $logicalBody = [regex]::Replace(
        [string]$procedure.bodyText,
        '\s+_\s*\r?\n\s*',
        ' '
    )
    foreach ($line in ($logicalBody -split "`r?`n")) {
        $runMatch = [regex]::Match(
            $line,
            'Application\.Run\s*(?:\(\s*)?(?<expression>.+)$',
            "IgnoreCase"
        )
        if (-not $runMatch.Success) {
            continue
        }
        $expression = $runMatch.Groups["expression"].Value.Trim()
        $literalMatch = [regex]::Match(
            $expression,
            '^"(?<target>(?:""|[^"])*)"(?<remainder>.*)$'
        )
        if ($literalMatch.Success) {
            $remainder = $literalMatch.Groups["remainder"].Value.TrimStart()
            if ([string]::IsNullOrWhiteSpace($remainder) -or
                $remainder.StartsWith(",") -or
                $remainder -eq ")") {
                $literalTargets.Add(
                    $literalMatch.Groups["target"].Value.Replace('""', '"')
                )
                continue
            }
        }
        if (-not [string]::IsNullOrWhiteSpace($expression)) {
            $dynamicExpressions.Add($expression)
        }
    }
    $procedure.literalApplicationRunTargets =
        @($literalTargets.ToArray() | Sort-Object -Unique)
    $procedure.unresolvedApplicationRunExpressions =
        @($dynamicExpressions.ToArray() | Sort-Object -Unique)
}

$ribbons = @(Get-RibbonDefinitions -Path $resolvedRibbonRoot -BuildDefinition $buildDefinition)
$dynamicRoots = New-Object System.Collections.Generic.List[object]
foreach ($procedure in $procedures) {
    foreach ($root in @(Get-RegistryRootsForProcedure -Procedure $procedure -Registry $registry)) {
        $dynamicRoots.Add($root)
    }
}
foreach ($root in @(Get-RibbonCallbackRecords -Ribbons $ribbons)) {
    $exists = @($dynamicRoots | Where-Object {
        $_.procedure -eq $root.procedure -and $_.rootKind -eq $root.rootKind
    }).Count -gt 0
    if (-not $exists) {
        $dynamicRoots.Add($root)
    }
}

foreach ($procedure in $procedures) {
    foreach ($target in @($procedure.literalApplicationRunTargets)) {
        $targetProcedure = [string]$target
        if ($targetProcedure.Contains("!")) {
            $targetProcedure = $targetProcedure.Split("!")[-1]
        }
        if ($targetProcedure.Contains(".")) {
            $targetProcedure = $targetProcedure.Split(".")[-1]
        }
        $targetProcedure = (
            [regex]::Match($targetProcedure, '[A-Za-z_][A-Za-z0-9_]*')
        ).Value
        if (-not [string]::IsNullOrWhiteSpace($targetProcedure)) {
            $dynamicRoots.Add([ordered]@{
                id = "literal-run:" +
                    (ConvertTo-IdToken ([string]$procedure.sourcePath)) + ":" +
                    $procedure.componentName + "." + $procedure.name + ":" +
                    $targetProcedure
                procedure = $targetProcedure
                rootKind = "CROSS_XLAM_BRIDGE"
                reason = "Literal Application.Run target discovered in exported VBA."
                source = [string]$procedure.sourcePath
            })
        }
    }

    foreach ($match in [regex]::Matches(
        [string]$procedure.bodyText,
        '\bAddressOf\s+([A-Za-z_][A-Za-z0-9_]*)',
        "IgnoreCase"
    )) {
        $callbackName = $match.Groups[1].Value
        $dynamicRoots.Add([ordered]@{
            id = "address-of:" +
                (ConvertTo-IdToken ([string]$procedure.sourcePath)) + ":" +
                $procedure.componentName + "." + $procedure.name + ":" +
                $callbackName
            procedure = $callbackName
            rootKind = "WINDOWS_CALLBACK"
            reason = "Procedure is passed with AddressOf and is not reachable as a normal direct call."
            source = [string]$procedure.sourcePath
        })
    }
}

$testFiles = @(
    Get-ChildItem -LiteralPath $resolvedTestRoot -Recurse -File |
        Where-Object { $_.Extension.ToLowerInvariant() -in @(".bas", ".cls", ".frm", ".ps1") } |
        Sort-Object FullName
)
$testTexts = @()
foreach ($testFile in $testFiles) {
    $testTexts += [ordered]@{
        path = Get-DisplayPath $testFile.FullName
        text = Get-Content -Raw -LiteralPath $testFile.FullName
    }
}

foreach ($testFile in $testFiles | Where-Object {
    $_.Extension.ToLowerInvariant() -in @(".bas", ".cls", ".frm")
}) {
    $testLines = @(Get-Content -LiteralPath $testFile.FullName)
    $testText = $testLines -join "`n"
    $testComponentName = [IO.Path]::GetFileNameWithoutExtension($testFile.Name)
    $testNameMatch = [regex]::Match(
        $testText,
        'Attribute\s+VB_Name\s*=\s*"([^"]+)"'
    )
    if ($testNameMatch.Success) {
        $testComponentName = $testNameMatch.Groups[1].Value
    }
    $testPath = Get-DisplayPath $testFile.FullName
    foreach ($testProcedure in @(
        Get-ProcedureRecords `
            -Lines $testLines `
            -ComponentName $testComponentName `
            -SourcePath $testPath
    )) {
        if ($testProcedure.visibility -in @("PUBLIC", "DEFAULT") -and
            $testProcedure.name -match '^(Run|Test)') {
            $dynamicRoots.Add([ordered]@{
                id = "test-entry:" + (ConvertTo-IdToken $testPath) + ":" +
                    $testComponentName + "." + $testProcedure.name
                procedure = [string]$testProcedure.name
                rootKind = "TEST_ENTRY"
                reason = "Public test harness entry point discovered under the configured test root."
                source = $testPath
            })
        }
    }
}

$dynamicRoots = @(
    $dynamicRoots.ToArray() |
        Sort-Object {
            ([string]$_.procedure) + "|" + ([string]$_.rootKind) + "|" + ([string]$_.id)
        } -Unique
)

foreach ($procedure in $procedures) {
    $procedure.rootIds = @(
        $dynamicRoots |
            Where-Object { $_.procedure -eq $procedure.name } |
            ForEach-Object { $_.id } |
            Sort-Object -Unique
    )
}

$testReferences = New-Object System.Collections.Generic.List[object]
foreach ($procedureName in $procedureNames) {
    $paths = @(
        $testTexts |
            Where-Object {
                [regex]::IsMatch($_.text, ('\b' + [regex]::Escape($procedureName) + '\b'), "IgnoreCase")
            } |
            ForEach-Object { $_.path } |
            Sort-Object -Unique
    )
    if ($paths.Count -gt 0) {
        $testReferences.Add([ordered]@{
            entryPoint = $procedureName
            testPaths = $paths
        })
    }
}

$managedHeaderNames = @(
    "System_Key", "SKU", "ITEM_CODE", "Qty", "QtyOnHand", "QtyAvailable",
    "QtyDelta", "Location", "Condition", "InventoryState", "AttributesJson",
    "LastAppliedUTC", "LastRefreshUTC", "SnapshotId", "SourceType", "IsStale"
)
$tableMap = @{}
$configKeys = New-Object System.Collections.Generic.List[string]
$eventTypes = New-Object System.Collections.Generic.List[string]
$capabilities = New-Object System.Collections.Generic.List[string]
$warnings = New-Object System.Collections.Generic.List[object]
$knownEventTypes = @(
    "RECEIVE", "SHIP", "PROD", "PROD_CONSUME", "PROD_COMPLETE",
    "ADJUST_INVENTORY", "UNDO", "DESIGN_CREATE", "DESIGN_RELEASE",
    "DESIGN_OBSOLETE"
)

foreach ($component in $componentInternals) {
    if (-not $allStringsByPath.ContainsKey([string]$component.sourcePath)) {
        continue
    }
    $strings = @($allStringsByPath[[string]$component.sourcePath])
    $componentText = $strings -join "`n"
    $tableNames = @(
        [regex]::Matches($componentText, '\btbl[A-Za-z0-9_]+\b', "IgnoreCase") |
            ForEach-Object { $_.Value } |
            Sort-Object -Unique
    )
    $headers = @(
        $strings |
            Where-Object { $_ -in $managedHeaderNames } |
            Sort-Object -Unique
    )
    foreach ($tableName in $tableNames) {
        $key = $tableName.ToLowerInvariant()
        if (-not $tableMap.ContainsKey($key)) {
            $tableMap[$key] = [ordered]@{
                name = $tableName
                headers = New-Object System.Collections.Generic.List[string]
                locations = New-Object System.Collections.Generic.List[string]
            }
        }
        foreach ($header in $headers) {
            if ($header -notin $tableMap[$key].headers) {
                $tableMap[$key].headers.Add($header)
            }
        }
        if ($component.sourcePath -notin $tableMap[$key].locations) {
            $tableMap[$key].locations.Add([string]$component.sourcePath)
        }
    }

    if ("ROW" -in $strings) {
        $warnings.Add([ordered]@{
            code = "RETIRED_ROW_HEADER"
            severity = "ERROR"
            message = "Retired managed header ROW is referenced by runtime source."
            sourcePath = [string]$component.sourcePath
        })
    }
    foreach ($value in $strings) {
        if ($value -match '^(FF_|Path)' -or
            $value -match '(Seconds|Minutes|Cadence|Enabled)$' -or
            $value -in @(
                "WarehouseId", "StationId", "Timezone", "DefaultLocation",
                "UomCatalog", "BatchSize", "PoisonRetryMax", "AuthCacheTTLSeconds"
            )) {
            if ($value -notin $configKeys) { $configKeys.Add($value) }
        }
        if ($value -in $knownEventTypes -and $value -notin $eventTypes) {
            $eventTypes.Add($value)
        }
        if ($value -match '^[A-Z][A-Z0-9_]*(POST|MAINT|READ|RUN|MANAGE)$' -and
            $value -notin $capabilities) {
            $capabilities.Add($value)
        }
    }
}

foreach ($procedure in $procedures) {
    if (@($procedure.unresolvedApplicationRunExpressions).Count -gt 0) {
        $warnings.Add([ordered]@{
            code = "UNRESOLVED_APPLICATION_RUN"
            severity = "WARNING"
            message = "Dynamic Application.Run expression requires registry or manual review."
            sourcePath = [string]$procedure.sourcePath
        })
    }
}

$tables = @(
    $tableMap.Values |
        Sort-Object { [string]$_.name } |
        ForEach-Object {
            [ordered]@{
                name = $_.name
                managedHeaders = @($_.headers | Sort-Object)
                unknownHeadersPolicy = "PRESERVE"
                locations = @($_.locations | Sort-Object)
            }
        }
)

$bridgeContracts = New-Object System.Collections.Generic.List[object]
foreach ($file in $sourceFiles | Where-Object { $_.Extension -ne ".frx" }) {
    $text = Get-Content -Raw -LiteralPath $file.FullName
    foreach ($match in [regex]::Matches(
        $text,
        '(?im)^\s*(?:Public\s+)?Const\s+([A-Za-z0-9_]*CONTRACT_VERSION[A-Za-z0-9_]*)' +
            '[^=]*=\s*"([^"]+)"'
    )) {
        $bridgeContracts.Add([ordered]@{
            name = $match.Groups[1].Value
            version = $match.Groups[2].Value
            sourcePath = Get-DisplayPath $file.FullName
        })
    }
}

foreach ($package in $packages) {
    $excludes = @($buildDefinition.excludes[[string]$package.key])
    $package.componentNames = @(
        $componentInternals.ToArray() |
            Where-Object {
                $_.packageKey -eq $package.key -and
                ([IO.Path]::GetFileName([string]$_.sourcePath) -notin $excludes)
            } |
            ForEach-Object { $_.name } |
            Sort-Object -Unique
    )
}

$manifestComponents = @(
    $componentInternals.ToArray() |
        Sort-Object { ([string]$_.sourcePath) + "|" + ([string]$_.kind) } |
        ForEach-Object {
            [ordered]@{
                name = $_.name
                kind = $_.kind
                sourcePath = $_.sourcePath
                lineCount = [int]$_.lineCount
                procedureNames = @($_.procedureNames)
                metadata = $_.metadata
            }
        }
)

$manifestProcedures = @(
    $procedures.ToArray() |
        Sort-Object {
            ([string]$_.sourcePath) + "|" + ([int]$_.startLine).ToString("D8")
        } |
        ForEach-Object {
            [ordered]@{
                componentName = $_.componentName
                name = $_.name
                visibility = $_.visibility
                kind = $_.kind
                startLine = [int]$_.startLine
                endLine = [int]$_.endLine
                lineCount = [int]$_.lineCount
                directCalls = @($_.directCalls)
                literalApplicationRunTargets = @($_.literalApplicationRunTargets)
                unresolvedApplicationRunExpressions = @($_.unresolvedApplicationRunExpressions)
                rootIds = @($_.rootIds)
            }
        }
)

$manifest = [ordered]@{
    schemaVersion = "1.0.0"
    reportType = "implementation-manifest"
    generatedAtUtc = $ReportTimestampUtc
    sourceRoot = Get-DisplayPath $resolvedSourceRoot
    packages = @($packages)
    components = $manifestComponents
    procedures = $manifestProcedures
    ribbons = @($ribbons)
    tables = $tables
    configKeys = @($configKeys.ToArray() | Sort-Object)
    eventTypes = @($eventTypes.ToArray() | Sort-Object)
    capabilities = @($capabilities.ToArray() | Sort-Object)
    bridgeContracts = @(
        $bridgeContracts.ToArray() |
            Sort-Object { ([string]$_.name) + "|" + ([string]$_.sourcePath) }
    )
    forms = @($forms.ToArray() | Sort-Object { [string]$_.sourcePath })
    testReferences = @(
        $testReferences.ToArray() |
            Sort-Object { [string]$_.entryPoint }
    )
    dynamicRoots = @($dynamicRoots)
    warnings = @(
        $warnings.ToArray() |
            Sort-Object { ([string]$_.code) + "|" + ([string]$_.sourcePath) } -Unique
    )
}

$incomingCalls = @{}
foreach ($procedureName in $procedureNames) { $incomingCalls[$procedureName] = 0 }
foreach ($procedure in $procedures) {
    foreach ($calledName in @($procedure.directCalls)) {
        if ($incomingCalls.ContainsKey([string]$calledName)) {
            $incomingCalls[[string]$calledName] += 1
        }
    }
}
$rootedNames = @($dynamicRoots | ForEach-Object { $_.procedure } | Sort-Object -Unique)
$testedNames = @(
    $testReferences.ToArray() |
        ForEach-Object { $_.entryPoint } |
        Sort-Object -Unique
)
$candidates = New-Object System.Collections.Generic.List[object]

foreach ($procedure in $procedures) {
    $name = [string]$procedure.name
    $isRoot = $name -in $rootedNames
    $isTested = $name -in $testedNames
    $incoming = [int]$incomingCalls[$name]

    if ($isRoot) {
        $candidates.Add([ordered]@{
            id = "root:" + (ConvertTo-IdToken ([string]$procedure.sourcePath)) + ":" +
                $procedure.componentName + "." + $name
            candidateType = "RETAIN_DYNAMIC_ROOT"
            confidence = "LOW"
            componentName = [string]$procedure.componentName
            procedureNames = @($name)
            sourcePaths = @([string]$procedure.sourcePath)
            reason = "Procedure is a registered or discovered dynamic root."
            protectingTests = @(
                $testReferences |
                    Where-Object { $_.entryPoint -eq $name } |
                    ForEach-Object { $_.testPaths } |
                    Sort-Object -Unique
            )
            reviewRequired = $true
        })
    }
    elseif ($incoming -eq 0 -and -not $isTested) {
        $isPrivate = ([string]$procedure.visibility -eq "PRIVATE")
        $candidates.Add([ordered]@{
            id = "reachability:" +
                (ConvertTo-IdToken ([string]$procedure.sourcePath)) + ":" +
                $procedure.componentName + "." + $name
            candidateType = $(if ($isPrivate) { "REMOVE" } else { "UNRESOLVED" })
            confidence = $(if ($isPrivate) { "HIGH" } else { "MEDIUM" })
            componentName = [string]$procedure.componentName
            procedureNames = @($name)
            sourcePaths = @([string]$procedure.sourcePath)
            reason = $(if ($isPrivate) {
                "No direct, dynamic-root, or test reference was found for a private procedure."
            } else {
                "No direct, dynamic-root, or test reference was found, but public visibility requires review."
            })
            protectingTests = @()
            reviewRequired = $true
        })
    }

    if ([int]$procedure.lineCount -gt 200) {
        $candidates.Add([ordered]@{
            id = "procedure-size:" +
                (ConvertTo-IdToken ([string]$procedure.sourcePath)) + ":" +
                $procedure.componentName + "." + $name
            candidateType = "SPLIT_MODULE"
            confidence = "MEDIUM"
            componentName = [string]$procedure.componentName
            procedureNames = @($name)
            sourcePaths = @([string]$procedure.sourcePath)
            reason = "Procedure exceeds the 200-line new-procedure threshold."
            protectingTests = @(
                $testReferences |
                    Where-Object { $_.entryPoint -eq $name } |
                    ForEach-Object { $_.testPaths } |
                    Sort-Object -Unique
            )
            reviewRequired = $true
        })
    }

    foreach ($expression in @($procedure.unresolvedApplicationRunExpressions)) {
        $candidates.Add([ordered]@{
            id = "dynamic-call:" +
                (ConvertTo-IdToken ([string]$procedure.sourcePath)) + ":" +
                $procedure.componentName + "." + $name + ":" +
                (Get-StringSha256 $expression).Substring(0, 12)
            candidateType = "UNRESOLVED"
            confidence = "LOW"
            componentName = [string]$procedure.componentName
            procedureNames = @($name)
            sourcePaths = @([string]$procedure.sourcePath)
            reason = "Application.Run target is dynamic and cannot be guessed safely: $expression"
            protectingTests = @()
            reviewRequired = $true
        })
    }

    foreach ($target in @($procedure.literalApplicationRunTargets)) {
        if ($target -match '(?i)\.xlam.*!') {
            continue
        }
        $targetName = $target
        if ($targetName.Contains("!")) { $targetName = $targetName.Split("!")[-1] }
        if ($targetName.Contains(".")) { $targetName = $targetName.Split(".")[-1] }
        if ($targetName -in $procedureNames) {
            $candidates.Add([ordered]@{
                id = "same-project-run:" +
                    (ConvertTo-IdToken ([string]$procedure.sourcePath)) + ":" +
                    $procedure.componentName + "." + $name + ":" + $targetName
                candidateType = "REPLACE_SAME_PROJECT_LATE_BINDING"
                confidence = "HIGH"
                componentName = [string]$procedure.componentName
                procedureNames = @($name, $targetName)
                sourcePaths = @([string]$procedure.sourcePath)
                reason = "Literal Application.Run resolves to a procedure in the scanned project surface."
                protectingTests = @()
                reviewRequired = $true
            })
        }
    }
}

foreach ($component in $componentInternals | Where-Object { $_.lineCount -gt 1000 }) {
    $candidates.Add([ordered]@{
        id = "module-size:" + (ConvertTo-IdToken ([string]$component.sourcePath)) +
            ":" + $component.name
        candidateType = "SPLIT_MODULE"
        confidence = "MEDIUM"
        componentName = [string]$component.name
        procedureNames = @($component.procedureNames)
        sourcePaths = @([string]$component.sourcePath)
        reason = "Runtime component exceeds the 1000-line new-module threshold."
        protectingTests = @()
        reviewRequired = $true
    })
}

$duplicateGroups = @(
    $procedures.ToArray() |
        ForEach-Object {
            [ordered]@{
                procedure = $_
                normalizedBody = Get-NormalizedProcedureBody $_
            }
        } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_.normalizedBody) } |
        Group-Object { [string]$_.normalizedBody } |
        Where-Object { $_.Count -gt 1 }
)
foreach ($group in $duplicateGroups) {
    $groupProcedures = @($group.Group | ForEach-Object { $_.procedure })
    $names = @($groupProcedures | ForEach-Object { $_.name } | Sort-Object)
    $paths = @($groupProcedures | ForEach-Object { $_.sourcePath } | Sort-Object -Unique)
    $isDynamic = @($names | Where-Object { $_ -in $rootedNames }).Count -gt 0
    $bodyHash = Get-StringSha256 ([string]$group.Name)
    $candidates.Add([ordered]@{
        id = "duplicate:" + $bodyHash.Substring(0, 16) + ":" + ($names -join "+")
        candidateType = "REPLACE_DUPLICATE"
        confidence = $(if ($isDynamic) { "LOW" } else { "HIGH" })
        componentName = [string]$groupProcedures[0].componentName
        procedureNames = $names
        sourcePaths = $paths
        reason = "Normalized procedure bodies are identical; review before consolidation."
        protectingTests = @()
        reviewRequired = $true
    })
}

$totalLineCount = 0
foreach ($component in $manifestComponents) {
    $totalLineCount += [int]$component.lineCount
}

$maintenance = [ordered]@{
    schemaVersion = "1.0.0"
    reportType = "maintenance-candidates"
    generatedAtUtc = $ReportTimestampUtc
    baseline = [ordered]@{
        componentCount = $manifestComponents.Count
        procedureCount = $manifestProcedures.Count
        lineCount = $totalLineCount
        literalApplicationRunCount = [int]((
            $manifestProcedures |
                ForEach-Object { @($_.literalApplicationRunTargets).Count } |
                Measure-Object -Sum
        ).Sum)
        unresolvedApplicationRunCount = [int]((
            $manifestProcedures |
                ForEach-Object { @($_.unresolvedApplicationRunExpressions).Count } |
                Measure-Object -Sum
        ).Sum)
        duplicateBodyCandidateCount = $duplicateGroups.Count
    }
    candidates = @(
        $candidates.ToArray() |
            Sort-Object { ([string]$_.candidateType) + "|" + ([string]$_.id) } -Unique
    )
    ratchets = [ordered]@{
        maxNewModuleLines = 1000
        maxNewProcedureLines = 200
        allowSameProjectApplicationRunGrowth = $false
        allowUnresolvedDynamicCallGrowth = $false
        allowDuplicateBodyGrowth = $false
    }
    warnings = @(
        [ordered]@{
            code = "REVIEW_REQUIRED"
            message = "Scanner candidates are evidence for review and are never automatic deletion authority."
        }
    )
}

$manifestPath = Join-Path $resolvedOutput "implementation-manifest.json"
$maintenancePath = Join-Path $resolvedOutput "maintenance-candidates.json"
$manifestJson = $manifest | ConvertTo-Json -Depth 100
$maintenanceJson = $maintenance | ConvertTo-Json -Depth 100
Write-Utf8NoBom -Path $manifestPath -Content $manifestJson
Write-Utf8NoBom -Path $maintenancePath -Content $maintenanceJson

Assert-JsonAgainstSchema -JsonPath $manifestPath -SchemaPath $resolvedImplementationSchema
Assert-JsonAgainstSchema -JsonPath $maintenancePath -SchemaPath $resolvedMaintenanceSchema

$manifestMarkdown = New-Object System.Collections.Generic.List[string]
$manifestMarkdown.Add('# invSys Static Implementation Manifest')
$manifestMarkdown.Add("")
$manifestMarkdown.Add("- Schema: 1.0.0")
$manifestMarkdown.Add("- Generated: " + $ReportTimestampUtc)
$manifestMarkdown.Add("- Packages: " + $manifest.packages.Count)
$manifestMarkdown.Add("- Components: " + $manifest.components.Count)
$manifestMarkdown.Add("- Procedures: " + $manifest.procedures.Count)
$manifestMarkdown.Add(
    "- Literal Application.Run targets: " + $maintenance.baseline.literalApplicationRunCount
)
$manifestMarkdown.Add(
    "- Unresolved dynamic calls: " + $maintenance.baseline.unresolvedApplicationRunCount
)
$manifestMarkdown.Add("")
$manifestMarkdown.Add("## Packages")
$manifestMarkdown.Add("")
$manifestMarkdown.Add("| Package | Project | Output | Components |")
$manifestMarkdown.Add("|---|---|---|---:|")
foreach ($package in @($manifest.packages)) {
    $manifestMarkdown.Add(
        "| $($package.key) | $($package.projectName) | $($package.outputFile) | " +
        "$(@($package.componentNames).Count) |"
    )
}
$manifestMarkdown.Add("")
$manifestMarkdown.Add("## Dynamic roots")
$manifestMarkdown.Add("")
$manifestMarkdown.Add("| Procedure | Kind | Source |")
$manifestMarkdown.Add("|---|---|---|")
foreach ($root in @($manifest.dynamicRoots)) {
    $manifestMarkdown.Add("| $($root.procedure) | $($root.rootKind) | $($root.source) |")
}
$manifestMarkdown.Add("")
$manifestMarkdown.Add("## Warnings")
$manifestMarkdown.Add("")
if (@($manifest.warnings).Count -eq 0) {
    $manifestMarkdown.Add("- None.")
}
else {
    foreach ($warning in @($manifest.warnings)) {
        $manifestMarkdown.Add(
            "- $($warning.code) - $($warning.message) [$($warning.sourcePath)]"
        )
    }
}

$maintenanceMarkdown = New-Object System.Collections.Generic.List[string]
$maintenanceMarkdown.Add('# invSys VBA Maintenance Candidates')
$maintenanceMarkdown.Add("")
$maintenanceMarkdown.Add("- Schema: 1.0.0")
$maintenanceMarkdown.Add("- Generated: " + $ReportTimestampUtc)
$maintenanceMarkdown.Add("- Total candidates: " + @($maintenance.candidates).Count)
$maintenanceMarkdown.Add(
    "- Duplicate-body groups: " + $maintenance.baseline.duplicateBodyCandidateCount
)
$maintenanceMarkdown.Add(
    "- Unresolved dynamic calls: " + $maintenance.baseline.unresolvedApplicationRunCount
)
$maintenanceMarkdown.Add("")
$maintenanceMarkdown.Add("Scanner output is review evidence only. It never authorizes automatic deletion.")
$maintenanceMarkdown.Add("")
$maintenanceMarkdown.Add("## Candidates")
$maintenanceMarkdown.Add("")
$maintenanceMarkdown.Add("| Type | Confidence | Component | Procedures | Reason |")
$maintenanceMarkdown.Add("|---|---|---|---|---|")
foreach ($candidate in @($maintenance.candidates)) {
    $safeReason = ([string]$candidate.reason).Replace("|", "\|")
    $maintenanceMarkdown.Add(
        "| $($candidate.candidateType) | $($candidate.confidence) | " +
        "$($candidate.componentName) | $($candidate.procedureNames -join ', ') | $safeReason |"
    )
}

Write-Utf8NoBom `
    -Path (Join-Path $resolvedOutput "implementation-manifest.md") `
    -Content ($manifestMarkdown -join "`n")
Write-Utf8NoBom `
    -Path (Join-Path $resolvedOutput "maintenance-candidates.md") `
    -Content ($maintenanceMarkdown -join "`n")

Write-Host "Static VBA surface inventory complete."
Write-Host ("Packages: " + $manifest.packages.Count)
Write-Host ("Components: " + $manifest.components.Count)
Write-Host ("Procedures: " + $manifest.procedures.Count)
Write-Host ("Candidates: " + @($maintenance.candidates).Count)
Write-Host ("Output: " + $resolvedOutput)
