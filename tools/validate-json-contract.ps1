[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$JsonPath,

    [Parameter(Mandatory = $true)]
    [string]$SchemaPath
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

function Read-Json {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "JSON file not found: $Path"
    }
    # Preserve JSON string literals on PowerShell 7.5+ without adding a
    # parameter unsupported by the Windows PowerShell 5.1 execution host.
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
        $Value -is [uint64]) {
        return "integer"
    }
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

function Test-SchemaNode {
    param(
        $Value,
        $Schema,
        $RootSchema,
        [string]$Path,
        [System.Collections.Generic.List[string]]$Errors
    )

    if (Test-HasProperty -Object $Schema -Name '$ref') {
        $resolved = Resolve-SchemaReference -RootSchema $RootSchema -Reference $Schema.'$ref'
        Test-SchemaNode -Value $Value -Schema $resolved -RootSchema $RootSchema `
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
        $matched = $false
        foreach ($candidate in @($Schema.enum)) {
            if ([string]$Value -ceq [string]$candidate) {
                $matched = $true
                break
            }
        }
        if (-not $matched) {
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

        $allowedNames = @()
        if (Test-HasProperty -Object $Schema -Name "properties") {
            $allowedNames = @($Schema.properties.PSObject.Properties.Name)
            foreach ($propertyName in $allowedNames) {
                if (Test-HasProperty -Object $Value -Name $propertyName) {
                    Test-SchemaNode -Value $Value.$propertyName `
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
                if ($actualName -notin $allowedNames) {
                    $Errors.Add("$Path contains undeclared property '$actualName'.")
                }
            }
        }
    }
    elseif ($valueType -eq "array" -and (Test-HasProperty -Object $Schema -Name "items")) {
        $index = 0
        foreach ($item in @($Value)) {
            Test-SchemaNode -Value $item -Schema $Schema.items -RootSchema $RootSchema `
                -Path ($Path + "[$index]") -Errors $Errors
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

$resolvedJson = (Resolve-Path -LiteralPath $JsonPath).Path
$resolvedSchema = (Resolve-Path -LiteralPath $SchemaPath).Path
$value = Read-Json $resolvedJson
$schema = Read-Json $resolvedSchema
$errors = New-Object System.Collections.Generic.List[string]
Test-SchemaNode -Value $value -Schema $schema -RootSchema $schema -Path '$' -Errors $errors

if ($errors.Count -gt 0) {
    throw (
        "JSON schema validation failed for $resolvedJson`n" +
        ($errors -join "`n")
    )
}

Write-Host "JSON contract valid: $resolvedJson"
