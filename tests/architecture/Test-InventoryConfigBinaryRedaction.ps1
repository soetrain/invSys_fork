[CmdletBinding()]
param([string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)))

# Synthetic MS-XLSB vectors only: no Excel, workbooks, user values or credentials.
# brtRevisionPtr: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/fae2cb92-4884-4752-97cb-0c1f31a31e0d
# XLWideString: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/5253755b-33cb-4796-835d-caf07bf70ad4
$ErrorActionPreference = 'Stop'
. (Join-Path $RepoRoot 'tests/tooling/Slice4beInventoryCommandBoundaries.ps1')
$script:failed = 0
$script:checked = 0
function Check-ConfigVector([string]$Name, [bool]$Passed) {
    $script:checked++
    if ($Passed) { Write-Output ('PASS ' + $Name) }
    else { $script:failed++; Write-Output ('FAIL ' + $Name) }
}
function Observed-Property($Result, [string]$Name) {
    $property = $Result.PSObject.Properties[$Name]
    if ($null -ne $property) { return $property.Value }
}
function Revision-Vector([string]$Document, [uint64]$Revision = 1,
                         [uint32]$Last = 2, [uint32]$Maximum = 3, [byte]$SaveByte = 4) {
    $bytes = [Collections.Generic.List[byte]]::new()
    $bytes.AddRange([BitConverter]::GetBytes($Revision))
    $bytes.AddRange([BitConverter]::GetBytes([uint32]$Document.Length))
    $bytes.AddRange([Text.Encoding]::Unicode.GetBytes($Document))
    $bytes.AddRange([BitConverter]::GetBytes($Last))
    $bytes.AddRange([BitConverter]::GetBytes($Maximum))
    $bytes.AddRange([byte[]](@($SaveByte) * 16))
    return ,$bytes.ToArray()
}
function Record-Vector([byte[]]$Payload, [int]$Type = 3073) {
    # Exercise the existing framing parser as well as the comparison boundary.
    $bytes = [Collections.Generic.List[byte]]::new()
    $bytes.Add([byte](($Type -band 127) -bor 128))
    $bytes.Add([byte]($Type -shr 7))
    $remaining = $Payload.Length
    do {
        $part = $remaining -band 127
        $remaining = $remaining -shr 7
        if ($remaining) { $part = $part -bor 128 }
        $bytes.Add([byte]$part)
    } while ($remaining)
    $bytes.AddRange($Payload)
    return Get-InventoryConfigBinaryRecords $bytes.ToArray()
}
function Compare-Vector([byte[]]$Before, [byte[]]$After) {
    return Compare-InventoryConfigBinaryRecords (Record-Vector $Before) (Record-Vector $After)
}
function Revision-Fields($Result) {
    return (@(Observed-Property $Result 'ChangedRevisionPointerFields') | Sort-Object) -join '|'
}

$base = Revision-Vector 'DOCUMENT-CANARY'
$same = Compare-Vector $base $base
Check-ConfigVector 'Unchanged.RecordFraming' ($same.Parsed -and $same.RecordOrderPreserved -and $same.RecordHeadersPreserved -and $same.ChangedRecordTypes.Count -eq 0)
Check-ConfigVector 'Unchanged.KnownRevisionLayout' ((Revision-Fields $same) -ceq '' -and (Observed-Property $same 'UnknownRevisionPointerChange') -ceq $false)

$save = Compare-Vector $base (Revision-Vector 'DOCUMENT-CANARY' -SaveByte 5)
Check-ConfigVector 'SaveIdentity.OnlyFixedFieldName' ((Revision-Fields $save) -ceq 'uidLastSave' -and (Observed-Property $save 'UnknownRevisionPointerChange') -ceq $false)
Check-ConfigVector 'SaveIdentity.OtherRecordEvidenceRetained' ($save.Parsed -and $save.RecordOrderPreserved -and $save.RecordHeadersPreserved -and ($save.ChangedRecordTypes -join '|') -ceq '3073')

$all = Compare-Vector $base (Revision-Vector 'DIFFERENT-LENGTH-DOCUMENT-CANARY' -Revision 7 -Last 8 -Maximum 9 -SaveByte 10)
Check-ConfigVector 'VariableLength.AllFixedFieldNames' ((Revision-Fields $all) -ceq 'CoauthVersionLast|CoauthVersionMax|documentId|uidLastSave|xrevidLastSaved')
Check-ConfigVector 'VariableLength.KnownLayout' ((Observed-Property $all 'UnknownRevisionPointerChange') -ceq $false)
$singleObservations = @()
foreach ($case in @(
    @('xrevidLastSaved', (Revision-Vector 'DOCUMENT-CANARY' -Revision 7)),
    @('CoauthVersionLast', (Revision-Vector 'DOCUMENT-CANARY' -Last 8)),
    @('CoauthVersionMax', (Revision-Vector 'DOCUMENT-CANARY' -Maximum 9)),
    @('documentId', (Revision-Vector 'OTHER-DOCUMENT-CANARY'))
)) {
    $single = Compare-Vector $base $case[1]
    $singleObservations += $single
    Check-ConfigVector ('SingleField.' + $case[0]) ((Revision-Fields $single) -ceq $case[0] -and (Observed-Property $single 'UnknownRevisionPointerChange') -ceq $false)
}

$short = Compare-Vector $base ([byte[]](1,2,3))
Check-ConfigVector 'TruncatedPayload.RemainsUnclassified' ((Observed-Property $short 'UnknownRevisionPointerChange') -ceq $true -and (Revision-Fields $short) -ceq '')
$badLength = [byte[]]$base.Clone()
[Array]::Copy([BitConverter]::GetBytes([uint32]::MaxValue), 0, $badLength, 8, 4)
$bad = Compare-Vector $base $badLength
Check-ConfigVector 'InvalidStringLength.RemainsUnclassified' ((Observed-Property $bad 'UnknownRevisionPointerChange') -ceq $true -and (Revision-Fields $bad) -ceq '')
$trailing = Compare-Vector $base ([byte[]]($base + @(0)))
Check-ConfigVector 'TrailingPayload.RemainsUnclassified' ((Observed-Property $trailing 'UnknownRevisionPointerChange') -ceq $true -and (Revision-Fields $trailing) -ceq '')

$missing = Compare-InventoryConfigBinaryRecords (Record-Vector $base) (Get-InventoryConfigBinaryRecords ([byte[]]@()))
Check-ConfigVector 'MissingRecord.RemainsUnclassified' ((Observed-Property $missing 'UnknownRevisionPointerChange') -ceq $true)
$other = Compare-InventoryConfigBinaryRecords (Record-Vector ([byte[]](1)) 778) (Record-Vector ([byte[]](2)) 778)
Check-ConfigVector 'OtherRecord.NotMisclassified' (($other.ChangedRecordTypes -join '|') -ceq '778' -and (Revision-Fields $other) -ceq '' -and (Observed-Property $other 'UnknownRevisionPointerChange') -ceq $false)

$observations = @($same, $save, $all, $short, $bad, $trailing, $missing, $other) + $singleObservations
$json = $observations | ConvertTo-Json -Depth 6 -Compress
Check-ConfigVector 'Redaction.NoPayloadValuesOrHashes' ($json -notmatch 'CANARY|"(?:Data|Hash|HeaderHash|Payload|DocumentValue)"\s*:')
$allowed = @('Parsed','BeforeRecordCount','AfterRecordCount','RecordOrderPreserved',
    'RecordHeadersPreserved','ChangedRecordTypes','ChangedBookViewFields',
    'UnknownBookViewChange','AllWorkbookSemanticsProven','ChangedRevisionPointerFields',
    'UnknownRevisionPointerChange','ListPrefixChanges')
$labels = @('xrevidLastSaved','documentId','CoauthVersionLast','CoauthVersionMax','uidLastSave')
$safeShape = $true
foreach ($observation in $observations) {
    foreach ($property in $observation.PSObject.Properties) {
        if ($property.Name -cnotin $allowed) { $safeShape = $false }
    }
    foreach ($label in @(Observed-Property $observation 'ChangedRevisionPointerFields')) {
        if ($label -cnotin $labels) { $safeShape = $false }
    }
}
Check-ConfigVector 'Redaction.OnlyApprovedSummaryProperties' $safeShape
Check-ConfigVector 'Scope.NoBlanketSemanticProof' (@($observations | Where-Object { $_.AllWorkbookSemanticsProven -ne $false }).Count -eq 0)
$partBefore=@{};$partAfter=@{}
foreach($path in @('xl/worksheets/sheet1.bin','xl/tables/table2.bin','xl/styles.bin','xl/CANARY.bin')){
    $partBefore[$path]=Record-Vector ([Text.Encoding]::UTF8.GetBytes('BEFORE-CANARY')) 778
    $partAfter[$path]=Record-Vector ([Text.Encoding]::UTF8.GetBytes('AFTER--CANARY')) 778
}
$parts=@(Compare-InventoryConfigPartRecords $partBefore $partAfter)
Check-ConfigVector 'Parts.FixedCategoriesOnly' (($parts.Category|Sort-Object) -join '|' -ceq 'Styles|Table|Worksheet')
Check-ConfigVector 'Parts.ChangedRecordTypes' (@($parts|Where-Object {($_.Comparison.ChangedRecordTypes -join '|') -ceq '778'}).Count -eq 3)
Check-ConfigVector 'Parts.UnchangedOmitted' (@(Compare-InventoryConfigPartRecords $partBefore $partBefore).Count -eq 0)
$missingPart=@(Compare-InventoryConfigPartRecords @{} @{'xl/styles.bin'=$partAfter['xl/styles.bin']})
Check-ConfigVector 'Parts.MissingSideUnclassified' ($missingPart.Count -eq 1 -and -not $missingPart[0].Comparison.Parsed)
$partsJson=(@($parts)+@($missingPart))|ConvertTo-Json -Depth 6 -Compress
Check-ConfigVector 'Parts.Redacted' ($partsJson -notmatch 'CANARY|xl/|"(?:Data|Hash|HeaderHash|Payload|Path)"\s*:')
Check-ConfigVector 'Parts.NoAcceptanceClaim' ($partsJson -notmatch '"(?:CauseEstablished|Accepted)"\s*:')
$listBase=[byte[]]::new(80);[Text.Encoding]::UTF8.GetBytes('LIST-CANARY').CopyTo($listBase,64)
$listChanged=[byte[]]$listBase.Clone();$listChanged[36]=1
$list=Compare-InventoryConfigBinaryRecords (Record-Vector $listBase 343) (Record-Vector $listChanged 343)
Check-ConfigVector 'List.FixedHeaderFormatField' ($list.ListPrefixChanges.Count -eq 1 -and ($list.ListPrefixChanges[0].ChangedFields -join '|') -ceq 'HeaderFormat' -and -not $list.ListPrefixChanges[0].TailChangedOrUnavailable)
$listChanged[64]=0
$listTail=Compare-InventoryListPrefix $listBase $listChanged
Check-ConfigVector 'List.TailRemainsUnclassified' $listTail.TailChangedOrUnavailable
$listShort=Compare-InventoryListPrefix ([byte[]](1)) $listChanged
Check-ConfigVector 'List.TruncatedPrefixUnavailable' (-not $listShort.FixedPrefixAvailable -and $listShort.TailChangedOrUnavailable)
$listJson=@($list,$listTail,$listShort)|ConvertTo-Json -Depth 8 -Compress
Check-ConfigVector 'List.Redacted' ($listJson -notmatch 'CANARY|"(?:Data|Hash|HeaderHash|Payload|Path)"\s*:')
$tableA=[byte[]]$listBase.Clone();$tableB=[byte[]]$listBase.Clone();$tableB[36]=1
$styles=[pscustomobject]@{Parsed=$true;Records=@{'507:1'=[pscustomobject]@{Hash='PRIVATE-HASH-CANARY'};'507:2'=[pscustomobject]@{Hash='PRIVATE-HASH-CANARY'}}}
$equivalent=Compare-InventoryHeaderFormatReferences (Record-Vector $tableA 343) (Record-Vector $tableB 343) $styles $styles
Check-ConfigVector 'HeaderReference.ExactEquivalentBytes' $equivalent.ExactFormatBytesEqual
$styles.Records['507:2'].Hash='DIFFERENT-HASH-CANARY'
$different=Compare-InventoryHeaderFormatReferences (Record-Vector $tableA 343) (Record-Vector $tableB 343) $styles $styles
Check-ConfigVector 'HeaderReference.DifferentBytesRefused' (-not $different.ExactFormatBytesEqual)
[Array]::Copy([BitConverter]::GetBytes([uint32]::MaxValue),0,$tableB,36,4)
$none=Compare-InventoryHeaderFormatReferences (Record-Vector $tableA 343) (Record-Vector $tableB 343) $styles $styles
Check-ConfigVector 'HeaderReference.NoneIsNotEquivalence' ($none.AfterKind -ceq 'None' -and -not $none.ExactFormatBytesEqual)
$missingStyle=Compare-InventoryHeaderFormatReferences (Record-Vector $tableA 343) (Record-Vector $tableA 343) $null $styles
Check-ConfigVector 'HeaderReference.MissingStyleUnclassified' ($missingStyle.BeforeKind -ceq 'Unavailable' -and -not $missingStyle.ExactFormatBytesEqual)
$referenceJson=@($equivalent,$different,$none,$missingStyle)|ConvertTo-Json -Depth 4 -Compress
Check-ConfigVector 'HeaderReference.Redacted' ($referenceJson -notmatch 'CANARY|"Hash"\s*:')
Write-Output ('CONFIG_BINARY_DIAGNOSTIC passed={0} failed={1} total={2}' -f ($script:checked-$script:failed),$script:failed,$script:checked)
if ($script:failed) { exit 1 }
