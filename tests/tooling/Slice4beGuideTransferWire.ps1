# Independent file assertions. Fixtures embed actual authored guide bytes; they
# never fabricate observations, successful executions or inventory evidence.
function TransferHash([string]$Text) {
    $sha=[Security.Cryptography.SHA256]::Create()
    try{[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::ASCII.GetBytes($Text))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
}
function TransferJson($Value) {
    if($Value -is [string]){
        return '"'+[regex]::Replace($Value,'["\\\x00-\x1f\x7f-\uffff]',{
            param($m)
            if($m.Value -ceq '"'){return '\"'}
            if($m.Value -ceq '\'){return '\\'}
            return '\u'+([int][char]$m.Value).ToString('X4')
        })+'"'
    }
    if($Value -is [bool]){if($Value){return 'true'}else{return 'false'}}
    if($Value -is [int] -or $Value -is [long]){return $Value.ToString([Globalization.CultureInfo]::InvariantCulture)}
    if($Value -is [Collections.IDictionary]){
        $parts=@(foreach($key in $Value.Keys){(TransferJson ([string]$key))+':'+(TransferJson $Value[$key])})
        return '{'+($parts -join ',')+'}'
    }
    if($Value -is [Collections.IEnumerable]){
        $parts=@(foreach($item in $Value){TransferJson $item})
        return '['+($parts -join ',')+']'
    }
    if($Value -is [pscustomobject]){
        $parts=@(foreach($property in $Value.PSObject.Properties){(TransferJson $property.Name)+':'+(TransferJson $property.Value)})
        return '{'+($parts -join ',')+'}'
    }
    throw 'Unsupported transfer fixture value.'
}
function TransferSeal([string]$Body) {
    if(-not $Body.EndsWith('}') -or $Body -match '[^\x00-\x7f]'){throw 'Invalid fixture body.'}
    $Body.Substring(0,$Body.Length-1)+',"ContentSha256":"'+(TransferHash $Body)+'"}'
}
function TransferRead([string]$Path,[string]$Kind,[int]$Schema,[string]$Fields) {
    if(-not (Test-Path -LiteralPath $Path -PathType Leaf)){return $null}
    $raw=[IO.File]::ReadAllText($Path)
    if($raw -match '[^\x00-\x7f]' -or $raw.Length -gt 1048576 -or (Get-Item -LiteralPath $Path).Length -ne $raw.Length){return $null}
    $match=[regex]::Match($raw,',"ContentSha256":"([0-9a-f]{64})"\}$')
    if(-not $match.Success){return $null}
    $body=$raw.Substring(0,$match.Index)+'}'
    if((TransferHash $body) -cne $match.Groups[1].Value){return $null}
    try{$model=$raw|ConvertFrom-Json}catch{return $null}
    if($model.RecordKind -cne $Kind -or $model.SchemaVersion -ne $Schema){return $null}
    if((@($model.PSObject.Properties.Name|Sort-Object) -join '|') -cne (@($Fields -split '\|'|Sort-Object) -join '|')){return $null}
    return $model
}
function TransferSetFile([string]$Path,[bool]$SignOut=$false) {
    if($guideTransferDialogProbeInstalled){[void](Run 'invSys.Operations.xlam' 'modGuideTransferUi.SetTransferFileForTest' @($Path,$SignOut))}
}
function TransferPins([string]$Root) {
    $pins=@{}
    if(Test-Path -LiteralPath $Root){
        foreach($file in Get-ChildItem -LiteralPath $Root -Recurse -File){
            # Ordinary activity is a separate allowed append-only effect.
            if($file.FullName -match '[\\/]Training[\\/]Activity[\\/]'){continue}
            $stream=[IO.File]::Open($file.FullName,'Open','Read','ReadWrite')
            try{$pins[$file.FullName]=(Get-FileHash -InputStream $stream).Hash}finally{$stream.Dispose()}
        }
    }
    return $pins
}
function TransferRetained($Before,$After) {
    foreach($path in $Before.Keys){if(-not $After.ContainsKey($path) -or $Before[$path] -cne $After[$path]){return $false}}
    return $true
}
function TransferOpen($Fixture,[string]$User='config-admin') {
    CloseRecordingViewer;SelectTarget $Fixture $User;OpenRecordingViewer
    if((BoundLibrary 'Open') -cne 'DELIVERED'){throw 'Existing Action Paths entry unavailable; not transfer RED.'}
    BoundOpen
}
function TransferImportAttempt([string]$Name,[string]$Path,$Source,$Destination,[bool]$LoseContext=$false) {
    $before=TransferPins $Destination.Root;$sourceBefore=TransferPins $Source.Root
    $fileHash=(Get-FileHash -LiteralPath $Path).Hash
    TransferSetFile $Path $LoseContext
    $delivered=(BoundControl 'btnImportGuide' 'Click') -ceq 'DELIVERED'
    $status=BoundControl 'lblPublishedGuideStatus' 'Label'
    $refused=$status -match '(?i)(unavailable|unsupported|invalid|exceed|already exists|refus|denied|permission|changed|hidden|requires|cannot|could not)'
    Check ('GuideTransfer.Import.Refuses.'+$Name) ($delivered -and $refused -and (BoundSame $before (TransferPins $Destination.Root)) -and (BoundSame $sourceBefore (TransferPins $Source.Root)) -and (Get-FileHash -LiteralPath $Path).Hash -ceq $fileHash)
    TransferSetFile ''
}
