# Shared actual-handler selection, integrity and policy fixture helpers.
function BoundControl([string]$Name,[string]$Action,[string]$Value='', [string]$Form='frmActionPathLibrary') {
        [string](Run 'invSys.Operations.xlam' 'modInventoryViewer.GuideDraftControlForTest' @($Form,$Name,$Action,$Value))
    }

function BoundLibrary([string]$Action,[string]$Value='') {
        [string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @($Action,$Value))
    }

function BoundOpen {[void](BoundControl 'btnPublishedGuides' 'Click' '' 'frmActionPaths')}

function BoundSelect($Guide) {
        $key=[string]$Guide.ActionPathId+'|'+[string]$Guide.Version+'|'+[string]$Guide.ContentSha256
        $keys=@((BoundControl 'lstPublishedGuides' 'Values') -split "`n"|Where-Object {$_ -cne ''})
        $index=[Array]::IndexOf($keys,$key)
        if($index -lt 0){throw 'Accepted exact-version reader fixture is unavailable; not guide-binding RED.'}
        if((BoundControl 'lstPublishedGuides' 'Select' ([string]$index)) -cne 'SELECTED'){throw 'Accepted guide selection fixture failed.'}
    }

function BoundPins([string]$Root) {
        $pins=@{}
        if(Test-Path -LiteralPath $Root){foreach($file in Get-ChildItem -LiteralPath $Root -Recurse -File){$pins[$file.FullName]=(Get-FileHash -LiteralPath $file.FullName).Hash}}
        return $pins
    }

function BoundSame($Before,$After) {
        if($Before.Count -ne $After.Count){return $false}
        foreach($path in $Before.Keys){if(-not $After.ContainsKey($path) -or $Before[$path] -cne $After[$path]){return $false}}
        return $true
    }

function ReadGuideExpectationRecord([string]$Path) {
        if(-not (Test-Path -LiteralPath $Path -PathType Leaf)){return $null}
        try {
            $text=[IO.File]::ReadAllText($Path);$length=(Get-Item -LiteralPath $Path).Length
            if($text -match '[^\x00-\x7f]' -or $length -ne $text.Length -or $length -gt 1048576){return $null}
            $value=$text|ConvertFrom-Json
            $match=[regex]::Match($text,',"ContentSha256":"([0-9a-f]{64})"\}$')
            if(-not $match.Success -or $value.RecordKind -cne 'Guide' -or $value.SchemaVersion -ne 1 -or @($value.PSObject.Properties).Count -ne 24){return $null}
            $body=$text.Substring(0,$match.Index)+'}';$sha=[Security.Cryptography.SHA256]::Create()
            try{$hash=[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($body))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
            if($value.ContentSha256 -cne $hash){return $null}
            return $value
        } catch {return $null}
    }

function SaveGuideExpectationVisibility([bool]$Visible) {
        try {
            [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
            $result=Run 'invSys.Admin.xlam' 'TestD5Commands.PublishedReadVisibilityForTest' @($Visible)
            if($result -isnot [bool] -or -not $result){throw 'Actual visibility command fixture failed; not expectation RED.'}
        } finally {[void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')}
    }
