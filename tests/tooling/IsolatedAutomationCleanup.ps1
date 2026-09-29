# Only for a completed, isolated automation worker that owns every supplied COM
# reference. Never use against an interactive/operational Excel instance.
function Release-IsolatedAutomationReferences {
    param([object[]]$Roots)
    $pending=[Collections.Generic.Stack[object]]::new()
    $alreadyReleased=0
    foreach($root in $Roots){
        try{$pending.Push($root)}
        catch [Runtime.InteropServices.InvalidComObjectException]{$alreadyReleased++}
    }
    $visited=[Collections.Generic.List[object]]::new()
    $com=[Collections.Generic.List[object]]::new()
    while($pending.Count -gt 0){
        $value=$pending.Pop()
        if($null -eq $value){continue}
        $isCom=[Runtime.InteropServices.Marshal]::IsComObject($value)
        if(-not $isCom -and $value -isnot [Collections.IDictionary] -and $value -isnot [Collections.IList]){continue}
        $known=$false
        foreach($item in $visited){if([object]::ReferenceEquals($item,$value)){$known=$true;break}}
        if($known){continue}
        $visited.Add($value)
        if($isCom){$com.Add($value);continue}
        if($value -is [Collections.IDictionary]){
            foreach($key in $value.Keys){
                try{$pending.Push($value[$key])}
                catch [Runtime.InteropServices.InvalidComObjectException]{$alreadyReleased++}
            }
        }else{
            foreach($child in $value){
                try{$pending.Push($child)}
                catch [Runtime.InteropServices.InvalidComObjectException]{$alreadyReleased++}
            }
        }
    }
    $released=0;$failed=0
    foreach($value in $com){
        try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value);$released++}
        catch [Runtime.InteropServices.InvalidComObjectException]{$alreadyReleased++}
        catch {$failed++}
    }
    # Counts only: never names, values, exception text, addresses or credentials.
    [pscustomobject]@{UniqueComReferences=$com.Count;Released=$released;AlreadyReleased=$alreadyReleased;ReleaseFailures=$failed}
}

function Release-IsolatedAutomationVariables {
    param([System.Management.Automation.PSVariable[]]$Variables)
    $roots=[Collections.Generic.List[object]]::new()
    $skipped=0
    foreach($variable in $Variables){
        try{$roots.Add($variable.Value)}
        catch [Runtime.InteropServices.InvalidComObjectException]{$skipped++}
    }
    $result=Release-IsolatedAutomationReferences -Roots $roots.ToArray()
    $result|Add-Member -NotePropertyName AlreadyReleasedVariables -NotePropertyValue $skipped
    $result
}
