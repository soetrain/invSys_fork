# Only for a completed Excel host owned by an isolated automation worker that
# owns every supplied COM reference. At restart, close all old workbooks and Quit
# before release; create the replacement afterward. Never use against an
# interactive/operational Excel instance.
function Release-IsolatedAutomationReferences {
    param([object[]]$Roots)
    $pending=[Collections.Generic.Stack[object]]::new()
    $alreadyReleased=0
    foreach($root in $Roots){
        try{$pending.Push($root)}
        catch [Runtime.InteropServices.InvalidComObjectException]{$alreadyReleased++}
    }
    $visited=[Runtime.Serialization.ObjectIDGenerator]::new()
    $com=[Collections.Generic.List[object]]::new()
    while($pending.Count -gt 0){
        try{
            $value=$pending.Pop()
            if($null -eq $value){continue}
            $isCom=[Runtime.InteropServices.Marshal]::IsComObject($value)
            if(-not $isCom -and $value -isnot [Collections.IDictionary] -and $value -isnot [Collections.IList]){continue}
        }catch [Runtime.InteropServices.InvalidComObjectException]{$alreadyReleased++;continue}
        # Keep identity comparison inside .NET; PowerShell's two-object binding
        # can fail when either stored reference is an already-released wrapper.
        $firstVisit=$false
        try{[void]$visited.GetId($value,[ref]$firstVisit)}
        catch [Runtime.InteropServices.InvalidComObjectException]{$alreadyReleased++;continue}
        if(-not $firstVisit){continue}
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
