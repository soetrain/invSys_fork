# Cleanup only for the disposable ProductionReusable validator host, after its
# workflow assertions. No operational workbook or runtime VBA changes.
. (Join-Path $PSScriptRoot 'IsolatedAutomationCleanup.ps1')

function Close-ReusableFixtureWorkbooks {
    param($Excel,[string]$RuntimeRoot)
    $root=[IO.Path]::GetFullPath($RuntimeRoot).TrimEnd('\')+'\'
    if(-not $root.StartsWith([IO.Path]::GetFullPath([IO.Path]::GetTempPath()),[StringComparison]::OrdinalIgnoreCase) -or
        (Split-Path -Leaf $RuntimeRoot) -notlike 'invsys-plan022-launcher-red-*'){
        throw 'Reusable cleanup requires its isolated temporary runtime.'
    }
    $books=@($Excel.Workbooks|Where-Object {-not [bool]$_.IsAddin})
    # Validate the entire set before closing any workbook.
    foreach($book in $books){
        if(-not [IO.Path]::GetFullPath([string]$book.FullName).StartsWith($root,[StringComparison]::OrdinalIgnoreCase)){
            throw 'Reusable cleanup found a workbook outside its isolated runtime.'
        }
    }
    foreach($book in $books){$book.Close($false)}
    $remaining=[int]$Excel.Workbooks.Count
    foreach($book in $books){Release-ComObject $book}
    if($remaining -ne 0){throw 'Reusable fixture workbook closure is incomplete.'}
    [pscustomobject]@{ClosedWorkbooks=$books.Count;RemainingWorkbooks=$remaining}
}

function Close-ReusableLoadedPackages {
    param([hashtable]$Packages,[string[]]$PackageNames,[string]$PackageRoot)
    foreach($name in $PackageNames){
        $expected=[IO.Path]::GetFullPath((Join-Path $PackageRoot $name))
        if(-not [IO.Path]::GetFullPath([string]$Packages[$name].FullName).Equals($expected,[StringComparison]::OrdinalIgnoreCase)){
            throw 'Reusable cleanup package ownership failed.'
        }
    }
    $reverse=[string[]]$PackageNames.Clone();[array]::Reverse($reverse)
    foreach($name in $reverse){
        $Packages[$name].Close($false)
        Release-ComObject $Packages[$name]
    }
    [pscustomobject]@{ClosedPackages=$reverse.Count;ReverseOrder=$true}
}

function Wait-ReusableAutomationExit {
    param([int]$ProcessId,[System.Management.Automation.PSVariable[]]$Variables)
    if($ProcessId -le 0){throw 'Reusable cleanup has no owned process identity.'}
    $owner=Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
    if($null -ne $owner -and $owner.ProcessName -cne 'EXCEL'){throw 'Reusable cleanup process identity changed.'}
    $references=$null;$failure=$null
    try{
        $references=Release-IsolatedAutomationVariables -Variables $Variables
        [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect()
    }catch{
        $failure=[pscustomobject]@{ExceptionType=$_.Exception.GetType().Name;HResult=$_.Exception.GetBaseException().HResult}
    }
    $start=[DateTimeOffset]::UtcNow
    $clock=[Diagnostics.Stopwatch]::StartNew()
    $exited=($null -eq $owner)
    if($null -ne $owner){$exited=$owner.WaitForExit(30000)}
    $clock.Stop()
    [pscustomobject]@{
        References=$references;Failure=$failure;UnassistedExitObserved=$exited
        WaitLimitMilliseconds=30000;WaitElapsedMilliseconds=$clock.ElapsedMilliseconds
        WaitStartUTC=$start.ToString('o');WaitEndUTC=[DateTimeOffset]::UtcNow.ToString('o')
    }
}
