# Save only instrumented disposable package copies, before fixture/form activity.
function Save-DisposableStartupProbes($Packages,[string]$ProbeDeploy,[string]$State) {
    if($State -notin @('SavedCopies','WritableCopies')){throw 'Explicit disposable package state required.'}
    foreach($name in @('invSys.Core.xlam','invSys.Inventory.Domain.xlam','invSys.Designs.Domain.xlam','invSys.Operations.xlam','invSys.Admin.xlam')){
        $book=$Packages[$name]
        if($null -eq $book -or $book.ReadOnly -isnot [bool] -or $book.ReadOnly -or
            -not [string]::Equals($book.FullName,(Join-Path $ProbeDeploy $name),[StringComparison]::OrdinalIgnoreCase)){
            throw 'Only the writable disposable probe package may be saved.'
        }
        if($State -eq 'SavedCopies'){
            $book.Save()
            if($book.Saved -isnot [bool] -or -not $book.Saved){throw 'Disposable probe package save is not verified.'}
        }
    }
    if($State -eq 'SavedCopies'){
        Check 'Harness.DisposableStartupProbesSaved' $true
    }else{
        $unsaved=$Packages['invSys.Operations.xlam'].Saved
        if($unsaved -isnot [bool] -or $unsaved){throw 'Writable Operations probes must remain unsaved for this comparison.'}
        Check 'Harness.DisposableStartupProbesUnsaved' $true
    }
}
