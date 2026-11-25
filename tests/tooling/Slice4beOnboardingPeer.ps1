# A second Excel process signs in to the same generated warehouse through the
# packaged form. Copy only the probes needed to deliver those operator actions.
function Test-OnboardingPeer($case,[string]$user,[string]$secret) {
    Invoke-OnboardingPeer $case $user $secret
    Invoke-OnboardingPeer $case $user $secret -VerifyRevoked
}

function Invoke-OnboardingPeer($case,[string]$user,[string]$secret,[switch]$VerifyRevoked) {
    $core='invSys.Core.xlam'
    $admin='invSys.Admin.xlam'
    $peer=$null
    $opened=@{}
    function Copy-ProbeProcedure($from,$to,[string]$procedure) {
        $start=$from.ProcBodyLine($procedure,0)
        $end=$start
        while($end -le $from.CountOfLines -and $from.Lines($end,1).Trim() -notin @('End Sub','End Function')){$end++}
        if($end -gt $from.CountOfLines){throw 'Incomplete peer form adapter.'}
        $to.AddFromString($from.Lines($start,$end-$start+1))
    }
    try {
        $peer=New-Object -ComObject Excel.Application
        if($peer.Hwnd -eq $excel.Hwnd){throw 'An independent Excel session is required.'}
        $peer.DisplayAlerts=$false
        $peer.EnableEvents=$false
        $priorSecurity=$peer.AutomationSecurity
        $peer.AutomationSecurity=1
        foreach($name in @($core,'invSys.Inventory.Domain.xlam','invSys.Designs.Domain.xlam','invSys.Operations.xlam',$admin)){
            $path=Join-Path (Join-Path $repo $DeployRoot) $name
            $opened[$name]=$peer.Workbooks.Open($path,0,$true)
            if($opened[$name].VBProject.FileName -ine $path){throw 'Unexpected peer package identity.'}
        }
        $peer.AutomationSecurity=$priorSecurity
        foreach($binding in @(@($core,'frmSignIn','EnterForTest'),@($admin,'frmCreateDeleteUser','GuideChoiceForTest'))){
            $name,$form,$procedure=$binding
            Copy-ProbeProcedure $packages[$name].VBProject.VBComponents.Item($form).CodeModule $opened[$name].VBProject.VBComponents.Item($form).CodeModule $procedure
        }
        foreach($name in @($core,$admin)){
            $module=$opened[$name].VBProject.VBComponents.Add(1)
            $module.Name='TestOnboardingPeer'
            $original=if($name -eq $core){'TestInitialAdmin'}else{'TestCreateOwner'}
            $procedures=if($name -eq $core){@('SignIn','HasGuide')}else{@('GuideChoice')}
            foreach($proc in $procedures){
                Copy-ProbeProcedure $packages[$name].VBProject.VBComponents.Item($original).CodeModule $module.CodeModule $proc
            }
        }
        $selected=[string]$peer.Run("'$core'!modNasConnection.SelectWarehouseTargetForAutomation",$case.Root,$case.Root,'S1',$false)
        if(-not $selected.StartsWith('OK|')){throw 'Peer could not select generated Training destination.'}
        [void]$peer.Run("'$core'!modNasConnection.SetCurrentTargetPathsForTest",'\\fixture-host\config-command',$case.Root)
        $signed=[bool]$peer.Run("'$core'!TestOnboardingPeer.SignIn",$user,$secret)
        if($VerifyRevoked){
            Check 'Onboarding.ColdReopen.OrdinarySignInAndRevokePersist' ($signed -and -not [bool]$peer.Run("'$core'!TestOnboardingPeer.HasGuide",$user))
            return
        }
        Check 'Onboarding.FreshExcel.CredentialAndGrantPersist' ($signed -and [bool]$peer.Run("'$core'!TestOnboardingPeer.HasGuide",$user))
        if(-not $signed){throw 'Peer ordinary destination sign-in failed.'}
        $allowed=[bool](Run $core 'TestInitialAdmin.TransferAllowed')
        Check 'Onboarding.GuideTransfer.GrantedBeforePeerRevocation' $allowed
        [void]$peer.Run("'$admin'!TestOnboardingPeer.GuideChoice",$user,$false)
        if([bool]$peer.Run("'$core'!TestOnboardingPeer.HasGuide",$user)){throw 'Peer Users & Roles revoke was not persisted.'}
        # Do not refresh the first session. Its next protected action owns the check.
        Check 'Onboarding.GuideTransfer.NextActionReadsRevocation' ($allowed -and -not [bool](Run $core 'TestInitialAdmin.TransferAllowed'))
        if($GuideDraftOnly){
            Test-OpenGuideRevocation $guide {
                param([bool]$grant)
                [void]$peer.Run("'$admin'!TestOnboardingPeer.GuideChoice",$user,$grant)
                if([bool]$peer.Run("'$core'!TestOnboardingPeer.HasGuide",$user) -ne $grant){throw 'Peer role Save did not persist its exact choice.'}
            }
        }
    } finally {
        if($peer){
            foreach($name in @($admin,'invSys.Operations.xlam','invSys.Designs.Domain.xlam','invSys.Inventory.Domain.xlam',$core)){
                if($opened.ContainsKey($name)){$opened[$name].Close($false)}
            }
            $peer.Quit()
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($peer)
        }
    }
}
