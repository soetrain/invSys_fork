function Test-GuideTransferPermissions($Source,$Destination,$Guide,[string]$InputPath,[string]$Files) {
    TransferOpen $Source;BoundSelect $Guide
    try {
        SaveGuideExpectationVisibility $false
        $before=TransferPins $Source.Root;$other=TransferPins $Destination.Root
        $path=Join-Path $Files 'hidden-export.json';TransferSetFile $path
        $clicked=BoundControl 'btnExportGuide' 'Click'
        $denied=$clicked -ceq 'DISABLED' -or ($clicked -ceq 'DELIVERED' -and (BoundControl 'lblPublishedGuideStatus' 'Label') -match '(?i)(hidden|policy|unavailable)')
        Check 'GuideTransfer.Export.RestrictedContentRefusedWhole' ($denied -and -not (Test-Path $path) -and (BoundSame $before (TransferPins $Source.Root)) -and (BoundSame $other (TransferPins $Destination.Root)))
    } finally {TransferSetFile '';CloseRecordingViewer;SaveGuideExpectationVisibility $true}
    TransferOpen $Destination
    try {
        SaveGuideExpectationVisibility $false
        TransferImportAttempt 'DestinationVisibility' $InputPath $Source $Destination
    } finally {CloseRecordingViewer;SaveGuideExpectationVisibility $true}
    foreach($mode in @('Export','Import')){
        $fixture=$Source;if($mode -ceq 'Import'){$fixture=$Destination}
        TransferOpen $fixture;if($mode -ceq 'Export'){BoundSelect $Guide}
        $authPath=Join-Path $fixture.Root ($fixture.Warehouse+'.invSys.Auth.xlsb')
        $authBytes=[IO.File]::ReadAllBytes($authPath)
        try {
            $auth=$excel.Workbooks.Open($authPath,0,$false)
            try {
                $caps=Table $auth 'tblCapabilities';$revoked=0
                foreach($row in $caps.ListRows){
                    if($row.Range.Cells.Item(1,$caps.ListColumns.Item('UserId').Index).Value2 -ceq 'config-admin' -and $row.Range.Cells.Item(1,$caps.ListColumns.Item('Capability').Index).Value2 -ceq 'ACTION_PATH_MAINT'){
                        $row.Range.Cells.Item(1,$caps.ListColumns.Item('Status').Index).Value2='Inactive';$revoked++
                    }
                }
                if($revoked -ne 1){throw 'Unique maintenance capability fixture unavailable.'}
                $auth.Save()
            } finally {$auth.Close($false)}
            [void](Run 'invSys.Core.xlam' 'modAuth.LoadAuth' @($fixture.Warehouse))
            $allowed=Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-admin',$fixture.Warehouse,'S1')
            $signed=Run 'invSys.Core.xlam' 'modAuth.IsSignedIn'
            if($allowed -isnot [bool] -or $allowed -or $signed -isnot [bool] -or -not $signed){throw 'Signed-in capability-revocation fixture failed.'}
            $before=TransferPins $Source.Root;$other=TransferPins $Destination.Root
            $path=$InputPath;if($mode -ceq 'Export'){$path=Join-Path $Files 'revoked-export.json'}
            TransferSetFile $path
            $clicked=BoundControl ('btn'+$mode+'Guide') 'Click'
            $denied=$clicked -ceq 'DISABLED' -or ($clicked -ceq 'DELIVERED' -and (BoundControl 'lblPublishedGuideStatus' 'Label') -match '(?i)(ACTION_PATH_MAINT|unavailable|permission|requires|denied)')
            $fileOK=$true;if($mode -ceq 'Export'){$fileOK=-not (Test-Path -LiteralPath $path)}
            Check ('GuideTransfer.'+$mode+'.MaintenanceRevocationRefused') ($denied -and $fileOK -and (BoundSame $before (TransferPins $Source.Root)) -and (BoundSame $other (TransferPins $Destination.Root)))
        } finally {
            TransferSetFile '';CloseRecordingViewer;[IO.File]::WriteAllBytes($authPath,$authBytes);SelectTarget $fixture 'config-admin'
        }
    }
}
