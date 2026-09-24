# Native interaction evidence only; compares the scrollbar interior, excluding
# borders, arrows and the guarded input point. No screenshot or source is edited.
function Measure-DetailScrollMotion([string]$Before,[string]$After,[string]$Geometry) {
    if($Geometry -cnotmatch '^\d+\|\d+\|(?<thickness>\d+)\|\d+\|\d+\|screen=\d+,\d+\|list=(?<left>\d+),(?<top>\d+),(?<right>\d+),(?<bottom>\d+)\|form=(?<formLeft>\d+),(?<formTop>\d+),(?<formRight>\d+),(?<formBottom>\d+)\|dpi=\d+$'){
        throw 'Native scrollbar evidence geometry is invalid.'
    }
    $thickness=[int]$Matches.thickness
    $formWidth=[int]$Matches.formRight-[int]$Matches.formLeft
    $formHeight=[int]$Matches.formBottom-[int]$Matches.formTop
    if($formWidth -le 0 -or $formHeight -le 0){throw 'Native form dimensions are invalid.'}
    Add-Type -AssemblyName System.Drawing
    $first=$null;$last=$null
    try {
        $first=[Drawing.Bitmap]::new($Before);$last=[Drawing.Bitmap]::new($After)
        # Input geometry may be DPI-virtualized; SaveWindow captures physical
        # pixels. Map the verified whole-window coordinates into that image.
        $scaleX=$first.Width/$formWidth;$scaleY=$first.Height/$formHeight
        $x1=[int][Math]::Ceiling(([int]$Matches.left-[int]$Matches.formLeft+$thickness+4)*$scaleX)
        $x2=[int][Math]::Floor(([int]$Matches.right-[int]$Matches.formLeft-2*$thickness-12)*$scaleX)
        $y1=[int][Math]::Ceiling(([int]$Matches.bottom-[int]$Matches.formTop-$thickness+4)*$scaleY)
        $y2=[int][Math]::Floor(([int]$Matches.bottom-[int]$Matches.formTop-4)*$scaleY)
        if($first.Width -ne $last.Width -or $first.Height -ne $last.Height -or $x1 -lt 0 -or $y1 -lt 0 -or $x2 -ge $first.Width -or $y2 -ge $first.Height -or $x2-$x1 -lt 100 -or $y2-$y1 -lt 3){
            throw 'Native scrollbar evidence bounds are unavailable.'
        }
        $changed=0;$leftChanged=0;$rightChanged=0;$middle=($x1+$x2)/2
        for($y=$y1;$y -lt $y2;$y++){
            for($x=$x1;$x -lt $x2;$x++){
                if($first.GetPixel($x,$y).ToArgb() -ne $last.GetPixel($x,$y).ToArgb()){
                    $changed++
                    if($x -lt $middle){$leftChanged++}else{$rightChanged++}
                }
            }
        }
        [pscustomobject]@{ChangedInteriorPixels=$changed;LeftHalfChanged=$leftChanged;RightHalfChanged=$rightChanged;MovementObserved=($leftChanged -ge 500 -and $rightChanged -ge 500);ComparedInteriorPixels=($x2-$x1)*($y2-$y1);ImageScaleX=$scaleX;ImageScaleY=$scaleY}
    } finally {
        if($null -ne $first){$first.Dispose()}
        if($null -ne $last){$last.Dispose()}
    }
}
