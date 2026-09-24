# Called inside the designer gate; shares its exact packaged Probe/Pair assertions.
function Test-ProductionRecipeActivity($RecipeFixture,[string]$Canary) {
    function SavedHash([string]$Path) {
        $stream=[IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
        try{(Get-FileHash -InputStream $stream -Algorithm SHA256).Hash}finally{$stream.Dispose()}
    }
    $recipeBook=$null
    try {
        [void](Probe 'Close')
        SelectTarget $RecipeFixture
        try {
            [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
            if(-not [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveSettings' @('DesignsEnabled','TRUE'))){throw 'Explicit Designs setup failed; not product RED.'}
        } finally {[void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')}
        SelectTarget $RecipeFixture 'config-producer'
        $recipeBook=$excel.Workbooks.Add()
        $path=Join-Path $runRoot 'recipe-observation.xlsb';$recipeBook.SaveAs($path,50)
        [void](Probe 'Open' @($recipeBook.Name))
        $setup=[string](Probe 'ReleasedProcess' @($Canary))
        if($setup -cnotmatch '^READY\|([^|]+)\|([^|]+)$'){throw ('Released-Process setup failed; fixed stage result='+$setup)}
        $identity=$Matches[1];$version=$Matches[2]
        $ready=[bool](Probe 'ReleasedRecipe' @($identity,$version,$Canary))
        Check 'ProductionDesigner.Recipe.Valid.RealReleasedProcess' $ready
        if(-not $ready){throw 'Released Process could not be read through the owning bridge; not product RED.'}
        # Setup intentionally saves a new Process; validation must preserve that saved authority.
        $pins=@{};foreach($file in @(Get-ChildItem $RecipeFixture.Root -Recurse -File|Where-Object {$_.Extension -in '.xlsb','.xlsm' -and -not $_.Name.StartsWith('~$')})){$pins[$file.FullName]=SavedHash $file.FullName}
        $bookPin=SavedHash $path;$state=[string](Probe 'State' @('Recipe'))
        $before=@(Get-Slice4beActivityFiles $RecipeFixture)
        $report=[string](Probe 'Act' @('Recipe','Validate'))
        Check 'ProductionDesigner.Recipe.Valid.ActualValidator' ($report -like 'Recipe graph is valid: nodes=1; connections=0.*')
        Pair $before 'PRODUCTION_RECIPE_VALIDATE' 'VALIDATED' 'ProductionDesigner.Recipe.Valid' -ObservedFixture $RecipeFixture
        Check 'ProductionDesigner.Recipe.Valid.DraftPreserved' ($state -ceq [string](Probe 'State' @('Recipe')))
        $same=(SavedHash $path) -ceq $bookPin;foreach($file in $pins.Keys){$same=$same -and (SavedHash $file) -ceq $pins[$file]}
        Check 'ProductionDesigner.Recipe.Valid.SavedAuthorityPreserved' $same
        if($CheckProductionDesignerPaths){
            . (Join-Path $PSScriptRoot 'Slice4beProductionPaths.ps1')
            Test-ProductionPaths $RecipeFixture $recipeBook.Name $identity $version $Canary
        }
    } finally {
        [void](Probe 'Close')
        if($null -ne $recipeBook){$recipeBook.Close($false)}
    }
}
