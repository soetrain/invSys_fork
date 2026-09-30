[CmdletBinding()]
param([string]$RepoRoot = ".")

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$docs = (Resolve-Path -LiteralPath (Join-Path $repo "..\invSys_docs")).Path

function Read-Text([string]$Path) {
    Get-Content -Raw -LiteralPath $Path
}

$form = Read-Text (Join-Path $repo "src\Production\Forms\frmProduction.frm")
$ordering = Read-Text (Join-Path $repo "src\Production\Modules\modProductionRecipeOrderActions.bas")
$validator = Read-Text (Join-Path $repo "tools\validate_plan022_packaged_launchers.ps1")
$spec = Read-Text (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Design-v4.11.md")
$plan = Read-Text (Join-Path $docs "expert guidance docs\022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md")
$controls = Read-Text (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Controls-v1.md")

$checks = @(
    [pscustomobject]@{ Name = "Docs.OutputFirstForkContract"; Pass =
        $spec -match 'Recipe Designer is output-first' -and
        $spec -match 'primary Recipe graph projection is \*\*Output Flow\*\*' -and
        $spec -match 'independent outputs may appear in the same stage and converge' -and
        $plan -match 'Slice 4aq -- Output-first Recipe routing and fork/convergence flow' -and
        $controls -match 'Slice 4aq output-first Recipe routing' },
    [pscustomobject]@{ Name = "Form.OutputFirstCompatibleTargetEditor"; Pass =
        $form -match 'ConfigureCompatibleTargetCombo' -and
        $form -match 'Private Sub RefreshCompatibleDownstreamChoices' -and
        $form -match 'Private Sub mCmbConnectionOutput_Change\(\)' -and
        $form -match 'ProcessRecordsHaveAlternativeItem' -and
        $form -match 'ConnectionRequirementId\(\)' },
    [pscustomobject]@{ Name = "Form.NoVisibleIngredientDropdown"; Pass =
        $form -match 'mCmbConnectionRequirement\.Visible = False' -and
        $form -match '"Produced by / Output / Feeds Process / Required Qty / Required % / UOM"' -and
        $form -notmatch '"Upstream Process / Output / Downstream Process / Input Requirement / Qty / % / UOM"' },
    [pscustomobject]@{ Name = "Form.OutputFlowProjection"; Pass =
        $form -match 'AddLabel pg, "Output Flow"' -and
        $form -match 'Array\("Stage", "Produced by", "Output", "Feeds Process", "Output Qty", "Yield %", "UOM", "", "", ""\)' -and
        $form -match 'mLstRecipeConnectionDisplay.*10' -and
        $form -match '"Finished inventory"' },
    [pscustomobject]@{ Name = "Form.DerivesForkStages"; Pass =
        $form -match 'Private Function BuildRecipeStageMap\(\) As Object' -and
        $form -match 'targetStage = sourceStage \+ 1' -and
        $form -match 'Stage " & CStr' -and
        $form -match 'Private Sub mBtnRecipeAutoOrder_Click\(\)\s+DesignerRecipeOrderAction RecipeOrderAuto\s+End Sub' -and
        $form -match 'Private Sub DesignerRecipeOrderAction[\s\S]{0,400}modProductionRecipeOrderActions.EditOrder\(Me, action,' -and
        $ordering -match 'owner.ApplyRecipeOrderAction\(action, report\)' -and
        $form -match 'Public Function ApplyRecipeOrderAction[\s\S]{0,1600}If action = RecipeOrderAuto Then\s+RefreshRecipeConnectionDisplay' },
    [pscustomobject]@{ Name = "Form.HiddenGraphIdentityPreserved"; Pass =
        $form -match 'fields = Array\(ComboText\(mCmbConnectionFromNode\), ConnectionOutputId\(\), _\s+ConnectionTargetNodeId\(\), ConnectionRequirementId\(\), _\s+Trim\$\(mTxtConnectionQty.Text\), Trim\$\(mTxtConnectionPercent.Text\), ComboText\(mCmbConnectionUom\)\)' -and
        $form -match 'For column = 0 To 6: \.List\(idx, column\) = fields\(column\): Next column' -and
        $form -match 'ToRequirementId' },
    [pscustomobject]@{ Name = "PublicAction.ExercisesOutputFlow"; Pass =
        $form -match '\|RecipeOutputFirstRouting=' -and
        $form -match '\|RecipeCompatibleTargetsOnly=' -and
        $form -match '\|RecipeRequirementInternallyBound=' -and
        $form -match '\|RecipeNoIngredientDropdown=' -and
        $form -match '\|RecipeForkConvergenceVisible=' -and
        $form -match '\|RecipeTerminalOutputVisible=' -and
        $form -match '\|RecipeStagesDerived=' },
    [pscustomobject]@{ Name = "Packaged.RequiresOutputFlow"; Pass =
        $validator -match 'RecipeOutputFirstRouting=True' -and
        $validator -match 'RecipeCompatibleTargetsOnly=True' -and
        $validator -match 'RecipeRequirementInternallyBound=True' -and
        $validator -match 'RecipeNoIngredientDropdown=True' -and
        $validator -match 'RecipeForkConvergenceVisible=True' -and
        $validator -match 'RecipeTerminalOutputVisible=True' -and
        $validator -match 'RecipeStagesDerived=True' }
)

$failed = @($checks | Where-Object { -not $_.Pass })
foreach ($check in $checks) {
    "{0} {1}" -f $(if ($check.Pass) { "PASS" } else { "FAIL" }), $check.Name
}

"PLAN022_SLICE4AQ_SOURCE passed=$($checks.Count - $failed.Count) red=$($failed.Count) total=$($checks.Count)"
if ($failed.Count -gt 0) {
    throw "Plan 022 Slice 4aq source contract RED: $($failed.Name -join ', ')"
}
