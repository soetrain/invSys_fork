# Production System - VBA Perspective Diagrams (Draft)

These diagrams describe *how the VBA behaves* (events, table generation, layout handling).
They are not user-facing workflows.

## 1) Production Run - Event & Generation Flow (VBA view)

```mermaid
flowchart TD
    classDef evt fill:#5b6b9a,stroke:#2f3b5a,color:#ffffff;
    classDef op fill:#2c7a9b,stroke:#195a73,color:#ffffff;
    classDef data fill:#8c6239,stroke:#4f341f,color:#ffffff;
    classDef note fill:#fff2cc,stroke:#b99a33,color:#000000;

    EvtSelect["Worksheet_SelectionChange\n(Production)"]:::evt
    EvtChange["Worksheet_Change\n(Production)"]:::evt
    EvtBeforeDbl["Worksheet_BeforeDoubleClick\n(Production)"]:::evt

    PickRecipe["Open Recipe picker\n(Recipe chooser cell)"]:::op
    LoadRecipe["Load recipe by RECIPE_ID"]:::op
    BuildProcTables["Generate process tables\nper PROCESS (RecipeChooser_generated)"]:::op
    ApplyTemplates["Apply TemplatesTable formulas\nby scope + process"]:::op
    ProcCheckboxes["Render Process selector checkboxes\n(adjacent to process tables)"]:::op

    BuildPaletteTables["Generate Inventory Palette tables\nper selected PROCESS"]:::op
    PalettePickers["Wire pickers:\n- IngredientPalette filter\n- invSys filter (+category)"]:::op

    OutputTables["Generate ProductionOutput table\nper PROCESS"]:::op
    BatchCodeCheckbox["Render Recall checkbox\n(per output row)"]:::op

    BatchLog["BatchCodesLog entry\n(on Send to MADE)"]:::data
    ProdLog["ProductionLog entry\n(on To USED / To MADE / To TOTAL INV)"]:::data

    EvtSelect --> PickRecipe --> LoadRecipe --> BuildProcTables --> ApplyTemplates
    ApplyTemplates --> ProcCheckboxes --> BuildPaletteTables --> PalettePickers
    BuildProcTables --> OutputTables --> BatchCodeCheckbox

    EvtBeforeDbl --> PickRecipe
    EvtChange --> BuildPaletteTables
    EvtChange --> OutputTables

    Note1["All PROCESS tables are generated;\ncheckboxes decide which palette tables appear."]:::note
    ProcCheckboxes -.-> Note1

    Note2["Batch code generated per PROCESS/BATCH\nwhen outputs are sent to MADE."]:::note
    BatchCodeCheckbox -.-> Note2

    ToMade["Send to MADE"]:::op
    ToUsed["To USED"]:::op
    ToTotal["Send to TOTAL INV"]:::op

    OutputTables --> ToMade --> BatchLog --> ProdLog
    BuildPaletteTables --> ToUsed --> ProdLog
    ToTotal --> ProdLog

    subgraph Legend
        L1["Event / Worksheet hook"]:::evt
        L2["Operation / VBA routine"]:::op
        L3["Data / Log table"]:::data
        L4["Note"]:::note
    end
```

## 2) Inventory Palette Auto-Expand (VBA layout detail)

```mermaid
flowchart TD
    classDef evt fill:#5b6b9a,stroke:#2f3b5a,color:#ffffff;
    classDef op fill:#2c7a9b,stroke:#195a73,color:#ffffff;
    classDef note fill:#fff2cc,stroke:#b99a33,color:#000000;

    ChangeEvt["Worksheet_Change\n(InventoryPalette table)"]:::evt
    DetectRowAdd["Detect new ListRow\nin InventoryPalette table"]:::op
    GetBand["Resolve process band\n(table anchor + bounds)"]:::op
    CalcShift["Calculate shift rows\n(only below in same band)"]:::op
    ShiftDown["Shift tables below down\n(no left/right edits)"]:::op
    ResizeTable["Resize table range\n+ re-anchor controls"]:::op

    ChangeEvt --> DetectRowAdd --> GetBand --> CalcShift --> ShiftDown --> ResizeTable

    Note1["Band = process block (recipe + palette + outputs)\nanchored to a top-left cell."]:::note
    GetBand -.-> Note1

    Note2["Only tables below *within the band* move;\nneighbor bands remain unchanged."]:::note
    ShiftDown -.-> Note2

    subgraph Legend
        L1["Event / Worksheet hook"]:::evt
        L2["Operation / VBA routine"]:::op
        L3["Note"]:::note
    end
```

## 3) Recipe Builder & Ingredient Palette Builder (VBA view)

```mermaid
flowchart TD
    classDef evt fill:#5b6b9a,stroke:#2f3b5a,color:#ffffff;
    classDef op fill:#2c7a9b,stroke:#195a73,color:#ffffff;
    classDef data fill:#8c6239,stroke:#4f341f,color:#ffffff;
    classDef note fill:#fff2cc,stroke:#b99a33,color:#000000;

    ToggleRB["Toggle Recipe Builder\n(show/hide tables)"]:::op
    LoadRB["Load Recipe for edit\n(populate builder tables)"]:::op
    SaveRB["Save Recipe\n-> Recipes sheet"]:::op
    RegisterTpl["Register Templates\n(scan formula columns)"]:::op
    Templates["TemplatesTable"]:::data

    ToggleIP["Toggle IngredientPalette Builder"]:::op
    PickRecipeIP["Recipe picker (IP)"]:::op
    PickIngredientIP["Ingredient picker\n(from Recipes)"]:::op
    PickItemIP["Item picker\n(invSys + category filters)"]:::op
    SaveIP["Save IngredientPalette\n-> IngredientPalette sheet"]:::op

    ToggleRB --> LoadRB --> SaveRB --> RegisterTpl --> Templates
    ToggleIP --> PickRecipeIP --> PickIngredientIP --> PickItemIP --> SaveIP

    Note1["Builder tables persist and generate GUIDs\non save."]:::note
    SaveRB -.-> Note1

    subgraph Legend
        L1["Event / Worksheet hook"]:::evt
        L2["Operation / VBA routine"]:::op
        L3["Data / Log table"]:::data
        L4["Note"]:::note
    end
```
