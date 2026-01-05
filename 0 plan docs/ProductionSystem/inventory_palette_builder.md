# Inventory Palette Builder (Production)

This doc captures System 2: Inventory Palette Builder (per recipe item, acceptable item list builder). It links recipe ingredients to real inventory items managed in invSys.

## What exists now
- `IP_ChooseRecipe` (picker table for selecting recipe)
- `IP_ChooseIngredient` (picker table for selecting ingredient from chosen recipe)
- `IP_ChooseItem` (picker table for selecting acceptable inventory items)
- `IngredientPalette` table headers: RECIPE_ID, INGREDIENT_ID, INPUT/OUTPUT, ITEM, PERCENT, UOM, AMOUNT, ROW, GUID
- Buttons:
  - `Save IngredientPalette` (save changes to IngredientPalette)
  - `Clear Inventory Palette Builder` (clear System 2 workspace)

## Builder flow (pickers + acceptable items)

```mermaid
flowchart TD
    classDef btn fill:#2f4e9c,stroke:#1f2f5c,color:#ffffff;
    classDef list fill:#2c7a9b,stroke:#195a73,color:#ffffff;
    classDef data fill:#8c6239,stroke:#4f341f,color:#ffffff;
    classDef note fill:#fff2cc,stroke:#b99a33,color:#000000;

    PickRecipe["Recipe picker\n(IP_ChooseRecipe)"]:::list
    PickIngredient["Ingredient picker\n(IP_ChooseIngredient)"]:::list
    PickItem["Inventory picker\n(IP_ChooseItem)"]:::list

    Recipes["Recipes table\n(Recipe rows)"]:::data
    Inventory["invSys inventory\n(real items)"]:::data
    Palette["Inventory Palette\n(acceptable items per ingredient)"]:::data

    PickRecipe --> PickIngredient --> PickItem
    Recipes --> PickRecipe
    Recipes --> PickIngredient
    Inventory --> PickItem
    PickItem --> Palette

    Note1["User selects recipe, then selects an ingredient\nfrom that recipe, then assigns acceptable inventory."]:::note
    PickRecipe -.-> Note1
```

## Map legend
- Blue nodes = picker tables on Production sheet.
- Brown nodes = source/target data tables.
- Yellow notes = user guidance.

## Data mapping (per ingredient)

```mermaid
flowchart LR
    classDef list fill:#2c7a9b,stroke:#195a73,color:#ffffff;
    classDef data fill:#8c6239,stroke:#4f341f,color:#ffffff;
    classDef note fill:#fff2cc,stroke:#b99a33,color:#000000;

    RecipeLine["Recipe ingredient\n(RECIPE_ID + INGREDIENT_ID)"]:::data
    IPChooseItem["IP_ChooseItem rows\n(acceptable inventory)"]:::list
    Output["Inventory Palette\n(per recipe ingredient)"]:::data

    RecipeLine --> IPChooseItem --> Output

    Note1["Many inventory items can map to one ingredient.\nThis builds an acceptable list per ingredient."]:::note
    Output -.-> Note1
```

## Notes / conventions
- System 2 is the Inventory Palette Builder in Production.
- Flow: pick recipe → pick ingredient → pick inventory items.
- Each ingredient can have multiple acceptable inventory items.
- Pickers are driven by existing data in Recipes + inventory tables.
