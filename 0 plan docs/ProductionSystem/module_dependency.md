# Production System - Module Dependencies (Draft)

| Module | Depends on | Purpose |
| --- | --- | --- |
| modPR_Production | modPR_Builders, cDynItemSearch, modInvMan, modGlobals | Core workflow |
| modPR_Builders | (none) | Recipe + acceptable item builders |
| cDynItemSearch | modTS_Received.LoadItemList | Filtered item picker |
| modInvMan | invSys tables | Apply/log inventory deltas |
| modGlobals | shared helpers | Keybindings and shared utils |
