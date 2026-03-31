Implement WAN support in these slices, in order.

| Slice | Primary purpose | Main areas touched |
|---|---|---|
| 1 | Freeze WAN contract in code-facing terms | docs, TODO map, test map |
| 2 | Config and path contract | `src/Core` |
| 3 | Warehouse publish engine | `src/Core`, `src/Admin` |
| 4 | Operator read-model freshness | `src/Core`, role workbooks |
| 5 | HQ aggregation hardening | HQ workbook modules, `src/Admin` |
| 6 | Recovery and interruption handling | `src/Core`, tests |
| 7 | Scheduler and operational wiring | scripts, admin commands, task setup docs |
| 8 | LAN + WAN proving | tests, evidence docs, smoke harnesses |