# Confirm Writes Tester Guide

Use this guide to prove the WH1 receiving flow on a new tester machine. You do not need to know the repo or any internal setup details before starting.

## What you need before you start

- Microsoft 365 Excel on Windows, with macros enabled when prompted
- Access to the SharePoint library location that contains `Addins/`
- About 10 minutes

You will use these two add-ins:

- `Addins/invSys.Admin.xlam`
- `Addins/invSys.Receiving.xlam`

This guide uses the standard proving scenario:

- Warehouse: `WH1`
- Station: `R1`
- SKU: `TEST-SKU-001`
- Starting Qty On Hand: `100`
- Test entry quantity: `10`
- Expected Qty On Hand after processing: `110`

## Step-by-step setup

1. Open the SharePoint folder that contains the `Addins/` library.
2. Copy `invSys.Admin.xlam` and `invSys.Receiving.xlam` from `Addins/` to a local folder such as `C:\invSys\Addins\`.
3. Open `invSys.Admin.xlam` in Excel and enable macros if Excel prompts you.
4. In the Excel ribbon, click `Setup Tester Station`.
5. Enter your `UserId` and enter a PIN of your choosing. You will use this PIN to log in later.
6. Leave `Warehouse = WH1` and `Station = R1` at their default values.
7. Click `Setup`, then wait until you see `Setup complete`.
8. Click `Open Receiving Workbook`.

## Running Confirm Writes

9. In the receiving workbook, click `Refresh Inventory`, then confirm that `TEST-SKU-001` shows `Qty On Hand = 100`.
10. In the first empty row, enter `SKU = TEST-SKU-001` and `Qty = 10`.
11. Click `Confirm Writes`.
12. Confirm that the row turns green and the row status shows `CONFIRMED`.
13. Wait for the processor to run, up to 60 seconds, then click `Refresh Inventory` again.
14. Confirm that `TEST-SKU-001` now shows `Qty On Hand = 110`.

## What success looks like

- The row status is `CONFIRMED`, not `PENDING` and not `ERROR`
- `Qty On Hand` updates from `100` to `110` after the refresh
- No red status panel is visible at the top of the receiving workbook

## If something goes wrong

- Red panel at the top: follow the message shown in the panel
- `MISSING_CAPABILITY`: contact the admin who set up `WH1`
- `STALE` snapshot: click `Refresh Inventory` and wait for the snapshot to update
- Setup fails: run `Setup Tester Station` again; it is safe to re-run
- Anything else: send `C:\invSys\WH1\config\invSys.log` to the developer
