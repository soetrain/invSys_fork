# Slice 4be-A operator comparison review

**Status:** the final 2026-11-16 candidate is open for a second human preview in disposable warehouse WHT7025AE. No second-preview result is recorded yet. This is an early usability review under Architecture v4.11 D18-REPLAY-01 and Plan 022, not 4be-A acceptance.

## Candidate and safe setup

- Use the five add-ins together from the final local review copy at `reports/runtime/4be-a-human-review-20261116b/addins/` in a dedicated Excel session. The matching inventory template is at `addins/templates/invSys.Data.Inventory.template.xlsb`. On 2026-11-16, all five XLAMs and the template matched PR17 commit `45d84ca8` byte for byte; the ignored local `candidate.json` records their hashes. The review host is a fresh copy of the prior saved host. Excel was closed during the copy and startup registration change. Recheck the candidate hashes and Excel ownership before any further build or test.
- Use the existing disposable NAS Training warehouse WHT7025AE and a disposable operator workbook. Do not use an operational warehouse or existing business workbook. The reviewer needs `RECEIVE_POST`, `ADMIN_MAINT` and `ACTION_PATH_MAINT` in that warehouse, either through one authorized account or separate authorized users. The current sign-in and warehouse must remain the same while a guide/run pair is open.
- Current bootstrap grants the initial user `ADMIN_MAINT` and `RECEIVE_POST`, but leaves that user's PIN hash empty and does not grant `ACTION_PATH_MAINT`. Normal sign-in rejects an empty hash; the current Users & Roles form cannot assign the guide right. Justin authorized private credential setup and a test-only guide-right grant in this disposable Training fixture. That enabled the first preview but does not qualify ordinary user-side onboarding. Credentials do not belong in project documents or reports.
- In Admin **Settings > Event Tracking**, verify that the reviewer's optional recording is enabled and the selected controls are eligible. Keep required audit enabled. Note the original personal **Preferred Action Path view** so it can be restored after review. Capture only redacted screenshots or comments; exclude credentials, paths, entered values and inventory details from shared evidence.
- Carry out both tasks through their ordinary controls. **Run How-To** is the separate executable replay path and is outside this 4be-A presentation review.

The first preview's stopped-run picker selection took over 10 seconds in a focused packaged RED. PR17's five-XLAM candidate passed the same real-handler sequence 17/17 with a 0.84-second first selection, both guide Save routes and Cancel; retained curation passed 70/70. On the final candidate, empty Shipping Holds publication passed 86/86 and an audited Release 1 chain passed 33/33 with no Excel Application Error/Hang; early Receiving replay passed 121/121 technically on the preceding picker build. The first preview's `VBE7.DLL` crash remains unassigned; these tests do not substitute for the human comparison below.

## Operations: Receiving task

1. Open Operations **Viewer > Events** and **Start Recording**. Use the ordinary Receiving form to **Add Selected** and **Confirm Writes**, then **Stop Recording**. In **Viewer > Action Paths**, select this saved source recording and **Create guide**. Give it a training name, plain-language instructions and an **Expected conclusion** whose terminal step is **Confirm Writes** with command completion. Save the guide. A confirmed command may conclude with **Domain application not asserted**. If you separately choose the stricter source-events-applied terminal, pending publication must remain **Awaiting** until exact application evidence appears.
2. Start a **separate** recording of the same Receiving task. Use Add Selected and Confirm Writes again, then Stop. Select this observed run in Action Paths. Open **Published guides**, select the exact guide version from step 1 and choose **Use for selected run**. Back in Action Paths, use **Evaluate** and **View guide and run**. Inspect **How-To**, **Diagnostic** and **Compare both** without changing the selected guide/version or observed run. Check that authored instructions, original control order, repetitions and the saved result have distinct provenance. A command-completion conclusion must not claim Domain application.
3. Start a third recording, use Add Selected, then Stop **without Confirm Writes**. Bind the same guide to this incomplete run and Evaluate. Compare all three views. The required terminal action should be missing and the run must not say **Conclusion observed**. Clear disposable staging afterward through the ordinary control.

## Admin: personal preference task

1. In the same Training warehouse, start a new Viewer recording. In Admin **Settings > Event Tracking > Action Paths**, change **Preferred Action Path view** and click **Save My Preference**. Stop recording. Select that source run in Action Paths, **Create guide**, author the steps and set **Save My Preference** with a completed terminal outcome, then save the guide. This is a personal preference action; it does not grant tracking rights.
2. Record a separate run of the same Admin task, choosing a different view and clicking **Save My Preference** so the save is a changed-value result. Stop, bind the exact guide to this observed run, Evaluate, and inspect How-To, Diagnostic and Compare both. Check the exact guide/run labels, the original observed action order and the saved result. If the save reports **Unchanged** or **Unavailable**, record that outcome instead of treating it as the expected changed save.
3. Record an incomplete Admin run that stages a different preferred view but stops **before Save My Preference**. Bind the same guide and Evaluate. The missing save must remain visible and must not conclude. Restore the original preference through Settings after the comparison.

## Review record

For each row, record **Pass**, **Issue**, or **Not tried**, with a short redacted note. Record a screenshot only when it helps explain a finding. This record is human feedback; the automated gates remain separate.

| Observation | Operations complete | Operations incomplete | Admin complete | Admin incomplete |
|---|---|---|---|---|
| Guide version and observed run remain the exact selected pair |  |  |  |  |
| How-To instructions are readable and clearly authored |  |  |  |  |
| Diagnostic shows original actions and truthful result |  |  |  |  |
| Compare both labels applicable matches, missing steps and unavailable evidence |  |  |  |  |
| Switching views, scrolling and resizing retain the pair and readable text |  |  |  |  |
| No incorrect conclusion or implied replay of controls |  |  |  |  |

Also note whether **Viewer > Events > Event Detail** is readable for the related published event, whether reopening the saved workbook retains the intended warehouse binding, and whether repeated Operations launcher entry reuses its existing form. Report any disabled control, clipped text, surprising outcome, modal or changed selection with the step at which it occurred. Do not infer acceptance from an untried row.

**After the review:** capture Justin's verdict and issues separately. Completion still requires the full five-package build/compile/initialization, layout/header, static maintenance, retained current GREEN regressions, live-role/full-chain and saved-workbook restart/binding gates named in Architecture v4.11. Technical Receiving B0 replay passed 121/121 on this candidate, while its human acceptance marker remains open.
