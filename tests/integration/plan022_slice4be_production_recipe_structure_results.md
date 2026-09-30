# Slice 4be Recipe structure observations

Last verified:2026-09-30 UTC. **Baseline RED and released-data Update RED reproduced;
connection-write correction approved. No runtime implementation or GREEN yet.**

Architecture v4.11 D18, Plan022, controls1.288 and coverage1.31 specified five
existing Recipe structure observations in docs commit da7bab7 before runtime
edits. Runtime remains cd2edc0, frozen catalog17 candidate
`deploy/validation-production-recipe-order-final`; registration remains32/68.
The current controls1.290 and normative decision explicitly record the discovered
Update conflict instead of treating a changed write algorithm as instrumentation.
The user approved the correction on2026-09-30 UTC; docs82bd41a records approval in
Architecture/Plan/controls/coverage before any runtime edit.

Supplemental released-data RED: **52 PASS/3 FAIL**,55 unique checks. The fixture
reuses the existing packaged reusable setup, including Process Save/Release and
compatible alternatives; both Recipe nodes are backed by released Domain records.
Changing the loaded connection to quantity4 and percentage75 and invoking the
actual Update handler fails the all-fields/quantity/percentage checks. Routing,
UOM, nodes, instructions, saved authority and unknown columns remain intact.
An unchanged-value Update passes. The older reusable GREEN clicked Update without
editing the loaded values and therefore did not protect this changed-value case.

Controller `production-recipe-structure-controller/347450280f194bb1b70b03fa08c7ace6`
and result `slice4be-production-recipe-structure-released/ffd8e452ce7143e286487ce1c1b11f8c`
record04:28:57.5563073--04:30:41.1055581 UTC, five instrumented compiles,
settings/package preservation, normal unassisted closure and zero Excel
Application1000/1001/1002 events. The principal capture shows the two released
Processes, intact routing, the old required quantity and staged-success status.
It is reviewed diagnostic evidence, not acceptance. The first supplemental run
had the same three behavioral failures plus a missing screenshot-adapter error;
that harness error is corrected and is not product RED.

The protecting gate enters the actual packaged Add Process, Remove Process,
Connect, Update and Disconnect Click handlers. Disposable adapters supply local
drafts and failure/nested-entry hooks; no runtime test backdoor was introduced.
The final baseline run has **182 PASS /604 FAIL**,786 unique checks:

- 599 failures demonstrate absent catalog18 metadata/observations, terminal facts,
  captured-context/current-capability/loading/nested guards, owned failure handling
  and visible non-blocking tracking failure behavior.
- Five failures concern Update's existing behavior: seven-field replacement,
  positive-percentage edits with blank/nonnumeric/negative other quantity, and an
  unchanged-value Update. They protect the approved correction, not silently
  classified as missing tracking or accepted preservation.

Controller `production-recipe-structure-controller/c710125a2a824300aed3deef1b8e8000`
and result `slice4be-production-recipe-structure/353ed89d061f48af90fd7b896927d961`
are under ignored `reports/runtime`. Controller `verification.json` records
04:17:13.3177374--04:19:40.6576169 UTC, all five instrumented compiles, preserved
settings and all five package hashes, normal unassisted closure and zero Excel
Application1000/1001/1002 events through the following12seconds. Runtime source
still matches cd2edc0. Saved fixture authority, workbook bytes, unknown columns
and existing activity records remain unchanged. No visible acceptance is claimed.

Fixed diagnostic booleans confirm all seven intended editor values immediately
before Update. One hidden connection-list Click occurs during the action. In the
result, output/target/requirement fields are empty and quantity/percentage retain
the old values; source and UOM match. The existing hidden-list Click calls
`LoadConnectionEditorFromIndex`, which reloads the same editor the ongoing writer
reads for later fields. This supports the explicit proposal to snapshot validated
inputs once before writing, preserving validation and separate Save authority.
No entered data or source paths are emitted by the diagnostic comparisons.

The initial184/602 run used a faulty selection fixture: selecting the hidden list
ran a callback that cleared the fixture loading flag, allowing the following
visible selection to change its target. Restoring suppression after that existing
callback corrects the fixture; that initial result is not the accepted RED.
The corrected preliminary and final diagnostic runs retain182/604. Do not weaken
the five Update assertions to conceal this discovered conflict.

Next: implement the explicitly approved connection-write stability and structure
observations with the existing786 checks plus55 released-data checks protecting
RED/GREEN. Build a new isolated candidate;
preserve the catalog17 baseline. Require recording/publication/How-To/Diagnostic/
Compare, compile, layout, static limits, live roles, full Release1 chain, reusable
Production and visible evidence before completing this bounded group.
