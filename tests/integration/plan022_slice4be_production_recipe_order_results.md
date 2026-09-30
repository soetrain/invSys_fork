# Slice 4be Recipe ordering observations

Last verified:2026-09-30 UTC. **Focused RED established; implementation pending.**
Architecture v4.11 D18, Plan022, controls1.281 and coverage1.24 specify the three
Recipe Designer ordering observations under approved semantic inheritance, committed
as docs61f1c79 before runtime changes. This preserves the existing local ordering
semantics and introduces no authority or permission amendment.

`Test-Slice4beProductionRecipeOrder.ps1` invokes the actual packaged Move Up,
Move Down and Auto Order Click handlers through disposable typed adapters. On
the frozen catalog16 component candidate, **146 PASS /317 expected FAIL** gives
463 unique check identities, with all five instrumented package compiles passing.
Existing successful/bounds/empty/ordered/cyclic/self-edge/unresolved-endpoint/case-
insensitive ordering examples pass. Failures are confined to catalog17 registration,
new observations/terminal facts, context/permission/loading/nested guards, visible
tracking refusal and owned handling of a synthetic partial failure. No unrelated
ordering failure or harness exception is counted as RED.

The gate protects identity-associated fields and connection values, execution
renumbering, the existing instruction-ordinal side effect, selection/choice refresh,
and partial local changes before rejection/failure. It also covers disabled and
unavailable optional tracking, current actor/warehouse context, redaction/integrity,
saved authority, unknown workbook columns and immutable earlier activity.
New controls can only conclude local CommandCompleted through STAGED; no Domain
application is asserted.

- Controller: `reports/runtime/production-recipe-order-controller/cab9ea7c691b452a9f43161f6982a1b6`.
- Result: `reports/runtime/slice4be-production-recipe-order/91ec2bd2a5ad4ac1ae76d36375ca8d42/red.json`.
- Verification: `reports/runtime/production-recipe-order-red-verification.json`.
- Run interval:02:03:16--02:05:40 UTC.
- Settings and all five package bytes are preserved; Excel exits without an
  explicit termination path, and zero Excel Application1000/1001/1002 failures occur.
- Runtime source still matches115b616. No desktop error5 occurs during this gate.

The first adapter installation targeted a nonexistent `mLstRecipeNodes_Click`;
VBE procedure lookup returned "Sub or Function not defined" before instrumented
compilation. That1 PASS/1 harness failure is **not product RED**. The nested-call
hook now enters the existing `RenumberRecipeExecutionOrder` routine and invokes
the actual handler while its outer action is active. Initial controller
`production-recipe-order-controller/3daa314ba1394863adc1572d58892382` and result
`slice4be-production-recipe-order/1b4220800f1d4fc6adacd73244f0c977` are under runtime
reports; settings/packages restore and no Excel host remains.

Next: implement against these463 identities, then prove packaged GREEN,
recording/publication and both Action Path methods, compile/layout/static, retained
regressions/full chain/reusable behavior and visible evidence. Registration remains
29/68; Slice4be and Release1 acceptance remain incomplete.
