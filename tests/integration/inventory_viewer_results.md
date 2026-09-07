# Inventory Viewer Packaged Results

- Status: **PASS**
- Runtime: isolated generated test warehouse
- ConfigLoaded: True
- AuthLoaded: True
- TargetSelected: True
- TargetPathsSet: True
- SignedIn: True
- SnapshotCreated: True
- FirstActionRows: 3
- RepeatedLaunchReusedGeneration: True
- FilterVisibleRows: 1
- ListBoxTableExport: True
- EventsVisibleRows: 6
- RefreshedEventsVisibleRows: 7
- NewestPublishedReference: BOL-VIEWER-NEW
- ReadableEventDates: True
- ViewerTabCount: 3
- ViewerTabCaptions: Inventory,Events,ListBox->Table
- SelectedViewerTab: Events
- RemoveEventsVisible: True
- InternalReservationHidden: True
- ProductionInputEventsVisible: True
- ProductionOutputEventsVisible: True
- EventsReadOnly: True
- RollingDateFilters: True
- RememberedRangeAfterReopen: True
- InvalidRememberedRangeFallsBackToAll: True
- SnapshotHashUnchanged: True
- NewPublicationChangedSnapshot: True

## Observed result

The public Operations Viewer action exported the displayed ListBox to a new unsaved worksheet table, loaded readable Receipt, Production input/output, and Shipping Remove events, excluded the internal SHIP_RESERVE fixture from the operator-action log, refreshed the already-open Events page to show a newly published receipt first, applied All/Day/Week/Month/custom rolling-day filters, restored custom 14 days after form close/reopen, kept Events read-only, and left the new snapshot byte-for-byte unchanged.
