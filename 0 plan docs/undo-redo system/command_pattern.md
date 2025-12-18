Great research. Let me give you a comprehensive breakdown of the most reliable approaches for implementing undo-redo in VBA Excel, which I'll organize by complexity and reliability.

## The Core Challenge

When a VBA macro writes anything to the worksheet, **Excel clears its entire undo stack**. This is a fundamental VBA limitation—there's no way around it. So you need a **custom undo-redo system independent of Excel's native stack**.

***

## 1. **Command Pattern + State Stack** (Most Reliable for Complex Operations)

This is the industry-standard approach used by most professional applications. It's ideal for your inventory system where you need to track multiple interdependent changes.

### Architecture

```vba
'ICommand Interface (pseudo-interface using naming conventions)
'File: ICommand.cls

Option Explicit

Public Function CanExecute() As Boolean
    'Override in concrete command
End Function

Public Sub Execute()
    'Override in concrete command
End Sub

Public Sub Undo()
    'Override in concrete command
End Sub

Public Sub Redo()
    'Override in concrete command
End Sub

Public Function Description() As String
    'Override in concrete command
End Function
```

```vba
'CommandManager - manages undo/redo stacks
'File: CommandManager.cls

Option Explicit

Private undoStack As Collection
Private redoStack As Collection

Public Sub New()
    Set undoStack = New Collection
    Set redoStack = New Collection
End Sub

Public Sub ExecuteCommand(cmd As Object)
    'Execute command and add to undo stack
    If cmd.CanExecute Then
        cmd.Execute
        undoStack.Add cmd
        ClearRedoStack  'Clear redo when new command executed
    End If
End Sub

Public Sub Undo()
    If undoStack.Count > 0 Then
        Dim cmd As Object
        Set cmd = undoStack(undoStack.Count)
        cmd.Undo
        redoStack.Add cmd
        undoStack.Remove undoStack.Count
    Else
        MsgBox "Nothing to undo"
    End If
End Sub

Public Sub Redo()
    If redoStack.Count > 0 Then
        Dim cmd As Object
        Set cmd = redoStack(redoStack.Count)
        cmd.Redo
        undoStack.Add cmd
        redoStack.Remove redoStack.Count
    Else
        MsgBox "Nothing to redo"
    End If
End Sub

Private Sub ClearRedoStack()
    Set redoStack = New Collection
End Sub

Public Property Get CanUndo() As Boolean
    CanUndo = (undoStack.Count > 0)
End Property

Public Property Get CanRedo() As Boolean
    CanRedo = (redoStack.Count > 0)
End Property

Public Function GetUndoDescription() As String
    If undoStack.Count > 0 Then
        Dim cmd As Object
        Set cmd = undoStack(undoStack.Count)
        GetUndoDescription = "Undo: " & cmd.Description
    End If
End Function

Public Function GetRedoDescription() As String
    If redoStack.Count > 0 Then
        Dim cmd As Object
        Set cmd = redoStack(redoStack.Count)
        GetRedoDescription = "Redo: " & cmd.Description
    End If
End Function
```

### Concrete Command Example (for your inventory)

```vba
'UpdateInventoryCommand.cls - Concrete command implementation

Option Explicit
Implements ICommand

Private before_Quantity As Long
Private before_Location As String
Private after_Quantity As Long
Private after_Location As String
Private itemID As String
Private ws As Worksheet

'Constructor-like initialization
Public Sub Initialize(id As String, oldQty As Long, oldLoc As String, _
                     newQty As Long, newLoc As String, sheet As Worksheet)
    itemID = id
    before_Quantity = oldQty
    before_Location = oldLoc
    after_Quantity = newQty
    after_Location = newLoc
    Set ws = sheet
End Sub

Public Sub Execute()
    'Update inventory to new values
    Call UpdateInventoryItem(itemID, after_Quantity, after_Location, ws)
End Sub

Public Sub Undo()
    'Revert to old values
    Call UpdateInventoryItem(itemID, before_Quantity, before_Location, ws)
End Sub

Public Sub Redo()
    'Same as Execute
    Call Execute
End Sub

Public Function Description() As String
    Description = "Update item " & itemID & " from " & before_Quantity & " to " & after_Quantity
End Function

Public Function CanExecute() As Boolean
    CanExecute = True 'Add validation logic if needed
End Function

Private Sub UpdateInventoryItem(id As String, qty As Long, loc As String, sheet As Worksheet)
    'Actual inventory update logic
    Dim row As Long
    row = FindItemRow(id, sheet)
    If row > 0 Then
        sheet.Cells(row, 3).Value = qty  'Quantity column
        sheet.Cells(row, 4).Value = loc  'Location column
    End If
End Sub

Private Function FindItemRow(id As String, sheet As Worksheet) As Long
    Dim i As Long
    For i = 2 To sheet.UsedRange.Rows.Count
        If sheet.Cells(i, 2).Value = id Then
            FindItemRow = i
            Exit Function
        End If
    Next
End Function
```

### Usage in Your Inventory Module

```vba
'Global command manager (in your main module or workbook_open)
Public cmdManager As CommandManager

Sub InitializeUndoRedo()
    Set cmdManager = New CommandManager
End Sub

Sub UpdateInventoryWithUndo(itemID As String, newQty As Long, newLoc As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Inventory")
    
    'Get current values
    Dim oldQty As Long, oldLoc As String
    oldQty = GetCurrentQuantity(itemID, ws)
    oldLoc = GetCurrentLocation(itemID, ws)
    
    'Create and execute command
    Dim cmd As UpdateInventoryCommand
    Set cmd = New UpdateInventoryCommand
    cmd.Initialize itemID, oldQty, oldLoc, newQty, newLoc, ws
    
    cmdManager.ExecuteCommand cmd
End Sub

Sub OnUndoButtonClick()
    cmdManager.Undo
End Sub

Sub OnRedoButtonClick()
    cmdManager.Redo
End Sub
```

***

## 2. **Memento Pattern + Stack** (For State-Heavy Operations)

Better when you need to capture entire object states, not individual changes.

```vba
'Memento.cls - Captures entire state snapshot
Option Explicit

Private state As Object  'Dictionary with all values

Public Sub SaveState(stateDict As Object)
    Set state = stateDict
End Sub

Public Function RestoreState() As Object
    Set RestoreState = state
End Function
```

### Advantages vs Command:
- ✅ Simpler for capturing complex multi-field states
- ❌ Uses more memory (stores full state, not just deltas)
- ✅ Easier to debug (can inspect entire state snapshots)

***

## 3. **Snapshot Sheet Method** (Simplest but Memory-Intensive)

Save entire worksheet states to hidden sheets:

```vba
Sub SaveSnapshot()
    Dim newSheet As Worksheet
    Set newSheet = Sheets.Add(After:=Sheets(Sheets.Count))
    newSheet.Name = "Snapshot_" & Format(Now(), "yyyymmdd_hhmmss")
    newSheet.Visible = xlSheetHidden
    
    'Copy current data
    Sheets("Inventory").Cells.Copy
    newSheet.Cells.PasteSpecial xlPasteAll
    
    'Store in collection or array
    undoStack.Add newSheet.Name
End Sub

Sub UndoSnapshot()
    If undoStack.Count > 0 Then
        Dim snapshotName As String
        snapshotName = undoStack(undoStack.Count)
        Sheets(snapshotName).Visible = xlSheetVisible
        Sheets(snapshotName).Cells.Copy
        Sheets("Inventory").Cells.PasteSpecial xlPasteAll
        Application.CutCopyMode = False
    End If
End Sub
```

### When to use:
- ✅ Simple operations, small datasets
- ✅ Quick prototyping
- ❌ Memory-intensive for large inventories
- ❌ Slow with complex workbooks

***

## 4. **Application.OnUndo** (Limited but Native)

For single-step operations, hook into Excel's native undo:

```vba
Sub ChangeInventoryItem()
    Dim itemID As String, newQty As Long
    itemID = InputBox("Item ID")
    newQty = InputBox("New Quantity")
    
    'Store original value
    Dim originalQty As Long
    originalQty = GetItemQuantity(itemID)
    
    'Update value
    UpdateInventoryItem itemID, newQty
    
    'Register undo
    Application.OnUndo "Undo inventory update", "UndoInventoryChange"
    'Note: UndoInventoryChange procedure needs parameter storage
End Sub
```

### Limitations:
- ❌ Only handles ONE level of undo (not a stack)
- ❌ Can't access previous parameters directly
- ⚠️ Unreliable for complex multi-step operations
- ✅ Integrates with Excel's UI

***

## **My Recommendation for Your Inventory System**

**Use Command Pattern (#1) because:**

1. **Scalability** - Handle unlimited undo/redo levels
2. **Granularity** - Track individual inventory changes precisely
3. **Composability** - Combine multiple commands (move 5 items from one location to another)
4. **Debuggability** - Each command logs exactly what changed and why
5. **Memory efficient** - Store only deltas, not full states
6. **Professional** - Standard architecture used in production systems

### Quick Implementation Path

```vba
'Create in your workbook startup:
Sub Workbook_Open()
    Set cmdManager = New CommandManager
    Call UpdateUIForUndoRedoState
End Sub

Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    'Optional: Auto-update UI buttons when undo state changes
    If Not Intersect(Target, Sh.Range("A1")) Is Nothing Then
        Call UpdateUIForUndoRedoState
    End If
End Sub

Sub UpdateUIForUndoRedoState()
    'Update button enabling/tooltips based on cmdManager states
    ThisWorkbook.Sheets("UI").CommandButton_Undo.Enabled = cmdManager.CanUndo
    ThisWorkbook.Sheets("UI").CommandButton_Redo.Enabled = cmdManager.CanRedo
End Sub
```

**Want me to create a complete, ready-to-use VBA undo-redo system module tailored specifically for your inventory project?** I can build out the full Command Pattern implementation with example commands for typical inventory operations (add item, update quantity, relocate item, delete item, etc.).

[1](https://wellsr.com/vba/2019/excel/how-to-undo-a-macro-with-vba-onundo-and-onrepeat/)
[2](https://rubberduckvba.blog/2020/11/19/from-macros-to-objects-the-command-pattern/)
[3](https://jkp-ads.com/articles/undowithvba03.aspx)
[4](https://www.msofficeforums.com/excel-programming/21312-enabling-undo-function-macro-enabled-workbook.html)
[5](https://www.imcoded.com/docs/c++/design-patterns/command-pattern)
[6](https://github.com/OfficeDev/Excel-Custom-Functions/issues/346)
[7](https://devblogs.microsoft.com/vbteam/implementing-infinite-undoredo-matt-gertz/)
[8](https://stackoverflow.com/questions/18762350/how-to-implement-state-in-command-pattern)
[9](https://www.reddit.com/r/vba/comments/1nm6k3e/vba_any_hacks_to_preserve_undo_stack_by/)
[10](https://dev.to/isaachagoel/you-dont-know-undoredo-4hol)
[11](https://learn.microsoft.com/en-us/shows/visual-studio-toolbox/design-patterns-commandmemento)
[12](https://www.excelforum.com/excel-programming-vba-macros/1409649-undo-function-not-working-with-ctrl-z-after-i-run-the-macro.html)
[13](https://rubberduckvba.blog/2025/05/31/undoing-and-redoing-stuff/)
[14](http://exceldevelopmentplatform.blogspot.com/2016/09/use-raii-design-pattern-to-tidy-your.html)
[15](https://www.youtube.com/watch?v=_cLY-_qVKXc)
[16](https://stackoverflow.com/questions/7004754/how-to-programmatically-code-an-undo-function-in-excel-vba)
[17](https://www.youtube.com/watch?v=mSZuEbAkJCo)
[18](https://stackoverflow.com/questions/19973394/run-vba-code-and-keep-excels-undo-redo-intact)
[19](https://www.reddit.com/r/programming/comments/1lidz3y/an_indepth_look_at_the_implementation_of_an/)
[20](https://cloudaffle.com/series/behavioral-design-patterns/command-pattern-application/)