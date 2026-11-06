VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "cInventoryEditRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Core-only submitted input. Never pass this object across an XLAM boundary.
' Read the captured row itself, not an Admin-supplied Dictionary.
Private mValues As Object, mCustom As Object, mSource As Variant
Private mQuantity As Double, mHasQuantity As Boolean, mHash As String
Private mResult As String, mImported As Boolean

Public Function Read(ByVal context As String, ByVal row As Long, ByRef report As String) As Boolean
    Dim book As Workbook, table As ListObject, column As ListColumn
    Dim names As Object, inputs As Object, key As String, text As String, name As Variant
    Dim mode As String, configured As Boolean, packed As Variant, part As Variant
    On Error GoTo Invalid
    mQuantity = 0: mHasQuantity = False: mHash = "": mResult = "": mImported = False
    mSource = modInventoryCommands.SourceValues(context)
    If Not IsArray(mSource) Then GoTo Invalid
    text = modInventoryCommands.OperatorWorkbookName(context)
    If text = "" Then GoTo Invalid
    Set book = Application.Workbooks(text)
    Set table = modExecutionTarget.Table(book, modInventoryCommands.StagingTableName(context))
    If table Is Nothing Then GoTo Invalid
    If row < 1 Or row > table.ListRows.Count Then GoTo Invalid
    Set names = CreateObject("Scripting.Dictionary")
    Set mValues = CreateObject("Scripting.Dictionary")
    Set mCustom = CreateObject("Scripting.Dictionary"): mCustom.CompareMode = vbTextCompare
    Set inputs = CreateObject("Scripting.Dictionary"): inputs.CompareMode = vbBinaryCompare
    For Each column In table.ListColumns
        key = modInventoryEditWire.WorksheetHeader(column.Name)
        If key = "" Or key = "ROW" Or key = "SYSTEMKEY" Or key = "CONDITION" Then GoTo Invalid
        If names.Exists(key) Then GoTo Invalid
        names.Add key, True
        text = Trim$(CStr(table.DataBodyRange.Cells(row, column.Index).Value2))
        mValues.Add key, text
        If key <> "UPLOADSTATUS" And key <> "UPLOADRESULT" Then
            If IsManaged(key) Then
                inputs.Add key, text
            Else
                ' Bind custom names and values independently of column order.
                inputs.Add "CUSTOM:" & Trim$(column.Name), text
                If text <> "" Then mCustom.Add Trim$(column.Name), text
            End If
        End If
    Next column
    For Each name In Split("ACTION|ITEMCODE|ITEMNAME|UOM|QTYMODE|QUANTITY|DEFAULTLOCATION|CATEGORY|DESCRIPTION|VENDORS|VENDORCODE|EXTERNALCODE|PICTUREPATHURL|EDITREASON|UPLOADSTATUS|UPLOADRESULT", "|")
        If Not mValues.Exists(CStr(name)) Then GoTo Invalid
    Next name
    If UCase$(mValues("ACTION")) <> "EDIT" Then GoTo Invalid
    If mValues("ITEMCODE") = "" Or mValues("ITEMNAME") = "" Or mValues("EDITREASON") = "" Then GoTo Invalid
    text = UCase$(mValues("UOM"))
    packed = Split(modUomSettings.GetConfiguredUomsPackedText(), "|")
    For Each part In packed
        If UCase$(Trim$(CStr(part))) = text And text <> "" Then configured = True
    Next part
    If Not configured Then GoTo Invalid
    mode = UCase$(mValues("QTYMODE"))
    If mode = "" Then mode = "COUNTED"
    Select Case mode
        Case "COUNTED", "UTILITY", "SERVICE", "NOT COUNTED"
        Case Else: GoTo Invalid
    End Select
    text = mValues("QUANTITY")
    If text <> "" Then
        If Not IsNumeric(text) Then GoTo Invalid
        mQuantity = CDbl(text): mHasQuantity = True
    End If
    If mode = "COUNTED" Then
        If mQuantity < 0 Then GoTo Invalid
        mCustom("TRACK_QTY") = "TRUE": mCustom("ITEM_KIND") = "INVENTORY"
    Else
        mQuantity = 0: mHasQuantity = False
        mCustom("TRACK_QTY") = "FALSE"
        If mode = "NOT COUNTED" Then mCustom("ITEM_KIND") = "NON_COUNTED" Else mCustom("ITEM_KIND") = mode
        If mValues("CATEGORY") = "" Then mValues("CATEGORY") = mCustom("ITEM_KIND")
        If mValues("VENDORS") = "" And mode = "UTILITY" Then mValues("VENDORS") = "Utility"
    End If
    inputs.Add "SOURCE:WAREHOUSE", CStr(mSource(0))
    inputs.Add "SOURCE:STATION", CStr(mSource(1))
    inputs.Add "SOURCE:USER", CStr(mSource(2))
    text = modInventoryCommands.SourceInventoryPath(context)
    If text = "" Then GoTo Invalid
    inputs.Add "SOURCE:INVENTORY", LCase$(modAggregationFiles.Canonical(text))
    mHash = modInventoryEditWire.HashFields(inputs.Keys, inputs.Items)
    If mHash = "" Then GoTo Invalid
    mResult = mValues("UPLOADRESULT")
    mImported = (UCase$(mValues("UPLOADSTATUS")) = "IMPORTED")
    Read = modInventoryCommands.CheckForm(context, report)
    Exit Function
Invalid:
    report = "The worksheet EDIT inputs could not be verified. No inventory changes were made."
End Function

Public Property Get InputHash() As String
    InputHash = mHash
End Property

Public Property Get Source() As Variant
    Source = mSource
End Property

Public Function SubmittedText(ByVal normalizedHeader As String) As String
    SubmittedText = CStr(mValues(normalizedHeader))
End Function

Public Property Get HasQuantity() As Boolean
    HasQuantity = mHasQuantity
End Property

Public Property Get TargetQuantity() As Double
    TargetQuantity = mQuantity
End Property

Public Property Get ResultText() As String
    ResultText = mResult
End Property

Public Property Get Imported() As Boolean
    Imported = mImported
End Property

Public Property Get Status() As String
    Status = UCase$(mValues("UPLOADSTATUS"))
End Property

Public Property Get CatalogValues() As Variant
    CatalogValues = Array(mValues("ITEMCODE"), mValues("ITEMNAME"), UCase$(mValues("UOM")), _
        mValues("DEFAULTLOCATION"), mValues("DESCRIPTION"), mValues("VENDORS"), _
        mValues("VENDORCODE"), mValues("CATEGORY"), mValues("EXTERNALCODE"), mValues("PICTUREPATHURL"))
End Property

Public Function CatalogCustomFields() As Variant
    Dim fields As Object, key As Variant, stamp As String
    Set fields = CreateObject("Scripting.Dictionary"): fields.CompareMode = vbTextCompare
    For Each key In mCustom.Keys: fields.Add CStr(key), mCustom(key): Next key
    stamp = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    fields("LAST_EDIT_REASON") = mValues("EDITREASON"): fields("LAST_EDIT_AT") = stamp
    fields("LAST_EDIT_USER") = CStr(mSource(2))
    fields("EDIT_HISTORY_APPEND") = stamp & " | User=" & CStr(mSource(2)) & _
        " | SKU=" & mValues("ITEMCODE") & " | Reason=" & mValues("EDITREASON")
    CatalogCustomFields = Array(fields.Keys, fields.Items)
End Function

Private Function IsManaged(ByVal name As String) As Boolean
    Select Case name
        Case "ACTION", "ITEMCODE", "ITEMNAME", "UOM", "QTYMODE", "QUANTITY", _
             "DEFAULTLOCATION", "CATEGORY", "DESCRIPTION", "VENDORS", "VENDORCODE", _
             "EXTERNALCODE", "PICTUREPATHURL", "EDITREASON", "UPLOADSTATUS", "UPLOADRESULT"
            IsManaged = True
    End Select
End Function
