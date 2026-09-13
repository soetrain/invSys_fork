Attribute VB_Name = "modEventDetailModel"
Option Explicit
Option Private Module

Public Function Defaults() As Object
    Dim model As Object, rows As Collection, catalog As Object, family As Variant, id As Variant, ordinal As Long
    Set model = CreateObject("Scripting.Dictionary")
    Set rows = New Collection
    Set catalog = modEventDetailCatalog.Fields()
    model.Add "SchemaVersion", 1
    For Each family In modEventDetailCatalog.Families()
        ordinal = 0
        For Each id In catalog.Keys
            ordinal = ordinal + 1
            rows.Add MakeRow(CStr(family), CStr(id), CBool(catalog(id)("DefaultEnabled")), ordinal)
        Next id
    Next family
    model.Add "Fields", rows
    Set Defaults = model
End Function

Public Function MakeRow(ByVal family As String, ByVal id As String, ByVal enabled As Boolean, ByVal ordinal As Long) As Object
    Dim row As Object
    Set row = CreateObject("Scripting.Dictionary")
    row.Add "EventFamily", family
    row.Add "FieldId", id
    row.Add "Enabled", enabled
    row.Add "DisplayOrder", ordinal
    Set MakeRow = row
End Function

Public Function Decode(ByVal request As String) As Object
    Dim model As Object, row As Variant, catalog As Object, families As Object, fieldsSeen As Object, orderSeen As Object
    Dim family As Variant, id As String, order As Long, definition As Object
    On Error GoTo Invalid
    If LenB(request) > 1048576 Then Exit Function
    Set model = modTrainingJson.DecodeObject(request)
    If Not ExactFields(model, Array("SchemaVersion", "Fields")) Then Exit Function
    If Not IntegerInRange(model("SchemaVersion"), 1, 1) Then Exit Function
    If TypeName(model("Fields")) <> "Collection" Then Exit Function
    Set catalog = modEventDetailCatalog.Fields()
    Set families = CreateObject("Scripting.Dictionary")
    For Each family In modEventDetailCatalog.Families()
        families.Add CStr(family), True
    Next family
    Set fieldsSeen = CreateObject("Scripting.Dictionary")
    Set orderSeen = CreateObject("Scripting.Dictionary")
    For Each row In model("Fields")
        If Not IsObject(row) Then Exit Function
        If Not ExactFields(row, Array("EventFamily", "FieldId", "Enabled", "DisplayOrder")) Then Exit Function
        If VarType(row("EventFamily")) <> vbString Or VarType(row("FieldId")) <> vbString Then Exit Function
        family = row("EventFamily"): id = row("FieldId")
        If Not families.Exists(family) Or Not catalog.Exists(id) Then Exit Function
        If VarType(row("Enabled")) <> vbBoolean Then Exit Function
        If Not IntegerInRange(row("DisplayOrder"), 1, catalog.Count) Then Exit Function
        order = CLng(row("DisplayOrder"))
        Set definition = catalog(id)
        If definition("Required") And Not row("Enabled") Then Exit Function
        If fieldsSeen.Exists(family & "|" & id) Or orderSeen.Exists(family & "|" & CStr(order)) Then Exit Function
        fieldsSeen.Add family & "|" & id, True
        orderSeen.Add family & "|" & CStr(order), True
    Next row
    If model("Fields").Count <> families.Count * catalog.Count Then Exit Function
    Set Decode = model
Invalid:
End Function

Public Function ExactFields(ByVal value As Object, ByVal names As Variant) As Boolean
    Dim name As Variant
    If value Is Nothing Then Exit Function
    If TypeName(value) <> "Dictionary" Then Exit Function
    If value.Count <> UBound(names) + 1 Then Exit Function
    For Each name In names
        If Not value.Exists(CStr(name)) Then Exit Function
    Next name
    ExactFields = True
End Function

Public Function IntegerInRange(ByVal value As Variant, ByVal minimum As Long, ByVal maximum As Long) As Boolean
    Select Case VarType(value)
        Case vbInteger, vbLong, vbDouble
            If value >= minimum And value <= maximum Then IntegerInRange = (value = Fix(value))
    End Select
End Function
