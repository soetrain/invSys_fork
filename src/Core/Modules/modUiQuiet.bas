Attribute VB_Name = "modUiQuiet"
Option Explicit

Private mQuietDepth As Long
Private mQuietOwnerKey As String
Private mPrevScreenUpdating As Boolean
Private mPrevEnableEvents As Boolean
Private mPrevDisplayAlerts As Boolean
Private mPrevCalculation As XlCalculation

Public Sub BeginQuietUi(Optional ByVal ownerWb As Workbook = Nothing)
    If mQuietDepth = 0 Then
        mQuietOwnerKey = BuildQuietWorkbookKey(ownerWb)
        mPrevScreenUpdating = Application.ScreenUpdating
        mPrevEnableEvents = Application.EnableEvents
        mPrevDisplayAlerts = Application.DisplayAlerts
        mPrevCalculation = Application.Calculation

        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.DisplayAlerts = False
        Application.Calculation = xlCalculationManual
    End If
    mQuietDepth = mQuietDepth + 1
End Sub

Public Sub EndQuietUi()
    If mQuietDepth <= 0 Then
        mQuietDepth = 0
        mQuietOwnerKey = vbNullString
        Exit Sub
    End If

    mQuietDepth = mQuietDepth - 1
    If mQuietDepth = 0 Then
        On Error Resume Next
        Application.Calculation = mPrevCalculation
        Application.DisplayAlerts = mPrevDisplayAlerts
        Application.EnableEvents = mPrevEnableEvents
        Application.ScreenUpdating = mPrevScreenUpdating
        On Error GoTo 0
        mQuietOwnerKey = vbNullString
    End If
End Sub

Public Function QuietUiIsActive() As Boolean
    QuietUiIsActive = (mQuietDepth > 0)
End Function

Public Sub ReactivateQuietOwner()
    Dim wb As Workbook

    If Not QuietUiIsActive() Then Exit Sub
    Set wb = ResolveQuietOwnerWorkbook()
    If wb Is Nothing Then Exit Sub

    On Error Resume Next
    wb.Activate
    On Error GoTo 0
End Sub

Private Function ResolveQuietOwnerWorkbook() As Workbook
    Dim wb As Workbook

    If Trim$(mQuietOwnerKey) = "" Then Exit Function

    For Each wb In Application.Workbooks
        If StrComp(BuildQuietWorkbookKey(wb), mQuietOwnerKey, vbTextCompare) = 0 Then
            Set ResolveQuietOwnerWorkbook = wb
            Exit Function
        End If
    Next wb
End Function

Private Function BuildQuietWorkbookKey(ByVal wb As Workbook) As String
    If wb Is Nothing Then Exit Function

    If Trim$(wb.FullName) <> "" Then
        BuildQuietWorkbookKey = LCase$(Trim$(wb.FullName))
    Else
        BuildQuietWorkbookKey = LCase$(Trim$(wb.Name))
    End If
End Function
