Attribute VB_Name = "modTrainingReadContext"
Option Explicit
Option Private Module

' Common private read guard for the recording and published-guide libraries.
Public Function Read(ByVal context As String, ByRef target As WarehouseTarget, ByRef policy As Object, ByRef notice As String) As Boolean
    Dim version As Long, collect As Boolean, visible As Boolean
    notice = "Unavailable: the invSys session or warehouse changed. Reopen Viewer."
    If context = "" Or context <> modActivity.CaptureContext() Then Exit Function
    Set target = modNasConnection.GetCurrentTarget(): Set policy = CreateObject("Scripting.Dictionary")
    If Not modActivityPolicy.ReadPolicy(target, "ADMIN_SETTINGS_SAVE_VALUE", version, collect, visible, notice, policy) Then Exit Function
    Read = (context = modActivity.CaptureContext())
End Function
