Attribute VB_Name = "TestWanWh2SetupEntry"
Option Explicit

Private mLastError As String

Public Function RunWanWh2SetupProof() As Long
    On Error GoTo FailRun
    mLastError = vbNullString
    RunWanWh2SetupProof = prove_wan_wh2_setup.SetupVerification_WH2()
    Exit Function

FailRun:
    mLastError = Err.Description
End Function

Public Function GetWanWh2SetupContext() As String
    GetWanWh2SetupContext = prove_wan_wh2_setup.GetWanWh2SetupContextPacked()
    If mLastError <> "" Then GetWanWh2SetupContext = GetWanWh2SetupContext & "|Error=" & mLastError
End Function

Public Function GetWanWh2SetupRows() As String
    GetWanWh2SetupRows = prove_wan_wh2_setup.GetWanWh2SetupEvidenceRows()
End Function
