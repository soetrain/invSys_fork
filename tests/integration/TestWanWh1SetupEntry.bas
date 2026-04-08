Attribute VB_Name = "TestWanWh1SetupEntry"
Option Explicit

Private mLastError As String

Public Function RunWanWh1SetupProof() As Long
    On Error GoTo FailRun
    mLastError = vbNullString
    RunWanWh1SetupProof = prove_wan_wh1_setup.SetupVerification_WH1()
    Exit Function

FailRun:
    mLastError = Err.Description
End Function

Public Function GetWanWh1SetupContext() As String
    GetWanWh1SetupContext = prove_wan_wh1_setup.GetWanWh1SetupContextPacked()
    If mLastError <> "" Then GetWanWh1SetupContext = GetWanWh1SetupContext & "|Error=" & mLastError
End Function

Public Function GetWanWh1SetupRows() As String
    GetWanWh1SetupRows = prove_wan_wh1_setup.GetWanWh1SetupEvidenceRows()
End Function
