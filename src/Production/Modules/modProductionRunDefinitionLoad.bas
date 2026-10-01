Attribute VB_Name = "modProductionRunDefinitionLoad"
Option Explicit
Option Private Module

' Preserve the owning collections and partial-load semantics across released reads.
Public Function LoadNodeProcessDefinitions(ByVal nodes As Collection, ByVal requirements As Collection, _
                                          ByVal alternatives As Collection, ByVal outputs As Collection, _
                                          ByVal instructions As Collection, ByRef report As String, _
                                          ByVal action As cProductionWorksheetAction) As Boolean
    Dim rawNode As Variant
    Dim node As Object
    Dim jsonText As String
    Dim parseReport As String
    Dim records As Collection
    Dim rawRecord As Variant
    Dim record As Object
    Dim enriched As Object
    Dim processName As String
    Dim statusValue As String

    For Each rawNode In nodes
        Set node = rawNode
        jsonText = modOperationsPrimitiveBridge.GetProcessVersion( _
            modProductionReusableRun.RunRecordText(node, "ProcessId"), modProductionReusableRun.RunRecordText(node, "ProcessVersion"))
        If Not modProductionRunLoadActions.CanContinue(action, report) Then Exit Function
        If jsonText = "" Then
            report = "Process " & modProductionReusableRun.RunRecordText(node, "ProcessId") & " version " & _
                     modProductionReusableRun.RunRecordText(node, "ProcessVersion") & " could not be read."
            Exit Function
        End If
        parseReport = ""
        Set records = modProductionReusableDesigns.ParseReusableDefinitionRecords(jsonText, parseReport)
        If records Is Nothing Then
            report = parseReport
            Exit Function
        End If
        For Each rawRecord In records
            Set record = rawRecord
            If StrComp(modProductionReusableRun.RunRecordText(record, "RecordType"), "PROCESS", vbTextCompare) = 0 Then
                processName = modProductionReusableRun.RunRecordText(record, "ProcessName")
                statusValue = modProductionReusableRun.RunRecordText(record, "Status")
            End If
        Next rawRecord
        If StrComp(statusValue, "RELEASED", vbTextCompare) <> 0 Then
            report = "Recipe references a Process version that is not released: " & _
                     modProductionReusableRun.RunRecordText(node, "ProcessId") & " v" & modProductionReusableRun.RunRecordText(node, "ProcessVersion") & "."
            Exit Function
        End If
        node("ProcessName") = processName
        For Each rawRecord In records
            Set record = rawRecord
            Select Case UCase$(modProductionReusableRun.RunRecordText(record, "RecordType"))
                Case "REQUIREMENT", "ALTERNATIVE", "OUTPUT", "INSTRUCTION"
                    Set enriched = modProductionReusableRun.CloneRunRecord(record)
                    enriched("ProcessNodeId") = modProductionReusableRun.RunRecordText(node, "ProcessNodeId")
                    enriched("ProcessId") = modProductionReusableRun.RunRecordText(node, "ProcessId")
                    enriched("ProcessVersion") = modProductionReusableRun.RunRecordText(node, "ProcessVersion")
                    enriched("ProcessName") = processName
                    Select Case UCase$(modProductionReusableRun.RunRecordText(record, "RecordType"))
                        Case "REQUIREMENT": requirements.Add enriched
                        Case "ALTERNATIVE": alternatives.Add enriched
                        Case "OUTPUT": outputs.Add enriched
                        Case "INSTRUCTION": instructions.Add enriched
                    End Select
            End Select
        Next rawRecord
    Next rawNode
    LoadNodeProcessDefinitions = True
End Function
