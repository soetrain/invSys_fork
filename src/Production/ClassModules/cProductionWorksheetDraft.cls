VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cProductionWorksheetDraft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Operations-local snapshot; never canonical Designs authority.
Public TableName As String
Public ProcessId As String
Public ProcessVersion As String
Public ProcessName As String
Public Description As String
Public Payload As String
