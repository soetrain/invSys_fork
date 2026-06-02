VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WarehouseTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public WarehouseId As String
Public WarehouseName As String
Public StationId As String
Public HubRoot As String
Public RuntimeRoot As String
Public ConfigPath As String
Public AuthPath As String
Public InboxRoot As String
Public SourceType As WH_SourceType
Public LastResolvedUTC As Date
