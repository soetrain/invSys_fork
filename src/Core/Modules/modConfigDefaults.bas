Attribute VB_Name = "modConfigDefaults"
Option Explicit

Public Const CONFIG_SCOPE_WAREHOUSE As String = "WAREHOUSE"
Public Const CONFIG_SCOPE_STATION As String = "STATION"

Public Const CONFIG_TYPE_STRING As String = "STRING"
Public Const CONFIG_TYPE_LONG As String = "LONG"
Public Const CONFIG_TYPE_BOOLEAN As String = "BOOLEAN"
Public Const CONFIG_TYPE_DATETIME As String = "DATETIME"

Public Type ConfigKeyDef
    Key As String
    DataType As String
    DefaultVal As String
    Required As Boolean
    Scope As String
End Type

Public Function GetConfigSchema(ByRef defs() As ConfigKeyDef) As Long
    Dim idx As Long
    ReDim defs(1 To 25)
    idx = 0

    AddConfigKey defs, idx, "WarehouseId", CONFIG_TYPE_STRING, "", True, CONFIG_SCOPE_WAREHOUSE
    AddConfigKey defs, idx, "WarehouseName", CONFIG_TYPE_STRING, "", True, CONFIG_SCOPE_WAREHOUSE
    AddConfigKey defs, idx, "Timezone", CONFIG_TYPE_STRING, "UTC", False, CONFIG_SCOPE_WAREHOUSE
    AddConfigKey defs, idx, "DefaultLocation", CONFIG_TYPE_STRING, "", False, CONFIG_SCOPE_WAREHOUSE
    AddConfigKey defs, idx, "BatchSize", CONFIG_TYPE_LONG, "500", False, CONFIG_SCOPE_WAREHOUSE
    AddConfigKey defs, idx, "LockTimeoutMinutes", CONFIG_TYPE_LONG, "3", False, CONFIG_SCOPE_WAREHOUSE
    AddConfigKey defs, idx, "HeartbeatIntervalSeconds", CONFIG_TYPE_LONG, "30", False, CONFIG_SCOPE_WAREHOUSE
    AddConfigKey defs, idx, "MaxLockHoldMinutes", CONFIG_TYPE_LONG, "2", False, CONFIG_SCOPE_WAREHOUSE
    AddConfigKey defs, idx, "SnapshotCadence", CONFIG_TYPE_STRING, "PER_BATCH", False, CONFIG_SCOPE_WAREHOUSE
    AddConfigKey defs, idx, "BackupCadence", CONFIG_TYPE_STRING, "DAILY", False, CONFIG_SCOPE_WAREHOUSE
    AddConfigKey defs, idx, "PathDataRoot", CONFIG_TYPE_STRING, "C:\invSys\{WarehouseId}\", False, CONFIG_SCOPE_WAREHOUSE
    AddConfigKey defs, idx, "PathBackupRoot", CONFIG_TYPE_STRING, "C:\invSys\Backups\{WarehouseId}\", False, CONFIG_SCOPE_WAREHOUSE
    AddConfigKey defs, idx, "PathSharePointRoot", CONFIG_TYPE_STRING, "", False, CONFIG_SCOPE_WAREHOUSE
    AddConfigKey defs, idx, "DesignsEnabled", CONFIG_TYPE_BOOLEAN, "FALSE", False, CONFIG_SCOPE_WAREHOUSE
    AddConfigKey defs, idx, "PoisonRetryMax", CONFIG_TYPE_LONG, "3", False, CONFIG_SCOPE_WAREHOUSE
    AddConfigKey defs, idx, "AuthCacheTTLSeconds", CONFIG_TYPE_LONG, "300", False, CONFIG_SCOPE_WAREHOUSE
    AddConfigKey defs, idx, "ProcessorServiceUserId", CONFIG_TYPE_STRING, "svc_processor", False, CONFIG_SCOPE_WAREHOUSE

    AddConfigKey defs, idx, "StationId", CONFIG_TYPE_STRING, "", True, CONFIG_SCOPE_STATION
    AddConfigKey defs, idx, "StationName", CONFIG_TYPE_STRING, "", False, CONFIG_SCOPE_STATION
    AddConfigKey defs, idx, "PathInboxRoot", CONFIG_TYPE_STRING, "", False, CONFIG_SCOPE_STATION
    AddConfigKey defs, idx, "RoleDefault", CONFIG_TYPE_STRING, "RECEIVE", False, CONFIG_SCOPE_STATION

    AddConfigKey defs, idx, "FF_DesignsEnabled", CONFIG_TYPE_BOOLEAN, "FALSE", False, CONFIG_SCOPE_WAREHOUSE
    AddConfigKey defs, idx, "FF_OutlookAlerts", CONFIG_TYPE_BOOLEAN, "FALSE", False, CONFIG_SCOPE_WAREHOUSE
    AddConfigKey defs, idx, "FF_AutoSnapshot", CONFIG_TYPE_BOOLEAN, "TRUE", False, CONFIG_SCOPE_WAREHOUSE
    AddConfigKey defs, idx, "AutoRefreshIntervalSeconds", CONFIG_TYPE_LONG, "0", False, CONFIG_SCOPE_WAREHOUSE

    GetConfigSchema = idx
End Function

Private Sub AddConfigKey(ByRef defs() As ConfigKeyDef, _
                         ByRef idx As Long, _
                         ByVal key As String, _
                         ByVal dataType As String, _
                         ByVal defaultVal As String, _
                         ByVal required As Boolean, _
                         ByVal scope As String)
    idx = idx + 1
    defs(idx).Key = key
    defs(idx).DataType = dataType
    defs(idx).DefaultVal = defaultVal
    defs(idx).Required = required
    defs(idx).Scope = scope
End Sub
