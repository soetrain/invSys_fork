Attribute VB_Name = "test_RetireMigrateSpec"
Option Explicit

Public Function TestValidateRetireMigrateSpec_TrimsAndAcceptsArchiveOnly() As Long
    Dim spec As modWarehouseRetire.RetireMigrateSpec
    Dim report As String

    spec.SourceWarehouseId = "  WH-RET-01  "
    spec.TargetWarehouseId = "   "
    spec.OperationMode = modWarehouseRetire.MODE_ARCHIVE_ONLY
    spec.AdminUser = "  admin.user  "
    spec.ConfirmedByUser = True
    spec.ArchiveDestPath = "  C:/invSys/Archive/WH-RET-01  "
    spec.PublishTombstone = False

    On Error GoTo CleanFail
    If Not modWarehouseRetire.ValidateRetireMigrateSpec(spec, report) Then GoTo CleanExit

    If StrComp(spec.SourceWarehouseId, "WH-RET-01", vbTextCompare) = 0 _
       And spec.TargetWarehouseId = "" _
       And StrComp(spec.AdminUser, "admin.user", vbTextCompare) = 0 _
       And StrComp(spec.ArchiveDestPath, "C:\invSys\Archive\WH-RET-01", vbTextCompare) = 0 _
       And StrComp(report, "OK", vbTextCompare) = 0 Then
        TestValidateRetireMigrateSpec_TrimsAndAcceptsArchiveOnly = 1
    End If

CleanExit:
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestValidateRetireMigrateSpec_RejectsEmptySourceWarehouseId() As Long
    Dim spec As modWarehouseRetire.RetireMigrateSpec
    Dim report As String

    spec.SourceWarehouseId = "   "
    spec.OperationMode = modWarehouseRetire.MODE_ARCHIVE_ONLY
    spec.AdminUser = "admin.user"
    spec.ConfirmedByUser = True
    spec.ArchiveDestPath = "C:\invSys\Archive\WH-RET-01"

    On Error GoTo CleanFail
    If modWarehouseRetire.ValidateRetireMigrateSpec(spec, report) Then GoTo CleanExit

    If InStr(1, report, "SourceWarehouseId is required", vbTextCompare) > 0 Then
        TestValidateRetireMigrateSpec_RejectsEmptySourceWarehouseId = 1
    End If

CleanExit:
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestValidateRetireMigrateSpec_RejectsMissingTargetForMigrate() As Long
    Dim spec As modWarehouseRetire.RetireMigrateSpec
    Dim report As String

    spec.SourceWarehouseId = "WH-RET-01"
    spec.TargetWarehouseId = "   "
    spec.OperationMode = modWarehouseRetire.MODE_ARCHIVE_MIGRATE
    spec.AdminUser = "admin.user"
    spec.ConfirmedByUser = True
    spec.ArchiveDestPath = "C:\invSys\Archive\WH-RET-01"

    On Error GoTo CleanFail
    If modWarehouseRetire.ValidateRetireMigrateSpec(spec, report) Then GoTo CleanExit

    If InStr(1, report, "TargetWarehouseId is required for MODE_ARCHIVE_MIGRATE", vbTextCompare) > 0 Then
        TestValidateRetireMigrateSpec_RejectsMissingTargetForMigrate = 1
    End If

CleanExit:
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestValidateRetireMigrateSpec_RejectsEqualSourceAndTarget() As Long
    Dim spec As modWarehouseRetire.RetireMigrateSpec
    Dim report As String

    spec.SourceWarehouseId = "WH-RET-01"
    spec.TargetWarehouseId = "WH-RET-01"
    spec.OperationMode = modWarehouseRetire.MODE_ARCHIVE_MIGRATE
    spec.AdminUser = "admin.user"
    spec.ConfirmedByUser = True
    spec.ArchiveDestPath = "C:\invSys\Archive\WH-RET-01"

    On Error GoTo CleanFail
    If modWarehouseRetire.ValidateRetireMigrateSpec(spec, report) Then GoTo CleanExit

    If InStr(1, report, "must not be the same", vbTextCompare) > 0 Then
        TestValidateRetireMigrateSpec_RejectsEqualSourceAndTarget = 1
    End If

CleanExit:
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestValidateRetireMigrateSpec_RejectsUnconfirmedWriteOperation() As Long
    Dim spec As modWarehouseRetire.RetireMigrateSpec
    Dim report As String

    spec.SourceWarehouseId = "WH-RET-01"
    spec.OperationMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE
    spec.AdminUser = "admin.user"
    spec.ConfirmedByUser = False
    spec.ArchiveDestPath = "C:\invSys\Archive\WH-RET-01"
    spec.PublishTombstone = True

    On Error GoTo CleanFail
    If modWarehouseRetire.ValidateRetireMigrateSpec(spec, report) Then GoTo CleanExit

    If InStr(1, report, "ConfirmedByUser must be True", vbTextCompare) > 0 Then
        TestValidateRetireMigrateSpec_RejectsUnconfirmedWriteOperation = 1
    End If

CleanExit:
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestValidateRetireMigrateSpec_RejectsInvalidArchiveDestPath() As Long
    Dim spec As modWarehouseRetire.RetireMigrateSpec
    Dim report As String

    spec.SourceWarehouseId = "WH-RET-01"
    spec.OperationMode = modWarehouseRetire.MODE_ARCHIVE_ONLY
    spec.AdminUser = "admin.user"
    spec.ConfirmedByUser = True
    spec.ArchiveDestPath = "\\server\share\archive"

    On Error GoTo CleanFail
    If modWarehouseRetire.ValidateRetireMigrateSpec(spec, report) Then GoTo CleanExit

    If InStr(1, report, "ArchiveDestPath must be a valid local path format", vbTextCompare) > 0 Then
        TestValidateRetireMigrateSpec_RejectsInvalidArchiveDestPath = 1
    End If

CleanExit:
    Exit Function
CleanFail:
    Resume CleanExit
End Function
