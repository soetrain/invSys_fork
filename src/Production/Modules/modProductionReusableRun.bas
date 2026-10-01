Attribute VB_Name = "modProductionReusableRun"
Option Explicit

Private Const EVENT_TYPE_PROD_CONSUME As String = "PROD_CONSUME"
Private Const EVENT_TYPE_PROD_COMPLETE As String = "PROD_COMPLETE"
Private Const QTY_TOLERANCE As Double = 0.0000001
Private Const RUN_CHECK_SOURCE_PROCESS_OUTPUT As String = "Source Process / Output"
Private Const RUN_CHECK_REMAINING_BALANCE As String = "Remaining Balance"

Private mLoaded As Boolean
Private mCheckedIn As Boolean
Private mCompleted As Boolean
Private mRecipeId As String
Private mRecipeVersion As String
Private mRecipeName As String
Private mScalePercent As Double
Private mBatchNumber As Long
Private mNodes As Collection
Private mRequirements As Collection
Private mAlternatives As Collection
Private mOutputs As Collection
Private mConnections As Collection
Private mOutputRegulations As Collection
Private mInstructions As Collection
Private mAllocations As Object
Private mAllocationNativeQuantities As Object
Private mAllocationConversionAudit As Object
Private mOutputKeys As Object
Private mActualOutputQty As Object
Private mLastOutputQty As Object
Private mOutputHistory As Collection
Private mCompletedNodes As Object
Private mCheckedInNodeId As String
Private mRunId As String
Private mEventIds As String
Private mAppliedCount As Long
Private mProcessorReports As String
Private mLastSummary As String
Private mBatchNote As String
Private mBatchNoteFrozen As Boolean

Public Sub ClearReusableRun()
    mLoaded = False
    mCheckedIn = False
    mCheckedInNodeId = ""
    mCompleted = False
    mRecipeId = ""
    mRecipeVersion = ""
    mRecipeName = ""
    mScalePercent = 100#
    mBatchNumber = 1
    Set mNodes = New Collection
    Set mRequirements = New Collection
    Set mAlternatives = New Collection
    Set mOutputs = New Collection
    Set mConnections = New Collection
    Set mOutputRegulations = New Collection
    Set mInstructions = New Collection
    Set mAllocations = NewTextDictionary()
    Set mAllocationNativeQuantities = NewTextDictionary()
    Set mAllocationConversionAudit = NewTextDictionary()
    Set mOutputKeys = NewTextDictionary()
    Set mActualOutputQty = NewTextDictionary()
    Set mLastOutputQty = NewTextDictionary()
    Set mOutputHistory = New Collection
    Set mCompletedNodes = NewTextDictionary()
    mCheckedInNodeId = ""
    mRunId = ""
    mEventIds = ""
    mAppliedCount = 0
    mProcessorReports = ""
    mLastSummary = ""
    mBatchNote = ""
    mBatchNoteFrozen = False
End Sub

Public Function LoadReleasedReusableRecipe(ByVal recipeId As String, _
                                           ByVal recipeVersion As String, _
                                           ByVal scalePercent As Double, _
                                           Optional ByRef report As String = "", _
                                           Optional ByVal action As cProductionWorksheetAction = Nothing) As Boolean
    On Error GoTo Failed

    Dim validation As String
    Dim jsonText As String
    Dim parseReport As String
    Dim recipeRecords As Collection
    Dim rawRecord As Variant
    Dim record As Object
    Dim node As Object

    ClearReusableRun
    recipeId = Trim$(recipeId)
    recipeVersion = Trim$(recipeVersion)
    If recipeId = "" Or recipeVersion = "" Then
        report = "Recipe ID and version are required."
        Exit Function
    End If
    If scalePercent < 0.001 Or scalePercent > 1000# Then
        report = "Batch scale must be from 0.001% through 1000%."
        Exit Function
    End If
    validation = modOperationsPrimitiveBridge.ValidateReleasedRecipe(recipeId, recipeVersion)
    If Not modProductionRunLoadActions.CanContinue(action, report) Then Exit Function
    If Left$(validation, 2) <> "1" & vbTab Then
        report = "Released Recipe validation failed: " & Replace$(validation, vbTab, " ")
        Exit Function
    End If
    jsonText = modOperationsPrimitiveBridge.GetRecipeGraph(recipeId, recipeVersion)
    If Not modProductionRunLoadActions.CanContinue(action, report) Then Exit Function
    If jsonText = "" Then
        report = "Released Recipe graph could not be read."
        Exit Function
    End If
    Set recipeRecords = modProductionReusableDesigns.ParseReusableDefinitionRecords(jsonText, parseReport)
    If recipeRecords Is Nothing Then
        report = parseReport
        Exit Function
    End If

    mRecipeId = recipeId
    mRecipeVersion = recipeVersion
    mScalePercent = scalePercent
    For Each rawRecord In recipeRecords
        Set record = rawRecord
        Select Case UCase$(modProductionReusableDesigns.ReusableRecordText(record, "RecordType"))
            Case "RECIPE"
                mRecipeName = modProductionReusableDesigns.ReusableRecordText(record, "RecipeName")
            Case "PROCESS_NODE"
                Set node = CloneRunRecord(record)
                AddNodeInExecutionOrder node
            Case "CONNECTION"
                mConnections.Add CloneRunRecord(record)
            Case "OUTPUT_REGULATION"
                mOutputRegulations.Add CloneRunRecord(record)
        End Select
    Next rawRecord
    If mNodes.Count = 0 Then
        report = "Released Recipe contains no Process nodes."
        Exit Function
    End If
    If Not LoadNodeProcessDefinitions(report, action) Then Exit Function
    If Not ApplyRecipeOutputRegulationOverrides(report) Then Exit Function
    If Not ValidateLoadedRunGraph(report) Then Exit Function

    mLoaded = True
    report = "Loaded released Recipe " & recipeId & " version " & recipeVersion & _
             " with " & CStr(mNodes.Count) & " Process(es) at " & _
             FormatRunNumberLocal(scalePercent) & "% scale."
    LoadReleasedReusableRecipe = True
    Exit Function

Failed:
    report = "Reusable Recipe load failed: " & Err.Description
End Function

Public Function ApplyReusableRunScale(ByVal scalePercent As Double, _
                                      Optional ByRef report As String = "") As Boolean
    If Not mLoaded Then
        report = "Load a released Recipe before applying Batch scale."
        Exit Function
    End If
    If scalePercent < 0.001 Or scalePercent > 1000# Then
        report = "Batch scale must be from 0.001% through 1000%."
        Exit Function
    End If
    If Not mCompletedNodes Is Nothing Then
        If mCompletedNodes.Count > 0 Then
            report = "Batch scale cannot change after a Process has completed. Finish or clear the batch."
            Exit Function
        End If
    End If
    mScalePercent = scalePercent
    Set mAllocations = NewTextDictionary()
    Set mAllocationNativeQuantities = NewTextDictionary()
    Set mAllocationConversionAudit = NewTextDictionary()
    Set mOutputKeys = NewTextDictionary()
    Set mActualOutputQty = NewTextDictionary()
    Set mCompletedNodes = NewTextDictionary()
    mCheckedInNodeId = ""
    mRunId = ""
    mEventIds = ""
    mAppliedCount = 0
    mProcessorReports = ""
    mCheckedIn = False
    mCheckedInNodeId = ""
    mCompleted = False
    mLastSummary = ""
    mBatchNote = ""
    mBatchNoteFrozen = False
    report = "Batch scale applied: " & FormatRunNumberLocal(scalePercent) & _
             "%. Exact-key allocations were cleared and recalculated requirements are ready."
    ApplyReusableRunScale = True
End Function

Public Function ReusableRunLoaderRows(Optional ByVal locationFilter As String = "", _
    Optional ByVal action As cProductionWorksheetAction = Nothing) As Variant
    Dim result() As Variant
    Dim totalRows As Long
    Dim rowIndex As Long
    Dim rawRecord As Variant
    Dim record As Object

    If Not mLoaded Then Exit Function
    totalRows = mRequirements.Count + mOutputs.Count
    If totalRows = 0 Then Exit Function
    ReDim result(1 To totalRows, 1 To 9)
    For Each rawRecord In mRequirements
        Set record = rawRecord
        rowIndex = rowIndex + 1
        result(rowIndex, 1) = RunRecordText(record, "ProcessName")
        result(rowIndex, 2) = RunRecordText(record, "ProcessNodeId")
        result(rowIndex, 3) = "INPUT"
        result(rowIndex, 4) = RunRecordText(record, "RequirementName")
        result(rowIndex, 5) = RunRecordNumber(record, "Percent")
        result(rowIndex, 6) = RunRecordText(record, "UOM")
        result(rowIndex, 7) = ScaledRecordQty(record)
        result(rowIndex, 8) = ReusableRunLineStatus( _
            RunRecordText(record, "ProcessNodeId"), "INPUT", _
            RunRecordText(record, "RequirementId"), locationFilter, action)
        If Not modProductionRunLoadActions.CanContinue(action) Then Exit Function
        result(rowIndex, 9) = RunRecordText(record, "RequirementId")
    Next rawRecord
    For Each rawRecord In mOutputs
        Set record = rawRecord
        rowIndex = rowIndex + 1
        result(rowIndex, 1) = RunRecordText(record, "ProcessName")
        result(rowIndex, 2) = RunRecordText(record, "ProcessNodeId")
        result(rowIndex, 3) = "OUTPUT"
        result(rowIndex, 4) = RunRecordText(record, "OutputName")
        result(rowIndex, 5) = RunRecordNumber(record, "Percent")
        result(rowIndex, 6) = RunRecordText(record, "UOM")
        result(rowIndex, 7) = ScaledRecordQty(record)
        result(rowIndex, 8) = ReusableRunLineStatus( _
            RunRecordText(record, "ProcessNodeId"), "OUTPUT", _
            RunRecordText(record, "OutputId"), locationFilter, action)
        If Not modProductionRunLoadActions.CanContinue(action) Then Exit Function
        result(rowIndex, 9) = RunRecordText(record, "OutputId")
    Next rawRecord
    ReusableRunLoaderRows = result
End Function

Public Function ReusableRunInstructionRows(ByVal processName As String) As Variant
    Dim result() As Variant
    Dim rawInstruction As Variant
    Dim instruction As Object
    Dim rowCount As Long
    Dim rowIndex As Long

    processName = Trim$(processName)
    If Not mLoaded Or processName = "" Then Exit Function
    For Each rawInstruction In mInstructions
        Set instruction = rawInstruction
        If StrComp(RunRecordText(instruction, "ProcessName"), processName, vbTextCompare) = 0 Then _
            rowCount = rowCount + 1
    Next rawInstruction
    If rowCount = 0 Then Exit Function
    ReDim result(1 To rowCount, 1 To 2)
    For Each rawInstruction In mInstructions
        Set instruction = rawInstruction
        If StrComp(RunRecordText(instruction, "ProcessName"), processName, vbTextCompare) = 0 Then
            rowIndex = rowIndex + 1
            result(rowIndex, 1) = RunRecordText(instruction, "InstructionOrdinal")
            result(rowIndex, 2) = RunRecordText(instruction, "Instruction")
        End If
    Next rawInstruction
    ReusableRunInstructionRows = result
End Function

Public Function ReusableRunPaletteRows(Optional ByVal locationFilter As String = "", _
    Optional ByVal action As cProductionWorksheetAction = Nothing) As Variant
    On Error GoTo CleanFail

    Dim entities As Variant
    Dim result() As Variant
    Dim trimmed() As Variant
    Dim buckets As Object
    Dim bucketOrder As Collection
    Dim rawRequirement As Variant
    Dim requirement As Object
    Dim entityRow As Long
    Dim outRow As Long
    Dim rowIndex As Long
    Dim c As Long
    Dim nodeId As String
    Dim requirementId As String
    Dim itemCode As String
    Dim bucketKey As String
    Dim rawBucketKey As Variant
    Dim bucket As Variant
    Dim allocatedQty As Double
    Dim requiredQty As Double
    Dim conversionFactor As Double
    Dim conversionVersion As String
    Dim conversionReport As String

    If Not mLoaded Then Exit Function
    entities = modInventoryDomainBridge.ListAvailableInventoryEntitiesBridge("")
    If Not modProductionRunLoadActions.CanContinue(action) Then Exit Function
    If Not IsArray(entities) Then Exit Function
    ReDim result(1 To mRequirements.Count * UBound(entities, 1), 1 To 10)
    locationFilter = Trim$(locationFilter)
    For Each rawRequirement In mRequirements
        Set requirement = rawRequirement
        nodeId = RunRecordText(requirement, "ProcessNodeId")
        requirementId = RunRecordText(requirement, "RequirementId")
        If NodeIsComplete(nodeId) Then GoTo NextRequirement
        If RequirementHasIncomingConnection(nodeId, requirementId) Then GoTo NextRequirement
        Set buckets = NewTextDictionary()
        Set bucketOrder = New Collection
        For entityRow = LBound(entities, 1) To UBound(entities, 1)
            itemCode = Trim$(CStr(entities(entityRow, 3)))
            If Not RequirementAllowsItem(nodeId, requirementId, itemCode, CStr(entities(entityRow, 2))) Then GoTo NextEntity
            If Not modUomSettings.GetUomConversion(CStr(entities(entityRow, 5)), _
                    RunRecordText(requirement, "UOM"), conversionFactor, conversionVersion, conversionReport) Then GoTo NextEntity
            If locationFilter <> "" Then
                If StrComp(locationFilter, Trim$(CStr(entities(entityRow, 7))), vbTextCompare) <> 0 Then GoTo NextEntity
            End If
            bucketKey = StockBucketKeyForEntity(entities, entityRow)
            If Not buckets.Exists(bucketKey) Then
                bucket = Array(CStr(entities(entityRow, 1)), CStr(entities(entityRow, 4)), _
                    CStr(entities(entityRow, 5)), CStr(entities(entityRow, 7)), 0#, _
                    EntityRowIsNonCounted(entities, entityRow), _
                    CStr(ExactEntityInventoryDisplay(entities, entityRow)))
                buckets.Add bucketKey, bucket
                bucketOrder.Add bucketKey
            End If
            bucket = buckets(bucketKey)
            If Not CBool(bucket(5)) And IsNumeric(entities(entityRow, 6)) Then _
                bucket(4) = CDbl(bucket(4)) + CDbl(entities(entityRow, 6))
            buckets(bucketKey) = bucket
NextEntity:
        Next entityRow
        requiredQty = ScaledRecordQty(requirement)
        For Each rawBucketKey In bucketOrder
            bucketKey = CStr(rawBucketKey)
            bucket = buckets(bucketKey)
            allocatedQty = StockBucketAllocatedQty(entities, nodeId, requirementId, bucketKey)
            outRow = outRow + 1
            result(outRow, 1) = nodeId
            result(outRow, 2) = requirementId
            result(outRow, 3) = RunRecordText(requirement, "RequirementName")
            result(outRow, 4) = bucket(0)
            result(outRow, 5) = bucket(1)
            If allocatedQty > QTY_TOLERANCE And requiredQty > 0 Then _
                result(outRow, 6) = allocatedQty / requiredQty * 100#
            If allocatedQty > QTY_TOLERANCE Then result(outRow, 7) = allocatedQty
            result(outRow, 8) = IIf(Trim$(CStr(bucket(2))) <> "", bucket(2), RunRecordText(requirement, "UOM")) & _
                                " / " & RunRecordText(requirement, "UOM")
            result(outRow, 9) = StockBucketAvailableDisplay(CDbl(bucket(4)), CBool(bucket(5)), CStr(bucket(6)))
            If CBool(bucket(5)) Then
                result(outRow, 9) = CStr(bucket(6))
            ElseIf modUomSettings.GetUomConversion(CStr(bucket(2)), RunRecordText(requirement, "UOM"), _
                    conversionFactor, conversionVersion, conversionReport) Then
                result(outRow, 9) = result(outRow, 9) & " / " & FormatRunNumberLocal(CDbl(bucket(4)) * conversionFactor)
            End If
            result(outRow, 10) = bucket(3)
        Next rawBucketKey
NextRequirement:
    Next rawRequirement
    If outRow = 0 Then Exit Function
    ReDim trimmed(1 To outRow, 1 To 10)
    For rowIndex = 1 To outRow
        For c = 1 To 10
            trimmed(rowIndex, c) = result(rowIndex, c)
        Next c
    Next rowIndex
    ReusableRunPaletteRows = trimmed
CleanFail:
End Function

Public Function ApplyReusableRunStockAllocation(ByVal processNodeId As String, _
                                                 ByVal requirementId As String, _
                                                 ByVal representativeSystemKey As String, _
                                                 ByVal qty As Double, _
                                                 Optional ByRef report As String = "", _
                                                 Optional ByVal action As cProductionWorksheetAction = Nothing) As Boolean
    On Error GoTo Failed

    Dim requirement As Object
    Dim entities As Variant
    Dim representativeRow As Long
    Dim entityRow As Long
    Dim bucketKey As String
    Dim entityKey As String
    Dim allocationId As String
    Dim currentAllocation As Double
    Dim currentBucketAllocation As Double
    Dim otherRequirementQty As Double
    Dim otherEntityQty As Double
    Dim availableForPlan As Double
    Dim bucketAvailable As Double
    Dim requiredQty As Double
    Dim remainingQty As Double
    Dim takeQty As Double
    Dim nonCounted As Boolean
    Dim planned As Object
    Dim removeIds As Collection
    Dim removeId As Variant
    Dim planId As Variant
    Dim uomReport As String
    Dim conversionFactor As Double
    Dim conversionVersion As String

    If Not mLoaded Then
        report = "Load a released Recipe first."
        Exit Function
    End If
    Set requirement = FindRequirement(processNodeId, requirementId)
    If requirement Is Nothing Or RequirementHasIncomingConnection(processNodeId, requirementId) Then
        report = "The selected external requirement could not be resolved."
        Exit Function
    End If
    If qty < 0 Then
        report = "Allocation quantity cannot be negative."
        Exit Function
    End If
    If Not modUomSettings.ValidateQuantityForUom(qty, _
            RunRecordText(requirement, "UOM"), uomReport) Then
        report = uomReport
        Exit Function
    End If

    If Not modProductionRunLoadActions.CanContinue(action, report) Then Exit Function
    entities = modInventoryDomainBridge.ListAvailableInventoryEntitiesBridge("")
    If Not modProductionRunLoadActions.CanContinue(action, report) Then Exit Function
    If Not IsArray(entities) Then
        If Not action Is Nothing Then action.OutcomeCode = "FAILED"
        report = "No managed inventory stock is available."
        Exit Function
    End If
    representativeRow = FindExactEntityRow(entities, representativeSystemKey)
    If representativeRow = 0 Then
        report = "The selected stock bucket is no longer available. Refresh Production Run."
        Exit Function
    End If
    If Not RequirementAllowsItem(processNodeId, requirementId, _
            CStr(entities(representativeRow, 3)), CStr(entities(representativeRow, 2))) Then
        report = "The selected stock bucket is not acceptable for this requirement."
        Exit Function
    End If
    If Not modUomSettings.GetUomConversion(CStr(entities(representativeRow, 5)), _
            RunRecordText(requirement, "UOM"), conversionFactor, conversionVersion, uomReport) Then
        report = uomReport
        Exit Function
    End If
    If StrComp(CStr(entities(representativeRow, 5)), RunRecordText(requirement, "UOM"), vbTextCompare) <> 0 Then
        If Not modProductionRunLoadActions.CanContinue(action, report) Then Exit Function
        ApplyReusableRunStockAllocation = ApplyReusableRunAllocation(processNodeId, requirementId, _
            representativeSystemKey, qty, report, action)
        Exit Function
    End If

    bucketKey = StockBucketKeyForEntity(entities, representativeRow)
    nonCounted = EntityRowIsNonCounted(entities, representativeRow)
    Set removeIds = New Collection
    For entityRow = LBound(entities, 1) To UBound(entities, 1)
        If StrComp(StockBucketKeyForEntity(entities, entityRow), bucketKey, vbTextCompare) = 0 Then
            entityKey = Trim$(CStr(entities(entityRow, 1)))
            allocationId = AllocationKey(processNodeId, requirementId, entityKey)
            currentAllocation = 0#
            If mAllocations.Exists(allocationId) Then
                currentAllocation = CDbl(mAllocations(allocationId))
                currentBucketAllocation = currentBucketAllocation + currentAllocation
                removeIds.Add allocationId
            End If
            If nonCounted Then
                bucketAvailable = qty
            Else
                otherEntityQty = AllocationTotalForEntity(entityKey, "") - currentAllocation
                availableForPlan = 0#
                If IsNumeric(entities(entityRow, 6)) Then _
                    availableForPlan = CDbl(entities(entityRow, 6)) - otherEntityQty
                If availableForPlan > 0 Then bucketAvailable = bucketAvailable + availableForPlan
            End If
        End If
    Next entityRow

    requiredQty = ScaledRecordQty(requirement)
    otherRequirementQty = AllocationTotalForRequirement(processNodeId, requirementId, "") - _
        currentBucketAllocation
    If Not RequirementIsActual(requirement) And otherRequirementQty + qty > requiredQty + QTY_TOLERANCE Then
        report = "Allocation exceeds the scaled requirement of " & FormatRunNumberLocal(requiredQty) & "."
        Exit Function
    End If
    If Not nonCounted And qty > bucketAvailable + QTY_TOLERANCE Then
        report = "Allocation exceeds stock bucket availability of " & _
            FormatRunNumberLocal(bucketAvailable) & "."
        Exit Function
    End If

    Set planned = NewTextDictionary()
    remainingQty = qty
    For entityRow = LBound(entities, 1) To UBound(entities, 1)
        If StrComp(StockBucketKeyForEntity(entities, entityRow), bucketKey, vbTextCompare) = 0 Then
            entityKey = Trim$(CStr(entities(entityRow, 1)))
            allocationId = AllocationKey(processNodeId, requirementId, entityKey)
            currentAllocation = 0#
            If mAllocations.Exists(allocationId) Then currentAllocation = CDbl(mAllocations(allocationId))
            If nonCounted Then
                If remainingQty > QTY_TOLERANCE Then
                    planned.Add allocationId, remainingQty
                    remainingQty = 0#
                End If
                Exit For
            End If
            otherEntityQty = AllocationTotalForEntity(entityKey, "") - currentAllocation
            availableForPlan = 0#
            If IsNumeric(entities(entityRow, 6)) Then _
                availableForPlan = CDbl(entities(entityRow, 6)) - otherEntityQty
            If availableForPlan > QTY_TOLERANCE And remainingQty > QTY_TOLERANCE Then
                takeQty = availableForPlan
                If takeQty > remainingQty Then takeQty = remainingQty
                planned.Add allocationId, takeQty
                remainingQty = remainingQty - takeQty
            End If
        End If
    Next entityRow
    If remainingQty > QTY_TOLERANCE Then
        report = "Allocation could not be expanded across sufficient exact inventory keys."
        Exit Function
    End If

    If Not modProductionRunLoadActions.CanContinue(action, report) Then Exit Function
    For Each removeId In removeIds
        RemoveAllocation CStr(removeId)
    Next removeId
    For Each planId In planned.Keys
        SaveAllocation CStr(planId), CDbl(planned(planId)), CDbl(planned(planId)), _
            conversionVersion, conversionFactor
    Next planId
    mCheckedIn = False
    mCheckedInNodeId = ""
    mCompleted = False
    report = "Exact allocation expansion saved: " & FormatRunNumberLocal(qty) & _
             " of " & FormatRunNumberLocal(requiredQty) & " across " & _
             CStr(planned.Count) & " System_Key entity(s)."
    ApplyReusableRunStockAllocation = True
    Exit Function

Failed:
    report = "Stock allocation failed: " & Err.Description
    If Not action Is Nothing Then action.OutcomeCode = "FAILED"
End Function

Public Function ApplyReusableRunAllocation(ByVal processNodeId As String, _
                                         ByVal requirementId As String, _
                                         ByVal systemKey As String, _
                                         ByVal qty As Double, _
                                         Optional ByRef report As String = "", _
                                         Optional ByVal action As cProductionWorksheetAction = Nothing) As Boolean
    Dim requirement As Object
    Dim availableQty As Double
    Dim requiredQty As Double
    Dim otherRequirementQty As Double
    Dim otherEntityQty As Double
    Dim allocationId As String
    Dim nonCounted As Boolean
    Dim uomReport As String
    Dim nativeUom As String
    Dim conversionFactor As Double
    Dim catalogVersion As String
    Dim nativeQty As Double

    If Not mLoaded Then
        report = "Load a released Recipe first."
        Exit Function
    End If
    Set requirement = FindRequirement(processNodeId, requirementId)
    If requirement Is Nothing Or RequirementHasIncomingConnection(processNodeId, requirementId) Then
        report = "The selected external requirement could not be resolved."
        Exit Function
    End If
    If qty < 0 Then
        report = "Allocation quantity cannot be negative."
        Exit Function
    End If
    If Not modUomSettings.ValidateQuantityForUom(qty, _
            RunRecordText(requirement, "UOM"), uomReport) Then
        report = uomReport
        Exit Function
    End If
    If Not modProductionRunLoadActions.CanContinue(action, report) Then Exit Function
    nativeUom = ExactEntityUom(systemKey, action)
    If Not modProductionRunLoadActions.CanContinue(action, report) Then Exit Function
    If nativeUom = "" Then
        report = "The selected System_Key is no longer available. Refresh Production Run."
        Exit Function
    End If
    If Not modUomSettings.GetUomConversion(nativeUom, RunRecordText(requirement, "UOM"), _
            conversionFactor, catalogVersion, uomReport) Then
        report = uomReport
        Exit Function
    End If
    nativeQty = qty / conversionFactor
    If Not modUomSettings.ValidateQuantityForUom(nativeQty, nativeUom, uomReport) Then
        report = uomReport
        Exit Function
    End If
    If Not modProductionRunLoadActions.CanContinue(action, report) Then Exit Function
    nonCounted = ExactEntityIsNonCounted(systemKey, action)
    If Not modProductionRunLoadActions.CanContinue(action, report) Then Exit Function
    availableQty = ExactEntityAvailableQty(systemKey, , action)
    If Not modProductionRunLoadActions.CanContinue(action, report) Then Exit Function
    If Not nonCounted And availableQty <= 0 And qty > 0 Then
        report = "The selected System_Key is no longer available. Refresh Production Run."
        Exit Function
    End If
    requiredQty = ScaledRecordQty(requirement)
    allocationId = AllocationKey(processNodeId, requirementId, systemKey)
    otherRequirementQty = AllocationTotalForRequirement(processNodeId, requirementId, allocationId)
    If Not RequirementIsActual(requirement) And otherRequirementQty + qty > requiredQty + QTY_TOLERANCE Then
        report = "Allocation exceeds the scaled requirement of " & FormatRunNumberLocal(requiredQty) & "."
        Exit Function
    End If
    otherEntityQty = AllocationTotalForEntity(systemKey, allocationId)
    If Not nonCounted And otherEntityQty + nativeQty > availableQty + QTY_TOLERANCE Then
        report = "Allocation exceeds exact entity availability of " & FormatRunNumberLocal(availableQty) & "."
        Exit Function
    End If
    If qty <= QTY_TOLERANCE Then
        RemoveAllocation allocationId
    Else
        SaveAllocation allocationId, qty, nativeQty, catalogVersion, conversionFactor
    End If
    mCheckedIn = False
    mCheckedInNodeId = ""
    mCompleted = False
    report = "Exact inventory allocation saved: " & FormatRunNumberLocal(qty) & " " & _
             RunRecordText(requirement, "UOM") & " requiring " & _
             FormatRunNumberLocal(nativeQty) & " " & nativeUom & _
             " from the exact System_Key."
    ApplyReusableRunAllocation = True
End Function

Public Function CheckInReusableRun(ByVal runLocation As String, _
                                   Optional ByRef report As String = "") As Boolean
    Dim rawRequirement As Variant
    Dim requirement As Object
    Dim requiredQty As Double
    Dim allocatedQty As Double
    Dim key As Variant
    Dim systemKey As String
    Dim liveQty As Double
    Dim allocatedForEntity As Double
    Dim liveLocation As String
    Dim nonCounted As Boolean

    If Not mLoaded Then
        report = "Load a released Recipe first."
        Exit Function
    End If
    If mCompleted Then
        report = "This batch is already complete. Click Next Batch."
        Exit Function
    End If
    If Not mCompletedNodes Is Nothing Then
        If mCompletedNodes.Count > 0 Then
            report = "Select one remaining Process; this batch is already in Process-scoped execution."
            Exit Function
        End If
    End If
    For Each rawRequirement In mRequirements
        Set requirement = rawRequirement
        If Not RequirementHasIncomingConnection(RunRecordText(requirement, "ProcessNodeId"), _
                                                 RunRecordText(requirement, "RequirementId")) Then
            requiredQty = ScaledRecordQty(requirement)
            allocatedQty = AllocationTotalForRequirement( _
                RunRecordText(requirement, "ProcessNodeId"), _
                RunRecordText(requirement, "RequirementId"), "")
            If Abs(allocatedQty - requiredQty) > QTY_TOLERANCE Then
                report = "Inventory is insufficient or unresolved for " & _
                         RunRecordText(requirement, "RequirementName") & ". Required=" & _
                         FormatRunNumberLocal(requiredQty) & "; allocated=" & _
                         FormatRunNumberLocal(allocatedQty) & "."
                Exit Function
            End If
        End If
    Next rawRequirement
    For Each key In mAllocations.Keys
        systemKey = AllocationSystemKey(CStr(key))
        liveQty = ExactEntityAvailableQty(systemKey, liveLocation)
        nonCounted = ExactEntityIsNonCounted(systemKey)
        allocatedForEntity = AllocationTotalForEntity(systemKey, "")
        If Not nonCounted And liveQty + QTY_TOLERANCE < allocatedForEntity Then
            report = "Stale allocation rejected for System_Key " & systemKey & _
                     ". Available=" & FormatRunNumberLocal(liveQty) & "; allocated=" & _
                     FormatRunNumberLocal(allocatedForEntity) & ". Refresh Production Run."
            Exit Function
        End If
        If Trim$(runLocation) <> "" And StrComp(Trim$(runLocation), liveLocation, vbTextCompare) <> 0 Then
            report = "System_Key " & systemKey & " is at " & liveLocation & _
                     "; the Production run location is " & Trim$(runLocation) & "."
            Exit Function
        End If
    Next key
    mCheckedIn = True
    mBatchNoteFrozen = True
    report = "Checked in " & FormatRunNumberLocal(TotalExternalAllocation()) & _
             " units using " & CStr(mAllocations.Count) & " exact System_Key allocation(s)."
    CheckInReusableRun = True
End Function

Public Function CheckInReusableProcess(ByVal processName As String, _
                                       ByVal runLocation As String, _
                                       Optional ByRef report As String = "") As Boolean
    Dim node As Object
    Dim nodeId As String

    If Not mLoaded Then
        report = "Load a released Recipe first."
        Exit Function
    End If
    If mCompleted Then
        report = "This batch is already complete. Click Next Batch."
        Exit Function
    End If
    Set node = FindNodeByProcessName(processName)
    If node Is Nothing Then
        report = "Choose one Process before Check In."
        Exit Function
    End If
    nodeId = RunRecordText(node, "ProcessNodeId")
    If NodeIsComplete(nodeId) Then
        report = "Process " & RunRecordText(node, "ProcessName") & " is already complete."
        Exit Function
    End If
    If Not ValidateProcessRequirementsReady(node, report) Then Exit Function
    If Not ValidateProcessAllocationsLive(node, runLocation, report) Then Exit Function

    mCheckedIn = True
    mBatchNoteFrozen = True
    mCheckedInNodeId = nodeId
    report = "Checked in Process " & RunRecordText(node, "ProcessName") & _
             " using " & CStr(AllocationCountForNode(nodeId)) & _
             " exact System_Key allocation(s)."
    CheckInReusableProcess = True
End Function

Public Function ReusableRunManagerCheckRows(Optional ByVal action As cProductionWorksheetAction = Nothing) As Variant
    Dim result() As Variant
    Dim key As Variant
    Dim rowIndex As Long
    Dim entity As Variant
    Dim systemKey As String
    Dim qty As Double
    Dim totalRows As Long
    Dim keyParts() As String
    Dim node As Object
    Dim requirement As Object
    Dim routedRow As Variant
    Dim columnIndex As Long
    Dim nativeQty As Double
    Dim requirementUom As String
    Dim stockUom As String

    If Not mCheckedIn Or mAllocations Is Nothing Then Exit Function
    If mCheckedInNodeId = "" Then
        totalRows = mAllocations.Count + RoutedInputCount("")
    Else
        totalRows = AllocationCountForNode(mCheckedInNodeId) + RoutedInputCount(mCheckedInNodeId)
    End If
    If totalRows = 0 Then Exit Function
    entity = modInventoryDomainBridge.ListAvailableInventoryEntitiesBridge("")
    If Not modProductionRunLoadActions.CanContinue(action) Then Exit Function
    ReDim result(1 To totalRows, 1 To 9)
    For Each key In mAllocations.Keys
        If mCheckedInNodeId <> "" Then
            If Not AllocationBelongsToNode(CStr(key), mCheckedInNodeId) Then GoTo NextAllocation
        End If
        rowIndex = rowIndex + 1
        keyParts = Split(CStr(key), vbTab)
        systemKey = AllocationSystemKey(CStr(key))
        qty = CDbl(mAllocations(key))
        Set node = FindNode(keyParts(0))
        Set requirement = FindRequirement(keyParts(0), keyParts(1))
        result(rowIndex, 1) = "EXTERNAL"
        result(rowIndex, 2) = RunRecordText(node, "ProcessName") & " / " & _
                              RunRecordText(requirement, "RequirementName")
        result(rowIndex, 4) = systemKey
        result(rowIndex, 5) = ExactEntityField(entity, systemKey, 3)
        result(rowIndex, 6) = ExactEntityField(entity, systemKey, 4)
        requirementUom = RunRecordText(requirement, "UOM")
        stockUom = ExactEntityField(entity, systemKey, 5)
        nativeQty = NativeAllocationQty(CStr(key))
        If StrComp(requirementUom, stockUom, vbTextCompare) = 0 Then
            result(rowIndex, 7) = stockUom
            result(rowIndex, 8) = qty
        Else
            result(rowIndex, 7) = requirementUom & " / " & stockUom
            result(rowIndex, 8) = FormatRunNumberLocal(qty) & " / " & _
                                  FormatRunNumberLocal(nativeQty)
        End If
        result(rowIndex, 9) = ExactEntityInventoryDisplayForKey(entity, systemKey)
NextAllocation:
    Next key
    For Each key In mRequirements
        Set requirement = key
        If mCheckedInNodeId <> "" Then
            If StrComp(RunRecordText(requirement, "ProcessNodeId"), mCheckedInNodeId, vbTextCompare) <> 0 Then GoTo NextRoutedRequirement
        End If
        If Not RequirementHasIncomingConnection(RunRecordText(requirement, "ProcessNodeId"), _
                                                 RunRecordText(requirement, "RequirementId")) Then GoTo NextRoutedRequirement
        routedRow = BuildRoutedInputCheckRow(requirement, entity)
        If IsEmpty(routedRow) Then GoTo NextRoutedRequirement
        rowIndex = rowIndex + 1
        For columnIndex = LBound(routedRow) To UBound(routedRow)
            result(rowIndex, columnIndex) = routedRow(columnIndex)
        Next columnIndex
NextRoutedRequirement:
    Next key
    If rowIndex = 0 Then Exit Function
    ReusableRunManagerCheckRows = result
End Function

Public Function ReusableRunOutputRows() As Variant
    Dim result() As Variant
    Dim rawHistory As Variant
    Dim history As Object
    Dim rawOutput As Variant
    Dim output As Object
    Dim totalRows As Long
    Dim rowIndex As Long
    Dim outputKey As String

    If Not mLoaded Or mOutputs.Count = 0 Then Exit Function
    totalRows = mOutputHistory.Count
    If Not mCompleted Then totalRows = totalRows + ActiveOutputCount()
    If totalRows = 0 Then Exit Function
    ReDim result(1 To totalRows, 1 To 9)
    For Each rawHistory In mOutputHistory
        Set history = rawHistory
        rowIndex = rowIndex + 1
        result(rowIndex, 1) = RunRecordText(history, "ProcessName")
        result(rowIndex, 2) = RunRecordText(history, "OutputName")
        result(rowIndex, 3) = RunRecordText(history, "UOM")
        result(rowIndex, 4) = RunRecordNumber(history, "ActualQty")
        result(rowIndex, 5) = CLng(RunRecordNumber(history, "BatchNumber"))
        result(rowIndex, 6) = RunRecordText(history, "UsedGoods")
        result(rowIndex, 7) = RunRecordNumber(history, "ProcessTotal")
        result(rowIndex, 8) = RunRecordText(history, "Recall")
        result(rowIndex, 9) = RunRecordText(history, "System_Key")
    Next rawHistory
    If Not mCompleted Then
        For Each rawOutput In mOutputs
            Set output = rawOutput
            If NodeIsComplete(RunRecordText(output, "ProcessNodeId")) Then GoTo NextActiveOutput
            rowIndex = rowIndex + 1
            outputKey = OutputIdentityKey(RunRecordText(output, "ProcessNodeId"), RunRecordText(output, "OutputId"))
            result(rowIndex, 1) = RunRecordText(output, "ProcessName")
            result(rowIndex, 2) = RunRecordText(output, "OutputName")
            result(rowIndex, 3) = RunRecordText(output, "UOM")
            result(rowIndex, 5) = mBatchNumber
            result(rowIndex, 6) = ProcessUsedGoodsDisplay( _
                RunRecordText(output, "ProcessNodeId"))
            result(rowIndex, 7) = ProcessOutputTotal( _
                RunRecordText(output, "ProcessName"), RunRecordText(output, "OutputName"), _
                RunRecordText(output, "UOM"))
            result(rowIndex, 8) = mRecipeId & "-B" & Format$(mBatchNumber, "0000")
            If mOutputKeys.Exists(outputKey) Then result(rowIndex, 9) = mOutputKeys(outputKey)
NextActiveOutput:
        Next rawOutput
    End If
    ReusableRunOutputRows = result
End Function

Public Function ReusableRunOutputDefinitionIndex(ByVal displayRowIndex As Long) As Long
    Dim firstCurrentRow As Long
    Dim rawOutput As Variant
    Dim output As Object
    Dim activeIndex As Long
    Dim definitionIndex As Long

    If Not mLoaded Or mCompleted Then Exit Function
    firstCurrentRow = mOutputHistory.Count + 1
    If displayRowIndex < firstCurrentRow Then Exit Function
    activeIndex = displayRowIndex - firstCurrentRow + 1
    For Each rawOutput In mOutputs
        definitionIndex = definitionIndex + 1
        Set output = rawOutput
        If Not NodeIsComplete(RunRecordText(output, "ProcessNodeId")) Then
            activeIndex = activeIndex - 1
            If activeIndex = 0 Then
                ReusableRunOutputDefinitionIndex = definitionIndex
                Exit Function
            End If
        End If
    Next rawOutput
End Function

Public Function ReusableRunPlannedOutput(ByVal outputIndex As Long) As String
    If Not mLoaded Then Exit Function
    If outputIndex < 1 Or outputIndex > mOutputs.Count Then Exit Function
    ReusableRunPlannedOutput = FormatRunNumberLocal(ScaledRecordQty(mOutputs(outputIndex)))
End Function

Public Function StageReusableRunActualOutput(ByVal outputIndex As Long, _
                                             ByVal quantityText As String, _
                                             Optional ByRef report As String = "") As Boolean
    Dim output As Object
    Dim outputKey As String
    Dim actualQty As Double
    Dim uomReport As String

    If Not mLoaded Then
        report = "Load a released Recipe before entering Actual Output."
        Exit Function
    End If
    If mCompleted Then
        report = "This batch is already complete. Click Next Batch before entering Actual Output."
        Exit Function
    End If
    If outputIndex < 1 Or outputIndex > mOutputs.Count Then
        report = "Select the Production Output row whose Actual Output is being entered."
        Exit Function
    End If

    Set output = mOutputs(outputIndex)
    outputKey = OutputIdentityKey(RunRecordText(output, "ProcessNodeId"), _
                                  RunRecordText(output, "OutputId"))
    quantityText = Trim$(quantityText)
    If quantityText = "" Then
        If mActualOutputQty.Exists(outputKey) Then mActualOutputQty.Remove outputKey
        report = "Actual Output cleared for " & RunRecordText(output, "OutputName") & "."
        StageReusableRunActualOutput = True
        Exit Function
    End If
    If Not IsNumeric(quantityText) Then
        report = "Actual Output must be a number greater than zero."
        Exit Function
    End If
    actualQty = CDbl(quantityText)
    If actualQty <= 0 Then
        report = "Actual Output must be a number greater than zero."
        Exit Function
    End If
    If Not modUomSettings.ValidateQuantityForUom(actualQty, _
            RunRecordText(output, "UOM"), uomReport) Then
        report = uomReport
        Exit Function
    End If

    mActualOutputQty(outputKey) = actualQty
    report = "Actual Output staged for " & RunRecordText(output, "OutputName") & _
             ": " & FormatRunNumberLocal(actualQty) & " " & RunRecordText(output, "UOM") & "."
    StageReusableRunActualOutput = True
End Function

Public Function ReusableRunActualOutput(ByVal outputIndex As Long) As String
    Dim output As Object
    Dim outputKey As String

    If Not mLoaded Then Exit Function
    If outputIndex < 1 Or outputIndex > mOutputs.Count Then Exit Function
    Set output = mOutputs(outputIndex)
    outputKey = OutputIdentityKey(RunRecordText(output, "ProcessNodeId"), _
                                  RunRecordText(output, "OutputId"))
    If mActualOutputQty.Exists(outputKey) Then _
        ReusableRunActualOutput = FormatRunNumberLocal(CDbl(mActualOutputQty(outputKey)))
End Function

Public Function CompleteReusableRun(ByVal runLocation As String, _
                                    Optional ByRef report As String = "") As Boolean
    On Error GoTo Failed

    Dim recheckReport As String
    Dim rawNode As Variant
    Dim node As Object
    Dim items As Collection
    Dim eventIds As String
    Dim eventId As String
    Dim queueError As String
    Dim processorReport As String
    Dim processorReports As String
    Dim appliedCount As Long
    Dim processedNow As Long
    Dim runId As String

    If Not mCheckedIn Then
        report = "Check exact inventory into Production before completing the run."
        Exit Function
    End If
    If mCompleted Then
        report = "This batch is already complete. Click Next Batch."
        Exit Function
    End If
    If Not CheckInReusableRun(runLocation, recheckReport) Then
        report = recheckReport
        Exit Function
    End If
    If Not ValidateReusableActualOutputs(report) Then Exit Function
    AssignFreshOutputKeys
    runId = "PROD-RUN-" & Replace$(modRoleEventWriter.CreateSystemKey(), "-", "")
    For Each rawNode In mNodes
        Set node = rawNode
        Set items = BuildNodeConsumeItems(node, runLocation, runId)
        If items.Count > 0 Then
            eventId = ""
            If Not modRoleEventWriter.QueuePayloadEventCurrent(EVENT_TYPE_PROD_CONSUME, "", _
                    modProductionJson.BuildJsonArray(items), RunEventNote(runId, node, "CONSUME"), _
                    eventId, queueError) Then
                report = "Production consume event was not queued: " & queueError
                Exit Function
            End If
            AppendEventId eventIds, eventId
            processedNow = modProcessor.RunBatch(Trim$(modConfig.GetWarehouseId()), 0, processorReport)
            If processedNow < 1 Then
                report = "Production consume event " & eventId & " was not applied. " & processorReport
                Exit Function
            End If
            appliedCount = appliedCount + processedNow
            AppendProcessorReport processorReports, processorReport
        End If
        Set items = BuildNodeCompleteItems(node, runLocation, runId)
        If items.Count = 0 Then
            report = "Process " & RunRecordText(node, "ProcessName") & " has no output to complete."
            Exit Function
        End If
        eventId = ""
        If Not modRoleEventWriter.QueuePayloadEventCurrent(EVENT_TYPE_PROD_COMPLETE, "", _
                modProductionJson.BuildJsonArray(items), RunEventNote(runId, node, "COMPLETE"), _
                eventId, queueError) Then
            report = "Production complete event was not queued: " & queueError
            Exit Function
        End If
        AppendEventId eventIds, eventId
        processedNow = modProcessor.RunBatch(Trim$(modConfig.GetWarehouseId()), 0, processorReport)
        If processedNow < 1 Then
            report = "Production complete event " & eventId & " was not applied. " & processorReport
            Exit Function
        End If
        appliedCount = appliedCount + processedNow
        AppendProcessorReport processorReports, processorReport
    Next rawNode

    If Not VerifyCompletedOutputBalances(report) Then
        report = report & " Processor applied=" & CStr(appliedCount) & ". " & processorReports
        Exit Function
    End If
    CaptureLastActualOutputs
    CaptureCompletedOutputHistory
    mCompleted = True
    mLastSummary = "RunId=" & runId & "; Recipe=" & mRecipeId & " v" & mRecipeVersion & _
                   "; Batch=" & CStr(mBatchNumber) & "; Scale=" & _
                   FormatRunNumberLocal(mScalePercent) & "%; ExactInputs=" & _
                   CStr(mAllocations.Count) & "; Outputs=" & CStr(mOutputs.Count) & _
                   "; Events=" & eventIds & "; ProcessorApplied=" & CStr(appliedCount) & "."
    report = "Production batch completed and persisted. " & mLastSummary
    CompleteReusableRun = True
    Exit Function

Failed:
    report = "Reusable Production completion failed: " & Err.Description
End Function

Public Function CompleteReusableProcess(ByVal processName As String, _
                                        ByVal runLocation As String, _
                                        Optional ByRef report As String = "") As Boolean
    On Error GoTo Failed

    Dim node As Object
    Dim nodeId As String
    Dim recheckReport As String
    Dim items As Collection
    Dim eventId As String
    Dim queueError As String
    Dim processorReport As String
    Dim processedNow As Long

    Set node = FindNodeByProcessName(processName)
    If node Is Nothing Then
        report = "Choose one Process before Complete Run."
        Exit Function
    End If
    nodeId = RunRecordText(node, "ProcessNodeId")
    If Not mCheckedIn Or StrComp(mCheckedInNodeId, nodeId, vbTextCompare) <> 0 Then
        report = "Check exact inventory into Process " & _
                 RunRecordText(node, "ProcessName") & " before completing it."
        Exit Function
    End If
    If Not CheckInReusableProcess(processName, runLocation, recheckReport) Then
        report = recheckReport
        Exit Function
    End If
    If Not ValidateReusableActualOutputsForNode(nodeId, report) Then Exit Function
    AssignFreshOutputKeysForNode nodeId
    If mRunId = "" Then _
        mRunId = "PROD-RUN-" & Replace$(modRoleEventWriter.CreateSystemKey(), "-", "")

    Set items = BuildNodeConsumeItems(node, runLocation, mRunId)
    If items.Count > 0 Then
        If Not modRoleEventWriter.QueuePayloadEventCurrent(EVENT_TYPE_PROD_CONSUME, "", _
                modProductionJson.BuildJsonArray(items), RunEventNote(mRunId, node, "CONSUME"), _
                eventId, queueError) Then
            report = "Production consume event was not queued: " & queueError
            Exit Function
        End If
        AppendEventId mEventIds, eventId
        processedNow = modProcessor.RunBatch(Trim$(modConfig.GetWarehouseId()), 0, processorReport)
        If processedNow < 1 Then
            report = "Production consume event " & eventId & " was not applied. " & processorReport
            Exit Function
        End If
        mAppliedCount = mAppliedCount + processedNow
        AppendProcessorReport mProcessorReports, processorReport
    End If

    Set items = BuildNodeCompleteItems(node, runLocation, mRunId)
    If items.Count = 0 Then
        report = "Process " & RunRecordText(node, "ProcessName") & " has no output to complete."
        Exit Function
    End If
    eventId = ""
    If Not modRoleEventWriter.QueuePayloadEventCurrent(EVENT_TYPE_PROD_COMPLETE, "", _
            modProductionJson.BuildJsonArray(items), RunEventNote(mRunId, node, "COMPLETE"), _
            eventId, queueError) Then
        report = "Production complete event was not queued: " & queueError
        Exit Function
    End If
    AppendEventId mEventIds, eventId
    processedNow = modProcessor.RunBatch(Trim$(modConfig.GetWarehouseId()), 0, processorReport)
    If processedNow < 1 Then
        report = "Production complete event " & eventId & " was not applied. " & processorReport
        Exit Function
    End If
    mAppliedCount = mAppliedCount + processedNow
    AppendProcessorReport mProcessorReports, processorReport
    If Not VerifyNodeOutputBalances(nodeId, False, report) Then Exit Function

    mCompletedNodes.Add nodeId, True
    CaptureCompletedNodeOutputHistory nodeId
    mCheckedIn = False
    mCheckedInNodeId = ""
    If AllNodesCompleted() Then
        If Not VerifyCompletedOutputBalances(report) Then Exit Function
        CaptureLastActualOutputs
        mCompleted = True
        mLastSummary = "RunId=" & mRunId & "; Recipe=" & mRecipeId & " v" & mRecipeVersion & _
                       "; Batch=" & CStr(mBatchNumber) & "; Scale=" & _
                       FormatRunNumberLocal(mScalePercent) & "%; ExactInputs=" & _
                       CStr(mAllocations.Count) & "; Outputs=" & CStr(mOutputs.Count) & _
                       "; Events=" & mEventIds & "; ProcessorApplied=" & CStr(mAppliedCount) & "."
        report = "Production batch completed and persisted. " & mLastSummary
    Else
        report = "Process " & RunRecordText(node, "ProcessName") & _
                 " completed and persisted. Select another READY Process. RunId=" & mRunId & "."
    End If
    CompleteReusableProcess = True
    Exit Function

Failed:
    report = "Reusable Process completion failed: " & Err.Description
End Function

Public Function BeginNextReusableBatch(Optional ByRef report As String = "") As Boolean
    If Not mLoaded Then
        report = "Load a released Recipe first."
        Exit Function
    End If
    If Not mCompleted Then
        report = "Complete the current batch before starting the next batch."
        Exit Function
    End If
    mBatchNumber = mBatchNumber + 1
    Set mAllocations = NewTextDictionary()
    Set mAllocationNativeQuantities = NewTextDictionary()
    Set mAllocationConversionAudit = NewTextDictionary()
    Set mOutputKeys = NewTextDictionary()
    Set mActualOutputQty = NewTextDictionary()
    Set mCompletedNodes = NewTextDictionary()
    mCheckedInNodeId = ""
    mRunId = ""
    mBatchNote = ""
    mBatchNoteFrozen = False
    mEventIds = ""
    mAppliedCount = 0
    mProcessorReports = ""
    mCheckedIn = False
    mCompleted = False
    mLastSummary = ""
    report = "Next Batch " & CStr(mBatchNumber) & " is ready; exact-key allocations were cleared."
    BeginNextReusableBatch = True
End Function

Public Function SetReusableRunBatchNote(ByVal batchNote As String, _
                                        Optional ByRef report As String = "") As Boolean
    batchNote = Trim$(batchNote)
    If Len(batchNote) > 500 Then
        report = "Batch Note must be 500 characters or fewer."
        Exit Function
    End If
    If mBatchNoteFrozen And StrComp(batchNote, mBatchNote, vbBinaryCompare) <> 0 Then
        report = "Batch Note is frozen after Check In for this batch. Start Next Batch to use a different note."
        Exit Function
    End If
    mBatchNote = batchNote
    SetReusableRunBatchNote = True
End Function

Public Function ReusableRunBatchNote() As String
    ReusableRunBatchNote = mBatchNote
End Function

Public Function ReusableRunBatchNoteFrozen() As Boolean
    ReusableRunBatchNoteFrozen = mBatchNoteFrozen
End Function

Public Function ReusableRunIsLoaded() As Boolean
    ReusableRunIsLoaded = mLoaded
End Function

Public Function ReusableRunIsCheckedIn() As Boolean
    ReusableRunIsCheckedIn = mCheckedIn
End Function

Public Function ReusableRunIsCompleted() As Boolean
    ReusableRunIsCompleted = mCompleted
End Function

Public Function ReusableRunIsProcessComplete(ByVal processName As String) As Boolean
    Dim node As Object
    Set node = FindNodeByProcessName(processName)
    If Not node Is Nothing Then _
        ReusableRunIsProcessComplete = NodeIsComplete(RunRecordText(node, "ProcessNodeId"))
End Function

Public Function ReusableRunBatchNumber() As Long
    ReusableRunBatchNumber = mBatchNumber
End Function

Public Function ReusableRunScalePercent() As Double
    ReusableRunScalePercent = mScalePercent
End Function

Public Function ReusableRunLastSummary() As String
    ReusableRunLastSummary = mLastSummary
End Function

Public Function ReusableRunIdentityForTest() As String
    ReusableRunIdentityForTest = "RecipeId=" & mRecipeId & _
        "|RecipeVersion=" & mRecipeVersion & _
        "|RunId=" & mRunId & _
        "|Batch=" & CStr(mBatchNumber)
End Function

Public Function ReusableRunOutputSystemKey(ByVal processNodeId As String, _
                                           ByVal outputId As String) As String
    Dim key As String
    key = OutputIdentityKey(processNodeId, outputId)
    If Not mOutputKeys Is Nothing Then
        If mOutputKeys.Exists(key) Then ReusableRunOutputSystemKey = CStr(mOutputKeys(key))
    End If
End Function

Public Function ReusableRunExactEntityQty(ByVal systemKey As String) As Double
    ReusableRunExactEntityQty = ExactEntityAvailableQty(systemKey)
End Function

Public Function ReusableRunRequirementQty(ByVal processNodeId As String, _
                                          ByVal requirementId As String) As Double
    Dim requirement As Object
    Set requirement = FindRequirement(processNodeId, requirementId)
    If Not requirement Is Nothing Then ReusableRunRequirementQty = ScaledRecordQty(requirement)
End Function

Private Function LoadNodeProcessDefinitions(ByRef report As String, ByVal action As cProductionWorksheetAction) As Boolean
    LoadNodeProcessDefinitions = modProductionRunDefinitionLoad.LoadNodeProcessDefinitions( _
        mNodes, mRequirements, mAlternatives, mOutputs, mInstructions, report, action)
End Function

Private Function ValidateLoadedRunGraph(ByRef report As String) As Boolean
    Dim rawRequirement As Variant
    Dim requirement As Object
    Dim rawConnection As Variant
    Dim connection As Object
    Dim sourceOutput As Object
    Dim sourceNode As Object
    Dim targetNode As Object
    Dim rawOutput As Variant
    Dim output As Object
    Dim floorQty As Double
    Dim ceilingQty As Double

    If mOutputs.Count < mNodes.Count Then
        report = "Every Process must declare at least one output."
        Exit Function
    End If
    For Each rawConnection In mConnections
        Set connection = rawConnection
        Set sourceNode = FindNode(RunRecordText(connection, "FromProcessNodeId"))
        Set targetNode = FindNode(RunRecordText(connection, "ToProcessNodeId"))
        Set sourceOutput = FindOutput(RunRecordText(connection, "FromProcessNodeId"), _
                                      RunRecordText(connection, "FromOutputId"))
        Set requirement = FindRequirement(RunRecordText(connection, "ToProcessNodeId"), _
                                          RunRecordText(connection, "ToRequirementId"))
        If sourceNode Is Nothing Or targetNode Is Nothing Or sourceOutput Is Nothing Or requirement Is Nothing Then
            report = "Recipe contains an unresolved Process connection."
            Exit Function
        End If
        If RunRecordNumber(sourceNode, "ExecutionOrdinal") >= RunRecordNumber(targetNode, "ExecutionOrdinal") Then
            report = "Recipe execution order is invalid or circular."
            Exit Function
        End If
        If Not UomCompatible(RunRecordText(sourceOutput, "UOM"), RunRecordText(requirement, "UOM")) Then
            report = "Recipe connection UOM is incompatible for " & RunRecordText(requirement, "RequirementName") & "."
            Exit Function
        End If
    Next rawConnection
    For Each rawRequirement In mRequirements
        Set requirement = rawRequirement
        If Not RequirementHasIncomingConnection(RunRecordText(requirement, "ProcessNodeId"), _
                                                 RunRecordText(requirement, "RequirementId")) Then
            If Not RequirementHasAlternative(RunRecordText(requirement, "ProcessNodeId"), _
                                             RunRecordText(requirement, "RequirementId")) Then
                report = "Unresolved external input: " & RunRecordText(requirement, "RequirementName") & "."
                Exit Function
            End If
        End If
        If Not RequirementIsActual(requirement) And ScaledRecordQty(requirement) <= 0 Then
            report = "Requirement quantity could not be resolved for " & RunRecordText(requirement, "RequirementName") & "."
            Exit Function
        End If
    Next rawRequirement
    For Each rawOutput In mOutputs
        Set output = rawOutput
        If OutputRegulationEnabled(output) Then
            floorQty = RunRecordNumber(output, "OutputFloorQty")
            ceilingQty = RunRecordNumber(output, "OutputCeilingQty")
            If floorQty <= 0 Or ceilingQty <= 0 Or floorQty - ceilingQty > QTY_TOLERANCE Then
                report = "Enabled output regulation requires a positive floor not above its ceiling."
                Exit Function
            End If
            If UCase$(Trim$(RunRecordText(output, "UOM"))) = "EA" _
               And (Abs(floorQty - Fix(floorQty)) > QTY_TOLERANCE _
                    Or Abs(ceilingQty - Fix(ceilingQty)) > QTY_TOLERANCE) Then
                report = "EA output regulation floor and ceiling must be whole units."
                Exit Function
            End If
            If ScaledOutputRegulationCeiling(output) + QTY_TOLERANCE < _
                    OutgoingQtyForOutput(RunRecordText(output, "ProcessNodeId"), _
                                         RunRecordText(output, "OutputId")) Then
                report = "Output regulation ceiling is below its routed downstream commitment."
                Exit Function
            End If
        End If
    Next rawOutput
    ValidateLoadedRunGraph = True
End Function

Public Function ReusableRunLineStatus(ByVal nodeId As String, _
                                      ByVal lineType As String, _
                                      ByVal recordId As String, _
                                      Optional ByVal locationFilter As String = "", _
                                      Optional ByVal action As cProductionWorksheetAction = Nothing) As String
    Dim requirement As Object
    Dim statusText As String

    If NodeIsComplete(nodeId) Then
        ReusableRunLineStatus = "COMPLETE"
        Exit Function
    End If
    If StrComp(lineType, "INPUT", vbTextCompare) = 0 Then
        Set requirement = FindRequirement(nodeId, recordId)
        statusText = RequirementReadinessStatus(requirement, locationFilter, action)
    Else
        statusText = ProcessReadinessStatus(nodeId, locationFilter, action)
    End If
    If Not modProductionRunLoadActions.CanContinue(action) Then Exit Function
    ReusableRunLineStatus = statusText
End Function

Private Function ProcessReadinessStatus(ByVal nodeId As String, _
                                        ByVal locationFilter As String, _
                                        Optional ByVal action As cProductionWorksheetAction = Nothing) As String
    Dim rawRequirement As Variant
    Dim requirement As Object
    Dim statusText As String
    Dim needsAllocation As Boolean
    Dim waitingUpstream As Boolean

    If NodeIsComplete(nodeId) Then
        ProcessReadinessStatus = "COMPLETE"
        Exit Function
    End If
    For Each rawRequirement In mRequirements
        Set requirement = rawRequirement
        If StrComp(RunRecordText(requirement, "ProcessNodeId"), nodeId, vbTextCompare) = 0 Then
            statusText = RequirementReadinessStatus(requirement, locationFilter, action)
            If Not modProductionRunLoadActions.CanContinue(action) Then Exit Function
            If statusText = "! INSUFFICIENT" Then
                ProcessReadinessStatus = statusText
                Exit Function
            ElseIf statusText = "WAITING UPSTREAM" Then
                waitingUpstream = True
            ElseIf statusText = "NEEDS ALLOCATION" Then
                needsAllocation = True
            End If
        End If
    Next rawRequirement
    If waitingUpstream Then
        ProcessReadinessStatus = "WAITING UPSTREAM"
    ElseIf needsAllocation Then
        ProcessReadinessStatus = "NEEDS ALLOCATION"
    Else
        ProcessReadinessStatus = "READY"
    End If
End Function

Private Function RequirementReadinessStatus(ByVal requirement As Object, _
                                            ByVal locationFilter As String, _
                                        Optional ByVal action As cProductionWorksheetAction = Nothing) As String
    Dim nodeId As String
    Dim requirementId As String
    Dim requiredQty As Double
    Dim allocatedQty As Double
    Dim connection As Object
    Dim sourceKey As String

    If requirement Is Nothing Then
        RequirementReadinessStatus = "! INSUFFICIENT"
        Exit Function
    End If
    nodeId = RunRecordText(requirement, "ProcessNodeId")
    requirementId = RunRecordText(requirement, "RequirementId")
    requiredQty = ScaledRecordQty(requirement)
    Set connection = IncomingConnectionForRequirement(nodeId, requirementId)
    If Not connection Is Nothing Then
        If Not NodeIsComplete(RunRecordText(connection, "FromProcessNodeId")) Then
            RequirementReadinessStatus = "WAITING UPSTREAM"
            Exit Function
        End If
        sourceKey = OutputKeyForConnection(connection)
        If sourceKey = "" Or _
           ExactEntityAvailableQty(sourceKey, , action) + QTY_TOLERANCE < ScaledConnectionQty(connection) Then
            RequirementReadinessStatus = "! INSUFFICIENT"
        Else
            RequirementReadinessStatus = "READY"
        End If
        Exit Function
    End If
    allocatedQty = AllocationTotalForRequirement(nodeId, requirementId, "")
    If RequirementIsActual(requirement) Then
        If allocatedQty > QTY_TOLERANCE Then
            RequirementReadinessStatus = "READY"
        Else
            RequirementReadinessStatus = "NEEDS ACTUAL INPUT"
        End If
    ElseIf Abs(allocatedQty - requiredQty) <= QTY_TOLERANCE Then
        RequirementReadinessStatus = "READY"
    ElseIf AvailableStockForRequirement(requirement, locationFilter, action) + QTY_TOLERANCE < requiredQty Then
        RequirementReadinessStatus = "! INSUFFICIENT"
    Else
        RequirementReadinessStatus = "NEEDS ALLOCATION"
    End If
End Function

Private Function RequirementIsActual(ByVal requirement As Object) As Boolean
    RequirementIsActual = (StrComp(RunRecordText(requirement, "RequirementQtyMode"), _
        "ACTUAL", vbTextCompare) = 0)
End Function

Private Function ValidateProcessRequirementsReady(ByVal node As Object, _
                                                  ByRef report As String) As Boolean
    Dim rawRequirement As Variant
    Dim requirement As Object
    Dim nodeId As String
    Dim requirementId As String
    Dim requiredQty As Double
    Dim allocatedQty As Double
    Dim connection As Object
    Dim sourceKey As String

    nodeId = RunRecordText(node, "ProcessNodeId")
    For Each rawRequirement In mRequirements
        Set requirement = rawRequirement
        If StrComp(RunRecordText(requirement, "ProcessNodeId"), nodeId, vbTextCompare) <> 0 Then GoTo NextRequirement
        requirementId = RunRecordText(requirement, "RequirementId")
        Set connection = IncomingConnectionForRequirement(nodeId, requirementId)
        If Not connection Is Nothing Then
            If Not NodeIsComplete(RunRecordText(connection, "FromProcessNodeId")) Then
                report = "Upstream output is not ready for " & _
                         RunRecordText(requirement, "RequirementName") & "."
                Exit Function
            End If
            sourceKey = OutputKeyForConnection(connection)
            If sourceKey = "" Or _
               ExactEntityAvailableQty(sourceKey) + QTY_TOLERANCE < ScaledConnectionQty(connection) Then
                report = "Upstream output is not ready or is insufficient for " & _
                         RunRecordText(requirement, "RequirementName") & "."
                Exit Function
            End If
        Else
            requiredQty = ScaledRecordQty(requirement)
            allocatedQty = AllocationTotalForRequirement(nodeId, requirementId, "")
            If RequirementIsActual(requirement) And allocatedQty <= QTY_TOLERANCE Then
                report = "Actual requirement requires a measured external allocation for " & _
                         RunRecordText(requirement, "RequirementName") & "."
                Exit Function
            End If
            If Not RequirementIsActual(requirement) And Abs(allocatedQty - requiredQty) > QTY_TOLERANCE Then
                report = "Inventory is insufficient or unresolved for " & _
                         RunRecordText(requirement, "RequirementName") & ". Required=" & _
                         FormatRunNumberLocal(requiredQty) & "; allocated=" & _
                         FormatRunNumberLocal(allocatedQty) & "."
                Exit Function
            End If
        End If
NextRequirement:
    Next rawRequirement
    ValidateProcessRequirementsReady = True
End Function

Private Function ValidateProcessAllocationsLive(ByVal node As Object, _
                                                ByVal runLocation As String, _
                                                ByRef report As String) As Boolean
    Dim key As Variant
    Dim nodeId As String
    Dim systemKey As String
    Dim liveQty As Double
    Dim liveLocation As String
    Dim allocatedForEntity As Double
    Dim nonCounted As Boolean

    nodeId = RunRecordText(node, "ProcessNodeId")
    For Each key In mAllocations.Keys
        If Not AllocationBelongsToNode(CStr(key), nodeId) Then GoTo NextAllocation
        systemKey = AllocationSystemKey(CStr(key))
        liveQty = ExactEntityAvailableQty(systemKey, liveLocation)
        nonCounted = ExactEntityIsNonCounted(systemKey)
        allocatedForEntity = AllocationTotalForEntityForNode(systemKey, nodeId)
        If Not nonCounted And liveQty + QTY_TOLERANCE < allocatedForEntity Then
            report = "Stale allocation rejected for System_Key " & systemKey & _
                     ". Available=" & FormatRunNumberLocal(liveQty) & "; allocated=" & _
                     FormatRunNumberLocal(allocatedForEntity) & ". Refresh Production Run."
            Exit Function
        End If
        If Trim$(runLocation) <> "" And _
           StrComp(Trim$(runLocation), liveLocation, vbTextCompare) <> 0 Then
            report = "System_Key " & systemKey & " is at " & liveLocation & _
                     "; the Production run location is " & Trim$(runLocation) & "."
            Exit Function
        End If
NextAllocation:
    Next key
    ValidateProcessAllocationsLive = True
End Function

Private Function AvailableStockForRequirement(ByVal requirement As Object, _
                                              ByVal locationFilter As String, _
                                              Optional ByVal action As cProductionWorksheetAction = Nothing) As Double
    Dim entities As Variant
    Dim entityRow As Long
    Dim nodeId As String
    Dim requirementId As String
    Dim systemKey As String
    Dim availableQty As Double
    Dim reservedOther As Double
    Dim currentAllocation As Double
    Dim allocationId As String
    Dim nativeUom As String
    Dim requirementUom As String
    Dim conversionFactor As Double
    Dim conversionReport As String

    entities = modInventoryDomainBridge.ListAvailableInventoryEntitiesBridge("")
    If Not modProductionRunLoadActions.CanContinue(action) Then Exit Function
    If Not IsArray(entities) Then Exit Function
    nodeId = RunRecordText(requirement, "ProcessNodeId")
    requirementId = RunRecordText(requirement, "RequirementId")
    requirementUom = RunRecordText(requirement, "UOM")
    locationFilter = Trim$(locationFilter)
    For entityRow = LBound(entities, 1) To UBound(entities, 1)
        If locationFilter <> "" Then
            If StrComp(locationFilter, Trim$(CStr(entities(entityRow, 7))), vbTextCompare) <> 0 Then GoTo NextEntity
        End If
        If Not RequirementAllowsItem(nodeId, requirementId, _
                CStr(entities(entityRow, 3)), CStr(entities(entityRow, 2))) Then GoTo NextEntity
        If EntityRowIsNonCounted(entities, entityRow) Then
            AvailableStockForRequirement = ScaledRecordQty(requirement)
            Exit Function
        End If
        systemKey = Trim$(CStr(entities(entityRow, 1)))
        nativeUom = Trim$(CStr(entities(entityRow, 5)))
        conversionFactor = 0#
        conversionReport = ""
        If Not modUomSettings.GetUomConversion(nativeUom, requirementUom, _
                conversionFactor, , conversionReport) Then GoTo NextEntity
        availableQty = 0#
        If IsNumeric(entities(entityRow, 6)) Then availableQty = CDbl(entities(entityRow, 6))
        allocationId = AllocationKey(nodeId, requirementId, systemKey)
        currentAllocation = NativeAllocationQty(allocationId)
        reservedOther = AllocationTotalForEntity(systemKey, "") - currentAllocation
        If availableQty > reservedOther Then _
            AvailableStockForRequirement = AvailableStockForRequirement + _
                (availableQty - reservedOther) * conversionFactor
NextEntity:
    Next entityRow
End Function

Private Function IncomingConnectionForRequirement(ByVal nodeId As String, _
                                                  ByVal requirementId As String) As Object
    Dim rawConnection As Variant
    Dim connection As Object
    For Each rawConnection In mConnections
        Set connection = rawConnection
        If StrComp(RunRecordText(connection, "ToProcessNodeId"), nodeId, vbTextCompare) = 0 _
           And StrComp(RunRecordText(connection, "ToRequirementId"), requirementId, vbTextCompare) = 0 Then
            Set IncomingConnectionForRequirement = connection
            Exit Function
        End If
    Next rawConnection
End Function

Private Function RoutedInputCount(ByVal nodeId As String) As Long
    Dim rawRequirement As Variant
    Dim requirement As Object

    For Each rawRequirement In mRequirements
        Set requirement = rawRequirement
        If nodeId = "" Or StrComp(RunRecordText(requirement, "ProcessNodeId"), nodeId, vbTextCompare) = 0 Then
            If RequirementHasIncomingConnection(RunRecordText(requirement, "ProcessNodeId"), _
                                                 RunRecordText(requirement, "RequirementId")) Then
                RoutedInputCount = RoutedInputCount + 1
            End If
        End If
    Next rawRequirement
End Function

Private Function BuildRoutedInputCheckRow(ByVal requirement As Object, _
                                           ByVal entityRows As Variant) As Variant
    Dim connection As Object
    Dim downstreamNode As Object
    Dim sourceNode As Object
    Dim sourceOutput As Object
    Dim sourceKey As String
    Dim result(1 To 9) As Variant

    Set connection = IncomingConnectionForRequirement( _
        RunRecordText(requirement, "ProcessNodeId"), RunRecordText(requirement, "RequirementId"))
    If connection Is Nothing Then Exit Function
    Set downstreamNode = FindNode(RunRecordText(requirement, "ProcessNodeId"))
    Set sourceNode = FindNode(RunRecordText(connection, "FromProcessNodeId"))
    Set sourceOutput = FindOutput(RunRecordText(connection, "FromProcessNodeId"), _
                                  RunRecordText(connection, "FromOutputId"))
    sourceKey = OutputKeyForConnection(connection)

    result(1) = "ROUTED"
    result(2) = RunRecordText(downstreamNode, "ProcessName") & _
                " / " & RunRecordText(requirement, "RequirementName")
    result(3) = RunRecordText(sourceNode, "ProcessName") & " / " & _
                RunRecordText(sourceOutput, "OutputName")
    result(4) = sourceKey
    result(5) = ExactEntityField(entityRows, sourceKey, 3)
    result(6) = ExactEntityField(entityRows, sourceKey, 4)
    result(7) = RunRecordText(sourceOutput, "UOM")
    result(8) = ScaledConnectionQty(connection)
    result(9) = ExactEntityInventoryDisplayForKey(entityRows, sourceKey)
    BuildRoutedInputCheckRow = result
End Function

Private Function OutputKeyForConnection(ByVal connection As Object) As String
    Dim key As String
    If connection Is Nothing Then Exit Function
    key = OutputIdentityKey(RunRecordText(connection, "FromProcessNodeId"), _
                            RunRecordText(connection, "FromOutputId"))
    If Not mOutputKeys Is Nothing Then
        If mOutputKeys.Exists(key) Then OutputKeyForConnection = CStr(mOutputKeys(key))
    End If
End Function

Private Function AllocationBelongsToNode(ByVal allocationId As String, _
                                         ByVal nodeId As String) As Boolean
    Dim parts() As String
    parts = Split(allocationId, vbTab)
    If UBound(parts) = 2 Then _
        AllocationBelongsToNode = _
            (StrComp(parts(0), nodeId, vbTextCompare) = 0)
End Function

Private Function AllocationCountForNode(ByVal nodeId As String) As Long
    Dim key As Variant
    For Each key In mAllocations.Keys
        If AllocationBelongsToNode(CStr(key), nodeId) Then _
            AllocationCountForNode = AllocationCountForNode + 1
    Next key
End Function

Private Function AllocationTotalForEntityForNode(ByVal systemKey As String, _
                                                 ByVal nodeId As String) As Double
    Dim key As Variant
    For Each key In mAllocations.Keys
        If AllocationBelongsToNode(CStr(key), nodeId) Then
            If StrComp(AllocationSystemKey(CStr(key)), systemKey, vbTextCompare) = 0 Then _
                AllocationTotalForEntityForNode = AllocationTotalForEntityForNode + NativeAllocationQty(CStr(key))
        End If
    Next key
End Function

Private Sub AddNodeInExecutionOrder(ByVal node As Object)
    Dim i As Long
    Dim existing As Object
    For i = 1 To mNodes.Count
        Set existing = mNodes(i)
        If RunRecordNumber(node, "ExecutionOrdinal") < RunRecordNumber(existing, "ExecutionOrdinal") Then
            mNodes.Add node, Before:=i
            Exit Sub
        End If
    Next i
    mNodes.Add node
End Sub

Private Function BuildNodeConsumeItems(ByVal node As Object, ByVal runLocation As String, _
                                       ByVal runId As String) As Collection
    Dim items As New Collection
    Dim rawRequirement As Variant
    Dim requirement As Object
    Dim rawConnection As Variant
    Dim connection As Object
    Dim key As Variant
    Dim allocationParts() As String
    Dim item As Object
    Dim sourceOutput As Object
    Dim qty As Double
    Dim systemKey As String
    Dim entityRows As Variant

    entityRows = modInventoryDomainBridge.ListAvailableInventoryEntitiesBridge("")
    For Each rawRequirement In mRequirements
        Set requirement = rawRequirement
        If StrComp(RunRecordText(requirement, "ProcessNodeId"), _
                   RunRecordText(node, "ProcessNodeId"), vbTextCompare) <> 0 Then GoTo NextRequirement
        If RequirementHasIncomingConnection(RunRecordText(node, "ProcessNodeId"), _
                                            RunRecordText(requirement, "RequirementId")) Then GoTo NextRequirement
        For Each key In mAllocations.Keys
            allocationParts = Split(CStr(key), vbTab)
            If UBound(allocationParts) = 2 Then
                If StrComp(allocationParts(0), RunRecordText(node, "ProcessNodeId"), vbTextCompare) = 0 _
                   And StrComp(allocationParts(1), RunRecordText(requirement, "RequirementId"), vbTextCompare) = 0 Then
                    qty = NativeAllocationQty(CStr(key))
                    Set item = modProductionJson.CreateProductionDeltaPayloadItem( _
                        allocationParts(2), CStr(ExactEntityField(entityRows, allocationParts(2), 2)), _
                        qty, CStr(ExactEntityField(entityRows, allocationParts(2), 7)), _
                        RunItemNote(runId, node, RunRecordText(requirement, "RequirementId")), "USED")
                    item("RequirementQty") = CDbl(mAllocations(key))
                    item("RequirementUOM") = RunRecordText(requirement, "UOM")
                    item("StockUOM") = CStr(ExactEntityField(entityRows, allocationParts(2), 5))
                    If Not mAllocationConversionAudit Is Nothing Then
                        If mAllocationConversionAudit.Exists(CStr(key)) Then
                            item("ConversionCatalogVersion") = Split(CStr(mAllocationConversionAudit(CStr(key))), vbTab)(0)
                            item("ConversionFactor") = CDbl(Split(CStr(mAllocationConversionAudit(CStr(key))), vbTab)(1))
                        End If
                    End If
                    items.Add item
                End If
            End If
        Next key
NextRequirement:
    Next rawRequirement
    For Each rawConnection In mConnections
        Set connection = rawConnection
        If StrComp(RunRecordText(connection, "ToProcessNodeId"), _
                   RunRecordText(node, "ProcessNodeId"), vbTextCompare) = 0 Then
            Set sourceOutput = FindOutput(RunRecordText(connection, "FromProcessNodeId"), _
                                          RunRecordText(connection, "FromOutputId"))
            qty = ScaledConnectionQty(connection)
            systemKey = CStr(mOutputKeys(OutputIdentityKey( _
                RunRecordText(connection, "FromProcessNodeId"), _
                RunRecordText(connection, "FromOutputId"))))
            Set item = modProductionJson.CreateProductionDeltaPayloadItem( _
                systemKey, RunRecordText(sourceOutput, "ITEM_CODE"), qty, runLocation, _
                RunItemNote(runId, node, RunRecordText(connection, "ToRequirementId")), "USED")
            items.Add item
        End If
    Next rawConnection
    Set BuildNodeConsumeItems = items
End Function

Private Function BuildNodeCompleteItems(ByVal node As Object, ByVal runLocation As String, _
                                        ByVal runId As String) As Collection
    Dim items As New Collection
    Dim rawOutput As Variant
    Dim output As Object
    Dim item As Object
    Dim outputKey As String
    Dim attributesJson As String

    For Each rawOutput In mOutputs
        Set output = rawOutput
        If StrComp(RunRecordText(output, "ProcessNodeId"), _
                   RunRecordText(node, "ProcessNodeId"), vbTextCompare) = 0 Then
            outputKey = CStr(mOutputKeys(OutputIdentityKey( _
                RunRecordText(node, "ProcessNodeId"), RunRecordText(output, "OutputId"))))
            attributesJson = "{""RunId"":""" & runId & """,""RecipeId"":""" & _
                             mRecipeId & """,""RecipeVersion"":""" & mRecipeVersion & _
                             """,""ProcessId"":""" & RunRecordText(node, "ProcessId") & _
                             """,""ProcessVersion"":""" & RunRecordText(node, "ProcessVersion") & _
                             """,""OutputId"":""" & RunRecordText(output, "OutputId") & _
                             """,""PlannedQty"":" & JsonRunNumber(ScaledRecordQty(output)) & _
                             ",""ActualQty"":" & JsonRunNumber(ActualOutputQty(output)) & "}"
            Set item = modProductionJson.CreateProductionInventoryEntityPayloadItem( _
                outputKey, RunRecordText(output, "ITEM_CODE"), ActualOutputQty(output), _
                runLocation, "GOOD", attributesJson, RunItemNote(runId, node, RunRecordText(output, "OutputId")))
            item("IoType") = "MADE"
            item("ITEM_CODE") = RunRecordText(output, "ITEM_CODE")
            item("ITEM") = RunRecordText(output, "OutputName")
            item("UOM") = RunRecordText(output, "UOM")
            items.Add item
        End If
    Next rawOutput
    Set BuildNodeCompleteItems = items
End Function

Private Sub AssignFreshOutputKeys()
    Dim rawOutput As Variant
    Dim output As Object
    Dim key As String
    Set mOutputKeys = NewTextDictionary()
    For Each rawOutput In mOutputs
        Set output = rawOutput
        key = OutputIdentityKey(RunRecordText(output, "ProcessNodeId"), RunRecordText(output, "OutputId"))
        mOutputKeys.Add key, modRoleEventWriter.CreateSystemKey()
    Next rawOutput
End Sub

Private Sub AssignFreshOutputKeysForNode(ByVal nodeId As String)
    Dim rawOutput As Variant
    Dim output As Object
    Dim key As String
    For Each rawOutput In mOutputs
        Set output = rawOutput
        If StrComp(RunRecordText(output, "ProcessNodeId"), nodeId, vbTextCompare) = 0 Then
            key = OutputIdentityKey(nodeId, RunRecordText(output, "OutputId"))
            If Not mOutputKeys.Exists(key) Then _
                mOutputKeys.Add key, modRoleEventWriter.CreateSystemKey()
        End If
    Next rawOutput
End Sub

Private Function VerifyCompletedOutputBalances(ByRef report As String) As Boolean
    Dim rawOutput As Variant
    Dim output As Object
    Dim key As String
    Dim expectedQty As Double
    Dim actualQty As Double

    For Each rawOutput In mOutputs
        Set output = rawOutput
        key = CStr(mOutputKeys(OutputIdentityKey(RunRecordText(output, "ProcessNodeId"), _
                                                RunRecordText(output, "OutputId"))))
        expectedQty = ActualOutputQty(output) - OutgoingQtyForOutput( _
            RunRecordText(output, "ProcessNodeId"), RunRecordText(output, "OutputId")) - _
            AllocationTotalForEntity(key, "")
        actualQty = ExactEntityAvailableQty(key)
        If Abs(actualQty - expectedQty) > QTY_TOLERANCE Then
            report = "Output persistence verification failed for System_Key " & key & _
                     ". Expected=" & FormatRunNumberLocal(expectedQty) & "; actual=" & _
                     FormatRunNumberLocal(actualQty) & "."
            Exit Function
        End If
    Next rawOutput
    VerifyCompletedOutputBalances = True
End Function

Private Function VerifyNodeOutputBalances(ByVal nodeId As String, _
                                          ByVal downstreamConsumed As Boolean, _
                                          ByRef report As String) As Boolean
    Dim rawOutput As Variant
    Dim output As Object
    Dim key As String
    Dim expectedQty As Double
    Dim actualQty As Double

    For Each rawOutput In mOutputs
        Set output = rawOutput
        If StrComp(RunRecordText(output, "ProcessNodeId"), nodeId, vbTextCompare) <> 0 Then GoTo NextOutput
        key = CStr(mOutputKeys(OutputIdentityKey(nodeId, RunRecordText(output, "OutputId"))))
        expectedQty = ActualOutputQty(output)
        If downstreamConsumed Then expectedQty = expectedQty - _
            OutgoingQtyForOutput(nodeId, RunRecordText(output, "OutputId"))
        actualQty = ExactEntityAvailableQty(key)
        If Abs(actualQty - expectedQty) > QTY_TOLERANCE Then
            report = "Output persistence verification failed for System_Key " & key & _
                     ". Expected=" & FormatRunNumberLocal(expectedQty) & "; actual=" & _
                     FormatRunNumberLocal(actualQty) & "."
            Exit Function
        End If
NextOutput:
    Next rawOutput
    VerifyNodeOutputBalances = True
End Function

Private Function ValidateReusableActualOutputs(ByRef report As String) As Boolean
    Dim rawOutput As Variant
    Dim output As Object
    Dim outputKey As String
    Dim actualQty As Double

    For Each rawOutput In mOutputs
        Set output = rawOutput
        outputKey = OutputIdentityKey(RunRecordText(output, "ProcessNodeId"), _
                                      RunRecordText(output, "OutputId"))
        If Not mActualOutputQty.Exists(outputKey) Then
            report = "Enter Actual Output for " & RunRecordText(output, "OutputName") & _
                     " before completing the run."
            Exit Function
        End If
        actualQty = ActualOutputQty(output)
        If actualQty <= 0 Then
            report = "Actual Output for " & RunRecordText(output, "OutputName") & _
                     " must be greater than zero."
            Exit Function
        End If
        If Not ValidateReusableActualOutput(output, actualQty, report) Then Exit Function
    Next rawOutput
    ValidateReusableActualOutputs = True
End Function

Private Function ValidateReusableActualOutputsForNode(ByVal nodeId As String, _
                                                      ByRef report As String) As Boolean
    Dim rawOutput As Variant
    Dim output As Object
    Dim outputKey As String
    Dim actualQty As Double
    Dim foundOutput As Boolean

    For Each rawOutput In mOutputs
        Set output = rawOutput
        If StrComp(RunRecordText(output, "ProcessNodeId"), nodeId, vbTextCompare) <> 0 Then GoTo NextOutput
        foundOutput = True
        outputKey = OutputIdentityKey(nodeId, RunRecordText(output, "OutputId"))
        If Not mActualOutputQty.Exists(outputKey) Then
            report = "Enter Actual Output for " & RunRecordText(output, "OutputName") & _
                     " before completing the selected Process."
            Exit Function
        End If
        actualQty = ActualOutputQty(output)
        If actualQty <= 0 Then
            report = "Actual Output for " & RunRecordText(output, "OutputName") & _
                     " must be greater than zero."
            Exit Function
        End If
        If Not ValidateReusableActualOutput(output, actualQty, report) Then Exit Function
NextOutput:
    Next rawOutput
    If Not foundOutput Then
        report = "The selected Process has no output to complete."
        Exit Function
    End If
    ValidateReusableActualOutputsForNode = True
End Function

Private Function ValidateReusableActualOutput(ByVal output As Object, _
                                              ByVal actualQty As Double, _
                                              ByRef report As String) As Boolean
    Dim committedQty As Double
    Dim effectiveFloor As Double
    Dim ceilingQty As Double

    committedQty = OutgoingQtyForOutput(RunRecordText(output, "ProcessNodeId"), _
                                        RunRecordText(output, "OutputId"))
    If actualQty + QTY_TOLERANCE < committedQty Then
        report = "Actual Output for " & RunRecordText(output, "OutputName") & _
                 " is below its routed downstream commitment. Actual=" & _
                 FormatRunNumberLocal(actualQty) & "; committed=" & _
                 FormatRunNumberLocal(committedQty) & "."
        Exit Function
    End If
    If OutputRegulationEnabled(output) Then
        effectiveFloor = EffectiveOutputRegulationFloor(output)
        ceilingQty = ScaledOutputRegulationCeiling(output)
        If actualQty + QTY_TOLERANCE < effectiveFloor Then
            report = "Actual Output for " & RunRecordText(output, "OutputName") & _
                     " is below its regulated floor. Actual=" & FormatRunNumberLocal(actualQty) & _
                     "; floor=" & FormatRunNumberLocal(effectiveFloor) & "."
            Exit Function
        End If
        If actualQty - QTY_TOLERANCE > ceilingQty Then
            report = "Actual Output for " & RunRecordText(output, "OutputName") & _
                     " is above its regulated ceiling. Actual=" & FormatRunNumberLocal(actualQty) & _
                     "; ceiling=" & FormatRunNumberLocal(ceilingQty) & "."
            Exit Function
        End If
    End If
    ValidateReusableActualOutput = True
End Function

Private Function OutputRegulationEnabled(ByVal output As Object) As Boolean
    Dim valueIn As Variant
    If output Is Nothing Then Exit Function
    If Not output.Exists("OutputRegulationEnabled") Then Exit Function
    valueIn = output("OutputRegulationEnabled")
    If VarType(valueIn) = vbBoolean Then
        OutputRegulationEnabled = CBool(valueIn)
    Else
        OutputRegulationEnabled = (StrComp(Trim$(CStr(valueIn)), "true", vbTextCompare) = 0 _
            Or Trim$(CStr(valueIn)) = "1")
    End If
End Function

Private Function EffectiveOutputRegulationFloor(ByVal output As Object) As Double
    Dim floorQty As Double
    Dim committedQty As Double
    floorQty = RunRecordNumber(output, "OutputFloorQty") * mScalePercent / 100#
    committedQty = OutgoingQtyForOutput(RunRecordText(output, "ProcessNodeId"), _
                                        RunRecordText(output, "OutputId"))
    If committedQty > floorQty Then floorQty = committedQty
    EffectiveOutputRegulationFloor = floorQty
End Function

Private Function ScaledOutputRegulationCeiling(ByVal output As Object) As Double
    ScaledOutputRegulationCeiling = RunRecordNumber(output, "OutputCeilingQty") * _
                                   mScalePercent / 100#
End Function

Private Function ApplyRecipeOutputRegulationOverrides(ByRef report As String) As Boolean
    Dim rawRegulation As Variant
    Dim regulation As Object
    Dim output As Object
    Dim overrideKey As String

    For Each rawRegulation In mOutputRegulations
        Set regulation = rawRegulation
        Set output = FindOutput(RunRecordText(regulation, "ProcessNodeId"), _
                                RunRecordText(regulation, "OutputId"))
        If output Is Nothing Then
            report = "Recipe output regulation references an output that is not declared by its pinned Process version."
            Exit Function
        End If
        If StrComp(RunRecordText(regulation, "ProcessId"), RunRecordText(output, "ProcessId"), vbTextCompare) <> 0 _
           Or StrComp(RunRecordText(regulation, "ProcessVersion"), RunRecordText(output, "ProcessVersion"), vbTextCompare) <> 0 Then
            report = "Recipe output regulation Process identity does not match its output."
            Exit Function
        End If
        output("OutputRegulationEnabled") = RunRecordValue(regulation, "OutputRegulationEnabled")
        output("OutputFloorQty") = RunRecordValue(regulation, "OutputFloorQty")
        output("OutputCeilingQty") = RunRecordValue(regulation, "OutputCeilingQty")
    Next rawRegulation
    ApplyRecipeOutputRegulationOverrides = True
End Function

Private Function ActualOutputQty(ByVal output As Object) As Double
    Dim outputKey As String

    outputKey = OutputIdentityKey(RunRecordText(output, "ProcessNodeId"), _
                                  RunRecordText(output, "OutputId"))
    If Not mActualOutputQty Is Nothing Then
        If mActualOutputQty.Exists(outputKey) Then ActualOutputQty = CDbl(mActualOutputQty(outputKey))
    End If
End Function

Private Sub CaptureLastActualOutputs()
    Dim key As Variant

    Set mLastOutputQty = NewTextDictionary()
    For Each key In mActualOutputQty.Keys
        mLastOutputQty(CStr(key)) = CDbl(mActualOutputQty(key))
    Next key
End Sub

Private Sub CaptureCompletedOutputHistory()
    Dim rawOutput As Variant
    Dim output As Object
    Dim record As Object
    Dim outputKey As String

    For Each rawOutput In mOutputs
        Set output = rawOutput
        outputKey = OutputIdentityKey(RunRecordText(output, "ProcessNodeId"), _
                                      RunRecordText(output, "OutputId"))
        Set record = NewTextDictionary()
        record("ProcessName") = RunRecordText(output, "ProcessName")
        record("OutputName") = RunRecordText(output, "OutputName")
        record("UOM") = RunRecordText(output, "UOM")
        record("ActualQty") = ActualOutputQty(output)
        record("BatchNumber") = mBatchNumber
        record("UsedGoods") = ProcessUsedGoodsDisplay( _
            RunRecordText(output, "ProcessNodeId"))
        record("Recall") = mRecipeId & "-B" & Format$(mBatchNumber, "0000")
        record("System_Key") = CStr(mOutputKeys(outputKey))
        mOutputHistory.Add record
        record("ProcessTotal") = ProcessOutputTotal( _
            RunRecordText(output, "ProcessName"), RunRecordText(output, "OutputName"), _
            RunRecordText(output, "UOM"))
    Next rawOutput
End Sub

Private Sub CaptureCompletedNodeOutputHistory(ByVal nodeId As String)
    Dim rawOutput As Variant
    Dim output As Object
    Dim record As Object
    Dim outputKey As String

    For Each rawOutput In mOutputs
        Set output = rawOutput
        If StrComp(RunRecordText(output, "ProcessNodeId"), nodeId, vbTextCompare) <> 0 Then GoTo NextOutput
        outputKey = OutputIdentityKey(nodeId, RunRecordText(output, "OutputId"))
        Set record = NewTextDictionary()
        record("ProcessName") = RunRecordText(output, "ProcessName")
        record("OutputName") = RunRecordText(output, "OutputName")
        record("UOM") = RunRecordText(output, "UOM")
        record("ActualQty") = ActualOutputQty(output)
        record("BatchNumber") = mBatchNumber
        record("UsedGoods") = ProcessUsedGoodsDisplay(nodeId)
        record("Recall") = mRecipeId & "-B" & Format$(mBatchNumber, "0000")
        record("System_Key") = CStr(mOutputKeys(outputKey))
        mOutputHistory.Add record
        record("ProcessTotal") = ProcessOutputTotal( _
            RunRecordText(output, "ProcessName"), RunRecordText(output, "OutputName"), _
            RunRecordText(output, "UOM"))
NextOutput:
    Next rawOutput
End Sub

Private Function ProcessUsedGoodsDisplay(ByVal nodeId As String) As String
    Dim groups As Object

    Set groups = UsedGoodsByNormalizedUom(nodeId)
    ProcessUsedGoodsDisplay = FormatUsedGoodsGroups(groups)
End Function

Public Function ReusableRunUsedGoodsDisplayForTest(ByVal firstUom As String, _
                                                    ByVal firstQty As Double, _
                                                    ByVal secondUom As String, _
                                                    ByVal secondQty As Double) As String
    Dim groups As Object

    Set groups = NewTextDictionary()
    groups(UCase$(Trim$(firstUom))) = firstQty
    groups(UCase$(Trim$(secondUom))) = secondQty
    ReusableRunUsedGoodsDisplayForTest = FormatUsedGoodsGroups(groups)
End Function

Private Function FormatUsedGoodsGroups(ByVal groups As Object) As String
    Dim keys As Variant
    Dim i As Long
    Dim j As Long
    Dim swapValue As Variant

    If groups Is Nothing Or groups.Count = 0 Then Exit Function
    keys = groups.Keys
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If StrComp(CStr(keys(i)), CStr(keys(j)), vbTextCompare) > 0 Then
                swapValue = keys(i)
                keys(i) = keys(j)
                keys(j) = swapValue
            End If
        Next j
    Next i
    For i = LBound(keys) To UBound(keys)
        If FormatUsedGoodsGroups <> "" Then FormatUsedGoodsGroups = FormatUsedGoodsGroups & "; "
        FormatUsedGoodsGroups = FormatUsedGoodsGroups & _
            FormatUsedGoodsQuantity(CDbl(groups(CStr(keys(i))))) & " " & CStr(keys(i))
    Next i
End Function

Private Function FormatUsedGoodsQuantity(ByVal quantity As Double) As String
    Dim decimalSeparator As String

    FormatUsedGoodsQuantity = FormatRunNumberLocal(quantity)
    decimalSeparator = CStr(Application.International(xlDecimalSeparator))
    If decimalSeparator <> "" Then
        If Right$(FormatUsedGoodsQuantity, Len(decimalSeparator)) = decimalSeparator Then _
            FormatUsedGoodsQuantity = Left$(FormatUsedGoodsQuantity, _
                                            Len(FormatUsedGoodsQuantity) - Len(decimalSeparator))
    End If
End Function

Private Function UsedGoodsByNormalizedUom(ByVal nodeId As String) As Object
    Dim rawRequirement As Variant
    Dim requirement As Object
    Dim requirementUom As String
    Dim groups As Object

    Set groups = NewTextDictionary()

    For Each rawRequirement In mRequirements
        Set requirement = rawRequirement
        If StrComp(RunRecordText(requirement, "ProcessNodeId"), nodeId, vbTextCompare) = 0 Then
            requirementUom = UCase$(Trim$(RunRecordText(requirement, "UOM")))
            If requirementUom <> "" Then
                If Not groups.Exists(requirementUom) Then groups.Add requirementUom, 0#
                groups(requirementUom) = CDbl(groups(requirementUom)) + ScaledRecordQty(requirement)
            End If
        End If
    Next rawRequirement
    Set UsedGoodsByNormalizedUom = groups
End Function

Private Function ProcessOutputTotal(ByVal processName As String, _
                                    ByVal outputName As String, _
                                    ByVal uom As String) As Double
    Dim rawHistory As Variant
    Dim history As Object

    For Each rawHistory In mOutputHistory
        Set history = rawHistory
        If StrComp(RunRecordText(history, "ProcessName"), processName, vbTextCompare) = 0 _
           And StrComp(RunRecordText(history, "OutputName"), outputName, vbTextCompare) = 0 _
           And StrComp(RunRecordText(history, "UOM"), uom, vbTextCompare) = 0 Then
            ProcessOutputTotal = ProcessOutputTotal + RunRecordNumber(history, "ActualQty")
        End If
    Next rawHistory
End Function

Private Function OutgoingQtyForOutput(ByVal nodeId As String, ByVal outputId As String) As Double
    Dim rawConnection As Variant
    Dim connection As Object
    For Each rawConnection In mConnections
        Set connection = rawConnection
        If StrComp(RunRecordText(connection, "FromProcessNodeId"), nodeId, vbTextCompare) = 0 _
           And StrComp(RunRecordText(connection, "FromOutputId"), outputId, vbTextCompare) = 0 Then
            OutgoingQtyForOutput = OutgoingQtyForOutput + ScaledConnectionQty(connection)
        End If
    Next rawConnection
End Function

Private Function FindNode(ByVal nodeId As String) As Object
    Dim rawRecord As Variant
    Dim record As Object
    For Each rawRecord In mNodes
        Set record = rawRecord
        If StrComp(RunRecordText(record, "ProcessNodeId"), nodeId, vbTextCompare) = 0 Then
            Set FindNode = record
            Exit Function
        End If
    Next rawRecord
End Function

Private Function FindNodeByProcessName(ByVal processName As String) As Object
    Dim rawRecord As Variant
    Dim record As Object
    processName = Trim$(processName)
    For Each rawRecord In mNodes
        Set record = rawRecord
        If StrComp(RunRecordText(record, "ProcessName"), processName, vbTextCompare) = 0 Then
            Set FindNodeByProcessName = record
            Exit Function
        End If
    Next rawRecord
End Function

Private Function NodeIsComplete(ByVal nodeId As String) As Boolean
    If mCompletedNodes Is Nothing Then Exit Function
    NodeIsComplete = mCompletedNodes.Exists(nodeId)
End Function

Private Function AllNodesCompleted() As Boolean
    If mCompletedNodes Is Nothing Then Exit Function
    AllNodesCompleted = (mNodes.Count > 0 And mCompletedNodes.Count = mNodes.Count)
End Function

Private Function ActiveOutputCount() As Long
    Dim rawOutput As Variant
    Dim output As Object
    For Each rawOutput In mOutputs
        Set output = rawOutput
        If Not NodeIsComplete(RunRecordText(output, "ProcessNodeId")) Then _
            ActiveOutputCount = ActiveOutputCount + 1
    Next rawOutput
End Function

Private Function FindRequirement(ByVal nodeId As String, ByVal requirementId As String) As Object
    Dim rawRecord As Variant
    Dim record As Object
    For Each rawRecord In mRequirements
        Set record = rawRecord
        If StrComp(RunRecordText(record, "ProcessNodeId"), nodeId, vbTextCompare) = 0 _
           And StrComp(RunRecordText(record, "RequirementId"), requirementId, vbTextCompare) = 0 Then
            Set FindRequirement = record
            Exit Function
        End If
    Next rawRecord
End Function

Private Function FindOutput(ByVal nodeId As String, ByVal outputId As String) As Object
    Dim rawRecord As Variant
    Dim record As Object
    For Each rawRecord In mOutputs
        Set record = rawRecord
        If StrComp(RunRecordText(record, "ProcessNodeId"), nodeId, vbTextCompare) = 0 _
           And StrComp(RunRecordText(record, "OutputId"), outputId, vbTextCompare) = 0 Then
            Set FindOutput = record
            Exit Function
        End If
    Next rawRecord
End Function

Private Function RequirementHasIncomingConnection(ByVal nodeId As String, _
                                                  ByVal requirementId As String) As Boolean
    Dim rawRecord As Variant
    Dim record As Object
    For Each rawRecord In mConnections
        Set record = rawRecord
        If StrComp(RunRecordText(record, "ToProcessNodeId"), nodeId, vbTextCompare) = 0 _
           And StrComp(RunRecordText(record, "ToRequirementId"), requirementId, vbTextCompare) = 0 Then
            RequirementHasIncomingConnection = True
            Exit Function
        End If
    Next rawRecord
End Function

Private Function RequirementHasAlternative(ByVal nodeId As String, _
                                           ByVal requirementId As String) As Boolean
    Dim rawRecord As Variant
    Dim record As Object
    For Each rawRecord In mAlternatives
        Set record = rawRecord
        If StrComp(RunRecordText(record, "ProcessNodeId"), nodeId, vbTextCompare) = 0 _
           And StrComp(RunRecordText(record, "RequirementId"), requirementId, vbTextCompare) = 0 _
           And RunRecordText(record, "ITEM_CODE") <> "" Then
            RequirementHasAlternative = True
            Exit Function
        End If
    Next rawRecord
End Function

Private Function RequirementAllowsItem(ByVal nodeId As String, ByVal requirementId As String, _
                                       ByVal itemCode As String, ByVal sku As String) As Boolean
    Dim rawRecord As Variant
    Dim record As Object
    Dim allowedCode As String
    For Each rawRecord In mAlternatives
        Set record = rawRecord
        If StrComp(RunRecordText(record, "ProcessNodeId"), nodeId, vbTextCompare) = 0 _
           And StrComp(RunRecordText(record, "RequirementId"), requirementId, vbTextCompare) = 0 Then
            allowedCode = RunRecordText(record, "ITEM_CODE")
            If StrComp(allowedCode, itemCode, vbTextCompare) = 0 _
               Or StrComp(allowedCode, sku, vbTextCompare) = 0 Then
                RequirementAllowsItem = True
                Exit Function
            End If
        End If
    Next rawRecord
End Function

Private Function ScaledRecordQty(ByVal record As Object) As Double
    Dim qty As Double
    Dim pct As Double
    Dim yieldBasis As Double
    qty = RunRecordNumber(record, "Qty")
    pct = RunRecordNumber(record, "Percent")
    yieldBasis = RunRecordNumber(record, "YieldBasis")
    If qty > 0 Then
        ScaledRecordQty = qty * mScalePercent / 100#
    ElseIf pct > 0 And yieldBasis > 0 Then
        ScaledRecordQty = yieldBasis * pct * mScalePercent / 10000#
    End If
End Function

Private Function ScaledConnectionQty(ByVal connection As Object) As Double
    Dim qty As Double
    Dim pct As Double
    Dim requirement As Object
    qty = RunRecordNumber(connection, "Qty")
    pct = RunRecordNumber(connection, "Percent")
    If qty > 0 Then
        ScaledConnectionQty = qty * mScalePercent / 100#
    ElseIf pct > 0 Then
        Set requirement = FindRequirement(RunRecordText(connection, "ToProcessNodeId"), _
                                          RunRecordText(connection, "ToRequirementId"))
        If Not requirement Is Nothing Then ScaledConnectionQty = ScaledRecordQty(requirement) * pct / 100#
    End If
End Function

Private Function AllocationKey(ByVal nodeId As String, ByVal requirementId As String, _
                               ByVal systemKey As String) As String
    AllocationKey = nodeId & vbTab & requirementId & vbTab & systemKey
End Function

Private Function AllocationSystemKey(ByVal allocationKeyValue As String) As String
    Dim parts() As String
    parts = Split(allocationKeyValue, vbTab)
    If UBound(parts) >= 2 Then AllocationSystemKey = parts(2)
End Function

Private Function AllocationQty(ByVal nodeId As String, ByVal requirementId As String, _
                               ByVal systemKey As String) As Variant
    Dim key As String
    key = AllocationKey(nodeId, requirementId, systemKey)
    If Not mAllocations Is Nothing Then
        If mAllocations.Exists(key) Then AllocationQty = mAllocations(key)
    End If
End Function

Private Function AllocationPercent(ByVal nodeId As String, ByVal requirementId As String, _
                                   ByVal systemKey As String) As Variant
    Dim requirement As Object
    Dim qty As Variant
    Dim requiredQty As Double
    qty = AllocationQty(nodeId, requirementId, systemKey)
    If IsEmpty(qty) Then Exit Function
    Set requirement = FindRequirement(nodeId, requirementId)
    If requirement Is Nothing Then Exit Function
    requiredQty = ScaledRecordQty(requirement)
    If requiredQty > 0 Then AllocationPercent = CDbl(qty) / requiredQty * 100#
End Function

Private Function StockBucketKey(ByVal sku As String, ByVal uom As String, _
                                ByVal locationValue As String, _
                                ByVal conditionValue As String) As String
    StockBucketKey = UCase$(Trim$(sku)) & vbTab & UCase$(Trim$(uom)) & vbTab & _
                     UCase$(Trim$(locationValue)) & vbTab & UCase$(Trim$(conditionValue))
End Function

Private Function StockBucketKeyForEntity(ByVal entities As Variant, _
                                         ByVal entityRow As Long) As String
    Dim conditionValue As String

    If UBound(entities, 2) >= 8 Then conditionValue = CStr(entities(entityRow, 8))
    StockBucketKeyForEntity = StockBucketKey(CStr(entities(entityRow, 2)), _
        CStr(entities(entityRow, 5)), CStr(entities(entityRow, 7)), conditionValue)
End Function

Private Function FindExactEntityRow(ByVal entities As Variant, _
                                    ByVal systemKey As String) As Long
    Dim entityRow As Long

    If Not IsArray(entities) Then Exit Function
    For entityRow = LBound(entities, 1) To UBound(entities, 1)
        If StrComp(Trim$(CStr(entities(entityRow, 1))), Trim$(systemKey), _
                   vbTextCompare) = 0 Then
            FindExactEntityRow = entityRow
            Exit Function
        End If
    Next entityRow
End Function

Private Function EntityRowIsNonCounted(ByVal entities As Variant, _
                                       ByVal entityRow As Long) As Boolean
    Dim trackQty As String
    Dim itemKind As String
    Dim categoryValue As String

    If UBound(entities, 2) >= 11 Then trackQty = UCase$(Trim$(CStr(entities(entityRow, 11))))
    If UBound(entities, 2) >= 12 Then itemKind = UCase$(Trim$(CStr(entities(entityRow, 12))))
    If UBound(entities, 2) >= 13 Then categoryValue = UCase$(Trim$(CStr(entities(entityRow, 13))))
    EntityRowIsNonCounted = (trackQty = "FALSE" Or trackQty = "NO" Or trackQty = "0" _
        Or itemKind = "UTILITY" Or itemKind = "SERVICE" Or itemKind = "NON_COUNTED" _
        Or categoryValue = "UTILITY" Or categoryValue = "SERVICE")
End Function

Private Function StockBucketAllocatedQty(ByVal entities As Variant, ByVal nodeId As String, _
                                         ByVal requirementId As String, _
                                         ByVal bucketKey As String) As Double
    Dim entityRow As Long
    Dim qty As Variant

    For entityRow = LBound(entities, 1) To UBound(entities, 1)
        If StrComp(StockBucketKeyForEntity(entities, entityRow), bucketKey, vbTextCompare) = 0 Then
            qty = AllocationQty(nodeId, requirementId, CStr(entities(entityRow, 1)))
            If IsNumeric(qty) Then StockBucketAllocatedQty = StockBucketAllocatedQty + CDbl(qty)
        End If
    Next entityRow
End Function

Private Function StockBucketAvailableDisplay(ByVal totalQty As Double, _
                                             ByVal nonCounted As Boolean, _
                                             ByVal nonCountedLabel As String) As Variant
    If nonCounted Then
        StockBucketAvailableDisplay = nonCountedLabel
    Else
        StockBucketAvailableDisplay = totalQty
    End If
End Function

Private Function AllocationTotalForRequirement(ByVal nodeId As String, ByVal requirementId As String, _
                                               ByVal excludedKey As String) As Double
    Dim key As Variant
    Dim parts() As String
    For Each key In mAllocations.Keys
        If StrComp(CStr(key), excludedKey, vbTextCompare) <> 0 Then
            parts = Split(CStr(key), vbTab)
            If UBound(parts) = 2 Then
                If StrComp(parts(0), nodeId, vbTextCompare) = 0 _
                   And StrComp(parts(1), requirementId, vbTextCompare) = 0 Then
                    AllocationTotalForRequirement = AllocationTotalForRequirement + CDbl(mAllocations(key))
                End If
            End If
        End If
    Next key
End Function

Private Function AllocationTotalForEntity(ByVal systemKey As String, ByVal excludedKey As String) As Double
    Dim key As Variant
    For Each key In mAllocations.Keys
        If StrComp(CStr(key), excludedKey, vbTextCompare) <> 0 _
           And StrComp(AllocationSystemKey(CStr(key)), systemKey, vbTextCompare) = 0 Then
            AllocationTotalForEntity = AllocationTotalForEntity + NativeAllocationQty(CStr(key))
        End If
    Next key
End Function

Private Function NativeAllocationQty(ByVal allocationId As String) As Double
    If Not mAllocationNativeQuantities Is Nothing Then
        If mAllocationNativeQuantities.Exists(allocationId) Then
            NativeAllocationQty = CDbl(mAllocationNativeQuantities(allocationId))
            Exit Function
        End If
    End If
    If Not mAllocations Is Nothing Then
        If mAllocations.Exists(allocationId) Then NativeAllocationQty = CDbl(mAllocations(allocationId))
    End If
End Function

Private Sub SaveAllocation(ByVal allocationId As String, ByVal requirementQty As Double, _
                           ByVal nativeQty As Double, ByVal catalogVersion As String, _
                           ByVal conversionFactor As Double)
    If mAllocations.Exists(allocationId) Then
        mAllocations(allocationId) = requirementQty
    Else
        mAllocations.Add allocationId, requirementQty
    End If
    If mAllocationNativeQuantities Is Nothing Then Set mAllocationNativeQuantities = NewTextDictionary()
    If mAllocationNativeQuantities.Exists(allocationId) Then
        mAllocationNativeQuantities(allocationId) = nativeQty
    Else
        mAllocationNativeQuantities.Add allocationId, nativeQty
    End If
    If mAllocationConversionAudit Is Nothing Then Set mAllocationConversionAudit = NewTextDictionary()
    If mAllocationConversionAudit.Exists(allocationId) Then
        mAllocationConversionAudit(allocationId) = catalogVersion & vbTab & CStr(conversionFactor)
    Else
        mAllocationConversionAudit.Add allocationId, catalogVersion & vbTab & CStr(conversionFactor)
    End If
End Sub

Private Sub RemoveAllocation(ByVal allocationId As String)
    If Not mAllocations Is Nothing Then If mAllocations.Exists(allocationId) Then mAllocations.Remove allocationId
    If Not mAllocationNativeQuantities Is Nothing Then If mAllocationNativeQuantities.Exists(allocationId) Then mAllocationNativeQuantities.Remove allocationId
    If Not mAllocationConversionAudit Is Nothing Then If mAllocationConversionAudit.Exists(allocationId) Then mAllocationConversionAudit.Remove allocationId
End Sub

Private Function ExactEntityUom(ByVal systemKey As String, Optional ByVal action As cProductionWorksheetAction = Nothing) As String
    Dim entities As Variant
    Dim rowIndex As Long
    entities = modInventoryDomainBridge.ListAvailableInventoryEntitiesBridge("")
    If Not modProductionRunLoadActions.CanContinue(action) Then Exit Function
    If Not IsArray(entities) Then Exit Function
    rowIndex = FindExactEntityRow(entities, systemKey)
    If rowIndex > 0 Then ExactEntityUom = Trim$(CStr(entities(rowIndex, 5)))
End Function

Public Function ReusableRunExactAllocationCountForRequirement(ByVal nodeId As String, _
                                                              ByVal requirementId As String) As Long
    Dim key As Variant
    Dim parts() As String

    If mAllocations Is Nothing Then Exit Function
    For Each key In mAllocations.Keys
        parts = Split(CStr(key), vbTab)
        If UBound(parts) = 2 Then
            If StrComp(parts(0), nodeId, vbTextCompare) = 0 _
               And StrComp(parts(1), requirementId, vbTextCompare) = 0 Then
                ReusableRunExactAllocationCountForRequirement = _
                    ReusableRunExactAllocationCountForRequirement + 1
            End If
        End If
    Next key
End Function

Private Function TotalExternalAllocation() As Double
    Dim key As Variant
    For Each key In mAllocations.Keys
        TotalExternalAllocation = TotalExternalAllocation + CDbl(mAllocations(key))
    Next key
End Function

Private Function ExactEntityAvailableQty(ByVal systemKey As String, _
                                         Optional ByRef locationOut As String = "", _
                                         Optional ByVal action As cProductionWorksheetAction = Nothing) As Double
    ExactEntityAvailableQty = modProductionRunEntityReads.AvailableQuantity(systemKey, locationOut, action)
End Function

Private Function ExactEntityIsNonCounted(ByVal systemKey As String, Optional ByVal action As cProductionWorksheetAction = Nothing) As Boolean
    ExactEntityIsNonCounted = modProductionRunEntityReads.IsNonCounted(systemKey, action)
End Function

Private Function ExactEntityField(ByVal entities As Variant, ByVal systemKey As String, _
                                  ByVal columnIndex As Long) As Variant
    Dim r As Long
    If Not IsArray(entities) Then Exit Function
    For r = LBound(entities, 1) To UBound(entities, 1)
        If StrComp(Trim$(CStr(entities(r, 1))), Trim$(systemKey), vbTextCompare) = 0 Then
            ExactEntityField = entities(r, columnIndex)
            Exit Function
        End If
    Next r
End Function

Private Function ExactEntityInventoryDisplayForKey(ByVal entities As Variant, _
                                                   ByVal systemKey As String) As Variant
    Dim r As Long

    If Not IsArray(entities) Then Exit Function
    For r = LBound(entities, 1) To UBound(entities, 1)
        If StrComp(Trim$(CStr(entities(r, 1))), Trim$(systemKey), vbTextCompare) = 0 Then
            ExactEntityInventoryDisplayForKey = ExactEntityInventoryDisplay(entities, r)
            Exit Function
        End If
    Next r
End Function

Private Function ExactEntityInventoryDisplay(ByVal entities As Variant, _
                                             ByVal entityRow As Long) As Variant
    Dim trackQty As String
    Dim itemKind As String
    Dim categoryValue As String

    If Not IsArray(entities) Then Exit Function
    If UBound(entities, 2) >= 11 Then trackQty = UCase$(Trim$(CStr(entities(entityRow, 11))))
    If UBound(entities, 2) >= 12 Then itemKind = UCase$(Trim$(CStr(entities(entityRow, 12))))
    If UBound(entities, 2) >= 13 Then categoryValue = UCase$(Trim$(CStr(entities(entityRow, 13))))
    If trackQty = "FALSE" Or trackQty = "NO" Or trackQty = "0" _
       Or itemKind = "UTILITY" Or itemKind = "SERVICE" Or itemKind = "NON_COUNTED" _
       Or categoryValue = "UTILITY" Or categoryValue = "SERVICE" Then
        Select Case itemKind
            Case "UTILITY": ExactEntityInventoryDisplay = "Utility"
            Case "SERVICE": ExactEntityInventoryDisplay = "Service"
            Case Else
                If categoryValue = "UTILITY" Then
                    ExactEntityInventoryDisplay = "Utility"
                ElseIf categoryValue = "SERVICE" Then
                    ExactEntityInventoryDisplay = "Service"
                Else
                    ExactEntityInventoryDisplay = "Not counted"
                End If
        End Select
    Else
        ExactEntityInventoryDisplay = entities(entityRow, 6)
    End If
End Function

Public Function ReusableRunInventoryDisplayForTest(ByVal trackQty As String, _
                                                   ByVal itemKind As String, _
                                                   ByVal categoryValue As String, _
                                                   ByVal qty As Double) As String
    Dim entities(1 To 1, 1 To 13) As Variant

    entities(1, 6) = qty
    entities(1, 11) = trackQty
    entities(1, 12) = itemKind
    entities(1, 13) = categoryValue
    ReusableRunInventoryDisplayForTest = CStr(ExactEntityInventoryDisplay(entities, 1))
End Function

Private Function OutputIdentityKey(ByVal nodeId As String, ByVal outputId As String) As String
    OutputIdentityKey = nodeId & vbTab & outputId
End Function

Private Function UomCompatible(ByVal sourceUom As String, ByVal targetUom As String) As Boolean
    sourceUom = Trim$(sourceUom)
    targetUom = Trim$(targetUom)
    UomCompatible = (sourceUom = "" Or targetUom = "" Or _
                     StrComp(sourceUom, targetUom, vbTextCompare) = 0)
End Function

Public Function CloneRunRecord(ByVal source As Object) As Object
    Dim result As Object
    Dim key As Variant
    Set result = NewTextDictionary()
    For Each key In source.Keys
        result(CStr(key)) = source(key)
    Next key
    Set CloneRunRecord = result
End Function

Private Function NewTextDictionary() As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    Set NewTextDictionary = result
End Function

Public Function RunRecordText(ByVal record As Object, ByVal fieldName As String) As String
    If record Is Nothing Then Exit Function
    If record.Exists(fieldName) Then
        If Not IsNull(record(fieldName)) And Not IsEmpty(record(fieldName)) Then _
            RunRecordText = Trim$(CStr(record(fieldName)))
    End If
End Function

Private Function RunRecordValue(ByVal record As Object, ByVal fieldName As String) As Variant
    If record Is Nothing Then Exit Function
    If record.Exists(fieldName) Then RunRecordValue = record(fieldName)
End Function

Private Function RunRecordNumber(ByVal record As Object, ByVal fieldName As String) As Double
    If record Is Nothing Then Exit Function
    If record.Exists(fieldName) Then
        If IsNumeric(record(fieldName)) Then RunRecordNumber = CDbl(record(fieldName))
    End If
End Function

Private Function FormatRunNumberLocal(ByVal valueIn As Double) As String
    FormatRunNumberLocal = Format$(valueIn, "0.#########")
End Function

Private Function JsonRunNumber(ByVal valueIn As Double) As String
    JsonRunNumber = Replace$(CStr(valueIn), Application.International(xlDecimalSeparator), ".")
End Function

Private Function RunEventNote(ByVal runId As String, ByVal node As Object, _
                              ByVal stageName As String) As String
    RunEventNote = "PRODUCTION_REUSABLE_RUN|RunId=" & runId & "|Recipe=" & mRecipeId & _
                   "|RecipeVersion=" & mRecipeVersion & "|Batch=" & CStr(mBatchNumber) & _
                   "|BatchNote=" & mBatchNote & _
                   "|ProcessNode=" & RunRecordText(node, "ProcessNodeId") & _
                   "|Process=" & RunRecordText(node, "ProcessId") & _
                   "|ProcessVersion=" & RunRecordText(node, "ProcessVersion") & _
                   "|Stage=" & stageName
End Function

Private Function RunItemNote(ByVal runId As String, ByVal node As Object, _
                             ByVal lineId As String) As String
    RunItemNote = RunEventNote(runId, node, "LINE") & "|Line=" & lineId
End Function

Private Sub AppendEventId(ByRef eventIds As String, ByVal eventId As String)
    If eventIds <> "" Then eventIds = eventIds & ","
    eventIds = eventIds & eventId
End Sub

Private Sub AppendProcessorReport(ByRef reports As String, ByVal report As String)
    If Trim$(report) = "" Then Exit Sub
    If reports <> "" Then reports = reports & " | "
    reports = reports & report
End Sub
