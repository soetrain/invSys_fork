Attribute VB_Name = "modLoadedEvents"
Option Explicit
Option Private Module

' Core retains the verified model. UI callers acknowledge only its descriptor.
Private mContext As String, mLoaded As String, mHash As String, mStale As Boolean
Private mModel As Object
Private mPendingContext As String, mPendingLoaded As String, mPendingHash As String
Private mPending As Object

Public Sub BeginRead(ByVal context As String)
    If context <> mContext Then ClearAll
    mStale = True: Set mPending = Nothing
    mPendingContext = context
End Sub

Public Sub StageRead(ByVal context As String, ByVal model As Object, ByVal hash As String, ByVal loaded As String)
    If context = "" Or context <> modActivity.CaptureContext() Or context <> mPendingContext Then Exit Sub
    Set mPending = model: mPendingHash = hash: mPendingLoaded = loaded
End Sub

Public Function Accept(ByVal context As String, ByVal publicationId As String, ByVal loaded As String) As Boolean
    If mPending Is Nothing Then Exit Function
    If context = "" Or context <> mPendingContext Or context <> modActivity.CaptureContext() Then Exit Function
    If publicationId <> CStr(mPending("PublicationId")) Or loaded <> mPendingLoaded Then Exit Function
    mContext = context: Set mModel = mPending: mHash = mPendingHash: mLoaded = loaded: mStale = False
    Set mPending = Nothing: mPendingContext = "": mPendingHash = "": mPendingLoaded = ""
    Accept = True
End Function

Public Sub MarkStale(ByVal context As String)
    If context = mContext Then mStale = True
    Set mPending = Nothing
End Sub

Public Sub ClearContext(ByVal context As String)
    If context = mContext Or context = mPendingContext Then ClearAll
End Sub

Public Sub ClearAll()
    Set mModel = Nothing: Set mPending = Nothing
    mContext = "": mLoaded = "": mHash = "": mStale = True
    mPendingContext = "": mPendingHash = "": mPendingLoaded = ""
End Sub

Public Function Read(ByVal context As String, ByRef model As Object, ByRef provenance As Object) As Boolean
    Dim field As Variant, coverage As Object
    Set model = Nothing: Set provenance = CreateObject("Scripting.Dictionary")
    provenance.Add "Availability", "Unavailable"
    For Each field In Array("WarehouseId", "PublicationId", "ContentSha256", "PackageSetVersion", "BuildIdentity", "PublishedAtUTC", "LoadedAtUTC")
        provenance.Add CStr(field), ""
    Next field
    provenance.Add "SchemaVersion", 0&: provenance.Add "PolicyVersion", 0&
    Set coverage = CreateObject("Scripting.Dictionary"): provenance.Add "Coverage", coverage
    If context = "" Or context <> modActivity.CaptureContext() Then ClearAll: Exit Function
    If context <> mContext Or mModel Is Nothing Then Exit Function
    For Each field In Array("WarehouseId", "PublicationId", "SchemaVersion", "PackageSetVersion", "BuildIdentity", "PublishedAtUTC", "PolicyVersion")
        provenance(field) = mModel(field)
    Next field
    provenance("ContentSha256") = mHash: provenance("LoadedAtUTC") = mLoaded
    Set provenance("Coverage") = mModel("Coverage")
    provenance("Availability") = IIf(mStale, "Stale", "Loaded")
    If mStale Then Exit Function
    Set model = mModel: Read = True
End Function
