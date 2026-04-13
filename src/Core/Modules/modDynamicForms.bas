Attribute VB_Name = "modDynamicForms"
Option Explicit

Public Function CreateDynItemSearch() As Object
    Set CreateDynItemSearch = New cDynItemSearch
End Function

Public Function CreatePickerRouter() As Object
    Set CreatePickerRouter = New cPickerRouter
End Function

Public Function CreateFormAnchorManager() As Object
    Set CreateFormAnchorManager = New cFormAnchorManager
End Function

