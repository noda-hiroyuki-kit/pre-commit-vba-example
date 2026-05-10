Attribute VB_Name = "VersionInfoModule"
'@Folder("utils")
Option Explicit
Option Private Module

Public Sub setVersion(Optional ByVal manager As VersionManager = Nothing)
    Dim response As String
    response = InputBox("Input Version")
    If Len(response) = 0 Then Exit Sub
    If manager Is Nothing Then Set manager = New VersionManager
    manager.setVersion response
End Sub

'@EntryPoint
Public Sub setVersionFromBranchIfReleaseOrHotfix()
    Dim manager As VersionManager: Set manager = New VersionManager
    manager.setVersionFromBranchIfReleaseOrHotfix
End Sub
