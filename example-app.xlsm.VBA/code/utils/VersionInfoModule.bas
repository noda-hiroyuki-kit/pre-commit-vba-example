Attribute VB_Name = "VersionInfoModule"
'@Folder("utils")
Option Explicit
Option Private Module

Public Sub setVersion()
    Dim response As String
    response = InputBox("Input Version")
    If Len(response) = 0 Then Exit Sub
    Dim manager As VersionManager: Set manager = New VersionManager
    manager.setDocumentVersion response
End Sub

'@EntryPoint
Public Sub setVersionFromBranchIfReleaseOrHotfix()
    Dim manager As VersionManager: Set manager = New VersionManager
    manager.setVersionFromBranchIfReleaseOrHotfix
End Sub
