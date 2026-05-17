Attribute VB_Name = "VersionInfoModule"
'@Folder("utils")
Option Explicit
Option Private Module

Public Sub SetVersion(Optional ByVal manager As VersionManager)
    Dim response As String
    response = InputBox("Input Version")
    If Len(response) = 0 Then Exit Sub
    Dim managerLocal As VersionManager: Set managerLocal = manager
    If managerLocal Is Nothing Then Set managerLocal = New VersionManager: managerLocal.Create ThisWorkbook
    managerLocal.SetVersion response
End Sub

'@EntryPoint
Public Sub SetVersionFromBranchIfReleaseOrHotfix()
    Dim manager As VersionManager: Set manager = New VersionManager: manager.Create ThisWorkbook
    Dim resolver As BranchVersionResolver: Set resolver = New BranchVersionResolver: resolver.Create manager
    resolver.SetVersionFromBranchIfReleaseOrHotfix
End Sub
