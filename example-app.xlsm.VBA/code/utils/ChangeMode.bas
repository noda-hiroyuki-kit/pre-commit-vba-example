Attribute VB_Name = "ChangeMode"
'@Folder("utils")
Option Explicit

'@EntryPoint
Public Sub prepareRelease()
    Dim manager As VersionManager: Set manager = New VersionManager
    manager.Create ThisWorkbook
    Dim resolver As BranchVersionResolver: Set resolver = New BranchVersionResolver
    resolver.Create manager
    resolver.setVersionFromBranchIfReleaseOrHotfix

    SetCanRefFalse
    RemoveRubberduckReference
End Sub

'@EntryPoint
Public Sub changeDevelop()
    SetCanRefTrue
    AddRubberduckReference
End Sub
