Attribute VB_Name = "ChangeMode"
'@Folder("utils")
Option Explicit

'@EntryPoint
Public Sub PrepareRelease()
    VersionInfoModule.SetVersionFromBranchIfReleaseOrHotfix
    SetCanRefFalse
    RemoveRubberduckReference
End Sub

'@EntryPoint
Public Sub ChangeDevelop()
    SetCanRefTrue
    AddRubberduckReference
End Sub
