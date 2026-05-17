Attribute VB_Name = "ChangeMode"
'@Folder("utils")
Option Explicit

'@EntryPoint
Public Sub PrepareRelease()
    Dim canRefDisabled As Boolean: canRefDisabled = False
    Dim rubberduckReferenceRemoved As Boolean: rubberduckReferenceRemoved = False

    On Error GoTo PrepareReleaseError
    VersionInfoModule.SetVersionFromBranchIfReleaseOrHotfix

    SetCanRefFalse
    canRefDisabled = True

    RemoveRubberduckReference
    rubberduckReferenceRemoved = True
    Exit Sub

PrepareReleaseError:
    Dim originalErrorNumber As Long:        originalErrorNumber = Err.Number
    Dim originalErrorDescription As String: originalErrorDescription = Err.Description

    On Error Resume Next
    If rubberduckReferenceRemoved Then AddRubberduckReference
    If canRefDisabled Then SetCanRefTrue
    On Error GoTo 0

    Err.Raise originalErrorNumber, "ChangeMode.PrepareRelease", _
              "PrepareRelease failed: " & originalErrorDescription
End Sub

'@EntryPoint
Public Sub ChangeDevelop()
    SetCanRefTrue
    AddRubberduckReference
End Sub
