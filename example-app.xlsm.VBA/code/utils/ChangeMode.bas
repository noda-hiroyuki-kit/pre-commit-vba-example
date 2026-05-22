Attribute VB_Name = "ChangeMode"
'@Folder("utils")
Option Explicit

'@EntryPoint
Public Sub PrepareRelease()
    '@Ignore AssignmentNotUsed
    Dim canRefDisabled As Boolean: canRefDisabled = False
    '@Ignore AssignmentNotUsed
    Dim rubberduckReferenceRemoved As Boolean: rubberduckReferenceRemoved = False

    On Error GoTo PrepareReleaseError
    VersionInfoModule.SetVersionFromBranchIfReleaseOrHotfix

    SetCanRefFalse
    canRefDisabled = True

    rubberduckReferenceRemoved = RubberduckReferenceModule.RemoveRubberduckReference
    Exit Sub
PrepareReleaseError:
    Dim originalErrorNumber As Long:        originalErrorNumber = Err.Number
    Dim originalErrorDescription As String: originalErrorDescription = Err.Description

    On Error Resume Next
    If rubberduckReferenceRemoved Then AddRubberduckReference showMessage:=False
    If canRefDisabled Then SetCanRefTrue
    On Error GoTo 0

    Err.Raise originalErrorNumber, "ChangeMode.PrepareRelease", _
              "PrepareRelease failed: " & originalErrorDescription
End Sub

'@EntryPoint
Public Sub ChangeDevelop()
    '@Ignore AssignmentNotUsed
    Dim referenceAdded As Boolean: referenceAdded = False
    '@Ignore AssignmentNotUsed
    Dim canRefEnabled As Boolean:  canRefEnabled = False

    On Error GoTo ChangeDevelopError
    AddRubberduckReference showMessage:=False

    referenceAdded = RubberduckReferenceModule.HasRubberduckReference()
    If Not referenceAdded Then
        Err.Raise AppError.RUBBERDUCK_TYPELIB_NOT_FOUND, "ChangeMode.ChangeDevelop", _
                  "Failed to add Rubberduck reference."
    End If

    SetCanRefTrue
    canRefEnabled = True
    Exit Sub
ChangeDevelopError:
    Dim originalErrorNumber As Long:        originalErrorNumber = Err.Number
    Dim originalErrorDescription As String: originalErrorDescription = Err.Description

    On Error Resume Next
    If canRefEnabled Then SetCanRefFalse
    '@Ignore FunctionReturnValueDiscarded
    If referenceAdded Then RemoveRubberduckReference showMessage:=False
    On Error GoTo 0

    Err.Raise originalErrorNumber, "ChangeMode.ChangeDevelop", _
              "ChangeDevelop failed: " & originalErrorDescription
End Sub
