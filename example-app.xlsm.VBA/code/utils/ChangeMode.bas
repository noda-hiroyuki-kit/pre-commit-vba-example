Attribute VB_Name = "ChangeMode"
'@Folder("utils")
Option Explicit

Private Const RUBBERDUCK_REGEX As String = "rubberduck\.x\d+\.tlb"

'@EntryPoint
Public Sub PrepareRelease()
    Dim canRefDisabled As Boolean: canRefDisabled = False
    Dim rubberduckReferenceRemoved As Boolean: rubberduckReferenceRemoved = False

    On Error GoTo PrepareReleaseError
    VersionInfoModule.SetVersionFromBranchIfReleaseOrHotfix

    SetCanRefFalse
    canRefDisabled = True

    rubberduckReferenceRemoved = RubberduckReferenceModule.RemoveRubberduckReference
    If Not rubberduckReferenceRemoved Then
        Err.Raise AppError.RUBBERDUCK_REFERENCE_REMOVE_FAILED, "ChangeMode.PrepareRelease", _
                  "Failed to remove Rubberduck reference."
    End If
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
    Dim referenceAdded As Boolean: referenceAdded = False
    Dim canRefEnabled As Boolean:  canRefEnabled = False

    On Error GoTo ChangeDevelopError
    AddRubberduckReference showMessage:=False

    referenceAdded = HasRubberduckReference()
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
    If referenceAdded Then RemoveRubberduckReference showMessage:=False
    On Error GoTo 0

    Err.Raise originalErrorNumber, "ChangeMode.ChangeDevelop", _
              "ChangeDevelop failed: " & originalErrorDescription
End Sub

Private Function HasRubberduckReference() As Boolean
    Dim ref As Reference
    For Each ref In ThisWorkbook.VBProject.References
        If IsRubberduckReferencePath(ref.FullPath) Then
            HasRubberduckReference = True
            Exit Function
        End If
    Next ref
End Function

Private Function IsRubberduckReferencePath(ByVal referencePath As String) As Boolean
    Dim regEx As RegExp: Set regEx = New RegExp
    regEx.Pattern = RUBBERDUCK_REGEX
    regEx.IgnoreCase = True
    IsRubberduckReferencePath = regEx.Test(referencePath)
End Function
