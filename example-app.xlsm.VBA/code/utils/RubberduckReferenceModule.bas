Attribute VB_Name = "RubberduckReferenceModule"
'@Folder("utils")
Option Explicit
Option Private Module

Private Const RUBBERDUCK_REGEX As String = "rubberduck\.x\d+\.tlb"
Private Const RUBBERDUCK_TYPELIB_GUID As String = "{7B73DBEA-A9EC-47C5-BF8B-FA7F32170024}"
Private Const RUBBERDUCK_TYPELIB_VERSION As String = "2.5"

'@EntryPoint
Public Sub AddRubberduckReference(Optional ByVal showMessage As Boolean = True)
    On Error GoTo ErrorHandler

    Dim oldPath As String
    oldPath = GetExistingRubberduckReferencePath()

    If Len(oldPath) > 0 Then
        If Not TryRemoveRubberduckReference() Then
            Err.Raise AppError.RUBBERDUCK_REFERENCE_REMOVE_FAILED, _
                      "RubberduckReferenceModule.AddRubberduckReference", _
                      "Failed to remove existing Rubberduck reference."
        End If
    End If

    addRubberduckReferenceByArchitecture
    If showMessage Then showAddedMessage
    Exit Sub

ErrorHandler:
    Dim originalErrorNumber As Long:        originalErrorNumber = Err.Number
    Dim addErrorDescription As String:      addErrorDescription = Err.Description
    Dim restored As Boolean:                restored = True
    ' Roll back if we removed an existing reference and re-add failed.
    If Len(oldPath) > 0 Then
        restored = TryRestoreRubberduckReference(oldPath)
        If Not restored Then _
           addErrorDescription = addErrorDescription & _
           " Restoration of the previous Rubberduck reference also failed."
    End If

    If showMessage Then showAddErrorMessage addErrorDescription
    Err.Raise originalErrorNumber, "RubberduckReferenceModule.AddRubberduckReference", addErrorDescription
End Sub

Public Function HasRubberduckReference() As Boolean
    On Error GoTo ErrorHandler
    HasRubberduckReference = Len(GetExistingRubberduckReferencePath()) > 0
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "RubberduckReferenceModule.HasRubberduckReference", _
              "Failed to determine whether Rubberduck reference exists: " & Err.Description
End Function

Private Function GetExistingRubberduckReferencePath() As String
    Dim ref As Reference
    For Each ref In ThisWorkbook.VBProject.References
        If IsRubberduckReference(ref) Then
            GetExistingRubberduckReferencePath = ref.FullPath
            Exit Function
        End If
    Next ref
End Function

Private Function TryRestoreRubberduckReference(ByVal path As String) As Boolean
    On Error GoTo Failed
    ThisWorkbook.VBProject.References.AddFromFile path
    TryRestoreRubberduckReference = True
    Exit Function
Failed:
    Debug.Print "TryRestoreRubberduckReference failed. path=" & path & _
                ", err=" & Err.Number & ": " & Err.Description
    Err.Clear
End Function

'@EntryPoint
Public Function RemoveRubberduckReference(Optional ByVal showMessage As Boolean = True) As Boolean
    On Error GoTo ErrorHandler
    If Not TryRemoveRubberduckReference Then
        RemoveRubberduckReference = False
        If showMessage Then showNotFoundMessage
        Exit Function
    End If
    RemoveRubberduckReference = True
    If showMessage Then showRemovedMessage
    Exit Function
ErrorHandler:
    If showMessage Then showRemoveErrorMessage Err.Description
    Err.Raise Err.Number, "RubberduckReferenceModule.RemoveRubberduckReference", Err.Description
End Function

Private Function TryRemoveRubberduckReference() As Boolean
    Dim ref As Reference
    Dim target As Reference

    For Each ref In ThisWorkbook.VBProject.References
        If IsRubberduckReference(ref) Then
            Set target = ref
            Exit For
        End If
    Next ref

    If target Is Nothing Then Exit Function

    ThisWorkbook.VBProject.References.Remove target
    TryRemoveRubberduckReference = True
End Function

Private Function IsRubberduckReference(ByVal ref As Reference) As Boolean
    Dim regEx As RegExp: Set regEx = New RegExp
    regEx.Pattern = RUBBERDUCK_REGEX
    regEx.IgnoreCase = True
    IsRubberduckReference = regEx.Test(ref.FullPath)
End Function

Private Sub addRubberduckReferenceByArchitecture()
    If TryAddRubberduckReference("win64") Then Exit Sub
    If TryAddRubberduckReference("win32") Then Exit Sub
    If TryAddRubberduckReferenceFromKnownPath() Then Exit Sub
    Err.Raise AppError.RUBBERDUCK_TYPELIB_NOT_FOUND, "RubberduckReferenceModule", "Rubberduck type library was not found in registry or known path."
End Sub

Private Function TryAddRubberduckReferenceFromFile(ByVal knownPath As String) As Boolean
    On Error GoTo Failed
    ThisWorkbook.VBProject.References.AddFromFile knownPath
    TryAddRubberduckReferenceFromFile = True
    Exit Function
Failed:
    Debug.Print "TryAddRubberduckReferenceFromFile failed. path=" & knownPath & _
                ", err=" & Err.Number & ": " & Err.Description
    Err.Clear
End Function

Private Function TryAddRubberduckReferenceFromKnownPath() As Boolean
    TryAddRubberduckReferenceFromKnownPath = False
    If TryAddRubberduckReferenceFromFile("C:\ProgramData\Rubberduck\Rubberduck.x64.tlb") Then
        TryAddRubberduckReferenceFromKnownPath = True
        Exit Function
    End If

    If TryAddRubberduckReferenceFromFile("C:\ProgramData\Rubberduck\Rubberduck.x86.tlb") Then
        TryAddRubberduckReferenceFromKnownPath = True
        Exit Function
    End If
End Function

Private Function TryAddRubberduckReference(ByVal platform As String) As Boolean
    On Error GoTo Failed
    Dim path As String: path = getPathRubberduckTlb(platform)
    ThisWorkbook.VBProject.References.AddFromFile path
    TryAddRubberduckReference = True
    Exit Function
Failed:
    Debug.Print "TryAddRubberduckReference failed. platform=" & platform & _
                ", err=" & Err.Number & ": " & Err.Description
    Err.Clear
End Function

Private Function getPathRubberduckTlb(ByVal platform As String) As String
    Dim shell As WshShell: Set shell = New WshShell
    getPathRubberduckTlb = shell.RegRead(getTypeLibKey(platform))
End Function

Private Function getTypeLibKey(ByVal platform As String) As String
    getTypeLibKey = "HKEY_LOCAL_MACHINE\Software\Classes\TypeLib\" & RUBBERDUCK_TYPELIB_GUID & "\" & RUBBERDUCK_TYPELIB_VERSION & "\" & platform & "\"
End Function

Private Sub showAddedMessage()
    MsgBox "Rubberduck reference added successfully.", vbOKOnly + vbInformation
End Sub

Private Sub showNotFoundMessage()
    MsgBox "Rubberduck reference not found.", vbOKOnly + vbInformation
End Sub

Private Sub showRemovedMessage()
    MsgBox "Rubberduck reference removed successfully.", vbOKOnly + vbInformation
End Sub

Private Sub showAddErrorMessage(ByVal errorText As String)
    MsgBox "Error adding Rubberduck reference: " & errorText, vbOKOnly + vbCritical
End Sub

Private Sub showRemoveErrorMessage(ByVal errorText As String)
    MsgBox "Error removing Rubberduck reference: " & errorText, vbOKOnly + vbCritical
End Sub
