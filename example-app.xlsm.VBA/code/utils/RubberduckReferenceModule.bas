Attribute VB_Name = "RubberduckReferenceModule"
'@Folder("utils")
Option Explicit
Option Private Module

Private Const RUBBERDUCK_REGEX As String = "rubberduck\.x\d+\.tlb"
Private Const RUBBERDUCK_TYPELIB_GUID As String = "{7B73DBEA-A9EC-47C5-BF8B-FA7F32170024}"
Private Const RUBBERDUCK_TYPELIB_VERSION As String = "2.5"

'@EntryPoint
Public Sub AddRubberduckReference()
    On Error GoTo ErrorHandler

    Dim oldPath As String
    oldPath = GetExistingRubberduckReferencePath()

    If Len(oldPath) > 0 Then
        If Not TryRemoveRubberduckReference() Then
            Err.Raise AppError.RUBBERDUCK_REFERENCE_RESTORE_FAILED, "AddRubberduckReference", "Failed to remove existing Rubberduck reference."
        End If
    End If

    addRubberduckReferenceByArchitecture
    showAddedMessage
    Exit Sub

ErrorHandler:
    ' Roll back if we removed an existing reference and re-add failed.
    If Len(oldPath) > 0 Then
        Call TryRestoreRubberduckReference(oldPath)
    End If
    showAddErrorMessage Err.Description
End Sub

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
    Err.Clear
End Function

'@EntryPoint
Public Sub RemoveRubberduckReference()
    On Error GoTo ErrorHandler
    If Not TryRemoveRubberduckReference() Then showNotFoundMessage: Exit Sub
    showRemovedMessage
    Exit Sub
ErrorHandler:
    showRemoveErrorMessage Err.Description
End Sub

Private Function hasRubberduckReference() As Boolean
    Dim ref As Reference
    For Each ref In ThisWorkbook.VBProject.References
        If IsRubberduckReference(ref) Then hasRubberduckReference = True: Exit Function
    Next ref
End Function

Private Function TryRemoveRubberduckReference() As Boolean
    Dim ref As Reference
    For Each ref In ThisWorkbook.VBProject.References
        If IsRubberduckReference(ref) Then ThisWorkbook.VBProject.References.Remove ref: TryRemoveRubberduckReference = True: Exit Function
    Next ref
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
    Err.Raise vbObjectError + 513, "AddRubberduckReference", "Rubberduck type library was not found in registry or known path."
End Sub

Private Function TryAddRubberduckReferenceFromKnownPath() As Boolean
     On Error GoTo TryX32
    Dim knownPath As String
    knownPath = "C:\ProgramData\Rubberduck\Rubberduck.x64.tlb"
    ThisWorkbook.VBProject.References.AddFromFile knownPath
    TryAddRubberduckReferenceFromKnownPath = True
    Exit Function
TryX32:
     Err.Clear
     On Error GoTo Failed
     knownPath = "C:\ProgramData\Rubberduck\Rubberduck.x32.tlb"
     ThisWorkbook.VBProject.References.AddFromFile knownPath
     TryAddRubberduckReferenceFromKnownPath = True
     Exit Function
Failed:
    Err.Clear
End Function

Private Function TryAddRubberduckReference(ByVal platform As String) As Boolean
    On Error GoTo Failed
    Dim path As String: path = getPathRubberduckTlb(platform)
    ThisWorkbook.VBProject.References.AddFromFile path
    TryAddRubberduckReference = True
    Exit Function
Failed:
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
