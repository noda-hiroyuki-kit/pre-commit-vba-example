Attribute VB_Name = "VersionInfoModule"
'@Folder("utils")
Option Explicit
Option Private Module

Public Sub setVersion()
    Dim response As String
    response = InputBox("Input Version")
    If Len(response) = 0 Then Exit Sub
    setDocumentVersion response
End Sub

'@EntryPoint
Public Sub setVersionFromBranchIfReleaseOrHotfix()
    Dim version As String: version = getVersionFromCurrentBranch()
    If Len(version) = 0 Then Exit Sub
    setDocumentVersion version
End Sub

Private Function getVersionFromCurrentBranch() As String
    getVersionFromCurrentBranch = extractVersionFromBranchName(getCurrentBranchName())
End Function

Private Function getCurrentBranchName() As String
    On Error GoTo ErrorHandler

    getCurrentBranchName = readStdOut("git rev-parse --abbrev-ref HEAD")
    Exit Function

ErrorHandler:
    getCurrentBranchName = vbNullString
End Function

Private Function readStdOut(ByVal commandText As String) As String
    Dim shell As WshShell: Set shell = New WshShell
    readStdOut = Trim$(shell.Exec(commandText).StdOut.ReadAll)
End Function

Private Function extractVersionFromBranchName(ByVal branchName As String) As String
    Dim regEx As RegExp: Set regEx = createTargetBranchRegex()
    If Not regEx.Test(branchName) Then Exit Function
    extractVersionFromBranchName = regEx.Execute(branchName).Item(0).SubMatches.Item(1)
End Function

Private Function createTargetBranchRegex() As RegExp
    Dim regEx As RegExp: Set regEx = New RegExp
    regEx.Pattern = "^(release|hotfix)/(v(0|[1-9]\\d*)\\.(0|[1-9]\\d*)\\.(0|[1-9]\\d*)(?:-(?:0|[1-9]\\d*|\\d*[A-Za-z-][0-9A-Za-z-]*)(?:\\.(?:0|[1-9]\\d*|\\d*[A-Za-z-][0-9A-Za-z-]*))*)?(?:\\+[0-9A-Za-z-]+(?:\\.[0-9A-Za-z-]+)*)?)$"
    regEx.IgnoreCase = False
    Set createTargetBranchRegex = regEx
End Function

Private Sub setDocumentVersion(ByVal version As String)
    ThisWorkbook.BuiltinDocumentProperties.Item("Document Version").Value = version
End Sub
