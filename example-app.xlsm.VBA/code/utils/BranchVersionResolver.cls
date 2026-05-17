VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BranchVersionResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "utils"
Option Explicit

Private Type TClass
    manager As VersionManager
    isSystemUnderTest As Boolean
    branchName As String
    Created As Boolean
End Type

Private this As TClass

Private Sub Class_Initialize()
    this.Created = False
End Sub

Public Sub Create(ByVal manager As VersionManager)
    checkNotCreated
    Set this.manager = manager
    this.isSystemUnderTest = False
    this.Created = True
End Sub

Public Sub CreateForTest(ByVal manager As VersionManager, ByVal branchName As String)
    checkNotCreated
    Set this.manager = manager
    this.isSystemUnderTest = True
    this.branchName = branchName
    this.Created = True
End Sub

Public Sub SetVersionFromBranchIfReleaseOrHotfix()
    checkCreated

    Dim branchName As String: branchName = getCurrentBranchName()

    Dim versionText As String: versionText = extractVersionFromBranchName(branchName)
    If Len(versionText) = 0 Then Exit Sub

    On Error GoTo SetVersionError
    this.manager.SetVersion versionText
    Exit Sub

SetVersionError:
    Err.Raise Err.Number, "BranchVersionResolver.SetVersionFromBranchIfReleaseOrHotfix", _
              "Failed to set version from branch """ & branchName & """: " & Err.Description
End Sub

Private Sub logBranchNameResolutionError(ByVal commandText As String, ByVal errorNumber As Long, ByVal errorDescription As String)
    Debug.Print "BranchVersionResolver.getCurrentBranchName failed. Command=""" & commandText & """ Error " & CStr(errorNumber) & ": " & errorDescription
End Sub

Private Function getCurrentBranchName() As String
    Dim commandText As String: commandText = "git rev-parse --abbrev-ref HEAD"
    On Error GoTo ErrorHandler
    getCurrentBranchName = readStdOut(commandText)
    Exit Function

ErrorHandler:
    logBranchNameResolutionError commandText, Err.Number, Err.Description
    getCurrentBranchName = vbNullString
End Function

Private Function readStdOut(ByVal commandText As String) As String
    If this.isSystemUnderTest Then readStdOut = this.branchName: Exit Function
    Dim shell As WshShell: Set shell = New WshShell

    Dim workbookPath As String: workbookPath = ThisWorkbook.path
    If Len(Trim$(workbookPath)) = 0 Then _
       Err.Raise AppError.INVALID_VALUE, "BranchVersionResolver.readStdOut", "Workbook path is empty."

    On Error GoTo CurrentDirectoryError
    shell.CurrentDirectory = workbookPath

    On Error GoTo ExecError
    readStdOut = Trim$(shell.Exec(commandText).StdOut.ReadAll)
    Exit Function
CurrentDirectoryError:
    Err.Raise AppError.INVALID_VALUE, "BranchVersionResolver.readStdOut", _
              "Failed to set shell current directory: " & workbookPath & " (" & Err.Description & ")"
    Exit Function
ExecError:
    Err.Raise AppError.INVALID_VALUE, "BranchVersionResolver.readStdOut", _
              "Failed to execute command: " & commandText & " (" & Err.Description & ")"
End Function

Private Function extractVersionFromBranchName(ByVal branchName As String) As String
    Dim regEx As RegExp: Set regEx = createTargetBranchRegex()
    If Not regEx.Test(branchName) Then Exit Function
    extractVersionFromBranchName = regEx.Execute(branchName).Item(0).SubMatches.Item(1)
End Function

Private Function createTargetBranchRegex() As RegExp
    Dim regEx As RegExp: Set regEx = New RegExp
    regEx.Pattern = "^(release/|hotfix/)(.*)"
    regEx.IgnoreCase = False
    Set createTargetBranchRegex = regEx
End Function

Private Sub checkCreated()
    If this.Created Then Exit Sub
    Err.Raise AppError.NOT_CREATED, "BranchVersionResolver", "BranchVersionResolver has not been created."
End Sub

Private Sub checkNotCreated()
    If Not this.Created Then Exit Sub
    Err.Raise AppError.HAS_CREATED, "BranchVersionResolver", "BranchVersionResolver has been created."
End Sub
