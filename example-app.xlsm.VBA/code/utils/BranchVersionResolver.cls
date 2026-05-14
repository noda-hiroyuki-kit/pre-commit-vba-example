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
End Type

Private this As TClass

Public Sub Create(ByVal manager As VersionManager)
    Set this.manager = manager
    this.isSystemUnderTest = False
End Sub

Public Sub CreateForTest(ByVal manager As VersionManager, ByVal branchName As String)
    Set this.manager = manager
    this.isSystemUnderTest = True
    this.branchName = branchName
End Sub
Public Sub setVersionFromBranchIfReleaseOrHotfix()
'    checkCreated
    Dim Version As String: Version = getVersionFromCurrentBranch()
    If Len(Version) = 0 Then Exit Sub
    this.manager.setVersion Version
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
    If this.isSystemUnderTest Then readStdOut = this.branchName: Exit Function
    Dim shell As WshShell: Set shell = New WshShell
    shell.CurrentDirectory = ThisWorkbook.path
    readStdOut = Trim$(shell.Exec(commandText).StdOut.ReadAll)
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
