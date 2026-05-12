VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VersionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "domain"
Option Explicit

Private Type TVersionManager
    currentVersion As String
    nextVersion As String
    targetBook As Workbook
    bookVersion As Object
    semVer As SemanticVersioning
    Created As Boolean
    isSystemUnderTest As Boolean
End Type

Private this As TVersionManager

Private Sub Class_Initialize()
    this.currentVersion = vbNullString
    this.nextVersion = vbNullString
    Set this.semVer = New SemanticVersioning
    this.Created = False
End Sub

Public Sub CreateForTest(ByVal currentVersion As String)
    On Error GoTo ErrorHandler
    checkNotCreated
    this.isSystemUnderTest = True
    this.semVer.ensureSemVer currentVersion, "current"
    this.currentVersion = currentVersion
    this.Created = True
Exit_Sub:
    On Error GoTo 0
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile
    Resume Exit_Sub
End Sub

Public Sub Create(ByVal targetBook As Workbook)
    On Error GoTo ErrorHandler
    checkNotCreated
    Set this.bookVersion = targetBook.BuiltinDocumentProperties("Document Version")
    Dim propertyVersion As String
    propertyVersion = CStr(this.bookVersion.Value)
    this.semVer.ensureSemVer propertyVersion, "current"
    this.currentVersion = propertyVersion
    this.Created = True
Exit_Sub:
    On Error GoTo 0
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile
    Resume Exit_Sub
End Sub

Public Property Get Version() As String
    checkCreated
    Version = getVersion
End Property

Private Function getVersion() As String
    If this.isSystemUnderTest Then getVersion = getVersionUnderTest: Exit Function
    getVersion = CStr(this.bookVersion.Value)
End Function

Private Function getVersionUnderTest() As String
    If this.nextVersion = vbNullString Then getVersionUnderTest = this.currentVersion: Exit Function
    getVersionUnderTest = this.nextVersion
End Function

Public Sub setVersion(ByVal Version As String)
    checkCreated
    ApplyToWorkbook Version
End Sub

Public Sub setVersionFromBranchIfReleaseOrHotfix()
    checkCreated
    Dim Version As String: Version = getVersionFromCurrentBranch()
    If Len(Version) = 0 Then Exit Sub
    Me.setVersion Version
End Sub

Private Sub ApplyToWorkbook(ByVal nextVersion As String)
    checkCreated
    this.semVer.ensureSemVer nextVersion, "next"
    If Len(Trim$(this.currentVersion)) > 0 Then
        this.semVer.ensureNoRollback this.currentVersion, nextVersion
    End If
    setVersion_ nextVersion
End Sub

Private Sub setVersion_(ByVal nextVersion As String)
    If this.isSystemUnderTest Then this.nextVersion = nextVersion: Exit Sub
    this.bookVersion.Value = nextVersion
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
    regEx.Pattern = "^(release|hotfix)/(v(0|[1-9]\\d*)\.(0|[1-9]\\d*)\.(0|[1-9]\\d*)(?:-(?:0|[1-9]\\d*|\\d*[A-Za-z-][0-9A-Za-z-]*)(?:\.(?:0|[1-9]\\d*|\\d*[A-Za-z-][0-9A-Za-z-]*))*)?(?:\+[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?)$"
    regEx.IgnoreCase = False
    Set createTargetBranchRegex = regEx
End Function

Private Function isNumericIdentifier(ByVal valueText As String) As Boolean
    Dim regEx As RegExp
    Set regEx = New RegExp
    regEx.Pattern = "^(0|[1-9]\\d*)$"
    isNumericIdentifier = regEx.Test(valueText)
End Function

Private Sub checkCreated()
    If this.Created Then Exit Sub
    Err.Raise AppError.NOT_CREATED, "VersionManager", "VersionManager has not been created."
End Sub

Private Sub checkNotCreated()
    If Not this.Created Then Exit Sub
    Err.Raise AppError.HAS_CREATED, "VersionManager", "VersionManager has been created."
End Sub
