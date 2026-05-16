VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VersionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "utils"
Option Explicit

Private Type TVersionManager
    currentVersion As String
    nextVersion As String
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
    checkNotCreated
    this.isSystemUnderTest = True
    this.semVer.ensureSemVer currentVersion, "current"
    this.currentVersion = currentVersion
    this.Created = True
End Sub

Public Sub Create(ByVal targetBook As Workbook)
    checkNotCreated
    this.isSystemUnderTest = False
    Set this.bookVersion = targetBook.BuiltinDocumentProperties.Item("Document Version")
    Dim propertyVersion As String
    propertyVersion = CStr(this.bookVersion.Value)
    this.semVer.ensureSemVer propertyVersion, "current"
    this.currentVersion = propertyVersion
    this.Created = True
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

Private Sub ApplyToWorkbook(ByVal nextVersion As String)
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

Private Sub checkCreated()
    If this.Created Then Exit Sub
    Err.Raise AppError.NOT_CREATED, "VersionManager", "VersionManager has not been created."
End Sub

Private Sub checkNotCreated()
    If Not this.Created Then Exit Sub
    Err.Raise AppError.HAS_CREATED, "VersionManager", "VersionManager has been created."
End Sub
