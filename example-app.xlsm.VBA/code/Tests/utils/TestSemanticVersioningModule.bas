Attribute VB_Name = "TestSemanticVersioningModule"
'@TestModule
'@Folder "Tests.utils"

Option Explicit
Option Private Module

Private testCon As TestController

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set testCon = New TestController
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set testCon = Nothing
End Sub

'@TestInitialize
'@Ignore EmptyMethod
Private Sub TestInitialize()
End Sub

'@TestCleanup
'@Ignore EmptyMethod
Private Sub TestCleanup()
End Sub

'@TestMethod("SemanticVersioning.ensureSemVer")
Private Sub TestEnsureSemVer_ValidSimpleVersion_DoesNotRaise()
    On Error GoTo TestFail

    AssertEnsureSemVerAccepts "10.20.30"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SemanticVersioning.ensureSemVer")
Private Sub TestEnsureSemVer_ValidPrereleaseWithMeta_DoesNotRaise()
    On Error GoTo TestFail

    AssertEnsureSemVerAccepts "1.1.2-prerelease+meta"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SemanticVersioning.ensureSemVer")
Private Sub TestEnsureSemVer_ValidAlphaDotBeta_DoesNotRaise()
    On Error GoTo TestFail

    AssertEnsureSemVerAccepts "1.0.0-alpha.beta.1"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SemanticVersioning.ensureSemVer")
Private Sub TestEnsureSemVer_ValidComplexPreAndBuild_DoesNotRaise()
    On Error GoTo TestFail

    AssertEnsureSemVerAccepts "1.0.0-alpha-a.b-c-somethinglong+build.1-aef.1-its-okay"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SemanticVersioning.ensureSemVer")
Private Sub TestEnsureSemVer_InvalidMajorOnly_Raises()
    On Error GoTo TestFail

    AssertEnsureSemVerRejects "1"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SemanticVersioning.ensureSemVer")
Private Sub TestEnsureSemVer_InvalidMissingPatch_Raises()
    On Error GoTo TestFail

    AssertEnsureSemVerRejects "1.2"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SemanticVersioning.ensureSemVer")
Private Sub TestEnsureSemVer_InvalidLeadingZeroInPrerelease_Raises()
    On Error GoTo TestFail

    AssertEnsureSemVerRejects "1.2.3-0123"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SemanticVersioning.ensureSemVer")
Private Sub TestEnsureSemVer_InvalidLeadingZeroInPrereleasePart_Raises()
    On Error GoTo TestFail

    AssertEnsureSemVerRejects "1.2.3-0123.0123"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SemanticVersioning.ensureSemVer")
Private Sub TestEnsureSemVer_InvalidEmptyPrereleaseIdentifier_Raises()
    On Error GoTo TestFail

    AssertEnsureSemVerRejects "1.0.0-alpha.."

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SemanticVersioning.ensureNoRollback")
Private Sub TestEnsureNoRollback_Beta2ToBeta11_DoesNotRaise()
    On Error GoTo TestFail

    AssertEnsureNoRollbackAccepts "1.0.0-beta.2", "1.0.0-beta.11"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SemanticVersioning.ensureNoRollback")
Private Sub TestEnsureNoRollback_Beta11ToRc1_DoesNotRaise()
    On Error GoTo TestFail

    AssertEnsureNoRollbackAccepts "1.0.0-beta.11", "1.0.0-rc.1"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SemanticVersioning.ensureNoRollback")
Private Sub TestEnsureNoRollback_Rc1ToStable_DoesNotRaise()
    On Error GoTo TestFail

    AssertEnsureNoRollbackAccepts "1.0.0-rc.1", "1.0.0"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SemanticVersioning.ensureNoRollback")
Private Sub TestEnsureNoRollback_100To110_DoesNotRaise()
    On Error GoTo TestFail

    AssertEnsureNoRollbackAccepts "1.0.0", "1.1.0"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SemanticVersioning.ensureNoRollback")
Private Sub TestEnsureNoRollback_1901To200_DoesNotRaise()
    On Error GoTo TestFail

    AssertEnsureNoRollbackAccepts "1.90.1", "2.0.0"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SemanticVersioning.ensureNoRollback")
Private Sub TestEnsureNoRollback_SameVersion_DoesNotRaise()
    On Error GoTo TestFail

    AssertEnsureNoRollbackAccepts "1.0.0", "1.0.0"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SemanticVersioning.ensureNoRollback")
Private Sub TestEnsureNoRollback_200To1901_Raises()
    On Error GoTo TestFail

    AssertEnsureNoRollbackRejects "2.0.0", "1.90.1"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SemanticVersioning.ensureNoRollback")
Private Sub TestEnsureNoRollback_StableToRc1_Raises()
    On Error GoTo TestFail

    AssertEnsureNoRollbackRejects "1.0.0", "1.0.0-rc.1"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

Private Sub AssertEnsureSemVerAccepts(ByVal versionText As String)
    Dim sut As SemanticVersioning
    Set sut = New SemanticVersioning
    On Error Resume Next
    sut.EnsureSemVer versionText, "current"
    Dim raisedNumber As Long
    raisedNumber = Err.Number
    Dim raisedDescription As String
    raisedDescription = Err.Description
    On Error GoTo 0
    testCon.Assert.AreEqual CLng(0), raisedNumber
    testCon.Assert.AreEqual vbNullString, raisedDescription
End Sub

Private Sub AssertEnsureSemVerRejects(ByVal versionText As String)
    Dim sut As SemanticVersioning
    Set sut = New SemanticVersioning
    On Error Resume Next
    sut.EnsureSemVer versionText, "current"
    Dim raisedNumber As Long
    raisedNumber = Err.Number
    Dim raisedDescription As String
    raisedDescription = Err.Description
    On Error GoTo 0
    testCon.Assert.AreEqual AppError.INVALID_SEMVER, raisedNumber
    testCon.Assert.AreEqual "current version must follow SemVer: " & versionText, raisedDescription
End Sub

Private Sub AssertEnsureNoRollbackAccepts(ByVal currentVersion As String, ByVal nextVersion As String)
    Dim sut As SemanticVersioning
    Set sut = New SemanticVersioning
    On Error Resume Next
    sut.EnsureNoRollback currentVersion, nextVersion
    Dim raisedNumber As Long
    raisedNumber = Err.Number
    Dim raisedDescription As String
    raisedDescription = Err.Description
    On Error GoTo 0
    testCon.Assert.AreEqual CLng(0), raisedNumber
    testCon.Assert.AreEqual vbNullString, raisedDescription
End Sub

Private Sub AssertEnsureNoRollbackRejects(ByVal currentVersion As String, ByVal nextVersion As String)
    Dim sut As SemanticVersioning
    Set sut = New SemanticVersioning
    On Error Resume Next
    sut.EnsureNoRollback currentVersion, nextVersion
    Dim raisedNumber As Long
    raisedNumber = Err.Number
    Dim raisedDescription As String
    raisedDescription = Err.Description
    On Error GoTo 0
    testCon.Assert.AreEqual AppError.VERSION_ROLLBACK_NOT_ALLOWED, raisedNumber
    testCon.Assert.AreEqual "Version rollback is not allowed: " & currentVersion & " -> " & nextVersion, raisedDescription
End Sub
