Attribute VB_Name = "TestBranchVersionResolver"
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

'@TestMethod("BranchVersionResolver.setVersionFromBranchIfReleaseOrHotfix")
Private Sub TestSetVersionFromBranchIfReleaseOrHotfix_ReleaseUpgrade_UpdatesVersion()
    On Error GoTo TestFail
    Dim sut As VersionManager: Set sut = New VersionManager
    sut.CreateForTest "v0.0.1"
    Dim resolver As BranchVersionResolver: Set resolver = New BranchVersionResolver
    resolver.CreateForTest sut, "release/v0.1.0"
    resolver.SetVersionFromBranchIfReleaseOrHotfix
    testCon.Assert.AreEqual "v0.1.0", sut.Version
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("BranchVersionResolver.setVersionFromBranchIfReleaseOrHotfix")
Private Sub TestSetVersionFromBranchIfReleaseOrHotfix_FeatureBranch_NoChange()
    On Error GoTo TestFail
    Dim sut As VersionManager: Set sut = New VersionManager
    sut.CreateForTest "v0.0.1"
    Dim resolver As BranchVersionResolver: Set resolver = New BranchVersionResolver
    resolver.CreateForTest sut, "feature/test"
    resolver.SetVersionFromBranchIfReleaseOrHotfix
    testCon.Assert.AreEqual "v0.0.1", sut.Version
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("BranchVersionResolver.setVersionFromBranchIfReleaseOrHotfix")
Private Sub TestSetVersionFromBranchIfReleaseOrHotfix_HotfixUpgrade_UpdatesVersion()
    On Error GoTo TestFail
    Dim sut As VersionManager: Set sut = New VersionManager
    sut.CreateForTest "v1.0.0"
    Dim resolver As BranchVersionResolver: Set resolver = New BranchVersionResolver
    resolver.CreateForTest sut, "hotfix/v1.0.1"
    resolver.SetVersionFromBranchIfReleaseOrHotfix
    testCon.Assert.AreEqual "v1.0.1", sut.Version
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("BranchVersionResolver.setVersionFromBranchIfReleaseOrHotfix")
Private Sub TestSetVersionFromBranchIfReleaseOrHotfix_ReleaseSameVersion_NoChange()
    On Error GoTo TestFail
    Dim sut As VersionManager: Set sut = New VersionManager
    sut.CreateForTest "v0.1.0"
    Dim resolver As BranchVersionResolver: Set resolver = New BranchVersionResolver
    resolver.CreateForTest sut, "release/v0.1.0"
    resolver.SetVersionFromBranchIfReleaseOrHotfix
    testCon.Assert.AreEqual "v0.1.0", sut.Version
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("BranchVersionResolver.setVersionFromBranchIfReleaseOrHotfix")
Private Sub TestSetVersionFromBranchIfReleaseOrHotfix_ReleaseLowerVersion_RaisesRollback()
    On Error GoTo TestFail
    Dim sut As VersionManager: Set sut = New VersionManager
    sut.CreateForTest "v0.1.0"
    Dim resolver As BranchVersionResolver: Set resolver = New BranchVersionResolver
    resolver.CreateForTest sut, "release/v0.0.9"
    resolver.SetVersionFromBranchIfReleaseOrHotfix
Assert:
    testCon.Assert.Fail "Expected error was not raised"
TestExit:
    Exit Sub
TestFail:
    If Err.Number = AppError.VERSION_ROLLBACK_NOT_ALLOWED Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub
