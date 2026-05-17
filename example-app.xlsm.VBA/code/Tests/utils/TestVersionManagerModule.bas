Attribute VB_Name = "TestVersionManagerModule"
'@TestModule
'@Folder("Tests.utils")

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

'@TestMethod("VersionManager")
Private Sub TestCreateForTestSetsVersion()
    On Error GoTo TestFail

    ' Arrange
    Dim sut As VersionManager: Set sut = New VersionManager

    ' Act
    sut.CreateForTest "v1.2.3"

    ' Assert
    testCon.Assert.AreEqual "v1.2.3", sut.Version

TestExit:
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VersionManager")
Private Sub TestVersionBeforeCreateRaisesNotCreated()
    On Error GoTo TestFail

    ' Arrange
    Dim sut As VersionManager: Set sut = New VersionManager

    ' Act
    On Error Resume Next
    '@Ignore VariableNotUsed
    Dim actual As String
    '@Ignore AssignmentNotUsed
    actual = sut.Version
    Dim raisedNumber As Long
    raisedNumber = Err.Number
    On Error GoTo TestFail

    ' Assert
    testCon.Assert.AreEqual AppError.NOT_CREATED, raisedNumber

TestExit:
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VersionManager")
Private Sub TestSetVersionUpdatesVersionInTestMode()
    On Error GoTo TestFail

    ' Arrange
    Dim sut As VersionManager: Set sut = New VersionManager
    sut.CreateForTest "v1.0.0"

    ' Act
    sut.SetVersion "v1.0.1"

    ' Assert
    testCon.Assert.AreEqual "v1.0.1", sut.Version

TestExit:
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VersionManager")
Private Sub TestCreateForTestTwiceRaisesHasCreated()
    On Error GoTo TestFail

    ' Arrange
    Dim sut As VersionManager: Set sut = New VersionManager
    sut.CreateForTest "v1.0.0"

    ' Act
    On Error Resume Next
    sut.CreateForTest "v1.0.1"
    Dim raisedNumber As Long
    raisedNumber = Err.Number
    On Error GoTo TestFail

    ' Assert
    testCon.Assert.AreEqual AppError.HAS_CREATED, raisedNumber

TestExit:
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
