Attribute VB_Name = "TestProductNameModule"
'@TestModule
'@Folder "Tests.domain.model"

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

'@TestMethod("ProductName.IsValid")
Private Sub TestIsValid_EmptyString_ReturnsInvalid()
    On Error GoTo TestFail

    Dim sut As ProductName: Set sut = New ProductName
    Dim result As ValidationResult
    Set result = sut.IsValid(vbNullString)
    testCon.Assert.AreEqual Invalid, result.result

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("ProductName.IsValid")
Private Sub TestIsValid_MinLength_ReturnsValid()
    On Error GoTo TestFail

    Dim sut As ProductName: Set sut = New ProductName
    Dim result As ValidationResult
    Set result = sut.IsValid("A")
    testCon.Assert.AreEqual Valid, result.result

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("ProductName.IsValid")
Private Sub TestIsValid_MaxLength_ReturnsValid()
    On Error GoTo TestFail

    Dim sut As ProductName: Set sut = New ProductName
    Dim result As ValidationResult
    Set result = sut.IsValid(String(255, "A"))
    testCon.Assert.AreEqual Valid, result.result

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("ProductName.IsValid")
Private Sub TestIsValid_TooLong_ReturnsInvalid()
    On Error GoTo TestFail

    Dim sut As ProductName: Set sut = New ProductName
    Dim result As ValidationResult
    Set result = sut.IsValid(String(256, "A"))
    testCon.Assert.AreEqual Invalid, result.result

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
