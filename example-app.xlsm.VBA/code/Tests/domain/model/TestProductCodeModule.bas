Attribute VB_Name = "TestProductCodeModule"
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

'@TestMethod("ProductCode.IsValid")
Private Sub TestIsValid_TooShort_ReturnsInvalid()
    On Error GoTo TestFail

    Dim sut As New ProductCode
    Dim result As ValidationResult
    Set result = sut.IsValid("ABC")
    testCon.Assert.AreEqual Invalid, result.result

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("ProductCode.IsValid")
Private Sub TestIsValid_MinLength_ReturnsValid()
    On Error GoTo TestFail

    Dim sut As New ProductCode
    Dim result As ValidationResult
    Set result = sut.IsValid("ABCD")
    testCon.Assert.AreEqual Valid, result.result

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("ProductCode.IsValid")
Private Sub TestIsValid_MaxLength_ReturnsValid()
    On Error GoTo TestFail

    Dim sut As New ProductCode
    Dim result As ValidationResult
    Set result = sut.IsValid("ABCDEFGHIJ0123456789")
    testCon.Assert.AreEqual Valid, result.result

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("ProductCode.IsValid")
Private Sub TestIsValid_TooLong_ReturnsInvalid()
    On Error GoTo TestFail

    Dim sut As New ProductCode
    Dim result As ValidationResult
    Set result = sut.IsValid("ABCDEFGHIJ01234567890")
    testCon.Assert.AreEqual Invalid, result.result

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("ProductCode.IsValid")
Private Sub TestIsValid_LowercaseChars_ReturnsInvalid()
    On Error GoTo TestFail

    Dim sut As New ProductCode
    Dim result As ValidationResult
    Set result = sut.IsValid("abcd")
    testCon.Assert.AreEqual Invalid, result.result

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("ProductCode.IsValid")
Private Sub TestIsValid_SpecialChars_ReturnsInvalid()
    On Error GoTo TestFail

    Dim sut As New ProductCode
    Dim result As ValidationResult
    Set result = sut.IsValid("AB-C")
    testCon.Assert.AreEqual Invalid, result.result

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("ProductCode.IsValid")
Private Sub TestIsValid_ValidValue_ReturnsValid()
    On Error GoTo TestFail

    Dim sut As New ProductCode
    Dim result As ValidationResult
    Set result = sut.IsValid("ABCD1234")
    testCon.Assert.AreEqual Valid, result.result

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
