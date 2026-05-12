Attribute VB_Name = "TestThisAppModule"
'@TestModule
'@Folder "Tests.src"


Option Explicit
Option Private Module

Private testCon As TestController

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set testCon = New TestController
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set testCon = Nothing
End Sub

'@TestInitialize
'@Ignore EmptyMethod
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
'@Ignore EmptyMethod
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("showAppIntroduction")
Private Sub TestShowAppIntroductionDisplaysExpectedMessage()
    On Error GoTo TestFail

    'Arrange:
    testCon.Fakes.MsgBox.Returns vbOK
    'Act:
    ThisAppModule.showAppIntroduction
    'Assert:
    With testCon.Fakes.MsgBox.Verify
        .Parameter testCon.Fakes.Params.MsgBox.Prompt, "This app is example-app."
    End With

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("showVersion")
Private Sub TestShowVersionDisplaysExpectedMessage()
    On Error GoTo TestFail

    'Arrange:
    Dim sut As VersionManager: Set sut = New VersionManager
    sut.CreateForTest "v0.9.0"
    testCon.Fakes.MsgBox.Returns vbOK
    testCon.Fakes.InputBox.Returns "v1.0.0"
    VersionInfoModule.setVersion sut
    'Act:
    ThisAppModule.showVersion sut
    'Assert:
    With testCon.Fakes.MsgBox.Verify
        .Parameter testCon.Fakes.Params.MsgBox.Prompt, _
            "example-app" & vbNewLine & "v1.0.0"
    End With

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
'    ThisWorkbook.BuiltinDocumentProperties.Item("Document Version").Value = backup
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("setVersion")
Private Sub TestSetVersionUpdatesDocumentVersion()
    On Error GoTo TestFail

    'Arrange:
    Dim sut As VersionManager: Set sut = New VersionManager
    sut.CreateForTest "v1.9.9"
    testCon.Fakes.InputBox.Returns "v2.0.0"
    'Act:
    VersionInfoModule.setVersion sut
    'Assert:
    testCon.Assert.AreEqual "v2.0.0", sut.Version

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("setVersion")
Private Sub TestSetVersionDoesNothingWhenInputIsEmpty()
    On Error GoTo TestFail

    'Arrange:
    Dim sut As VersionManager: Set sut = New VersionManager
    sut.CreateForTest "v3.0.0"
    testCon.Fakes.InputBox.Returns vbNullString
    Dim versionBefore As String: versionBefore = sut.Version
    'Act:
    VersionInfoModule.setVersion sut
    'Assert:
    testCon.Assert.AreEqual versionBefore, sut.Version

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("setVersion")
Private Sub TestSetVersionRejectsRollback()
    On Error GoTo TestFail

    'Arrange:
    Dim sut As VersionManager: Set sut = New VersionManager
    sut.CreateForTest "2.1.0"
    testCon.Fakes.InputBox.Returns "2.0.9"
    Dim versionBefore As String
    versionBefore = sut.Version
    'Act:
    On Error Resume Next
    VersionInfoModule.setVersion
    Dim raisedNumber As Long
    raisedNumber = Err.Number
    On Error GoTo TestFail
    'Assert:
    testCon.Assert.AreEqual True, raisedNumber <> 0
    testCon.Assert.AreEqual versionBefore, "2.1.0"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
