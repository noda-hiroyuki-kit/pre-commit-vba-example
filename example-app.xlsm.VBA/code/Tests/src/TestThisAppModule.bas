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
    Dim backup As String: backup = ThisWorkbook.BuiltinDocumentProperties.Item("Document Version").Value
    testCon.Fakes.MsgBox.Returns vbOK
    testCon.Fakes.InputBox.Returns "1.0.0"
    ThisAppModule.setVersion
    'Act:
    ThisAppModule.showVersion
    'Assert:
    With testCon.Fakes.MsgBox.Verify
        .Parameter testCon.Fakes.Params.MsgBox.Prompt, _
            "example-app" & vbNewLine & "1.0.0"
    End With

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    ThisWorkbook.BuiltinDocumentProperties.Item("Document Version").Value = backup
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("setVersion")
Private Sub TestSetVersionUpdatesDocumentVersion()
    On Error GoTo TestFail

    'Arrange:
    Dim backup As String: backup = ThisWorkbook.BuiltinDocumentProperties.Item("Document Version").Value
    testCon.Fakes.InputBox.Returns "2.0.0"
    'Act:
    ThisAppModule.setVersion
    'Assert:
    testCon.Assert.AreEqual "2.0.0", _
        ThisWorkbook.BuiltinDocumentProperties.Item("Document Version").Value

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    ThisWorkbook.BuiltinDocumentProperties.Item("Document Version").Value = backup
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("setVersion")
Private Sub TestSetVersionDoesNothingWhenInputIsEmpty()
    On Error GoTo TestFail

    'Arrange:
    Dim backup As String: backup = ThisWorkbook.BuiltinDocumentProperties.Item("Document Version").Value
    testCon.Fakes.InputBox.Returns ""
    Dim versionBefore As String
    versionBefore = ThisWorkbook.BuiltinDocumentProperties.Item("Document Version").Value
    'Act:
    ThisAppModule.setVersion
    'Assert:
    testCon.Assert.AreEqual versionBefore, _
        ThisWorkbook.BuiltinDocumentProperties.Item("Document Version").Value

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    ThisWorkbook.BuiltinDocumentProperties.Item("Document Version").Value = backup
    Exit Sub
TestFail:
    testCon.Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
