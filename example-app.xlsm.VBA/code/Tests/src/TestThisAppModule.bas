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
