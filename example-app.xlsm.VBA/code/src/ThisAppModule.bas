Attribute VB_Name = "ThisAppModule"
'@Folder("src")
Option Explicit
Option Private Module

Public Sub showAppIntroduction()
    MsgBox "This app is example-app.", vbOKOnly + vbSystemModal + vbInformation
End Sub

Public Sub showVersion()
    MsgBox "example-app" & vbNewLine & ThisWorkbook.BuiltinDocumentProperties.Item("Document Version").Value, _
        vbOKOnly + vbSystemModal + vbInformation
End Sub

Public Sub setVersion()
    Dim response As String
    response = InputBox("Input Version")
    If Len(response) = 0 Then Exit Sub
    setDocumentVersion response
End Sub

Private Sub setDocumentVersion(ByVal version As String)
    ThisWorkbook.BuiltinDocumentProperties.Item("Document Version").Value = version
End Sub
