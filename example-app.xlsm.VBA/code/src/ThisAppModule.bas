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
