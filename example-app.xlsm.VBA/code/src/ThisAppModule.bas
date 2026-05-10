Attribute VB_Name = "ThisAppModule"
'@Folder("src")
Option Explicit
Option Private Module

Public Sub showAppIntroduction()
    MsgBox "This app is example-app.", vbOKOnly + vbSystemModal + vbInformation
End Sub

Public Sub showVersion(Optional ByVal manager As VersionManager = Nothing)
    If manager Is Nothing Then Set manager = New VersionManager: manager.Create ThisWorkbook
    MsgBox "example-app" & vbNewLine & manager.Version, vbOKOnly + vbSystemModal + vbInformation
End Sub
