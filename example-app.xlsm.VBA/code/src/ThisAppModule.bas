Attribute VB_Name = "ThisAppModule"
'@Folder("src")
Option Explicit
Option Private Module

Public Sub ShowAppIntroduction()
    MsgBox "This app is example-app.", vbOKOnly + vbSystemModal + vbInformation
End Sub

Public Sub ShowVersion(Optional ByVal manager As VersionManager)
    If manager Is Nothing Then Set manager = New VersionManager: manager.Create ThisWorkbook
    MsgBox "example-app" & vbNewLine & manager.Version, vbOKOnly + vbSystemModal + vbInformation
End Sub
