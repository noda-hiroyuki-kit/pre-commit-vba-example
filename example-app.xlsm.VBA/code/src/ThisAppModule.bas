Attribute VB_Name = "ThisAppModule"
'@Folder("src")
Option Explicit
Option Private Module

Public Sub ShowAppIntroduction()
    MsgBox "This app is example-app.", vbOKOnly + vbSystemModal + vbInformation
End Sub

Public Sub ShowVersion(Optional ByVal manager As VersionManager)
    Dim managerLocal As VersionManager: Set managerLocal = manager
    If managerLocal Is Nothing Then Set managerLocal = New VersionManager: managerLocal.Create ThisWorkbook
    MsgBox "example-app" & vbNewLine & managerLocal.Version, vbOKOnly + vbSystemModal + vbInformation
End Sub
