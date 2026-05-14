Attribute VB_Name = "CanRefToggleModule"
'@Folder("utils")
Option Explicit
Option Private Module

' ===== Procedures to toggle #Const canRef in TestController =====

'@EntryPoint
Public Sub SetCanRefTrue()
    On Error GoTo ErrorHandler
    SetCanRefValue True
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description
End Sub

'@EntryPoint
Public Sub SetCanRefFalse()
    On Error GoTo ErrorHandler
    SetCanRefValue False
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description
End Sub

Private Sub SetCanRefValue(ByVal newValue As Boolean)
    Dim vbMod As VBIDE.CodeModule
    Set vbMod = ThisWorkbook.VBProject.VBComponents.Item("TestController").CodeModule

    Dim lineNum As Long
    For lineNum = 1 To vbMod.CountOfLines
        If InStr(vbMod.lines(lineNum, 1), "#Const canRef") = 0 Then GoTo continue
        vbMod.ReplaceLine lineNum, CreateNewCanRefLine(newValue)
        Exit For
continue:
    Next lineNum
End Sub

Private Function CreateNewCanRefLine(ByVal newValue As Boolean) As String
    CreateNewCanRefLine = "#Const canRef = " & Format$(newValue, "True/False")
End Function
