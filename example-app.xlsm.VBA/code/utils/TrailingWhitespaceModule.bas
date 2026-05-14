Attribute VB_Name = "TrailingWhitespaceModule"
'@Folder("utils")
Option Explicit
Option Private Module

'@EntryPoint
Public Sub CleanAllModulesInWorkbook()
    ProcessEachModule ThisWorkbook.VBProject
    MsgBox "すべてのモジュール内の末尾空白を削除しました。", vbInformation
End Sub

Private Sub ProcessEachModule(ByVal vbProj As VBProject)
    Dim component As VBComponent
    For Each component In vbProj.VBComponents
        CleanModuleCode component
    Next component
End Sub

Private Sub CleanModuleCode(ByRef component As VBComponent)
    Dim originalCode As String
    originalCode = component.CodeModule.lines(1, component.CodeModule.CountOfLines)
    Dim cleanedCode As String
    cleanedCode = RemoveTrailingWhitespaceFromText(originalCode)
    If originalCode <> cleanedCode Then
        ReplaceModuleCode component, cleanedCode
    End If
End Sub

Private Sub ReplaceModuleCode(ByRef component As VBComponent, ByVal newCode As String)
    component.CodeModule.DeleteLines 1, component.CodeModule.CountOfLines
    component.CodeModule.InsertLines 1, newCode
End Sub

Private Function RemoveTrailingWhitespaceFromText(ByVal Text As String) As String
    Dim lines() As String: lines = Split(Text, vbNewLine)
    RemoveTrailingWhitespaceFromText = JoinCleanedLines(lines)
End Function

Private Function JoinCleanedLines(ByRef lines() As String) As String
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim result As String
        result = result & IIf(i > LBound(lines), vbNewLine, vbNullString) & RemoveLineTrailingWhitespace(lines(i))
    Next i
    JoinCleanedLines = result
End Function

Private Function RemoveLineTrailingWhitespace(ByVal line As String) As String
    RemoveLineTrailingWhitespace = RegexReplace(line, "\s+$", vbNullString)
End Function

Private Function RegexReplace(ByVal inputText As String, ByVal Pattern As String, ByVal replacement As String) As String
    Dim regEx As RegExp: Set regEx = New RegExp
    regEx.Pattern = Pattern
    regEx.Global = True
    RegexReplace = regEx.Replace(inputText, replacement)
End Function
