Attribute VB_Name = "TrailingWhitespaceModule"
'@Folder("utils")
Option Explicit
Option Private Module

'@EntryPoint
Public Sub CleanAllModulesInWorkbook()
    On Error GoTo CleanAllModulesInWorkbookError
    ProcessEachModule ThisWorkbook.VBProject
    MsgBox "Removed trailing whitespace from all modules.", vbInformation + vbSystemModal + vbOKOnly
    Exit Sub
CleanAllModulesInWorkbookError:
    MsgBox "Unable to clean VBA modules in this workbook." & vbNewLine & vbNewLine & _
           "Programmatic access to the VBA project may be disabled, or one of the components could not be edited." & vbNewLine & vbNewLine & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbExclamation + vbSystemModal + vbOKOnly, _
           "Clean All Modules"
End Sub

Private Sub ProcessEachModule(ByVal vbProj As VBProject)
    Dim component As VBComponent
    For Each component In vbProj.VBComponents
        CleanModuleCode component
    Next component
End Sub

Private Sub CleanModuleCode(ByRef component As VBComponent)
    If component.CodeModule.CountOfLines = 0 Then Exit Sub
    Dim originalCode As String
    originalCode = component.CodeModule.lines(1, component.CodeModule.CountOfLines)

    Dim cleanedCode As String
    cleanedCode = RemoveTrailingWhitespaceFromText(originalCode)
    ReplaceModuleCode component, originalCode, cleanedCode
End Sub

'@Ignore ParameterCanBeByVal
Private Sub ReplaceModuleCode(ByRef component As VBComponent, ByVal originalCode As String, ByVal newCode As String)
    If originalCode = newCode Then Exit Sub
    component.CodeModule.DeleteLines 1, component.CodeModule.CountOfLines
    component.CodeModule.InsertLines 1, newCode
End Sub

Private Function RemoveTrailingWhitespaceFromText(ByVal Text As String) As String
    Dim lines() As String: lines = Split(Text, vbNewLine)
    RemoveTrailingWhitespaceFromText = JoinCleanedLines(lines)
End Function

Private Function JoinCleanedLines(ByRef lines() As String) As String
    Dim index As Long
    Dim cleaned() As String: ReDim cleaned(LBound(lines) To UBound(lines))
    For index = LBound(lines) To UBound(lines)
        cleaned(index) = RemoveLineTrailingWhitespace(lines(index))
    Next index
    JoinCleanedLines = Join(cleaned, vbNewLine)
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
