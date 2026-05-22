Attribute VB_Name = "WorkbookSaveModule"
'@Folder("utils")
Option Explicit
Option Private Module

'@EntryPoint
Public Sub SaveWorkbookSilently()
    SaveWorkbookSilentlyByName ThisWorkbook.Name
End Sub

Public Sub SaveWorkbookSilentlyByName(ByVal workbookName As String)
    On Error GoTo ErrorHandler
    Dim targetWb As Workbook: Set targetWb = Excel.Application.Workbooks.Item(workbookName)

    Application.DisplayAlerts = False
    targetWb.Save
    Application.DisplayAlerts = True
    Exit Sub
ErrorHandler:
    Application.DisplayAlerts = True
    Err.Raise Err.Number, "WorkbookSaveModule.SaveWorkbookSilentlyByName", Err.Description
End Sub
