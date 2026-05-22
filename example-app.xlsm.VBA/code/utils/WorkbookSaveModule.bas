Attribute VB_Name = "WorkbookSaveModule"
'@Folder("utils")
Option Explicit
Option Private Module

Public Sub SaveWorkbookSilently()
    SaveWorkbookSilentlyByName ThisWorkbook.Name
End Sub

Public Sub SaveWorkbookSilentlyByName(ByVal workbookName As String)
    Dim wb As Workbook
    Dim previousDisplayAlerts As Boolean

    On Error GoTo ErrorHandler
    Set wb = Application.Workbooks(workbookName)

    previousDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    wb.Save
    Application.DisplayAlerts = previousDisplayAlerts
    Exit Sub

ErrorHandler:
    Application.DisplayAlerts = previousDisplayAlerts
    Err.Raise Err.Number, "WorkbookSaveModule.SaveWorkbookSilentlyByName", Err.Description
End Sub
