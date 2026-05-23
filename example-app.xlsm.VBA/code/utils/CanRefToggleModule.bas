Attribute VB_Name = "CanRefToggleModule"
'@Folder("utils")
Option Explicit
Option Private Module

Private Const TARGET_COMPONENT As String = "TestController"
Private Const TARGET_CONST As String = "#Const canRef"

'@EntryPoint
Public Sub SetCanRefTrue()
    setCanRefValueWithContext True, "CanRefToggleModule.SetCanRefTrue"
End Sub

'@EntryPoint
Public Sub SetCanRefFalse()
    setCanRefValueWithContext False, "CanRefToggleModule.SetCanRefFalse"
End Sub

Private Sub setCanRefValueWithContext(ByVal newValue As Boolean, ByVal sourceName As String)
    On Error GoTo ErrorHandler
    SetCanRefValue newValue
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, sourceName, Err.Description
End Sub

Private Sub SetCanRefValue(ByVal newValue As Boolean)
    Dim moduleRef As VBIDE.CodeModule
    Set moduleRef = GetTestControllerCodeModule()

    Dim targetLine As Long
    targetLine = FindCanRefLine(moduleRef)

    If targetLine = 0 Then RaiseCanRefNotFound
    moduleRef.ReplaceLine targetLine, CreateNewCanRefLine(newValue)
End Sub

Private Function GetTestControllerCodeModule() As VBIDE.CodeModule
    On Error GoTo AccessError

    Dim projectRef As VBIDE.VBProject
    Set projectRef = ThisWorkbook.VBProject

    Set GetTestControllerCodeModule = GetComponentCodeModule(projectRef, TARGET_COMPONENT)
    Exit Function
AccessError:
    RaiseCodeModuleAccessError Err.Description
End Function

Private Function GetComponentCodeModule(ByVal projectRef As VBIDE.VBProject, ByVal componentName As String) As VBIDE.CodeModule
    Dim componentRef As VBIDE.VBComponent
    Set componentRef = projectRef.VBComponents.Item(componentName)
    Set GetComponentCodeModule = componentRef.CodeModule
End Function

Private Function FindCanRefLine(ByVal moduleRef As VBIDE.CodeModule) As Long
    Dim lineNum As Long
    For lineNum = 1 To moduleRef.CountOfLines
        If IsCanRefLine(moduleRef, lineNum) Then FindCanRefLine = lineNum: Exit Function
    Next lineNum
End Function

Private Function IsCanRefLine(ByVal moduleRef As VBIDE.CodeModule, ByVal lineNum As Long) As Boolean
    Dim oneLine As String
    oneLine = moduleRef.lines(lineNum, 1)
    IsCanRefLine = (InStr(oneLine, TARGET_CONST) > 0)
End Function

Private Function CreateNewCanRefLine(ByVal newValue As Boolean) As String
    CreateNewCanRefLine = TARGET_CONST & " = " & BoolText(newValue)
End Function

Private Function BoolText(ByVal Value As Boolean) As String
    If Value Then BoolText = "True": Exit Function
    BoolText = "False"
End Function

Private Sub RaiseCanRefNotFound()
    Err.Raise AppError.NOT_FOUND_CAN_REF, "CanRefToggleModule.SetCanRefValue", _
              "#Const canRef directive not found in TestController module."
End Sub

Private Sub RaiseCodeModuleAccessError(ByVal details As String)
    Err.Raise AppError.CAN_REF_CODE_MODULE_ACCESS_FAILED, "CanRefToggleModule.GetTestControllerCodeModule", _
              "Unable to access TestController code module. Ensure VBA project access settings are enabled. Details: " & details
End Sub
