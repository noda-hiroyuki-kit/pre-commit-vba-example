VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RegisterProductForm
   Caption         =   "RegisterProductForm"
   ClientHeight    =   2136
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3264
   OleObjectBlob   =   "RegisterProductForm.frx":0000
End
Attribute VB_Name = "RegisterProductForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("registerForm")
Option Explicit

Implements IForm

Private Type TProductCodeRegisterForm
    cancelled As Boolean
    Product As Product
    codeResult As ValidationResult
    nameResult As ValidationResult
End Type

Private GREEN As Long
Private RED As Long

Private this As TProductCodeRegisterForm

'@Ignore ProcedureNotUsed
Public Property Get IsCancelled() As Boolean
    IsCancelled = this.cancelled
End Property

Private Sub CancelCommandButton_Click()
    OnCancel
End Sub

Private Function IForm_ShowForm(ByVal viewModel As Object) As Boolean
    Set this.Product = viewModel
    this.cancelled = False
    textBoxChange
    Me.Show
    IForm_ShowForm = Not this.cancelled
End Function

Private Sub ProductCodeTextBox_Change()
    textBoxChange
End Sub

Private Sub ProductNameTextBox_Change()
    textBoxChange
End Sub

Private Sub textBoxChange()
    Set this.codeResult = this.Product.Code.IsValid(ProductCodeTextBox.Text)
    Set this.nameResult = this.Product.Name.IsValid(ProductNameTextBox.Text)
    ValidateForm
End Sub

Private Sub RegisterCommandButton_Click()
    On Error GoTo ErrorHandler
    this.Product.Code.Create ProductCodeTextBox.Text
    this.Product.Name.Create ProductNameTextBox.Text
    Me.Hide
ErrorHandler:
    On Error GoTo 0
    If Err.Number = 0 Then Exit Sub
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Sub UserForm_Initialize()
    GREEN = RGB(0, 255, 0)
    RED = RGB(255, 0, 0)
End Sub

Private Sub UserForm_QueryClose(ByRef Cancel As Integer, ByRef CloseMode As Integer)
    If Not CloseMode = VbQueryClose.vbFormControlMenu Then Exit Sub
    Cancel = True
    OnCancel
End Sub

Private Sub OnCancel()
    this.cancelled = True
    Me.Hide
End Sub

Private Sub ValidateForm()
    Me.RegisterCommandButton.Enabled = _
                                     (this.codeResult.result = Valid) _
                                     And (this.nameResult.result = Valid)
    setLabelProperties Me.CodeValidationMessageLabel, this.codeResult
    setLabelProperties Me.NameValidationMessageLabel, this.nameResult
End Sub

Private Sub setLabelProperties(ByVal control As Object, ByVal result As ValidationResult)
    With control
        .Caption = result.Message
        .ForeColor = setCaptionColor(result)
    End With
End Sub

Private Function setCaptionColor(ByVal result As ValidationResult) As Long
    If result.result = Valid Then setCaptionColor = GREEN: Exit Function
    setCaptionColor = RED
End Function
