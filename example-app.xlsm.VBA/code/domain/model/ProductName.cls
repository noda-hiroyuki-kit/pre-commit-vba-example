VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProductName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("domain.model")
Option Explicit
' Product Name Class

Private Const MIN_LENGTH As Long = 1
Private Const MAX_LENGTH As Long = 255

Private Type TProductName
    Value As String
    Created As Boolean
End Type

Private this As TProductName

Private Sub Class_Initialize()
    this.Created = False
    this.Value = vbNullString
End Sub

Public Property Get Value() As String
    If Not this.Created Then Err.Raise AppError.NOT_CREATED, , "Product name has not been created."
    Value = this.Value
End Property

Public Sub Create(ByVal ValueArg As String)
    If this.Created Then Err.Raise AppError.HAS_CREATED, , "Product name has already been created."
    Dim result As ValidationResult: Set result = IsValid(ValueArg)
    If Not result.result = Valid Then Err.Raise AppError.INVALID_VALUE, , result.Message
    this.Value = ValueArg
    this.Created = True
End Sub

Public Function IsValid(ByVal ValueArg As String) As ValidationResult
    Set IsValid = New ValidationResult
    Dim valueLength As Long: valueLength = LenB(StrConv(ValueArg, vbFromUnicode))
    If valueLength < MIN_LENGTH Then
        IsValid.SetInvalid
        IsValid.SetMessage "Name length is " & MIN_LENGTH & " Characters or more."
        Exit Function
    End If
    If valueLength > MAX_LENGTH Then
        IsValid.SetInvalid
        IsValid.SetMessage "Name length is " & MAX_LENGTH & " Characters or less."
        Exit Function
    End If
    IsValid.SetValid
    IsValid.SetMessage "OK"
End Function
