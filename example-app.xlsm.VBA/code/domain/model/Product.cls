VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Product"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("domain.model")
Option Explicit

Private Type TProduct
    Code As ProductCode
    Name As ProductName
End Type

Private this As TProduct

Public Property Get Code() As ProductCode
    Set Code = this.Code
End Property

Public Property Get Name() As ProductName
    Set Name = this.Name
End Property

Private Sub Class_Initialize()
    Set this.Code = New ProductCode
    Set this.Name = New ProductName
End Sub
