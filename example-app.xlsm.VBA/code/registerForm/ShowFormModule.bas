Attribute VB_Name = "ShowFormModule"
'@Folder "registerForm"
Option Explicit
Option Private Module

Public Sub registerProduct()
    Dim form As IForm
    Set form = New RegisterProductForm

    Dim Product As Product: Set Product = New Product
    If Not form.ShowForm(Product) Then Exit Sub
    MsgBox "Product registered as follows:" & vbNewLine _
         & " code: " & Product.code.Value & vbNewLine _
         & " name: " & Product.Name.Value, _
           vbSystemModal + vbInformation + vbOKOnly
End Sub
