Attribute VB_Name = "ShowFormModule"
'@Folder "registerForm"
Option Explicit
Option Private Module

Public Sub registerProduct()
    Dim form As IForm
    Set form = New RegisterProductForm

    Dim newProduct As Product: Set newProduct = New Product
    If Not form.ShowForm(newProduct) Then Exit Sub
    MsgBox "Product registered as follows:" & vbNewLine _
         & " code: " & newProduct.code.Value & vbNewLine _
         & " name: " & newProduct.Name.Value, _
           vbSystemModal + vbInformation + vbOKOnly
End Sub
