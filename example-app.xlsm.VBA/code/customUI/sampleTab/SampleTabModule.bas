Attribute VB_Name = "SampleTabModule"
'@IgnoreModule ParameterNotUsed, ParameterCanBeByVal
'@Folder "customUI.sampleTab"
Option Explicit
'Keep the contents of the text box to module variable.
'テキストボックスの内容は自分で保持する

Private Type TRibbon
    SampleText As String
End Type

Private this As TRibbon

'@Ignore ProcedureNotUsed
Public Sub SampleText_getText(ByRef control As IRibbonControl, ByRef Text As Variant)
    Text = this.SampleText
End Sub

'@Ignore ProcedureNotUsed
Public Sub SampleText_onChange(ByRef control As IRibbonControl, ByRef Text As Variant)
    this.SampleText = Text
End Sub

'@Ignore ProcedureNotUsed
Public Sub SampleButton_onAction(ByVal control As IRibbonControl)
    On Error GoTo ErrorHandler
    Excel.Application.EnableEvents = False
    MsgBox "Clicked Search Button!" & vbNewLine _
         & "Text Box value is " & this.SampleText, _
           vbSystemModal + vbInformation + vbOKOnly
ErrorHandler:
    Excel.Application.EnableEvents = True
    On Error GoTo 0
    If Err.Number = 0 Then Exit Sub
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

'@Ignore ProcedureNotUsed
Public Sub AppIntroductionButton_onAction(ByVal control As IRibbonControl)
    On Error GoTo ErrorHandler
    Excel.Application.EnableEvents = False
    ThisAppModule.showAppIntroduction
ErrorHandler:
    Excel.Application.EnableEvents = True
    On Error GoTo 0
    If Err.Number = 0 Then Exit Sub
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

'@Ignore ProcedureNotUsed
Public Sub AppVersionsButton_onAction(ByVal control As IRibbonControl)
    On Error GoTo ErrorHandler
    Excel.Application.EnableEvents = False
    ThisAppModule.showVersion
ErrorHandler:
    Excel.Application.EnableEvents = True
    On Error GoTo 0
    If Err.Number = 0 Then Exit Sub
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

'@Ignore ProcedureNotUsed
Public Sub RegisterProductButton_onAction(ByVal control As IRibbonControl)
    On Error GoTo ErrorHandler
    Excel.Application.EnableEvents = False
    ShowFormModule.registerProduct
ErrorHandler:
    Excel.Application.EnableEvents = True
    On Error GoTo 0
    If Err.Number = 0 Then Exit Sub
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub
