VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VersionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "domain"
Option Explicit

Public Sub ApplyToWorkbook(ByVal targetWorkbook As Workbook, ByVal nextVersion As String)
    Dim currentVersion As String
    currentVersion = CStr(targetWorkbook.BuiltinDocumentProperties.Item("Document Version").Value)
    ensureSemVer nextVersion, "next"
    If Len(Trim$(currentVersion)) > 0 Then
        ensureSemVer currentVersion, "current"
        ensureNoRollback currentVersion, nextVersion
    End If
    targetWorkbook.BuiltinDocumentProperties.Item("Document Version").Value = nextVersion
End Sub

Private Sub ensureSemVer(ByVal versionText As String, ByVal role As String)
    If isSemVer(versionText) Then Exit Sub
    Err.Raise vbObjectError + 1100, "VersionManager", role & " version must follow SemVer: " & versionText
End Sub

Private Sub ensureNoRollback(ByVal currentVersion As String, ByVal nextVersion As String)
    If Compare(nextVersion, currentVersion) >= 0 Then Exit Sub
    Err.Raise vbObjectError + 1101, "VersionManager", "Version rollback is not allowed: " & currentVersion & " -> " & nextVersion
End Sub

Private Function isSemVer(ByVal versionText As String) As Boolean
    Dim regEx As RegExp
    Set regEx = New RegExp
    regEx.Pattern = "^(?:v)?(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)(?:-(?:0|[1-9]\d*|\d*[A-Za-z-][0-9A-Za-z-]*)(?:\.(?:0|[1-9]\d*|\d*[A-Za-z-][0-9A-Za-z-]*))*)?(?:\+[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?$"
    isSemVer = regEx.Test(Trim$(versionText))
End Function

Private Function Compare(ByVal leftVersion As String, ByVal rightVersion As String) As Long
    Dim leftCore As String
    Dim rightCore As String
    Dim leftPre As String
    Dim rightPre As String
    splitVersion normalize(leftVersion), leftCore, leftPre
    splitVersion normalize(rightVersion), rightCore, rightPre
    Compare = compareCore(leftCore, rightCore)
    If Compare <> 0 Then Exit Function
    Compare = comparePreRelease(leftPre, rightPre)
End Function

Private Function normalize(ByVal versionText As String) As String
    normalize = Trim$(versionText)
    If LCase$(Left$(normalize, 1)) = "v" Then normalize = Mid$(normalize, 2)
End Function

Private Sub splitVersion(ByVal versionText As String, ByRef core As String, ByRef pre As String)
    Dim plusPos As Long
    plusPos = InStr(versionText, "+")
    If plusPos > 0 Then versionText = Left$(versionText, plusPos - 1)
    Dim dashPos As Long
    dashPos = InStr(versionText, "-")
    If dashPos = 0 Then
        core = versionText
        pre = vbNullString
        Exit Sub
    End If
    core = Left$(versionText, dashPos - 1)
    pre = Mid$(versionText, dashPos + 1)
End Sub

Private Function compareCore(ByVal leftText As String, ByVal rightText As String) As Long
    Dim leftParts() As String
    Dim rightParts() As String
    Dim i As Long
    leftParts = Split(leftText, ".")
    rightParts = Split(rightText, ".")
    For i = 0 To 2
        If CLng(leftParts(i)) < CLng(rightParts(i)) Then compareCore = -1: Exit Function
        If CLng(leftParts(i)) > CLng(rightParts(i)) Then compareCore = 1: Exit Function
    Next i
End Function

Private Function comparePreRelease(ByVal leftText As String, ByVal rightText As String) As Long
    If Len(leftText) = 0 And Len(rightText) = 0 Then Exit Function
    If Len(leftText) = 0 Then comparePreRelease = 1: Exit Function
    If Len(rightText) = 0 Then comparePreRelease = -1: Exit Function
    comparePreRelease = compareIdentifiers(leftText, rightText)
End Function

Private Function compareIdentifiers(ByVal leftText As String, ByVal rightText As String) As Long
    Dim leftParts() As String
    Dim rightParts() As String
    leftParts = Split(leftText, ".")
    rightParts = Split(rightText, ".")
    Dim maxIndex As Long
    maxIndex = UBound(leftParts)
    If UBound(rightParts) > maxIndex Then maxIndex = UBound(rightParts)
    Dim i As Long
    For i = 0 To maxIndex
        If i > UBound(leftParts) Then compareIdentifiers = -1: Exit Function
        If i > UBound(rightParts) Then compareIdentifiers = 1: Exit Function
        Dim c As Long
        c = compareIdentifier(leftParts(i), rightParts(i))
        If c <> 0 Then compareIdentifiers = c: Exit Function
    Next i
End Function

Private Function compareIdentifier(ByVal leftText As String, ByVal rightText As String) As Long
    Dim leftNum As Boolean
    Dim rightNum As Boolean
    leftNum = isNumericIdentifier(leftText)
    rightNum = isNumericIdentifier(rightText)
    If leftNum And rightNum Then compareIdentifier = Sgn(CLng(leftText) - CLng(rightText)): Exit Function
    If leftNum And Not rightNum Then compareIdentifier = -1: Exit Function
    If Not leftNum And rightNum Then compareIdentifier = 1: Exit Function
    compareIdentifier = Sgn(StrComp(leftText, rightText, vbBinaryCompare))
End Function

Private Function isNumericIdentifier(ByVal valueText As String) As Boolean
    Dim regEx As RegExp
    Set regEx = New RegExp
    regEx.Pattern = "^(0|[1-9]\\d*)$"
    isNumericIdentifier = regEx.Test(valueText)
End Function
