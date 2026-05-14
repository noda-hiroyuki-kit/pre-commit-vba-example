VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SemanticVersioning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "utils"
Option Explicit

Private Enum ComparisonResult
    LessThan = -1
    EqualTo = 0
    GreaterThan = 1
End Enum

Public Sub ensureSemVer(ByVal versionText As String, ByVal role As String)
    If isSemVer(StripVersionPrefix(versionText)) Then Exit Sub
    Err.Raise AppError.INVALID_SEMVER, "SemanticVersioning", role & " version must follow SemVer: " & versionText
End Sub

Public Sub ensureNoRollback(ByVal currentVersion As String, ByVal nextVersion As String)
    If Compare(StripVersionPrefix(nextVersion), StripVersionPrefix(currentVersion)) >= ComparisonResult.EqualTo Then Exit Sub
    Err.Raise AppError.VERSION_ROLLBACK_NOT_ALLOWED, "SemanticVersioning", "Version rollback is not allowed: " & currentVersion & " -> " & nextVersion
End Sub

Private Function StripVersionPrefix(ByVal versionText As String) As String
    If Len(versionText) > 0 And LCase$(Left$(versionText, 1)) = "v" Then
        StripVersionPrefix = Mid$(versionText, 2)
    Else
        StripVersionPrefix = versionText
    End If
End Function

Private Function isSemVer(ByVal versionText As String) As Boolean
    Dim regEx As RegExp
    Set regEx = New RegExp
    regEx.Pattern = "^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)(?:-(?:0|[1-9]\d*|\d*[A-Za-z-][0-9A-Za-z-]*)(?:\.(?:0|[1-9]\d*|\d*[A-Za-z-][0-9A-Za-z-]*))*)?(?:\+[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?$"
    isSemVer = regEx.Test(Trim$(versionText))
End Function

Private Function Compare(ByVal leftVersion As String, ByVal rightVersion As String) As ComparisonResult
    Dim leftParts() As String
    Dim rightParts() As String
    Dim leftCore As String
    Dim rightCore As String
    Dim leftPre As String
    Dim rightPre As String

    leftParts = splitVersion(normalize(leftVersion))
    leftCore = leftParts(0)
    leftPre = leftParts(1)

    rightParts = splitVersion(normalize(rightVersion))
    rightCore = rightParts(0)
    rightPre = rightParts(1)

    Compare = compareCore(leftCore, rightCore)
    If Compare <> ComparisonResult.EqualTo Then Exit Function
    Compare = comparePreRelease(leftPre, rightPre)
End Function

Private Function normalize(ByVal versionText As String) As String
    normalize = Trim$(versionText)
    If LCase$(Left$(normalize, 1)) = "v" Then normalize = Mid$(normalize, 2)
End Function

Private Function splitVersion(ByVal versionText As String) As String()
    Dim result(0 To 1) As String
    Dim plusPos As Long
    Dim dashPos As Long

    plusPos = InStr(versionText, "+")
    If plusPos > 0 Then versionText = Left$(versionText, plusPos - 1)

    dashPos = InStr(versionText, "-")
    If dashPos = 0 Then
        result(0) = versionText
        result(1) = vbNullString
    Else
        result(0) = Left$(versionText, dashPos - 1)
        result(1) = Mid$(versionText, dashPos + 1)
    End If

    splitVersion = result
End Function

Private Function compareCore(ByVal leftText As String, ByVal rightText As String) As ComparisonResult
    Dim leftParts() As String:  leftParts = Split(leftText, ".")
    Dim rightParts() As String: rightParts = Split(rightText, ".")

    Dim i As Long
    For i = 0 To 2
        If CLng(leftParts(i)) < CLng(rightParts(i)) Then compareCore = ComparisonResult.LessThan:    Exit Function
        If CLng(leftParts(i)) > CLng(rightParts(i)) Then compareCore = ComparisonResult.GreaterThan: Exit Function
    Next i
End Function

Private Function comparePreRelease(ByVal leftText As String, ByVal rightText As String) As ComparisonResult
    If Len(leftText) = 0 And Len(rightText) = 0 Then Exit Function
    If Len(leftText) = 0 Then comparePreRelease = ComparisonResult.GreaterThan: Exit Function
    If Len(rightText) = 0 Then comparePreRelease = ComparisonResult.LessThan: Exit Function
    comparePreRelease = compareIdentifiers(leftText, rightText)
End Function

Private Function compareIdentifiers(ByVal leftText As String, ByVal rightText As String) As ComparisonResult
    Dim leftParts() As String:  leftParts = Split(leftText, ".")
    Dim rightParts() As String: rightParts = Split(rightText, ".")
    Dim maxIndex As Long:       maxIndex = UBound(leftParts)
    If UBound(rightParts) > maxIndex Then maxIndex = UBound(rightParts)
    Dim i As Long
    For i = 0 To maxIndex
        If i > UBound(leftParts) Then compareIdentifiers = ComparisonResult.LessThan:     Exit Function
        If i > UBound(rightParts) Then compareIdentifiers = ComparisonResult.GreaterThan: Exit Function
        Dim c As ComparisonResult
        c = compareIdentifier(leftParts(i), rightParts(i))
        If c <> ComparisonResult.EqualTo Then compareIdentifiers = c: Exit Function
    Next i
End Function

Private Function compareIdentifier(ByVal leftText As String, ByVal rightText As String) As ComparisonResult
    Dim leftNum As Boolean:  leftNum = isNumericIdentifier(leftText)
    Dim rightNum As Boolean: rightNum = isNumericIdentifier(rightText)
    If leftNum And rightNum Then compareIdentifier = Sgn(CLng(leftText) - CLng(rightText)): Exit Function
    If leftNum And Not rightNum Then compareIdentifier = ComparisonResult.LessThan: Exit Function
    If Not leftNum And rightNum Then compareIdentifier = ComparisonResult.GreaterThan: Exit Function
    compareIdentifier = Sgn(StrComp(leftText, rightText, vbBinaryCompare))
End Function

Private Function isNumericIdentifier(ByVal valueText As String) As Boolean
    Dim regEx As RegExp
    Set regEx = New RegExp
    regEx.Pattern = "^(0|[1-9]\d*)$"
    isNumericIdentifier = regEx.Test(valueText)
End Function
