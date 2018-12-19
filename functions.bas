Function ArrayLen(arr As Variant) As Integer
    ArrayLen = UBound(arr) - LBound(arr) + 1
End Function

Function NumericOnly(strSource As String) As String
    Dim i As Integer
    Dim strResult As String
    For i = 1 To Len(strSource)
        Select Case Asc(Mid(strSource, i, 1))
            Case 48 To 57
                strResult = strResult & Mid(strSource, i, 1)
        End Select
    Next
    NumericOnly = strResult
End Function

Function IsInArray(arr As Variant,str As String) As Boolean
    Dim isVerified As Boolean
    isVerified = False
    For Each elem In arr
        If InStr(str, elem) <> 0 Then
            isVerified = True
        End If
    Next
    ArrayVerification = isVerified
End Function

Function LcFirst(str As String) As String
    LcFirst = LCase(Left(str, 1)) & Right(str, Len(str) - 1)
End Function

Function ColLetter(colNum As Long) As String
    Dim vArr
    vArr = Split(Cells(1, colNum).Address(True, False), "$")
    GetColumnLetter = vArr(0)
End Function