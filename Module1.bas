Attribute VB_Name = "Module1"
Public Function pReplace(strExpression As String, strFind As String, strReplace As String)
    Dim intX As Integer


    If (Len(strExpression) - Len(strFind)) >= 0 Then


        For intX = 1 To Len(strExpression)


            If Mid(strExpression, intX, Len(strFind)) = strFind Then
                strExpression = Left(strExpression, (intX - 1)) + strReplace + Mid(strExpression, intX + Len(strFind), Len(strExpression))
            End If
        Next
    End If
    pReplace = strExpression
End Function
