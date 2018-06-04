Attribute VB_Name = "modFunctions"
Function contains(ByVal subject As String, ByVal containsTxt As String) As Boolean
    contains = InStr(1, subject, containsTxt) > 0
End Function

Function formatPhoneNumber(ByVal phone As String) As String
    If Not contains(phone, "-") Then
        If Len(phone) = 7 Then
            formatPhoneNumber = Left(phone, 3) & "-" & Right(phone, 4)
        ElseIf Len(phone) = 10 Then
            formatPhoneNumber = "(" & Left(phone, 3) & ") " & MiD(phone, 4, 3) & "-" & Right(phone, 4)
        ElseIf Len(phone) = 11 Then
            formatPhoneNumber = Left(phone, 1) & "-" & MiD(phone, 2, 3) & "-" & MiD(phone, 5, 3) & "-" & Right(phone, 4)
        Else
            formatPhoneNumber = phone
        End If
    Else
        formatPhoneNumber = phone
    End If
End Function

Function truncatePhoneNumber(ByVal phone As String) As String
    If Len(phone) = 7 Then
        truncatePhoneNumber = formatPhoneNumber(phone)
    ElseIf Len(phone) = 8 Then
        truncatePhoneNumber = phone
    ElseIf Len(phone) = 10 Then
        If Not contains(phone, "-") Then
            truncatePhoneNumber = formatPhoneNumber(Right(phone, 7))
        Else
            truncatePhoneNumber = Right(phone, 8)
        End If
    ElseIf Len(phone) = 11 Then
        If contains(phone, "-") Then
            truncatePhoneNumber = Right(phone, 7)
        Else
            truncatePhoneNumber = formatPhoneNumber(Right(phone, 7))
        End If
    ElseIf Len(phone) = 12 Then
        If contains(phone, "-") Then
            truncatePhoneNumber = Right(phone, 8)
        Else
            truncatePhoneNumber = formatPhoneNumber(Right(phone, 7))
        End If
    Else
        truncatePhoneNumber = phone
    End If
End Function


