Attribute VB_Name = "modPrintCheques"
Option Explicit

Sub printTextChq(ByVal s As String, ByVal x As Double, ByVal y As Double, ByVal fontname As String, ByVal fontsize As Double, ByVal bold As Boolean, Optional justify As Byte = 0)
    'Dim textpic As PictureBox
    'Set textpic = frmMain.Picture1
    If justify <> 0 Then ' anything other than 0 is right justify
        'textpic.Width = 3000
        'textpic.Height = 320
        'textpic.Font = fontname
        'textpic.FontBold = bold
        'textpic.fontsize = fontsize
        'textpic.CurrentX = 0
        'textpic.CurrentY = 0
        'textpic.Print s;
        x = x - Printer.textwidth(s)
    End If
    
    Printer.fontname = fontname
    Printer.fontsize = fontsize
    Printer.FontBold = bold
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.forecolor = &H111111
    Printer.Print s;
    
    'Set textpic = Nothing
End Sub
Sub findCheque(ByVal num As Long)
'    With Cheques
'        Do Until .EOF
'            If !chqNumber = num Then Exit Do
'            .MoveNext
'        Loop
'        If .EOF Then
'        .MoveFirst
'        Do Until .EOF
'            If !chqNumber = num Then Exit Do
'            .MoveNext
'        Loop
'        End If
'    End With
End Sub
Function english3Digit(ByVal num As Long) As String
    Dim num1 As Byte
    Dim num2 As Byte
    num1 = val(Left$(Trim(str$(num)), 1))
    If num < 100 Then num1 = 0
    num2 = val(Right$(Trim(str$(num)), 2))
    
    If num1 > 0 Then
        english3Digit = englishDigit(num1) & " Hundred " & english2Digit(num2)
    Else
        english3Digit = english2Digit(num2)
    End If
    english3Digit = Trim(english3Digit)
End Function

Function englishDigit(ByVal num As Byte) As String
    englishDigit = ""
    If num = 1 Then
        englishDigit = "One"
    ElseIf num = 2 Then
        englishDigit = "Two"
    ElseIf num = 3 Then
        englishDigit = "Three"
    ElseIf num = 4 Then
        englishDigit = "Four"
    ElseIf num = 5 Then
        englishDigit = "Five"
    ElseIf num = 6 Then
        englishDigit = "Six"
    ElseIf num = 7 Then
        englishDigit = "Seven"
    ElseIf num = 8 Then
        englishDigit = "Eight"
    ElseIf num = 9 Then
        englishDigit = "Nine"
    ElseIf num = 0 Then
        englishDigit = ""
    End If
End Function
Function english2Digit(ByVal num As Byte) As String
    Dim num1 As Byte
    Dim num2 As Byte
    num1 = val(Left$(Trim(str$(num)), 1))
    If num < 10 Then num1 = 0
    num2 = val(Right$(Trim(str$(num)), 1))
    english2Digit = ""
    If num1 = 1 Then
        If num2 = 0 Then
            english2Digit = "Ten"
        ElseIf num2 = 1 Then
            english2Digit = "Eleven"
        ElseIf num2 = 2 Then
            english2Digit = "Twelve"
        ElseIf num2 = 3 Then
            english2Digit = "Thirteen"
        ElseIf num2 = 4 Then
            english2Digit = "Fourteen"
        ElseIf num2 = 5 Then
            english2Digit = "Fifteen"
        ElseIf num2 = 6 Then
            english2Digit = "Sixteen"
        ElseIf num2 = 7 Then
            english2Digit = "Seventeen"
        ElseIf num2 = 8 Then
            english2Digit = "Eighteen"
        ElseIf num2 = 9 Then
            english2Digit = "Nineteen"
        End If
    ElseIf num1 = 2 Then
        english2Digit = "Twenty " & englishDigit(num2)
    ElseIf num1 = 3 Then
        english2Digit = "Thirty " & englishDigit(num2)
    ElseIf num1 = 4 Then
        english2Digit = "Fourty " & englishDigit(num2)
    ElseIf num1 = 5 Then
        english2Digit = "Fifty " & englishDigit(num2)
    ElseIf num1 = 6 Then
        english2Digit = "Sixty " & englishDigit(num2)
    ElseIf num1 = 7 Then
        english2Digit = "Seventy " & englishDigit(num2)
    ElseIf num1 = 8 Then
        english2Digit = "Eighty " & englishDigit(num2)
    ElseIf num1 = 9 Then
        english2Digit = "Ninety " & englishDigit(num2)
    ElseIf num1 = 0 Then
        english2Digit = englishDigit(num2)
    End If
End Function


Function englishNumber(ByVal num As Double) As String
    Dim n As String
    n = Trim(str$(Int(num)))
    englishNumber = ""
    
    If Len(n) <= 3 Then
        englishNumber = english3Digit(val(n))
    ElseIf Len(n) <= 6 Then
        englishNumber = english3Digit(val(Left$(n, Len(n) - 3))) & " Thousand " & english3Digit(val(Right$(n, 3)))
    Else
        MsgBox "OMG!  A Million!"
        englishNumber = english3Digit(val(Left$(n, Len(n) - 6))) & " Million " & english3Digit(val(Left$(n, Len(n) - 3))) & " Thousand " & english3Digit(val(Right$(n, 3)))
    End If
End Function

Function getLastChqNum() As Long
    getLastChqNum = 100000
    'With Cheques
    '    If Not (.EOF And .BOF) Then .MoveFirst
    '    Do Until .EOF
    '        If !chqNumber > getLastChqNum Then getLastChqNum = !chqNumber
    '        .MoveNext
    '    Loop
    'End With
End Function

Sub printCheque(ByVal chqNumber As Long, ByVal pos As Byte, ByVal perPage As Byte)
    ' THIS IS NOT THE FUNCTION USED TO PRINT PAYCHEQUES.  SEE FUNCTION PRINTPAYCHEQUE FOR THAT.
    
    Dim offset As Long
    Dim address_offset As Long
    Dim chq1PaperSize As Long
    Dim chq2PaperSize As Long
    Dim chq1PrintOffset As Long
    Dim chq2PrintOffset As Long
    Dim chq3PrintOffset As Long
    Dim sysid_y As Long
    Dim overall_x_offset As Long
    Dim cheques As ADODB.Recordset
    
    'these offsets are to adjust for the idiots at nebs
    'none of the three cheques on a page are identical size/position
    'the top one has wording 1/32" closer to the top edge
    'the bottom one has wording 1/32" closer to the bottom edge
    '1/32" ~= 45 twips
    
    chq1PaperSize = 5000 ' twips   (height of first cheque on the sheet)
    chq2PaperSize = 5067 ' twips   (height of second cheque on the sheet)
    chq1PrintOffset = -70 ' twips   (70 twips closer to the top edge than the standard)
    chq2PrintOffset = 0 ' twips   (The bottom of the D in the Date field is exactly 1" from the top edge of cheque.  Align the other two chegues according to cheque 2)
    chq3PrintOffset = -25 'twips   (25 twips closer to the top edge than the standard)
    
    sysid_y = 550
    overall_x_offset = 200
    
    If perPage = 2 Then
    'if theres 2 cheques per page then pos1 needs to have the
    'paper size and offset for the second cheque on a sheet, etc
        If pos = 1 Then
            offset = chq2PrintOffset
        ElseIf pos = 2 Then
            offset = chq2PaperSize + chq3PrintOffset
        End If
    Else
    'otherwise pos1 = SheetChq1, pos2 = SheetChq2, pos 3 = SheetChq3
    'and pos 4 takes care of the 1 chq per page
        If pos = 1 Then
            offset = chq1PrintOffset
        ElseIf pos = 2 Then
            offset = chq1PaperSize + chq2PrintOffset
        ElseIf pos = 3 Then
            offset = chq1PaperSize + chq2PaperSize + chq3PrintOffset
        ElseIf pos = 4 Then ' special case for single cheques
            overall_x_offset = 200 + Printer.ScaleWidth - Printer.ScaleHeight
            offset = 3550 '7200
        End If
    End If
    Printer.ScaleMode = vbTwips
    Set cheques = db.Execute("SELECT * FROM cheques WHERE chqNumber = " & chqNumber)

    address_offset = 0
    With cheques
        printTextChq "MM/DD/YYYY", overall_x_offset + 9550, 900 + offset, "Courier New", 8.4, False
        printTextChq chequeDate(!Date), 9550 + overall_x_offset, 1050 + offset, "Verdana", 8.6, False
        printTextChq englishNumber(!amount), 1000 + overall_x_offset, 1880 + offset, "Verdana", 10, False
        printTextChq Format(!amount, "0.00"), 11350 + overall_x_offset, 1800 + offset, "Verdana", 10, False, 1
        printTextChq !payto, 1000 + overall_x_offset, 2650 + offset, "Verdana", 10, False
        
        printTextChq !Memo, 650 + overall_x_offset, 3500 + offset + (address_offset / 3), "Verdana", 10, False
        
        If val(Right$(Trim(Format(!amount, "0.00")), 2)) = 0 Then
            printTextChq "xx", 8110 + overall_x_offset, 1820 + offset, "Arial", 8, False
        Else
            printTextChq Right$(Trim(Format(!amount, "0.00")), 2), 8110 + overall_x_offset, 1820 + offset, "Arial", 8, False
        End If
        Printer.Line (8150 + overall_x_offset, 2120 + offset)-(8510 + overall_x_offset, 1860 + offset), vbBlack ' /
        
        Printer.Line (930 + overall_x_offset, 2150 + offset)-(8000 + overall_x_offset, 2150 + offset), vbBlack ' written amt
        Printer.Line (overall_x_offset + 930, 2950 + offset)-(overall_x_offset + 8000, 2950 + offset), vbBlack ' pay to
        Printer.Line -(overall_x_offset + 8000, 2650 + offset), vbBlack ' pay to | VERTICAL
    End With
End Sub

Function chequeDate(val As Variant) As String
    Dim d As Date
    d = CDate(val)
    
    chequeDate = Format(d, "mm dd yyyy")
End Function

