Attribute VB_Name = "modMain"
Public RIDE As Boolean
Public authenticated As Boolean
Public EPOCH As Date
Public FDOS As Date
Public LDOS As Date
Public Const HST_RATE = 0.15


Function getFDOS(ByVal yr As Long) As Date
    getFDOS = LabourDay(yr) + 2
End Function

Function getLDOS(ByVal yr As Long) As Date
    getLDOS = DateSerial(yr, 6, 22)
End Function

Function shortDate(var As Variant) As String
    Dim d As Date
        
    If IsNull(var) Then
        shortDate = ""
    Else
        d = CDate(var)
        shortDate = Format(d, "mmm dd, yyyy")
    End If
End Function

Function ansiDate(val As Variant) As String
    Dim d As Date
    
    If IsNull(val) Then
        ansiDate = ""
    Else
        d = CDate(val)
        ansiDate = Format(d, "YYYY-MM-DD")
    End If
End Function

Function MonthNumber(ByVal M As String) As Byte
    MonthNumber = month(CDate(M & " 1, 2016"))
End Function


Function timestamp() As String
    timestamp = CDbl(Now) * 10000000000#
    'timestamp = CLng(CDbl(Now) * 10000)
End Function

Function str_pad_left(ByVal str As String, ByVal ch As String, ByVal length As Byte) As String
    str_pad_left = str
    If Len(str) < length Then
        str_pad_left = String$(length - Len(str), ch) & str
    End If
End Function

Sub comboSelectItem(ByRef combo As ComboBox, ByVal value As String)
    Dim i As Long
    For i = 0 To combo.ListCount - 1
        If combo.List(i) = value Then
            combo.ListIndex = i
            Exit For
        End If
    Next i
End Sub

Function getAge(ByVal DOB As Date, ByVal atDate As Date) As Long
    getAge = year(atDate) - year(DOB)
    If CDate(year(atDate) & "-" & Format(month(DOB), "00") & "-" & Format(day(DOB), "00")) > atDate Then getAge = getAge - 1
End Function
Function getAgeM(ByVal DOB As Date, ByVal atDate As Date) As Long
    'returns the age in months
    getAgeM = DateDiff("m", DOB, atDate)
    If day(atDate) < day(DOB) Then getAgeM = getAgeM - 1
End Function

Function nearestFriday(ByVal d As Date) As Date
    If Weekday(d) = 6 Then
        nearestFriday = d
    ElseIf Weekday(d) < 6 Then
        If Weekday(d) > 3 Then
            nearestFriday = d + 6 - Weekday(d)
        Else
            nearestFriday = d - Weekday(d) - 1
        End If
    Else
        nearestFriday = d - 1
    End If
End Function

Public Function daysInMonth(ByVal d As Date) As Long
  ' Return the number of days in the specified month.
  daysInMonth = day(DateSerial(year(d), month(d) + 1, 1) - 1)
End Function

Function LabourDay(ByVal yr As Long) As Date
    LabourDay = DateSerial(yr, 9, 1)
    If Weekday(LabourDay) <> 2 Then 'monday
        If Weekday(LabourDay) < 2 Then 'sunday
            LabourDay = LabourDay + 1
        Else
            LabourDay = LabourDay + 7 - (Weekday(LabourDay) - 2)
        End If
    End If
End Function

Function getStatHolidays(ByVal year As Long) As Date()
    Dim ret(6) As Date
    ret(1) = CDate("01/01/" & year) 'New Years Day
    'ret(1) = CDate("01/01/" & year) 'May 24 (Variable Date, ignore)
    ret(2) = CDate("01/07/" & year) 'Canada Day
    ret(3) = LabourDay(year)        'Labour Day
    ret(4) = CDate("11/11/" & year) 'Rememberance Day
    ret(5) = CDate("12/25/" & year) 'Christmas Day
    ret(6) = CDate("12/26/" & year) 'Boxing Day
    getStatHolidays = ret
End Function

Function isStatHoliday(ByVal d As Date) As Boolean
    Dim hol() As Date
    Dim h As Byte
    isStatHoliday = False
    
    hol = getStatHolidays(year(d))
    For h = 1 To 6
        If hol(h) = d Then
            isStatHoliday = True
            Exit Function
        End If
    Next h
End Function

Public Function schoolAge(ByVal DOB As Date, Optional ByVal asOf As Date = 0) As Boolean
    If asOf = 0 Then asOf = Date
    If year(DOB) <= year(asOf) - 5 Then
        schoolAge = True
    Else
        schoolAge = False
    End If
End Function

Public Function kindergartenAge(ByVal DOB As Date, Optional ByVal asOf As Date = 0) As Boolean
    If asOf = 0 Then asOf = Date
    If getFDOS(year(asOf)) < asOf Then
        If year(DOB) = year(asOf) - 5 Then
            kindergartenAge = True
        Else
            kindergartenAge = False
        End If
    Else
        If year(DOB) = year(asOf) - 6 Then
            kindergartenAge = True
        Else
            kindergartenAge = False
        End If
    End If
End Function

Function getFeeClassAtDate(ByVal client As Long, ByVal d As Date) As Long
    Dim q As ADODB.Recordset
    Set q = db.Execute("SELECT * FROM client_changes WHERE idClient = " & client & " AND date <= " & sqlDate(d) & " ORDER BY date DESC, idChange DESC LIMIT 1;")
    If (q.EOF And q.BOF) Then Set q = db.Execute("SELECT * FROM client_changes WHERE idClient = " & client & " ORDER BY date ASC, idChange ASC LIMIT 1;")
    With q
        If Not (.EOF And .BOF) Then
            .MoveFirst
            getFeeClassAtDate = !feeClassID
        End If
    End With
    Set q = Nothing
End Function

Function getFeesAtDate(ByVal client As Long, ByVal d As Date) As Long
    Dim q As ADODB.Recordset
    Dim sc As ADODB.Recordset
    
    'Set q = db.Execute("SELECT * FROM client_changes WHERE idClient = " & client & " AND date <= " & sqlDate(d) & " ORDER BY date DESC, idChange DESC LIMIT 1;")
    Set q = db.Execute("SELECT client_changes.*, DOB FROM client_changes INNER JOIN clients ON client_changes.idClient = clients.idClient WHERE client_changes.idClient = " & client & " AND date <= " & sqlDate(d) & " ORDER BY date DESC, idChange DESC LIMIT 1;")
    Set sc = db.Execute("SELECT * FROM school_closures WHERE date = " & sqlDate(d) & " ORDER BY type DESC LIMIT 1;")
    If (q.EOF And q.BOF) Then Set q = db.Execute("SELECT * FROM client_changes WHERE idClient = " & client & " ORDER BY date ASC, idChange ASC LIMIT 1;")
    With q
        If Not (.EOF And .BOF) Then
            .MoveFirst
            getFeesAtDate = !fees
            If Not (sc.EOF And sc.BOF) Then
                sc.MoveFirst
                If isSchoolAgeClass(!feeClassID) Then
                    If sc!Type = "SC" Then
                        getFeesAtDate = 150
                    Else
                        If kindergartenAge(!DOB, d) Then
                            getFeesAtDate = 150
                        End If
                    End If
                End If
            End If
        End If
    End With
    Set q = Nothing
    Set sc = Nothing
End Function

Function getPayperiodAtDate(ByVal client As Long, ByVal d As Date) As Long
    Dim q As ADODB.Recordset
    Set q = db.Execute("SELECT * FROM client_changes WHERE idClient = " & client & " AND date <= " & sqlDate(d) & " ORDER BY date DESC, idChange DESC LIMIT 1;")
    If (q.EOF And q.BOF) Then Set q = db.Execute("SELECT * FROM client_changes WHERE idClient = " & client & " ORDER BY date ASC, idChange ASC LIMIT 1;")
    With q
        If Not (.EOF And .BOF) Then
            .MoveFirst
            getPayperiodAtDate = !payperiod
        End If
    End With
    Set q = Nothing
End Function

Function getParents(ByVal client As Long) As ADODB.Recordset
    Set getParents = db.Execute("SELECT * FROM client_contacts INNER JOIN contacts ON (client_contacts.idContact = contacts.idContact) WHERE idClient = " & client & " AND type=""Parent""")
End Function

Function getEmergency(ByVal client As Long) As ADODB.Recordset
    Set getEmergency = db.Execute("SELECT * FROM client_contacts INNER JOIN contacts ON (client_contacts.idContact = contacts.idContact) WHERE idClient = " & client & " AND type=""Emergency""")
End Function

Function getDoctor(ByVal client As Long) As ADODB.Recordset
    Set getDoctor = db.Execute("SELECT * FROM client_contacts INNER JOIN contacts ON (client_contacts.idContact = contacts.idContact) WHERE idClient = " & client & " AND type=""Doctor""")
End Function

Function getBestContactInfo(ByVal contact As Long) As String
    Dim c As ADODB.Recordset
    Set c = db.Execute("SELECT * FROM contacts WHERE idContact = " & contact)
    If Not (c.EOF And c.BOF) Then
        c.MoveFirst
        If ("" & c!cell) <> "" Then
            getBestContactInfo = c!cell
        ElseIf ("" & c!land) <> "" Then
            getBestContactInfo = c!land
        ElseIf ("" & c!email) <> "" Then
            getBestContactInfo = c!email
        Else
            getBestContactInfo = ""
        End If
    End If
End Function

Function getRoomAtDate(ByVal client As Long, ByVal d As Date) As String
    Dim q As ADODB.Recordset
    Set q = db.Execute("SELECT * FROM client_changes WHERE idClient = " & client & " AND date <= " & sqlDate(d) & " ORDER BY date DESC, idChange DESC LIMIT 1;")
    If (q.EOF And q.BOF) Then Set q = db.Execute("SELECT * FROM client_changes WHERE idClient = " & client & " ORDER BY date ASC, idChange ASC LIMIT 1;")
    With q
        If Not (.EOF And .BOF) Then
            .MoveFirst
            getRoomAtDate = !room
        End If
    End With
    Set q = Nothing
End Function

Function getSubsidizedAtDate(ByVal client As Long, ByVal d As Date) As Long
    Dim q As ADODB.Recordset
    Set q = db.Execute("SELECT * FROM client_changes WHERE idClient = " & client & " AND date <= " & sqlDate(d) & " ORDER BY date DESC, idChange DESC LIMIT 1;")
    If (q.EOF And q.BOF) Then Set q = db.Execute("SELECT * FROM client_changes WHERE idClient = " & client & " ORDER BY date ASC, idChange ASC LIMIT 1;")
    With q
        If Not (.EOF And .BOF) Then
            .MoveFirst
            getSubsidizedAtDate = !subsidized
        End If
    End With
    Set q = Nothing
End Function

Function getStartDateAtDate(ByVal client As Long, ByVal d As Date) As Date
    Dim q As ADODB.Recordset
    getStartDateAtDate = ""
    Set q = db.Execute("SELECT * FROM client_changes WHERE idClient = " & client & " AND date <= " & sqlDate(d) & " ORDER BY date DESC, idChange DESC LIMIT 1;")
    If (q.EOF And q.BOF) Then Set q = db.Execute("SELECT * FROM client_changes WHERE idClient = " & client & " ORDER BY date ASC, idChange ASC LIMIT 1;")
    With q
        If Not (.EOF And .BOF) Then
            .MoveFirst
            getStartDateAtDate = !startDate
        End If
    End With
    Set q = Nothing
End Function

Function getEndDateAtDate(ByVal client As Long, ByVal d As Date) As Date
    Dim q As ADODB.Recordset
    getEndDateAtDate = ""
    Set q = db.Execute("SELECT * FROM client_changes WHERE idClient = " & client & " AND date <= " & sqlDate(d) & " ORDER BY date DESC, idChange DESC LIMIT 1;")
    If (q.EOF And q.BOF) Then Set q = db.Execute("SELECT * FROM client_changes WHERE idClient = " & client & " ORDER BY date ASC, idChange ASC LIMIT 1;")
    With q
        If Not (.EOF And .BOF) Then
            .MoveFirst
            getEndDateAtDate = !endDate
        End If
    End With
    Set q = Nothing
End Function

Function getActiveAtDate(ByVal client As Long, ByVal d As Date) As Long
    Dim q As ADODB.Recordset
    Set q = db.Execute("SELECT * FROM client_changes WHERE idClient = " & client & " AND date <= " & sqlDate(d) & " ORDER BY date DESC, idChange DESC LIMIT 1;")
    If (q.EOF And q.BOF) Then Set q = db.Execute("SELECT * FROM client_changes WHERE idClient = " & client & " ORDER BY date ASC, idChange ASC LIMIT 1;")
    With q
        If Not (.EOF And .BOF) Then
            .MoveFirst
            getActiveAtDate = !active
        End If
    End With
    Set q = Nothing
End Function

Function getAuthorizationNumberAtDate(ByVal client As Long, ByVal d As Date) As String
    Dim q As ADODB.Recordset
    Set q = db.Execute("SELECT * FROM client_changes WHERE idClient = " & client & " AND date <= " & sqlDate(d) & " ORDER BY date DESC, idChange DESC LIMIT 1;")
    If (q.EOF And q.BOF) Then Set q = db.Execute("SELECT * FROM client_changes WHERE idClient = " & client & " ORDER BY date ASC, idChange ASC LIMIT 1;")
    With q
        If Not (.EOF And .BOF) Then
            .MoveFirst
            getAuthorizationNumberAtDate = "" & !authorizationNumber
        End If
    End With
    Set q = Nothing
End Function

Function getParentContributionAtDate(ByVal client As Long, ByVal d As Date) As Long
    Dim q As ADODB.Recordset
    Set q = db.Execute("SELECT * FROM client_changes WHERE idClient = " & client & " AND date <= " & sqlDate(d) & " ORDER BY date DESC, idChange DESC LIMIT 1;")
    If (q.EOF And q.BOF) Then Set q = db.Execute("SELECT * FROM client_changes WHERE idClient = " & client & " ORDER BY date ASC, idChange ASC LIMIT 1;")
    With q
        If Not (.EOF And .BOF) Then
            .MoveFirst
            getParentContributionAtDate = !parentalContribution
        End If
    End With
    Set q = Nothing
End Function


Sub ShellSort(arr As Variant, numEls As Long, descending As Boolean)
    Dim index As Long, index2 As Long, firstItem As Long
    Dim distance As Long, value As Variant
    
    ' Exit if it is not an array.
    If VarType(arr) < vbArray Then Exit Sub
    firstItem = LBound(arr)
    
    ' Find the best value for distance.
    Do
        distance = distance * 3 + 1
    Loop Until distance > numEls
    
    ' Sort the array.
    Do
        distance = distance \ 3
        For index = distance + firstItem To numEls + firstItem - 1
            value = arr(index)
            index2 = index
            Do While (arr(index2 - distance) > value) Xor descending
                arr(index2) = arr(index2 - distance)
                index2 = index2 - distance
                If index2 - distance < firstItem Then Exit Do
            Loop
            arr(index2) = value
        Next
    Loop Until distance = 1
End Sub

Sub check_for_client_changes()
    ' check for children aging out of their fee class.
    ' check for first and last day of school
    Dim q As ADODB.Recordset
    Dim fc As ADODB.Recordset
    
    Set q = db.Execute("SELECT * FROM clients WHERE active = 1")
    With q
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do Until .EOF
                Set fc = db.Execute("SELECT * FROM fee_classes WHERE idFeeClasses = " & !feeClassID)
                If Not (fc.EOF And fc.BOF) Then
                    If getAgeM(!DOB, Date) > fc!max_age Then
                        DoEvents
                        dlgAgeChange.Timer1.Tag = !idClient
                        dlgAgeChange.Timer2.Tag = fc!max_age + 1
                        DoEvents
                        dlgAgeChange.Show 1
                    End If
                    
                    
                    If Abs(DateDiff("d", LDOS, Date)) < 5 Or Abs(DateDiff("d", FDOS, Date)) < 5 Then
                        If (fc!idFeeClasses = 4 Or fc!idFeeClasses = 5 Or q!DOB < CDate("Dec 31, " & year(Date) - 5)) And Abs(DateDiff("d", LDOS, Date)) < 5 Then 'they are already in school (feeclass 4 or 5) or are going to kindergarten this year)
                            'School's out!  Set them to fee class 9
                            dlgSchoolChange.Timer1.Tag = !idClient
                            dlgSchoolChange.cboFeeClass.Tag = 9
                            DoEvents
                            dlgSchoolChange.Show 1
                        End If
                        If (fc!idFeeClasses = 7) And Abs(DateDiff("d", LDOS, Date)) < 5 Then
                            'School's out!  Set them to fee class 8
                            dlgSchoolChange.Timer1.Tag = !idClient
                            dlgSchoolChange.cboFeeClass.Tag = 8
                            DoEvents
                            dlgSchoolChange.Show 1
                        End If
                        
                        If (fc!idFeeClasses = 9) And Abs(DateDiff("d", FDOS, Date)) < 5 Then
                            'Back to school!  Set them to fee class 4
                            dlgSchoolChange.Timer1.Tag = !idClient
                            dlgSchoolChange.cboFeeClass.Tag = 4
                            DoEvents
                            dlgSchoolChange.Show 1
                        End If
                        If (fc!idFeeClasses = 8) And Abs(DateDiff("d", FDOS, Date)) < 5 Then
                            'Back to school!  Set them to fee class 7
                            dlgSchoolChange.Timer1.Tag = !idClient
                            dlgSchoolChange.cboFeeClass.Tag = 7
                            DoEvents
                            dlgSchoolChange.Show 1
                        End If
                    End If
                End If
                .MoveNext
            Loop
            
            Unload dlgAgeChange
        End If
    End With
End Sub

Sub insertClientChange(changedate As Date, idClient As Long, feeClassID As Long, fees As String, payperiod As Byte, room As String, subsidized As Byte, authorizationNumber As String, parentalContribution As Double, startDate As Variant, endDate As Variant, active As Byte)
    sql = "INSERT INTO client_changes (date, idClient, feeClassID, fees, payperiod, room, subsidized, authorizationNumber, parentalContribution, startDate, endDate, active) VALUES ("
    sql = sql & sqlDate(changedate) & ","
    sql = sql & idClient & ","
    sql = sql & feeClassID & ","
    sql = sql & fees & ","
    sql = sql & payperiod & ","
    sql = sql & """" & room & ""","
    sql = sql & subsidized & ","
    sql = sql & """" & authorizationNumber & ""","
    sql = sql & parentalContribution & ","
    sql = sql & sqlDate(startDate) & ","
    If IsNull(endDate) Then
        sql = sql & "NULL ,"
    Else
        sql = sql & sqlDate(endDate) & ","
    End If
    sql = sql & active & ")"
    db.Execute sql
End Sub

Function isSchoolAgeClass(ByVal feeClassID As Long) As Boolean
    isSchoolAgeClass = False
    Dim fc As ADODB.Recordset
    Set fc = db.Execute("SELECT * FROM fee_classes WHERE idFeeClasses = " & feeClassID)
    
    'old static method
    'If feeClassID = 4 Or feeClassID = 5 Or feeClassID = 7 Or feeClassID = 8 Or feeClassID = 9 Then
    
    If Not (fc.EOF And fc.BOF) Then
        fc.MoveFirst
        If fc!isSA = 1 Then
            isSchoolAgeClass = True
        End If
    End If
    Set fc = Nothing
End Function

Function isSummerClass(ByVal feeClassID As Long) As Boolean
    isSummerClass = False
    Dim fc As ADODB.Recordset
    Set fc = db.Execute("SELECT * FROM fee_classes WHERE idFeeClasses = " & feeClassID)
    
    If Not (fc.EOF And fc.BOF) Then
        fc.MoveFirst
        If fc!isSummer = 1 Then
            isSummerClass = True
        End If
    End If
    Set fc = Nothing
End Function

Function LDOM(ByVal Mo As Byte, ByVal yr As Long) As Date 'LAST DAY OF MONTH
    LDOM = CDate(yr & "-" & Format(Mo, "00") & "-15") + 30
    LDOM = CDate(year(LDOM) & "-" & Format(month(LDOM), "00") & "-01") - 1
End Function

Function weekdays(ByVal fromDate As Date, ByVal toDate As Date) As Long
    'Returns the number of weekdays between two dates
    Dim d1 As Long
    Dim d2 As Long
    Dim d3 As Long
    Dim d4 As Long
    
    d1 = DateDiff("d", fromDate, toDate) + 1
    d2 = d1 \ 7
    d3 = d1 Mod 7
    d4 = d2 * 5
    For d1 = 0 To d3 - 1
        If Not isWeekend(fromDate + d1) Then
            d4 = d4 + 1
        End If
    Next d1
    weekdays = d4
    
End Function

Function isWeekend(ByVal d As Date) As Boolean
    If Weekday(d) = 1 Or Weekday(d) = 7 Then
        isWeekend = True
    Else
        isWeekend = False
    End If
End Function

Function dayToAbbrev(ByVal d As Long) As String
    dayToAbbrev = Left$(WeekdayName(d), 2)
End Function

Function abbrevToDay(ByVal d As String) As Long
    abbrevToDay = 1
    For i = 1 To 7
        If LCase(Left$(WeekdayName(i), 2)) = LCase(Left$(d, 2)) Then
            abbrevToDay = i
            Exit For
        End If
    Next i
End Function

Function weekdayToLetter(ByVal d As Long) As String
    weekdayToLetter = Left$(WeekdayName(d), 1)
    If d = 5 Then weekdayToLetter = "H"
End Function

Function letterToWeekday(ByVal d As String) As Long
    letterToWeekday = 2
    For i = 2 To 6
        If LCase(Left$(WeekdayName(i), 1)) = LCase(Left$(d, 1)) Then
            letterToWeekday = i
            Exit For
        End If
    Next i
    If d = "H" Then letterToWeekday = 5
End Function

Function feeClassDaysPerWeek(ByVal feeClassID As Long) As Byte
    feeClassDaysPerWeek = 5
    Dim fc As ADODB.Recordset
    Set fc = db.Execute("SELECT * FROM fee_classes WHERE idFeeClasses = " & feeClassID)
    
    If Not (fc.EOF And fc.BOF) Then
        Dim d As Byte
        d = 0
        fc.MoveFirst
        For i = 2 To 6
            If fc.Fields(weekdayToLetter(i)) Then d = d + 1
        'M Then d = d + 1
        'If fc!T Then d = d + 1
        'If fc!W Then d = d + 1
        'If fc!h Then d = d + 1
        'If fc!f Then d = d + 1
        Next i
        feeClassDaysPerWeek = d
    End If
    Set fc = Nothing
End Function

