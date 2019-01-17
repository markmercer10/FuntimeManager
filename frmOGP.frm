VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmOGP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OGP"
   ClientHeight    =   10635
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   12960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10635
   ScaleWidth      =   12960
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton calcButn 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   4440
      TabIndex        =   17
      Top             =   480
      Width           =   1455
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   495
      Left            =   4200
      TabIndex        =   15
      Top             =   5040
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   873
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.TextBox totProjDays 
      Height          =   288
      Left            =   12360
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.CommandButton printButn 
      Caption         =   "Print"
      Height          =   855
      Left            =   6720
      TabIndex        =   8
      Top             =   0
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5880
      Top             =   480
   End
   Begin VB.TextBox totDays 
      Height          =   288
      Left            =   5400
      TabIndex        =   6
      Top             =   120
      Width           =   492
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   9615
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   16960
      _Version        =   393216
      Rows            =   39
      Cols            =   6
   End
   Begin VB.ComboBox cboQuarter 
      Height          =   315
      ItemData        =   "frmOGP.frx":0000
      Left            =   720
      List            =   "frmOGP.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   2172
   End
   Begin MSComCtl2.DTPicker dpToDate 
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "MMM d, yyyy"
      Format          =   136577027
      CurrentDate     =   42613
   End
   Begin MSComCtl2.DTPicker dpFromDate 
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "MMM d, yyyy"
      Format          =   136577027
      CurrentDate     =   42527
   End
   Begin VB.Label Label7 
      Caption         =   "Quarter"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   615
   End
   Begin VB.Label ProjToDate 
      Caption         =   "date"
      Height          =   375
      Left            =   9360
      TabIndex        =   14
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "To"
      Height          =   255
      Left            =   9360
      TabIndex        =   13
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label ProjFromDate 
      Caption         =   "date"
      Height          =   375
      Left            =   8160
      TabIndex        =   12
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "From"
      Height          =   255
      Left            =   8160
      TabIndex        =   11
      Top             =   240
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   8040
      X2              =   8040
      Y1              =   0
      Y2              =   840
   End
   Begin VB.Label Label4 
      Caption         =   "Total billable days in projected quarter"
      Height          =   495
      Left            =   10680
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Total billable days in  quarter"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "frmOGP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub calcButn_Click()
    Timer1.Enabled = True
End Sub

Private Sub cboQuarter_Change()
    If cboQuarter.ListIndex = 0 Then
        If month(Date) = 12 Then
            dpFromDate = CDate("Dec 1," & year(Date))
            dpToDate = CDate("Mar 1," & year(Date) + 1) - 1
        Else
            dpFromDate = CDate("Dec 1," & year(Date) - 1)
            dpToDate = CDate("Mar 1," & year(Date)) - 1
        End If
    ElseIf cboQuarter.ListIndex = 1 Then
        dpFromDate = CDate("Mar 1," & year(Date))
        dpToDate = CDate("May 31," & year(Date))
    ElseIf cboQuarter.ListIndex = 2 Then
        dpFromDate = CDate("Jun 1," & year(Date))
        dpToDate = CDate("Aug 31," & year(Date))
    ElseIf cboQuarter.ListIndex = 3 Then
        dpFromDate = CDate("Sep 1," & year(Date))
        dpToDate = CDate("Nov 30," & year(Date))
    End If
    
    'totDays = DateDiff("d", dpFromDate.value, dpToDate.value)
End Sub

Private Sub cboQuarter_Click()
    cboQuarter_Change
End Sub

Private Sub dpFromDate_Change()
    totDays = DateDiff("d", dpFromDate.value, dpToDate.value)
End Sub

Private Sub dpToDate_Change()
    dpToDate_Click
End Sub

Private Sub dpToDate_Click()
    totDays = DateDiff("d", dpFromDate.value, dpToDate.value)
    ProjFromDate = shortDate(dpToDate.value + 1)
    ProjToDate = shortDate(DateSerial(year(dpToDate.value + 1), month(dpToDate.value + 1) + 3, day(dpToDate.value + 1)) - 1)
End Sub

Private Sub Form_Load()
    cboQuarter.ListIndex = 2
    If dpFromDate.value < EPOCH Then dpFromDate = EPOCH
    DoEvents
    dpToDate_Click
    DoEvents
End Sub

Private Sub printButn_Click()
    Dim ori As Long
    Dim i As Long
    Dim r As Long
    
    
    ori = Printer.Orientation
    Printer.Orientation = vbPRORLandscape
    
    printText "Operating Grant Program", 50, 20, 10000, "Arial", 16, True, 0
    printText "Funtime Child Care Center - Enrollment Statistics " & shortDate(dpFromDate.value) & " - " & shortDate(dpToDate.value), 50, 400, 10000, "Arial", 11, False, 0
    printText "Attention: Regan Power - 729-1400", 50, 700, 10000, "Arial", 8, False, 0
    printFlexGrid Printer, Grid, 50, 1000, 1
    Printer.EndDoc
    
    
    Printer.Orientation = ori
End Sub

Private Sub Timer1_Timer_OLD()
    Timer1.Enabled = False
    Dim r As Long
    Dim c As Byte
    Dim q As ADODB.Recordset
    Dim d1 As Long
    Dim d2 As Long
    Dim d3 As Long
    Dim d4 As Long
    Dim tempdate As Date
    Dim feeclass_cats(1 To 255) As Byte
    Dim IC_index As Byte
    Dim PS_index As Byte
    Dim SA_index As Byte
    Dim cat As Byte
    Dim row_index As Byte
    Dim room As String
    Dim age As Long
    Dim stop_counting As Boolean
    
    ProgressBar.Visible = True
    Grid.Visible = False
    
    Set q = db.Execute("SELECT * FROM fee_classes;")
    With q
        If Not (.EOF And .BOF) Then
            Do Until .EOF
                feeclass_cats(!idFeeClasses) = !ogp_cat
                .MoveNext
            Loop
        End If
    End With
        
    
    d1 = DateDiff("d", dpFromDate, dpToDate) + 1
    d2 = d1 \ 7
    d3 = d1 Mod 7
    d4 = d2 * 5
    For d1 = 0 To d3 - 1
        If Weekday(dpFromDate + d1) > 1 And Weekday(dpFromDate + d1) < 7 Then
            d4 = d4 + 1
        End If
    Next d1
    totDays = d4
    
    
    Grid.ColWidth(0) = 2500
    Grid.ColWidth(1) = 1200
    Grid.ColWidth(2) = 1200
    Grid.ColWidth(3) = 1200
    Grid.ColWidth(4) = 1200
    Grid.ColWidth(6) = 2000
    Grid.ColWidth(7) = 2000
    Grid.ColWidth(8) = 2000
    Grid.ColWidth(9) = 2000
    Grid.ColWidth(10) = 1500
    Grid.ColAlignment(0) = 0
    Grid.ColAlignment(1) = 3
    Grid.ColAlignment(2) = 3
    Grid.ColAlignment(3) = 3
    Grid.ColAlignment(4) = 6
    Grid.ColAlignment(6) = 0
    Grid.ColAlignment(7) = 3
    Grid.ColAlignment(8) = 3
    Grid.ColAlignment(9) = 3
    Grid.ColAlignment(10) = 6
    Grid.TextMatrix(0, 1) = "0-23 Months"
    Grid.TextMatrix(0, 2) = "24-59 Months"
    Grid.TextMatrix(0, 3) = "60+ Months"
    Grid.TextMatrix(0, 4) = "Amount"
    Grid.TextMatrix(0, 6) = "Child's Name" 'This wont show on the submission but it will be handy for the person to see to adjust projected enrollment
    Grid.TextMatrix(0, 7) = "Projected 0-23 Months"
    Grid.TextMatrix(0, 8) = "Projected 24-59 Months"
    Grid.TextMatrix(0, 9) = "Projected 60+ Months"
    Grid.TextMatrix(0, 10) = "Projected Amount"
    For r = 1 To 15
        Grid.TextMatrix(r, 0) = "School Age Room Slot " & r
    Next r
    For r = 1 To 14
        Grid.TextMatrix(r + 15, 0) = "Preschool Room Slot " & r
    Next r
    For r = 1 To 7
        Grid.TextMatrix(r + 29, 0) = "Infant-Care/Mixed Room Slot " & r
    Next r
    
    For r = 1 To 36
        Grid.row = r
        Grid.col = 2
        Grid.CellAlignment = 1
        Grid.col = 3
        Grid.CellAlignment = 1
        Grid.col = 4
        Grid.CellAlignment = 2
    Next r
    
    Grid.TextMatrix(38, 0) = "Totals"
    
    For tempdate = dpFromDate.value To dpToDate.value
        ProgressBar.value = (tempdate - dpFromDate.value) / (dpToDate.value - dpFromDate.value) * 50
        DoEvents
        
        If Weekday(tempdate) <> 1 And Weekday(tempdate) <> 7 Then ' if it's a weekday
            IC_index = 0
            PS_index = 0
            SA_index = 0
            Set q = db.Execute("SELECT * FROM clients WHERE startDate <= " & sqlDate(tempdate) & " AND active = 1;")
            With q
                If Not (.EOF And .BOF) Then
                    Do Until .EOF
                        stop_counting = False
                        If !idClient = 29 Or !idClient = 30 Then
                            'Do Nothing... This is Amelia and Catherine
                        Else
                            room = getRoomAtDate(!idClient, tempdate)
                            'cat = feeclass_cats(getFeeClassAtDate(!idClient, tempdate))
                            age = getAgeM(!DOB, tempdate)
                            If age < 155 Then cat = 3
                            If age < 60 Then cat = 2
                            If age < 24 Then cat = 1
                            
                            If room = "IC" Then
                                IC_index = IC_index + 1
                                If IC_index > 7 Then stop_counting = True
                                row_index = IC_index + 29
                            ElseIf room = "PS" Then
                                PS_index = PS_index + 1
                                If PS_index > 14 Then stop_counting = True
                                row_index = PS_index + 15
                            Else
                                SA_index = SA_index + 1
                                If SA_index > 15 Then stop_counting = True
                                row_index = SA_index
                            End If
                            'If room = "SA" And cat = 2 Then MsgBox !First & !Last & "  " & getFeeClassAtDate(!idclient, tempdate) & "  Really... This should be based on age at date "
                            If Not stop_counting Then Grid.TextMatrix(row_index, cat) = val(Grid.TextMatrix(row_index, cat)) + 1
                        End If
                        .MoveNext
                    Loop
                End If
            End With
        End If
    Next tempdate
    
    For r = 1 To 36
        Grid.TextMatrix(r, 4) = Format(val(Grid.TextMatrix(r, 1)) * 14 + val(Grid.TextMatrix(r, 2)) * 8 + val(Grid.TextMatrix(r, 3)) * 3, "0.00")
    Next r
    
    For c = 1 To 4
        d1 = 0
        For r = 1 To 36
            d1 = d1 + val(Grid.TextMatrix(r, c))
        Next r
        Grid.TextMatrix(38, c) = d1
    Next c
    Grid.TextMatrix(38, 4) = Format(val(Grid.TextMatrix(38, 4)), "$0.00")
    
    
    ' PROJECTED
    For tempdate = CDate(ProjFromDate) To CDate(ProjToDate)
        ProgressBar.value = (tempdate - CDate(ProjFromDate)) / (CDate(ProjToDate) - CDate(ProjFromDate)) * 50 + 50
        DoEvents
        If Weekday(tempdate) <> 1 And Weekday(tempdate) <> 7 Then ' if it's a weekday
            IC_index = 0
            PS_index = 0
            SA_index = 0
            Set q = db.Execute("SELECT * FROM clients WHERE startDate <= " & sqlDate(tempdate) & " AND active = 1;")
            With q
                If Not (.EOF And .BOF) Then
                    Do Until .EOF
                        stop_counting = False
                        If !idClient = 29 Or !idClient = 30 Then
                            'Do Nothing... This is Amelia and Catherine
                        Else
                            room = getRoomAtDate(!idClient, tempdate)
                            'cat = feeclass_cats(getFeeClassAtDate(!idClient, tempdate))
                            age = getAgeM(!DOB, tempdate)
                            If age < 155 Then cat = 9
                            If age < 60 Then cat = 8
                            If age < 24 Then cat = 7
                            
                            If room = "IC" Then
                                IC_index = IC_index + 1
                                If IC_index > 7 Then stop_counting = True
                                row_index = IC_index + 29
                            ElseIf room = "PS" Then
                                PS_index = PS_index + 1
                                If PS_index > 14 Then stop_counting = True
                                row_index = PS_index + 15
                            Else
                                SA_index = SA_index + 1
                                If SA_index > 15 Then stop_counting = True
                                row_index = SA_index
                            End If
                            'If room = "SA" And cat = 2 Then MsgBox !First & !Last & "  " & getFeeClassAtDate(!idclient, tempdate) & "  Really... This should be based on age at date "
                            If Not stop_counting Then
                                Grid.TextMatrix(row_index, cat) = val(Grid.TextMatrix(row_index, cat)) + 1
                                Grid.TextMatrix(row_index, 6) = !Last & ", " & !First
                            End If
                        End If
                        .MoveNext
                    Loop
                End If
            End With
        End If
    Next tempdate
    
    For r = 1 To 36
        Grid.TextMatrix(r, 10) = Format(val(Grid.TextMatrix(r, 7)) * 14 + val(Grid.TextMatrix(r, 8)) * 8 + val(Grid.TextMatrix(r, 9)) * 3, "0.00")
    Next r
    
    For c = 1 To 4
        d1 = 0
        For r = 1 To 36
            d1 = d1 + val(Grid.TextMatrix(r, c + 6))
        Next r
        Grid.TextMatrix(38, c + 6) = d1
    Next c
    Grid.TextMatrix(38, 10) = Format(val(Grid.TextMatrix(38, 10)), "$0.00")
    
    ProgressBar.Visible = False
    Grid.Visible = True
    
    Set q = Nothing
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Dim r As Long
    Dim c As Byte
    Dim q As ADODB.Recordset
    Dim d1 As Long 'a temp variable
    Dim tempdate As Date
    Dim feeclass_cats(1 To 255) As Byte
    Dim IC_index As Byte
    Dim PS_index As Byte
    Dim SA_index As Byte
    Dim cat As Byte
    Dim row_index As Byte
    Dim room As String
    Dim age As Long
    Dim stop_counting As Boolean
    Dim fc As Long
    Dim slot_quantity As Double
    
    ProgressBar.Visible = True
    Grid.Visible = False
    
    Set q = db.Execute("SELECT * FROM fee_classes;")
    With q
        If Not (.EOF And .BOF) Then
            Do Until .EOF
                feeclass_cats(!idFeeClasses) = !ogp_cat
                .MoveNext
            Loop
        End If
    End With
        
    
    totDays = weekdays(dpFromDate, dpToDate)
    Grid.Clear
    
    Grid.ColWidth(0) = 4000
    Grid.ColWidth(1) = 1600
    Grid.ColWidth(2) = 1600
    Grid.ColWidth(3) = 1600
    Grid.ColWidth(4) = 1600
    Grid.ColWidth(5) = 1200
    Grid.ColAlignment(0) = 0
    Grid.ColAlignment(1) = 3
    Grid.ColAlignment(2) = 3
    Grid.ColAlignment(3) = 3
    Grid.ColAlignment(4) = 3
    Grid.ColAlignment(5) = 6
    Grid.TextMatrix(0, 1) = "0-23 Months"
    Grid.TextMatrix(0, 2) = "24-59 Months"
    Grid.TextMatrix(0, 3) = "60+ Mo. Full Day"
    Grid.TextMatrix(0, 4) = "60+ Mo. AS"
    Grid.TextMatrix(0, 5) = "Amount"
    For r = 1 To 15
        Grid.TextMatrix(r, 0) = "School Age Room Slot " & r
    Next r
    For r = 1 To 14
        Grid.TextMatrix(r + 15, 0) = "Preschool Room Slot " & r
    Next r
    For r = 1 To 7
        Grid.TextMatrix(r + 29, 0) = "Infant-Care/Mixed Room Slot " & r
    Next r
    
    For r = 1 To 36
        Grid.row = r
        Grid.col = 2
        Grid.CellAlignment = 1
        Grid.col = 3
        Grid.CellAlignment = 1
        Grid.col = 4
        Grid.CellAlignment = 1
        Grid.col = 5
        Grid.CellAlignment = 2
    Next r
    
    Grid.TextMatrix(38, 0) = "Totals"
    DoEvents
    
    For tempdate = dpFromDate.value To dpToDate.value
        ProgressBar.value = (tempdate - dpFromDate.value) / (dpToDate.value - dpFromDate.value) * 100
        DoEvents
        
        If Weekday(tempdate) <> 1 And Weekday(tempdate) <> 7 Then ' if it's a weekday
            IC_index = 0
            PS_index = 0
            SA_index = 0
            Set q = db.Execute("SELECT * FROM clients WHERE startDate <= " & sqlDate(tempdate) & " AND active = 1;")
            With q
                If Not (.EOF And .BOF) Then
                    Do Until .EOF
                        stop_counting = False
                        If !idClient = 29 Or !idClient = 30 Then
                            'Do Nothing... This is Amelia and Catherine
                        Else
                            room = getRoomAtDate(!idClient, tempdate)
                            fc = getFeeClassAtDate(!idClient, tempdate)
                            'cat = feeclass_cats(getFeeClassAtDate(!idClient, tempdate))
                            age = getAgeM(!DOB, tempdate)
                            If age < 155 Then cat = 4
                            If age < 60 Then cat = 2
                            If age < 24 Then cat = 1
                            If cat = 4 Then
                                If fc = 8 Or fc = 9 Then 'THESE ARE FULL TIME SUMMER CARE FOR SCHOOL AGERS
                                    cat = 3
                                End If
                            End If
                            
                            If fc = 6 Or fc = 7 Or fc = 8 Or fc = 10 Then 'THESE ARE PART TIME FEE CLASSES
                                slot_quantity = 0.5
                            Else
                                slot_quantity = 1
                            End If
                            
                            If room = "IC" Then
                                IC_index = IC_index + 1
                                If IC_index > 7 Then stop_counting = True
                                row_index = IC_index + 29
                            ElseIf room = "PS" Then
                                PS_index = PS_index + 1
                                If PS_index > 14 Then stop_counting = True
                                row_index = PS_index + 15
                            Else
                                SA_index = SA_index + 1
                                If SA_index > 15 Then stop_counting = True
                                row_index = SA_index
                            End If
                            'If room = "SA" And cat = 2 Then MsgBox !First & !Last & "  " & getFeeClassAtDate(!idclient, tempdate) & "  Really... This should be based on age at date "
                            If Not stop_counting Then Grid.TextMatrix(row_index, cat) = val(Grid.TextMatrix(row_index, cat)) + slot_quantity
                            If cat = 2 And row_index = 13 Then MsgBox !First & " " & !Last & " " & !DOB & " " & row_index
                        End If
                        .MoveNext
                    Loop
                End If
            End With
        End If
    Next tempdate
    
    For c = 1 To 4
        For r = 1 To 36
            If Grid.TextMatrix(r, c) <> "" Then
                Grid.TextMatrix(r, c) = Int(val(Grid.TextMatrix(r, c)))
            End If
        Next r
    Next c
    
    For r = 1 To 36
        Grid.TextMatrix(r, 5) = Format(val(Grid.TextMatrix(r, 1)) * 14 + val(Grid.TextMatrix(r, 2)) * 8 + val(Grid.TextMatrix(r, 3)) * 8 + val(Grid.TextMatrix(r, 4)) * 3, "0.00")
    Next r
    
    For c = 1 To 5
        d1 = 0
        For r = 1 To 36
            d1 = d1 + val(Grid.TextMatrix(r, c))
        Next r
        Grid.TextMatrix(38, c) = d1
    Next c
    Grid.TextMatrix(38, 5) = Format(val(Grid.TextMatrix(38, 5)), "$0.00")
    
    ProgressBar.Visible = False
    Grid.Visible = True
    
    Set q = Nothing
End Sub

