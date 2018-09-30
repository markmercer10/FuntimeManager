VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmAttendanceEntry 
   Caption         =   "Attendance Entry"
   ClientHeight    =   10215
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   15285
   LinkTopic       =   "Form1"
   ScaleHeight     =   10215
   ScaleWidth      =   15285
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame SaveFrame 
      BorderStyle     =   0  'None
      Caption         =   "Viewing records for"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Left            =   0
      TabIndex        =   16
      Top             =   2520
      Width           =   2655
      Begin VB.ListBox lstMissing 
         Height          =   4155
         Left            =   0
         TabIndex        =   23
         Top             =   2640
         Width           =   2535
      End
      Begin VB.CommandButton saveButn 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   612
         Left            =   0
         TabIndex        =   18
         Top             =   960
         Width           =   2532
      End
      Begin VB.TextBox txtMissing 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   2160
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Viewing records for"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   22
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label LabDate 
         Alignment       =   2  'Center
         Caption         =   "Jun 4, 2016"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   16.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   21
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Missing Days"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   20
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label saved 
         Alignment       =   2  'Center
         Caption         =   "Save Complete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   0
         TabIndex        =   19
         Top             =   1680
         Visible         =   0   'False
         Width           =   2535
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2760
      Top             =   4920
   End
   Begin VB.VScrollBar VScroll 
      Height          =   10212
      Left            =   15000
      TabIndex        =   4
      Top             =   0
      Width           =   252
   End
   Begin VB.Frame DataFrame 
      BorderStyle     =   0  'None
      Height          =   4812
      Left            =   2760
      TabIndex        =   1
      Top             =   0
      Width           =   12252
      Begin VB.CheckBox chkAttended 
         Caption         =   "Absent"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Index           =   0
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   615
         Width           =   1092
      End
      Begin VB.CheckBox chkSick 
         Caption         =   "Sick"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Index           =   0
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   615
         Width           =   975
      End
      Begin VB.CheckBox chkExistsOld 
         ForeColor       =   &H0000FF00&
         Height          =   288
         Index           =   0
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2280
         Visible         =   0   'False
         Width           =   612
      End
      Begin VB.CheckBox chkPaid 
         Caption         =   "Not Paid"
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   405
         Index           =   0
         Left            =   10320
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   615
         Width           =   852
      End
      Begin VB.TextBox fees 
         Alignment       =   1  'Right Justify
         Height          =   288
         Index           =   0
         Left            =   7440
         TabIndex        =   3
         Text            =   "0"
         Top             =   2280
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.ComboBox cboFeeClass 
         Height          =   288
         Index           =   0
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2280
         Visible         =   0   'False
         Width           =   3612
      End
      Begin MSComCtl2.DTPicker signin 
         Height          =   405
         Index           =   0
         Left            =   5400
         TabIndex        =   7
         Top             =   615
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   714
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "hh:mm tt"
         Format          =   124649475
         UpDown          =   -1  'True
         CurrentDate     =   42533
      End
      Begin MSComCtl2.DTPicker signout 
         Height          =   405
         Index           =   0
         Left            =   6360
         TabIndex        =   8
         Top             =   615
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   714
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "hh:mm tt"
         Format          =   124649475
         UpDown          =   -1  'True
         CurrentDate     =   42533
      End
      Begin VB.Label labSADELETE 
         Caption         =   "  > School Age Room"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   15
         Top             =   2760
         Visible         =   0   'False
         Width           =   12255
      End
      Begin VB.Label labPSDELETE 
         Caption         =   "  > Preschool Room"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   14
         Top             =   1800
         Visible         =   0   'False
         Width           =   12255
      End
      Begin VB.Label roomLabels 
         Caption         =   "  > Infant Care Room"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   13
         Top             =   1320
         Width           =   12255
      End
      Begin VB.Image chkExists 
         Height          =   315
         Index           =   0
         Left            =   11400
         Top             =   675
         Width           =   375
      End
      Begin VB.Label labFeeClass 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fee Class"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   305
         Index           =   0
         Left            =   7320
         TabIndex        =   12
         Top             =   645
         Width           =   3375
      End
      Begin VB.Label Headers 
         Caption         =   $"dlgAttendance.frx":0000
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   12372
      End
      Begin VB.Label labClient 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   0
         TabIndex        =   5
         Top             =   615
         Width           =   3375
      End
   End
   Begin MSComCtl2.MonthView MonthView 
      Height          =   2520
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4445
      _Version        =   393216
      ForeColor       =   0
      BackColor       =   4210752
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthBackColor  =   16777215
      StartOfWeek     =   124583937
      TitleBackColor  =   16755302
      CurrentDate     =   42533
   End
   Begin VB.Label weekend 
      Alignment       =   2  'Center
      Caption         =   "Weekend"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   48
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2880
      TabIndex        =   24
      Top             =   1200
      Width           =   11655
   End
End
Attribute VB_Name = "frmAttendanceEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Changed As Boolean
Private prevent As Boolean
Private prevent_date As Date
Private autofilling As Boolean
Private lineheight As Long
Private lastline As Long
Private lastloaded As Long

Public selected As Long

Sub fillClientList(ByVal d As Date)
    Dim clients As ADODB.Recordset
    Dim index As Long
    Dim i As Long
    Dim fc As ADODB.Recordset
    Dim rm As ADODB.Recordset
    Dim section As Byte
    Dim client_hash() As Byte
    Dim room As String
    
    clearClientList
        
    Set rm = db.Execute("SELECT * FROM rooms")
    With rm
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do Until .EOF
                If !idroom > 1 Then
                    
                    'If Not (roomLabels(!idroom - 1) Is LOADED) Then Load roomLabels(!idroom - 1)
                    If roomLabels.count < !idroom Then Load roomLabels(!idroom - 1)
                End If
                roomLabels(!idroom - 1).backcolor = val("&H" & !color1)
                roomLabels(!idroom - 1).Tag = !color2
                roomLabels(!idroom - 1).Caption = "  > " & !name
                roomLabels(!idroom - 1).Visible = True
                .MoveNext
            Loop
        End If
    End With
    
    'Set clients = db.Execute("SELECT * FROM clients WHERE active=1 ORDER BY room, last, first")
    Set clients = db.Execute("SELECT * FROM clients WHERE startDate <= " & sqlDate(d) & " AND (endDate >= " & sqlDate(d) & " OR active=1 ) ORDER BY idClient DESC")
    If Not (clients.EOF And clients.BOF) Then
        clients.MoveFirst
        ReDim client_hash(clients!idClient) As Byte
    End If
    
    
    
    ' if its the weekend... empty list.
    If isWeekend(d) Then
        DataFrame.Visible = False
    Else
        DataFrame.Visible = True
        VScroll.Visible = True
        
        'create hash table to store room at date info
        With clients
            If Not (.EOF And .BOF) Then
                .MoveFirst
                Do Until .EOF
                    'find out if client is scheduled today
                    Set fc = db.Execute("SELECT * FROM fee_classes WHERE idFeeClasses = " & getFeeClassAtDate(!idClient, d))
                    If Not (fc.EOF And fc.BOF) Then
                        If fc.fields(weekdayToLetter(Weekday(d))) > 0 Then
                            
                            'find what room the client is in
                            room = getRoomAtDate(!idClient, d)
                            
                            'new
                            If Not (rm.EOF And rm.BOF) Then
                                rm.MoveFirst
                                Do Until rm.EOF
                                    If rm!Abbreviation = room Then
                                        client_hash(!idClient) = rm!idroom
                                        Exit Do
                                    End If
                                    rm.MoveNext
                                Loop
                            End If
                        End If
                    End If
                    .MoveNext
                Loop
            End If
        End With
    End If
    
    Set fc = db.Execute("SELECT * FROM fee_classes")
    Set clients = db.Execute("SELECT * FROM clients WHERE startDate <= " & sqlDate(d) & " AND (endDate >= " & sqlDate(d) & " OR active=1 ) ORDER BY last, first")
    
    index = 0
    'labIC.Top = 240 'Headers.height
    roomLabels(0).Top = 240
    'labIC.backcolor = &HBBFFFF
    'labPS.backcolor = &HBBFFBB
    'labSA.backcolor = &HFFDDBB
    
    If Not (rm.EOF And rm.BOF) Then
        rm.MoveFirst
        Do Until rm.EOF
            section = rm!idroom
            'For section = 1 To 3
            With clients
                If Not (.EOF And .BOF) Then
                    .MoveFirst
                    Do Until .EOF
                        If client_hash(!idClient) = section Then
                            If index > lastloaded Then
                                newLine index, section
                            Else
                                reactivateLine index, section
                            End If
                            
                            With fc
                                If Not (.BOF And .EOF) Then
                                    .MoveFirst
                                    Do Until .EOF
                                        'cboFeeClass(index).AddItem !Description & " - $" & !charge, !idFeeClasses - 1
                                        If !idFeeClasses = getFeeClassAtDate(clients!idClient, d) Then labFeeClass(index) = !Description & " - $" & !charge
                                        'If !idFeeClasses = clients!feeClassID Then cboFeeClass(index).ListIndex = !idFeeClasses - 1
                                        .MoveNext
                                    Loop
                                End If
                            End With
                            labClient(index) = !Last & ", " & !First
                            labClient(index).Tag = !idClient
                            labClient(index).backcolor = val("&H" & roomLabels(section - 1).Tag)
                            'If section = 1 Then labClient(Index).backcolor = IC_color
                            'If section = 2 Then labClient(Index).backcolor = PS_color
                            'If section = 3 Then labClient(Index).backcolor = SA_color
                            lastline = index
                            index = index + 1
                        End If
                        
                        .MoveNext
                    Loop
                End If
            End With
            'Next section
            rm.MoveNext
        Loop
    End If
    
    For i = index To lastloaded
        labClient(i).Caption = ""
        labFeeClass(i).Caption = ""
        chkAttended(i).value = 0
        chkPaid(i).value = 0
    Next i
    
    DataFrame.height = (index + 1) * lineheight + roomLabels(0).height * roomLabels.count
    
    If chkAttended(0).Visible Then chkAttended(0).SetFocus
    
    'largechg = content height
    'max = window height - content height
    VScroll.max = DataFrame.height - (Me.height - 575)
    If VScroll.max < 0 Then VScroll.max = 0
    VScroll.LargeChange = DataFrame.height  'Me.height
    If VScroll.max >= 30 Then VScroll.SmallChange = VScroll.max / 30
    
    Set fc = Nothing
    Set clients = Nothing
End Sub

Sub clearClientList()
    Dim i As Long
    For i = 0 To lastloaded
        labClient(i).Caption = ""
        labClient(i).Visible = False
        chkAttended(i).Visible = False
        signin(i).Visible = False
        signout(i).Visible = False
        labFeeClass(i).Visible = False
        chkPaid(i).Visible = False
        chkExists(i).Visible = False

    Next i
End Sub

Sub fillAttendanceData()
    Dim att As ADODB.Recordset
    Dim avgTime As ADODB.Recordset
    Dim index As Long
    'Dim intime As Double
    'Dim outtime As Double
    Dim intimes(1 To 20) As Double
    Dim outtimes(1 To 20) As Double
    Dim countTimes As Byte
    
    autofilling = True
    Changed = False
    
    Set att = db.Execute("SELECT * FROM attendance WHERE Date = " & sqlDate(MonthView.value))
    
    For index = 0 To labClient.count - 1
        chkExists(index).Tag = 0
        chkExists_chg (index)
        chkAttended(index) = 0
        chkSick(index) = 0
        signin(index) = CDate("8:00")
        signout(index) = CDate("17:00")
        chkPaid(index) = 0
    Next index
    
    With att
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do Until .EOF
                For index = 0 To labClient.count - 1
                    If labClient(index).Tag = !idClient Then
                        chkExists(index).Tag = 1
                        chkExists_chg (index)
                        chkAttended(index) = !attended
                        chkSick(index) = !sick
                        signin(index) = !signin
                        signout(index) = !signout
                        chkPaid(index) = !paid
                    End If
                Next index
                .MoveNext
            Loop
        End If
    End With
    
    '**********  Average Times  **********
    If labClient.count <= 1 And labClient(0).Tag = "" Then
        'do nothing
    Else
        For index = 0 To labClient.count - 1
            If chkExists(index).Tag = 0 Then
                Set avgTime = db.Execute("SELECT * FROM attendance WHERE idclient = " & labClient(index).Tag & " AND attended = 1 ORDER BY date DESC LIMIT 10")
                countTimes = 0
                'intime = 0
                'outtime = 0
                With avgTime
                    If Not (.EOF And .BOF) Then
                        .MoveFirst
                        Do Until .EOF
                            countTimes = countTimes + 1
                            'intime = intime + CDbl(!signin)
                            'outtime = outtime + CDbl(!signout)
                            intimes(countTimes) = CDbl(!signin)
                            outtimes(countTimes) = CDbl(!signout)
                            .MoveNext
                        Loop
                    End If
                End With
                If countTimes > 0 Then
                    ShellSort intimes, 20, True
                    ShellSort outtimes, 20, True
                    'intime = intime / countTimes
                    'outtime = outtime / countTimes
                    'signin(index) = CDate(Format(intime, "hh:mm:00"))
                    'signout(index) = CDate(Format(outtime, "hh:mm:00"))
                    signin(index) = CDate(Format(intimes(5), "hh:mm:00"))
                    signout(index) = CDate(Format(outtimes(5), "hh:mm:00"))
                End If
            End If
        Next index
    End If
    
    autofilling = False
    
    Set att = Nothing
    Set avgTime = Nothing
End Sub

Private Sub cboFeeClass_Change(index As Integer)
    cboFeeClass_Click index
End Sub

Private Sub cboFeeClass_Click(index As Integer)
    'Dim fc As ADODB.Recordset
    'Set fc = db.Execute("SELECT * FROM fee_classes WHERE idFeeClasses = " & cboFeeClass(index).ListIndex + 1)
    'With fc
    '    If Not (.BOF And .EOF) Then
    '        .MoveFirst
    '        fees(index) = !charge
    '    End If
    'End With
    'Set fc = Nothing
End Sub


Private Sub chkAttended_Click(index As Integer)
    If chkAttended(index) = 1 Then
        selected = index
        chkAttended(index).Caption = "Present"
        chkAttended(index).forecolor = &HC000&
        signin(index).Enabled = True
        signout(index).Enabled = True
        chkPaid(index).Enabled = True
        chkAttended(index).width = 2067
        
        'chk if this attended day is already paid on a receipt.  if so automatically set it to paid.
        Dim q As ADODB.Recordset
        Set q = db.Execute("SELECT * FROM payments WHERE idClient = " & labClient(index).Tag & " AND fromdate <= " & sqlDate(MonthView.value) & " AND todate >= " & sqlDate(MonthView.value))
        If Not (q.EOF And q.BOF) Then
            chkPaid(index).value = 1
        End If
        Set q = Nothing
        DoEvents
        
        If Not autofilling Then dlgSetTimes.Show 1
    Else
        chkAttended(index).Caption = "Absent"
        chkAttended(index).forecolor = vbRed
        signin(index).Enabled = False
        signout(index).Enabled = False
        chkPaid(index).Enabled = False
        chkAttended(index).width = 1092
    End If
    Changed = True
    saved.Visible = False
End Sub

Private Sub chkExists_chg(index As Integer)
    If chkExists(index).Tag = 1 Then
        chkExists(index).Picture = frmMain.ImageList.ListImages("check").Picture
    Else
        chkExists(index).Picture = Nothing
    End If
End Sub



Private Sub chkPaid_Click(index As Integer)
    If chkPaid(index) = 1 Then
        chkPaid(index).Caption = "Paid"
        chkPaid(index).forecolor = vbGreen
    Else
        chkPaid(index).Caption = "Not Paid"
        chkPaid(index).forecolor = vbRed
    End If
    Changed = True
    saved.Visible = False
End Sub




Private Sub chkSick_Click(index As Integer)
    If chkSick(index).value = 1 Then
        chkSick(index).forecolor = vbGreen
    Else
        chkSick(index).forecolor = vbRed
    End If
End Sub

Private Sub Form_Load()
    lastloaded = 0
    lineheight = labClient(0).height
    MonthView.value = Date
    
    updateMissingDays
End Sub

Sub updateMissingDays()
    Dim q As ADODB.Recordset
    Dim d As Date
    Dim registered(60) As Long
    Dim r As Byte
    
    lstMissing.Clear
    txtmissingdays = ""
    For r = 1 To 60
        registered(r) = 0
    Next r
    r = 0
    For d = Date - 60 To Date - 1
        r = r + 1
        If Not isWeekend(d) Then
            Set q = db.Execute("SELECT * FROM attendance WHERE date = " & sqlDate(d))
            If q.EOF And q.BOF Then
                If Not isStatHoliday(d) Then
                    txtMissing = txtMissing & shortDate(d) & vbCrLf
                    lstMissing.AddItem shortDate(d)
                End If
            Else
                q.MoveFirst
                Do Until q.EOF
                    registered(r) = registered(r) + 1
                    q.MoveNext
                Loop
            End If
        End If
    Next d
    
    lstMissing.AddItem ""
    lstMissing.AddItem "Suspected Partly Missing"
    r = 0
    For d = Date - 60 To Date - 1
        r = r + 1
        If Not isStatHoliday(d) Then
            If Not isWeekend(d) Then
                'Set q = db.Execute("SELECT * FROM attendance WHERE date = " & sqlDate(d))
                Set q = db.Execute("SELECT SUM(attended) AS attended FROM attendance WHERE date = " & sqlDate(d))
                'MsgBox d & " " & q!attended & " / " & registered(r)
                If q!attended < registered(r) * 0.4 Then
                    lstMissing.AddItem shortDate(d) & " Low"
                Else
                    Set q = db.Execute("SELECT room, SUM(attended) AS attended FROM attendance INNER JOIN clients ON clients.idClient = attendance.idClient WHERE date = " & sqlDate(d) & " GROUP BY room")
                    With q
                        If Not (.EOF And .BOF) Then
                            .MoveFirst
                            Do Until .EOF
                                'MsgBox d & " - " & !room & " " & !attended
                                If !attended <= 1 Then
                                    If !room = "SA" Then
                                        lstMissing.AddItem shortDate(d) & " SA room"
                                    ElseIf !room = "PS" Then
                                        lstMissing.AddItem shortDate(d) & " PS room"
                                    ElseIf !room = "IC" Then
                                        lstMissing.AddItem shortDate(d) & " IC room"
                                    End If
                                    Exit Do
                                End If
                                .MoveNext
                            Loop
                        End If
                    End With
                End If
            End If
        End If
    Next d
    
    
    Set q = Nothing
End Sub

Private Sub Form_Resize()
    '15495 min wid
    '29000 max wid
    If Me.width < 15495 Then Me.width = 15495
    DoEvents
    'MonthView.Font.Size = CLng(Me.width / 2000)
    MonthView.Font.Size = CLng((Me.width - 6000) / 1200)
    DoEvents
    SaveFrame.Top = MonthView.height + 100
    DataFrame.Left = MonthView.width + 200
    If DataFrame.Left < SaveFrame.width Then DataFrame.Left = SaveFrame.width
    VScroll.height = Me.height - 575
    VScroll.Left = DataFrame.Left + DataFrame.width
    'VScroll.height = DataFrame.height
    If DataFrame.height > Me.height - 575 Then
        
        VScroll.max = DataFrame.height - (Me.height - 575)
        If VScroll.max < 0 Then VScroll.max = 0
        VScroll.LargeChange = DataFrame.height  'Me.height
        If VScroll.max >= 30 Then VScroll.SmallChange = VScroll.max / 30
        
        VScroll.Visible = True
    Else
        VScroll.value = 0
        VScroll.Visible = False
    End If
    
    If MonthView.width > SaveFrame.width Then
        SaveFrame.Left = MonthView.width / 2 - SaveFrame.width / 2
    Else
        SaveFrame.Left = 0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Changed Then
        If MsgBox("Some data for this date has been modified but not saved.  If you close the window the changes will be lost.  Are you sure you want to close the window?", vbYesNo) = vbNo Then
            Cancel = True
        Else
            Cancel = False
        End If
    End If
End Sub

Private Sub MonthView_DateClick(ByVal DateClicked As Date)
    'MsgBox Changed
    MonthView_Validate False
    If prevent Then
        'MsgBox MonthView.value & "    " & DateClicked
        MonthView.value = prevent_date
        chkAttended(0).SetFocus
    Else
        If isWeekend(DateClicked) Or DateClicked > Date Or isStatHoliday(DateClicked) Then
            SaveButn.Enabled = False
            If DateClicked > Date Then
                SaveButn.Caption = "Future!"
            ElseIf isWeekend(DateClicked) Then
                SaveButn.Caption = "Weekend"
                weekend.Caption = WeekdayName(Weekday(DateClicked))
            Else
                SaveButn.Caption = "Stat Holiday"
            End If
        Else
            SaveButn.Enabled = True
            SaveButn.Caption = "Save"
        End If
        LabDate = shortDate(MonthView.value)
        'Timer1.Enabled = True
        DoEvents
        fillClientList DateClicked
        DoEvents
        fillAttendanceData
        If chkAttended(0).Visible Then chkAttended(0).SetFocus
        Changed = False
        prevent_date = DateClicked
    End If
End Sub

Private Sub MonthView_Validate(Cancel As Boolean)
    If Changed And MonthView.value <> prevent_date Then
        If MsgBox("Some data for this date has been modified but not saved.  If you select another date the changes will be lost.  Are you sure you want to select another date?", vbYesNo) = vbNo Then
            Cancel = True
            prevent = True
        Else
            Cancel = False
            prevent = False
        End If
    End If
End Sub


Private Sub SaveButn_Click()
    Dim index As Long
    Dim sql As String
    
    For index = 0 To lastline
        If chkExists(index).Tag = 1 Then
            'update
            sql = "UPDATE attendance SET "
            sql = sql & "attended=" & chkAttended(index).value & ","
            sql = sql & "sick=" & chkSick(index).value & ","
            sql = sql & "signin=" & sqlTime(signin(index).value) & ","
            sql = sql & "signout=" & sqlTime(signout(index).value) & ","
            sql = sql & "paid=" & chkPaid(index).value
            sql = sql & " WHERE idClient=" & labClient(index).Tag & " AND date=" & sqlDate(MonthView.value)
            
        Else
            'insert
            sql = "INSERT INTO attendance "
            sql = sql & "(idClient,date,attended,sick,signin,signout,paid)"
            sql = sql & " VALUES ("
            sql = sql & labClient(index).Tag & ","
            sql = sql & sqlDate(MonthView.value) & ","
            sql = sql & chkAttended(index).value & ","
            sql = sql & chkSick(index).value & ","
            sql = sql & sqlTime(signin(index).value) & ","
            sql = sql & sqlTime(signout(index).value) & ","
            sql = sql & chkPaid(index).value
            sql = sql & ")"
            
        End If
        
        db.Execute sql
        'db.Execute "SET @attid = LAST_INSERT_ID();"
        'Dim q As ADODB.Recordset
        'Set q = db.Execute("SELECT @attid AS att;")
        'MsgBox q!att
        
    Next index

    fillAttendanceData
    updateMissingDays
    
    Changed = False
    prevent = False
    saved.Visible = True
End Sub



Private Sub signin_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
    If Hour(signin(index).value) < 7 Or Hour(signin(index).value) > 18 Or (Hour(signin(index).value) = 18 And Minute(signin(index).value) > 0) Then
        signin(index).Font.bold = True
        signin(index).Font.Italic = True
        signin(index).ToolTipText = "The selected time is outside the normal licensed hours! (Check AM/PM)"
    Else
        signin(index).Font.bold = False
        signin(index).Font.Italic = False
        signin(index).ToolTipText = ""
    End If
    Changed = True
    saved.Visible = False
End Sub


Private Sub signout_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
    If Hour(signout(index).value) < 7 Or Hour(signout(index).value) > 18 Or (Hour(signout(index).value) = 18 And Minute(signout(index).value) > 0) Then
        signout(index).Font.bold = True
        signout(index).Font.Italic = True
        signout(index).ToolTipText = "The selected time is outside the normal licensed hours! (Check AM/PM)"
    Else
        signout(index).Font.bold = False
        signout(index).Font.Italic = False
        signout(index).ToolTipText = ""
    End If
    Changed = True
    saved.Visible = False
End Sub


Private Sub Timer1_Timer()
    Timer1.Enabled = False
    MonthView_DateClick Date
    
    'fillClientList
End Sub

Sub newLine(ByVal index As Long, ByVal section As Byte)
    Dim lineTop As Long
    Dim headerheight As Long
    'headerheight = labIC.height
    headerheight = roomLabels(0).height 'labIC.height
    
    'lineTop = labClient(index - 1).Top + lineheight
    lineTop = labClient(0).Top + index * lineheight + headerheight * (section - 1)
    
    'If section = 1 Then labPS.Top = lineTop + lineheight
    'If section = 2 Then labSA.Top = lineTop + lineheight
    If section > 0 And section < roomLabels.count Then roomLabels(section).Top = lineTop + lineheight
    
    Load labClient(index)
        labClient(index).Top = lineTop
        labClient(index).Visible = True
    Load chkAttended(index)
        chkAttended(index).Top = lineTop
        chkAttended(index).Visible = True
    Load chkSick(index)
        chkSick(index).Top = lineTop
        chkSick(index).Visible = True
    Load signin(index)
        signin(index).Top = lineTop
        signin(index).Visible = True
    Load signout(index)
        signout(index).Top = lineTop
        signout(index).Visible = True
    Load labFeeClass(index)
        labFeeClass(index).Top = lineTop + 30
        labFeeClass(index).Visible = True
    'Load fees(index)
    '    fees(index).Top = lineTop
    '    fees(index).Visible = True
    Load chkPaid(index)
        chkPaid(index).Top = lineTop
        chkPaid(index).Visible = True
    Load chkExists(index)
        chkExists(index).Top = lineTop + 60
        chkExists(index).Visible = True
    
    lastloaded = index
End Sub

Sub reactivateLine(ByVal index As Long, ByVal section As Byte)
    Dim lineTop As Long
    Dim headerheight As Long
    headerheight = roomLabels(0).height 'labIC.height
    'lineTop = labClient(index - 1).Top + lineheight
    lineTop = labClient(0).Top + index * lineheight + headerheight * (section - 1)
    
    'If section = 1 Then labPS.Top = lineTop + lineheight
    'If section = 2 Then labSA.Top = lineTop + lineheight
    If section > 0 And section < roomLabels.count Then roomLabels(section).Top = lineTop + lineheight
    
    'Load labClient(Index)
        labClient(index).Top = lineTop
        labClient(index).Visible = True
    'Load chkAttended(Index)
        chkAttended(index).Top = lineTop
        chkAttended(index).Visible = True
        
        chkSick(index).Top = lineTop
        chkSick(index).Visible = True
    'Load signin(Index)
        signin(index).Top = lineTop
        signin(index).Visible = True
    'Load signout(Index)
        signout(index).Top = lineTop
        signout(index).Visible = True
    'Load labFeeClass(Index)
        labFeeClass(index).Top = lineTop + 30
        labFeeClass(index).Visible = True
    'Load fees(index)
    '    fees(index).Top = lineTop
    '    fees(index).Visible = True
    'Load chkPaid(Index)
        chkPaid(index).Top = lineTop
        chkPaid(index).Visible = True
    'Load chkExists(Index)
        chkExists(index).Top = lineTop + 60
        chkExists(index).Visible = True
    
    
End Sub


Sub clearLines() 'NOT USED
    'Dim L As Long
    'For L = labClient.count - 1 To 1 Step -1
    '    Unload labClient(L)
    '    Unload chkAttended(L)
    '    Unload signin(L)
    '    Unload signout(L)
    '    Unload labFeeClass(L)
    '    Unload chkPaid(L)
    '    Unload chkExists(L)
    'Next L
End Sub

Private Sub VScroll_Change()
    'when value = 0 then top should be 0
    'when value = max then top should be -frame.height - me.height
    If VScroll.max = 0 Then
        DataFrame.Top = 0
    Else
        DataFrame.Top = -(VScroll.value / VScroll.max) * (DataFrame.height - (Me.height - 500))
    End If
End Sub

