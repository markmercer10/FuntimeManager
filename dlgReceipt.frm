VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form dlgReceipt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receipt"
   ClientHeight    =   7425
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   11100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   11100
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox custom_prompt 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   288
      Left            =   1560
      TabIndex        =   28
      Text            =   "Describe why a custom amount was entered."
      Top             =   2880
      Visible         =   0   'False
      Width           =   3612
   End
   Begin VB.CommandButton updateButn 
      Caption         =   "Update"
      Height          =   252
      Left            =   1320
      TabIndex        =   27
      Top             =   3960
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.TextBox txtDetails 
      Height          =   1452
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   24
      Top             =   2400
      Width           =   6012
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3600
      Top             =   720
   End
   Begin VB.TextBox ID 
      Height          =   288
      Left            =   4200
      TabIndex        =   20
      Top             =   1560
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.CommandButton printButn 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   6600
      TabIndex        =   19
      Top             =   3360
      Width           =   1692
   End
   Begin VB.TextBox attendanceText 
      Height          =   3132
      Left            =   240
      MultiLine       =   -1  'True
      OLEDragMode     =   1  'Automatic
      TabIndex        =   17
      Top             =   4200
      Visible         =   0   'False
      Width           =   4452
   End
   Begin VB.TextBox txtWeeks 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   6600
      TabIndex        =   14
      Text            =   "0"
      Top             =   2520
      Width           =   492
   End
   Begin VB.CommandButton saveButn 
      Caption         =   "Save && Close"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   8400
      TabIndex        =   10
      Top             =   3360
      Width           =   2412
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   624
      Left            =   8400
      TabIndex        =   9
      Top             =   2520
      Width           =   2412
   End
   Begin MSComCtl2.DTPicker dpDate 
      Height          =   372
      Left            =   8880
      TabIndex        =   6
      Top             =   720
      Width           =   2052
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTitleBackColor=   12632064
      CustomFormat    =   "MMM d, yyyy"
      Format          =   123863043
      CurrentDate     =   42533
   End
   Begin MSComCtl2.DTPicker dpTodate 
      Height          =   312
      Left            =   5520
      TabIndex        =   5
      Top             =   1080
      Width           =   1692
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTitleBackColor=   12632064
      CustomFormat    =   "MMM d, yyyy"
      Format          =   123863043
      CurrentDate     =   42533
   End
   Begin MSComCtl2.DTPicker dpFromdate 
      Height          =   312
      Left            =   5520
      TabIndex        =   4
      Top             =   720
      Width           =   1692
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTitleBackColor=   12632064
      CustomFormat    =   "MMM d, yyyy"
      Format          =   123863043
      CurrentDate     =   42533
   End
   Begin VB.ComboBox cboFeeClass 
      Height          =   288
      Left            =   6600
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1680
      Width           =   4212
   End
   Begin VB.ListBox attendanceList 
      Height          =   2985
      Left            =   240
      TabIndex        =   1
      Top             =   4200
      Width           =   4452
   End
   Begin VB.ComboBox cboClient 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1680
      Width           =   4212
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "To Date"
      Height          =   252
      Left            =   4560
      TabIndex        =   8
      Top             =   1080
      Width           =   852
   End
   Begin VB.Label labSignature 
      Alignment       =   1  'Right Justify
      Caption         =   "Signature: ___________________________________________________"
      Height          =   372
      Left            =   4920
      TabIndex        =   26
      Top             =   4200
      Visible         =   0   'False
      Width           =   5892
   End
   Begin VB.Label labDetails 
      Caption         =   "Other Details"
      Height          =   252
      Left            =   240
      TabIndex        =   25
      Top             =   2160
      Width           =   2052
   End
   Begin VB.Label Label1 
      Caption         =   "Received From"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   240
      TabIndex        =   23
      Top             =   960
      Width           =   1572
   End
   Begin VB.Label labFor 
      Caption         =   "For Child Care of"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   240
      TabIndex        =   22
      Top             =   1680
      Width           =   1812
   End
   Begin VB.Label labReceivedFrom 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1800
      TabIndex        =   21
      Top             =   960
      Width           =   3132
   End
   Begin VB.Label Label8 
      Caption         =   "Funtime Child Care Center - Receipt of Payment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   240
      TabIndex        =   18
      Top             =   120
      Width           =   10692
   End
   Begin VB.Label labFinePrint 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   6600
      TabIndex        =   16
      Top             =   4560
      Width           =   4212
   End
   Begin VB.Label Label7 
      Caption         =   "Weeks billed"
      Height          =   252
      Left            =   6600
      TabIndex        =   15
      Top             =   2280
      Width           =   1332
   End
   Begin VB.Label Label6 
      Caption         =   "Amount"
      Height          =   252
      Left            =   8400
      TabIndex        =   13
      Top             =   2280
      Width           =   1572
   End
   Begin VB.Label Label5 
      Caption         =   "Fee Class"
      Height          =   252
      Left            =   6720
      TabIndex        =   12
      Top             =   1440
      Width           =   2172
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Receipt Dated"
      Height          =   372
      Left            =   7320
      TabIndex        =   11
      Top             =   720
      Width           =   1452
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "From Date"
      Height          =   252
      Left            =   4560
      TabIndex        =   7
      Top             =   720
      Width           =   852
   End
   Begin VB.Label labAttendance 
      Caption         =   "Attendence"
      Height          =   252
      Left            =   240
      TabIndex        =   2
      Top             =   3960
      Width           =   2292
   End
End
Attribute VB_Name = "dlgReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rate As Long
Private doublepay As Boolean
Private desc_manual_edit As Boolean

Sub calculate_amount()
    Dim d As Date
    Dim tot As Double
    Dim c As Long
    Dim lastfeeclass As Long
    Dim lastfees As Double
    Dim thisfees As Double
    Dim subcount As Long
    Dim editing_desc As Boolean
    
    c = getClientID(cboClient)
    txtDetails = ""
    tot = 0
    thisfees = 0
    lastfees = 0
    lastfeeclass = 0
    subcount = 0
    editing_desc = False
    'MsgBox c
    For d = dpFromDate.value To dpToDate.value
        If Weekday(d) <> 1 And Weekday(d) <> 7 Then ' if it's not the weekend
            'MsgBox d & "  " & getFeesAtDate(c, d)
            If isStatHoliday(d) Then ' parents pay minimum $30 for a stat holiday regardless of regular fees and regardless of attendance.
                If getFeesAtDate(c, d) / 5# <= 30 Then
                    thisfees = 30
                Else
                    thisfees = getFeesAtDate(c, d) / 5#
                End If
                txtDetails = txtDetails & "STAT HOLIDAY " & shortDate(d) & vbCrLf
            End If
            
            'This section adds automatic comments to a receipt if there are two different fee classes present
            If getFeeClassAtDate(c, d) <> lastfeeclass Or (editing_desc And d = dpToDate.value) Then
                If lastfeeclass = 0 Then
                    lastfeeclass = getFeeClassAtDate(c, d)
                Else
                    If Not desc_manual_edit Then
                        If Not editing_desc Then txtDetails = ""
                        editing_desc = True
                        If d = dpToDate.value Then subcount = subcount + 1 ' adjust for the last entry
                        txtDetails = txtDetails & subcount & " days @ $" & lastfees & vbCrLf
                        subcount = 0
                    End If
                    lastfeeclass = getFeeClassAtDate(c, d)
                End If
            End If
            subcount = subcount + 1
            tot = tot + thisfees
            lastfees = thisfees
        Else
            If (editing_desc And d = dpToDate.value) Then
                txtDetails = txtDetails & subcount & " days @ $" & lastfees & vbCrLf
            End If
        End If
    Next d
    txtAmount = "$" & Format(tot, "0.00")
    
    'the old method
    'txtAmount = "$" & Format(rate * val(txtWeeks), "0.00")
End Sub
Sub calculate_attendance()
    Dim al As ADODB.Recordset
    
    attendanceList.Clear
    doublepay = False
    
    Set al = db.Execute("SELECT * FROM attendance WHERE idClient = " & getClientID(cboClient) & " AND date >= " & sqlDate(dpFromDate.value) & " AND date <= " & sqlDate(dpToDate.value) & " ORDER BY date ASC")
    With al
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do Until .EOF
                s = shortDate(!Date)
                If !attended = 1 Then
                    s = s & "  -  " & Format(!signin, "h:nn AMPM") & " - " & Format(!signout, "h:nn AMPM")
                Else
                    s = s & "  -  Did not attend"
                End If
                attendanceList.AddItem s
                If !paid Then doublepay = True
                .MoveNext
            Loop
        End If
    End With
    
    attendanceList.Visible = True
    attendanceText.Visible = False
    
    Set al = Nothing
End Sub

Sub calculate_weeks()
    Dim ddiff As Long
    ddiff = dpToDate.value - dpFromDate.value + 2
    txtWeeks = CLng(ddiff / 7#)
End Sub



Sub guessDates()
    Dim ddiff As Long
    Dim tempdate As Date
    'Dec 29, 2016 changed this function to return dates Sunday through Saturday
    
    If ID = "" Then ' if we're editing stop guessing!
        If cboClient.ListIndex = -1 Then
            ddiff = 6 - Weekday(dpDate.value)
            If ddiff > 2 Then ddiff = ddiff - 7
            'monday is 2
            'friday is 6
            
            'sun 1 5 -2 -2
            'mon 2 4  4 -3
            'tue 3 3  3 -4
            'wed 4 2  2  2
            'thu 5 1  1  1
            'fri 6 7  0  0
            'sat 7 6 -1 -1
            dpToDate.value = dpDate.value + ddiff
            dpFromDate.value = dpToDate.value - 4
            
        Else
            Dim q As ADODB.Recordset
            Set q = db.Execute("SELECT * FROM payments WHERE idClient = " & getClientID(cboClient.Text) & " ORDER BY todate DESC LIMIT 1")
            With q
                If .EOF And .BOF Then
                    'sun 1  1
                    'mon 2  0
                    'tue 3 -1
                    'wed 4 -2
                    'thu 5 -3
                    'fri 6 -4
                    'sat 7  2
                    dpFromDate.value = dpToDate.value + 3 - (val(txtWeeks) * 7)
                    ddiff = 2 - Weekday(dpFromDate.value)
                    If ddiff > 6 Then ddiff = ddiff + 7
                    dpToDate.value = dpFromDate.value + ddiff
                Else
                    tempdate = !toDate + 3
                    ddiff = 2 - Weekday(tempdate)
                    dpFromDate.value = tempdate + ddiff
                    dpToDate.value = dpFromDate.value + (val(txtWeeks) * 7) - 3
                End If
            End With
            Set q = Nothing
        End If
    End If
    
End Sub

Private Sub cboClient_Change()
    cboClient_Click
End Sub

Private Sub cboClient_Click()
    Dim c As Long
    Dim q As ADODB.Recordset
    Dim fc As ADODB.Recordset
    Dim fcd As String
    Dim s As String
    
    
    Set q = db.Execute("SELECT * FROM clients WHERE idClient = " & getClientID(cboClient))
    q.MoveFirst
    'labReceivedFrom = q!receivedFrom
    'MsgBox "|" & labReceivedFrom & "|" & (labReceivedFrom = "")
    'If labReceivedFrom = "" Then
    If ID = "" Then
        Dim p As ADODB.Recordset
        Dim parentname As String
        Set p = getParents(q!idClient)
        If Not (p.EOF And p.BOF) Then
            p.MoveFirst
            parentname = "Parent: " & p!name
        End If
        labReceivedFrom = parentname
        txtWeeks = q!payperiod
        If InStr(1, labReceivedFrom, " ") = 0 And ID = "" Or labReceivedFrom = "" Then
            labReceivedFrom = InputBox("Enter Received From", "Receipt", parentname)
        End If
    End If
    
    Set fc = db.Execute("SELECT * FROM fee_classes WHERE idFeeClasses = " & q!feeClassID)
    If Not (fc.EOF And fc.BOF) Then
        fc.MoveFirst
        fcd = fc!Description & " - $" & fc!charge
        'MsgBox fcd
        
        'Set fc = db.Execute("SELECT * FROM fee_classes")
        For c = 0 To cboFeeClass.ListCount - 1
            If cboFeeClass.List(c) = fcd Then
                cboFeeClass.ListIndex = c
                rate = fc!charge
            End If
        Next c
    End If
    guessDates
        
    calculate_attendance
    calculate_weeks
    calculate_amount
    
    Set q = Nothing
    Set fc = Nothing
End Sub

Private Sub cboFeeClass_Change()
    cboFeeClass_Click
End Sub

Private Sub cboFeeClass_Click()
    rate = val(MiD$(cboFeeClass.Text, InStr(1, cboFeeClass.Text, "$") + 1))
    calculate_amount
End Sub

Private Sub custom_prompt_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    custom_prompt.Visible = False
End Sub

Private Sub dpDate_Change()
    dpDate_Click
End Sub


Private Sub dpDate_Click()
    guessDates
End Sub


Private Sub dpFromDate_Change()
    dpFromdate_Click
End Sub

Private Sub dpFromdate_Click()
    calculate_attendance
    calculate_weeks
    calculate_amount
End Sub

Private Sub dpToDate_Change()
    dpToDate_Click
End Sub

Private Sub dpToDate_Click()
    calculate_attendance
    calculate_weeks
    calculate_amount
End Sub

Function getAttendanceText() As String
    Dim count As Long
    getAttendanceText = ""
    For count = 0 To attendanceList.ListCount - 1
        getAttendanceText = getAttendanceText & attendanceList.List(count) & vbCrLf
    Next count
End Function

Function getClientID(ByVal s As String) As Long
    If s <> "" Then
        getClientID = val(MiD$(Left$(s, InStr(1, s, ")") - 1), 2))
    End If
End Function

Private Sub Form_Load()
    desc_manual_edit = False
End Sub

Private Sub labReceivedFrom_DblClick()
    labReceivedFrom = InputBox("Enter Received From", "Receipt", labReceivedFrom)
End Sub

Private Sub printButn_Click()
        If InStr(1, labReceivedFrom, " ") = 0 Then
            labReceivedFrom = InputBox("Enter Received From", "Receipt", labReceivedFrom)
        End If
        Me.backcolor = vbWhite
        SaveButn.Visible = False
        printButn.Visible = False
        labSignature.Visible = True
        txtDetails.BorderStyle = 0
        attendanceText.BorderStyle = 0
        cboClient.Top = labFor.Top '1440
        cboClient.Left = 2040
        attendanceText = getAttendanceText
        attendanceList.Visible = False
        txtDetails.backcolor = vbWhite
        custom_prompt.Visible = False
        
        If MsgBox("Detailed Receipt?", vbYesNo) = vbYes Then
            labAttendance.Visible = True
            attendanceText.Visible = True
        Else
            labAttendance.Visible = False
            attendanceText.Visible = False
        End If
        
        formPrint Me
        Printer.EndDoc
        
        SaveButn.Visible = True
        printButn.Visible = True
        labSignature.Visible = False
        txtDetails.BorderStyle = 1
        attendanceText.BorderStyle = 1
        labAttendance.Visible = True
        attendanceText.Visible = True

End Sub

Private Sub SaveButn_Click()
    Dim guid As String
    Dim sql As String
    Dim q As ADODB.Recordset
    Dim tempdate As Date
    
    'do checks and if something comes up prompt the user to see if they want to cancel
    'Set chk = db.Execute("SELECT * FROM attendance WHERE paid = 0 and date < " & sqlDate(dpFromdate.value))
    'If Not (chk.EOF And chk.BOF) Then
    '    If MsgBox("Some attendance records predating this receipt have not yet been paid!  Do you wish to save this receipt anyway?", vbYesNo) = vbNo Then Exit Sub
    'End If
    If ID = "" Then 'dont do this check if we are editing a saved receipt
        If doublepay Then
            If MsgBox("Some attendance listed on this receipt have already been paid!  Do you wish to save this receipt anyway?", vbYesNo) = vbNo Then Exit Sub
        End If
    End If
    For tempdate = dpFromDate.value To dpToDate.value
        If Weekday(tempdate) = 1 Or Weekday(tempdate) = 7 Then
            'do nothing, it's the weekend.... lol
        Else
            Set q = db.Execute("SELECT * FROM attendance WHERE idClient = " & getClientID(cboClient) & " AND date = " & sqlDate(tempdate))
            If q.EOF And q.BOF Then
                If MsgBox("You are entering a receipt with dates for which attendance has not been entered!  (" & shortDate(tempdate) & ")" & vbCrLf & vbCrLf & "Do you wish to save this receipt anyway?", vbYesNo) = vbNo Then Exit Sub
                Exit For
            End If
        End If
    Next tempdate
    
    attendanceText = getAttendanceText
    attendanceText.Visible = True
    attendanceList.Visible = False
    If Left$(txtAmount, 1) <> "$" Then txtAmount = Format(txtAmount, "$0.00")
    
    If ID = "" Then
        guid = createGUID
        labFinePrint = "#" & guid
        sql = "INSERT INTO payments ("
        sql = sql & "guid,idClient,receivedFrom,date,fromdate,todate,amount,attendance,details) VALUES ("
        sql = sql & """" & guid & ""","
        sql = sql & getClientID(cboClient) & ","
        sql = sql & """" & labReceivedFrom & ""","
        sql = sql & sqlDate(dpDate.value) & ","
        sql = sql & sqlDate(dpFromDate.value) & ","
        sql = sql & sqlDate(dpToDate.value) & ","
        sql = sql & MiD$(txtAmount, 2) & ","
        sql = sql & """" & attendanceText & ""","
        sql = sql & """" & txtDetails & """)"
        db.Execute sql
        
        db.Execute ("UPDATE attendance SET paid = 1 WHERE idClient = " & getClientID(cboClient) & " AND date >= " & sqlDate(dpFromDate) & " AND date <= " & sqlDate(dpToDate))
        
        create_gnc_receipt guid, Trim(MiD$(cboClient, InStr(1, cboClient, ")") + 1)) & " -- " & shortDate(dpFromDate) & " - " & shortDate(dpToDate), val(MiD$(txtAmount, 2)), dpDate.value
    Else
        'EDIT RECEIPT
        'set previous attendance records NOT paid.
        Set q = db.Execute("SELECT * FROM payments WHERE guid = """ & ID & """")
        
        db.Execute ("UPDATE attendance SET paid = 0 WHERE idClient = " & q!idClient & " AND date >= " & sqlDate(q!fromDate) & " AND date <= " & sqlDate(q!toDate))
        
        sql = "UPDATE payments SET "
        sql = sql & "idClient=" & getClientID(cboClient) & ","
        sql = sql & "receivedFrom=" & """" & labReceivedFrom & ""","
        sql = sql & "date=" & sqlDate(dpDate.value) & ","
        sql = sql & "fromdate=" & sqlDate(dpFromDate.value) & ","
        sql = sql & "todate=" & sqlDate(dpToDate.value) & ","
        sql = sql & "amount=" & MiD$(txtAmount, 2) & ","
        sql = sql & "attendance=" & """" & attendanceText & ""","
        sql = sql & "details=" & """" & txtDetails & """"
        sql = sql & " WHERE guid=" & """" & ID & """"
        db.Execute sql
        
        db.Execute ("UPDATE attendance SET paid = 1 WHERE idClient = " & getClientID(cboClient) & " AND date >= " & sqlDate(dpFromDate) & " AND date <= " & sqlDate(dpToDate))
        
        update_gnc_receipt ID, Trim(MiD$(cboClient, InStr(1, cboClient, ")") + 1)) & " -- " & shortDate(dpFromDate) & " - " & shortDate(dpToDate), val(MiD$(txtAmount, 2)), dpDate.value
        
    End If
        
    If MsgBox("Do you want to print this receipt?", vbYesNo) = vbYes Then printButn_Click
    
    Set q = Nothing
    Unload Me
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    'Dim D As Date
    Dim fc As ADODB.Recordset
    Dim q As ADODB.Recordset
    Dim li As ListItem
    Dim i As Long
    
    
    Set fc = db.Execute("SELECT * FROM fee_classes")
    With fc
        If Not (.BOF And .EOF) Then
            .MoveFirst
            Do Until .EOF
                cboFeeClass.AddItem !Description & " - $" & !charge, !idFeeClasses - 1
                .MoveNext
            Loop
        End If
    End With
    'Set q = db.Execute("SELECT * FROM clients WHERE active=1 ORDER BY last, first") ' SOMETIMES WE NEED TO ENTER A RECEIPT FOR A CLIENT WHOS NOT ACTIVE
    Set q = db.Execute("SELECT * FROM clients ORDER BY last, first")
    With q
        If Not (.BOF And .EOF) Then
            .MoveFirst
            Do Until .EOF
                cboClient.AddItem "(" & Trim(str(!idClient)) & ") " & !Last & ", " & !First
                .MoveNext
            Loop
        End If
    End With
    
    If ID = "" Then 'NEW ENTRY!
        dpDate = Date
        guessDates
    Else            'EDIT ENTRY!
        'saveButn.Caption = "Close"
        'txtAmount.Locked = True
        'txtWeeks.Locked = True
        'txtDetails.Locked = True
        'cboClient.Enabled = False
        'cboFeeClass.Enabled = False
        'dpFromdate.Enabled = False
        'dpTodate.Enabled = False
        'dpDate.Enabled = False
        'attendanceText.Locked = True
        attendanceText.Visible = True
        attendanceList.Visible = False
        
        Set q = db.Execute("SELECT * FROM payments WHERE guid = """ & ID & """")
        labReceivedFrom = "" & q!receivedFrom
        labFinePrint = "#" & ID

        updateButn.Visible = True
        For i = 0 To cboClient.ListCount - 1
            If getClientID(cboClient.List(i)) = q!idClient Then
                cboClient.ListIndex = i
            End If
        Next i
        dpToDate.value = q!toDate
        dpFromDate.value = q!fromDate
        dpDate = q!Date
        
        calculate_weeks

        attendanceText = "" & q!attendance
        txtDetails = "" & q!details
        
        DoEvents
        txtAmount = Format(q!amount, "$0.00")
        
    End If
    
    Set li = Nothing
    Set fc = Nothing
    Set q = Nothing

    If frmReceipts.chkFilterClient.value = 1 Then
        cboClient.ListIndex = frmReceipts.cboClient.ListIndex
    End If

End Sub

Private Sub txtAmount_Change()
    If val(MiD$(txtAmount, 2)) > 0 Then
        SaveButn.Enabled = True
    Else
        SaveButn.Enabled = False
    End If
    If Me.ActiveControl = txtAmount Then
        If txtDetails = "" Then
            txtDetails.backcolor = vbYellow
            custom_prompt.Visible = True
        End If
    End If
End Sub

Private Sub txtDetails_Change()
    If txtDetails <> "" Then
        txtDetails.backcolor = vbWhite
    End If

End Sub

Private Sub txtDetails_KeyPress(KeyAscii As Integer)
    desc_manual_edit = True
End Sub

Private Sub txtDetails_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    custom_prompt.Visible = False
End Sub

Private Sub txtWeeks_Change()
    calculate_amount
End Sub

Private Sub updateButn_Click()
    calculate_attendance
End Sub
