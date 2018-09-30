VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Funtime Child Care Center Manager"
   ClientHeight    =   5310
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   11640
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   4  'Icon
   ScaleHeight     =   5310
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageForms 
      Left            =   120
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   1936
      ImageHeight     =   149
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11B5B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   36000
      Left            =   840
      Top             =   120
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   120
      Top             =   120
      _ExtentX        =   979
      _ExtentY        =   979
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E924A
            Key             =   "check"
         EndProperty
      EndProperty
   End
   Begin VB.Label chequesButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cheques"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   7320
      TabIndex        =   21
      Top             =   2160
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label qifButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Import QIF"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   7320
      TabIndex        =   20
      Top             =   1200
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label schoolcloseButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "School Closures"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8800&
      Height          =   375
      Left            =   4680
      TabIndex        =   19
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label expensesButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Expenses By Category"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   7320
      TabIndex        =   18
      Top             =   720
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label thisweekButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Who's Paying This Week"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   7320
      TabIndex        =   17
      Top             =   3120
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label graphsButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Chart: Income / Expenses"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   7320
      TabIndex        =   16
      Top             =   1680
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label ledgerButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Client Ledger"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8800&
      Height          =   375
      Left            =   4680
      TabIndex        =   15
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label birthdaysButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Birthday Calendar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8800&
      Height          =   375
      Left            =   4680
      TabIndex        =   14
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label budgetButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Budget  (Incomplete)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   7320
      TabIndex        =   13
      Top             =   2640
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label LabExp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Expenses"
      Height          =   252
      Left            =   4440
      TabIndex        =   12
      Top             =   4560
      Width           =   852
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   4140
      X2              =   4200
      Y1              =   4636
      Y2              =   4596
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   4116
      X2              =   4146
      Y1              =   4682
      Y2              =   4632
   End
   Begin VB.Shape ShapeExp 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   252
      Left            =   4080
      Shape           =   3  'Circle
      Top             =   4560
      Width           =   252
   End
   Begin VB.Label LabRec 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Receipts"
      Height          =   252
      Left            =   3120
      TabIndex        =   11
      Top             =   4560
      Width           =   852
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   2820
      X2              =   2880
      Y1              =   4636
      Y2              =   4596
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   2796
      X2              =   2826
      Y1              =   4682
      Y2              =   4632
   End
   Begin VB.Shape ShapeRec 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   252
      Left            =   2760
      Shape           =   3  'Circle
      Top             =   4560
      Width           =   252
   End
   Begin VB.Label LabAtt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Attendance"
      Height          =   252
      Left            =   1680
      TabIndex        =   10
      Top             =   4560
      Width           =   852
   End
   Begin VB.Line Shine2 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   1380
      X2              =   1440
      Y1              =   4630
      Y2              =   4590
   End
   Begin VB.Line Shine1 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   1350
      X2              =   1380
      Y1              =   4680
      Y2              =   4630
   End
   Begin VB.Shape ShapeAtt 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   252
      Left            =   1320
      Shape           =   3  'Circle
      Top             =   4560
      Width           =   252
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Data Health:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   9
      Top             =   4560
      Width           =   1092
   End
   Begin VB.Label adminButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Admin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8800&
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label ogpButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OGP Forms"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   7320
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label receiptButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Payments"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8800&
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label subButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Subsidy Forms"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8800&
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label attendButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Attendance"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8800&
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label clientButn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clients"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8800&
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label labVersion 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ver"
      Height          =   255
      Left            =   9000
      TabIndex        =   2
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Label quote 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   16.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   4575
      Left            =   7200
      TabIndex        =   1
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label byline 
      BackColor       =   &H00FFFFFF&
      Caption         =   "This software written and maintained by Mark Mercer.  Copyright 2016"
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   5040
      Width           =   8052
   End
   Begin VB.Image Image1 
      Height          =   4440
      Left            =   60
      Picture         =   "frmMain.frx":EC384
      Stretch         =   -1  'True
      Top             =   60
      Width           =   4440
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *******************************
' ****    Funtime Manager    ****
' ****    Version History    ****
' *******************************

' Earlier versions, changelog not tracked

' v2.1.8 - Feb 8, 2017
' -Added fourth room
' -Rooms now read from the db and autogenerate in the attendance list
' -Attendance list now generates with each date clicked on calendar and only
'  adds those who were registered on that day and puts them in the right rm.

' v2.2.0 - Jan 28, 2018
' -Subsidy auto 4 wk pp
' -Popup when selecting fee class that doesn't match age
' -Ability to enter and print cheques added
' -Afterschoolers not showing as $70.  Subsidy not working?  Wrong Date?
' -Rename rooms to front, small and back because their designations change
' -Custom fees, specific week days, in feeclasschanges
' --Don 't show in attendance window on days not scheduled
' --show / on subsidy form if not scheduled.
' -Runtime error middle initial too many characters
' -Fee class update popup doesn't change amount!

' v2.2.1 - Apr 17, 2018
' -Added sick as option in attendance.
'  ALTER TABLE `funtime`.`attendance` ADD COLUMN `sick` TINYINT(1) NULL DEFAULT 0 AFTER `attended`;

' v2.3.0 - Jun 4, 2018
' -Added contacts and client_contacts table.  Now instead of the clients table
'   having the phone number, parent1, parent2, emergency fields, I've created
'   the above tables so that any number of contacts can be added to a single
'   client and they can be designated as parent/guardian, emergency, doctor, or
'   other.
' -Added MCP and Allergies to the clients table
' -Added the ability to print a list of clients formatted for sticky labels
' -Added a reactivate button on client form
' -Added tracking for start and end date in client_changes
' -On client create, prompt if duplicate client
' -Fixed subsidy calculation: Sick + SC = $30

' v2.3.1 - Jun 18, 2018
' -Fixed 'schools out' runtime error.  getStartDateAtDate and getEndDateAtDate were
'   failing because they were trying to initialize a date with empty string.
' -Formatted phone numbers throughout application
' -Land line was showing as duplicate of cell phone
' -Fixed a bug that allowed the selected index variable for contacts being added
'   to clients change after it is supposed to be set, which was causing weird
'   overwriting of database entries.

' v2.3.2 - Jul 15, 2018
' -Fixed some bugs on attendance entry form.  "sick" showing on room title line
'   and made the DTPickers show the UpDown arrows rather than calendar dropdown

' v2.3.3 - Sep 30, 2018
' -minor tweaks, to ensure the subsidy save button gets enabled

' *********************************************************************

Private timeout_counter As Long
Private listButtons(30) As Control

Private Sub Form_GotFocus()

    'THIS IS THE GOT FOCUS METHOD
    updateDataHealth

End Sub

Sub admin_enable()
    ogpButn.Visible = True
    expensesButn.Visible = True
    budgetButn.Visible = True
    graphsButn.Visible = True
    thisweekButn.Visible = True
    qifButn.Visible = True
    chequesButn.Visible = True
    quote.Visible = False
End Sub


Private Sub adminButn_Click()
    If Not RIDE Then
        dlgPassword.Show 1
    Else
        authenticated = True
    End If
    
    If authenticated Then admin_enable
End Sub

Private Sub adminButn_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    mouseoutButtons
    adminButn.backcolor = &HFFCC99
End Sub

Private Sub attendButn_Click()
    frmAttendance.Show 1
End Sub

Private Sub attendButn_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    mouseoutButtons
    attendButn.backcolor = &HFFCC99
End Sub

Private Sub birthdaysButn_Click()
    frmCalendar.Show 1
End Sub

Private Sub birthdaysButn_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    mouseoutButtons
    birthdaysButn.backcolor = &HFFCC99
End Sub


Private Sub budgetButn_Click()
    frmBudget.Show 1
End Sub

Private Sub budgetButn_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    mouseoutButtons
    budgetButn.backcolor = &H80FF&
End Sub


Private Sub chequesButn_Click()
    frmCheques.Show 1
End Sub

Private Sub chequesButn_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    mouseoutButtons
    chequesButn.backcolor = &H80FF&
End Sub

Private Sub clientButn_Click()
    frmClients.Show 1
End Sub


Private Sub clientButn_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    mouseoutButtons
    clientButn.backcolor = &HFFCC99
End Sub

Private Sub expensesButn_Click()
    frmExpenses.Show 1
End Sub

Private Sub expensesButn_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    mouseoutButtons
    expensesButn.backcolor = &H80FF&
End Sub

Private Sub Form_Load()
    Dim c As Control
    Dim i As Long
    Dim q As ADODB.Recordset
    Dim d As Date
    
    Randomize Timer
    authenticated = False
    RIDE = (App.EXEName = "FuntimeManagerProj")
    
    EPOCH = CDate("June 6, 2016")
    FDOS = LabourDay(year(Date)) + 2 'First Day of School
    LDOS = DateSerial(year(Date), 6, 22) 'Last Day of School
    
    If year(Date) > 2016 Then byline = byline & " - " & year(Date)
    labVersion = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    i = 1
    For Each c In Me.Controls
        If TypeOf c Is Label Then
            Set listButtons(i) = c
            i = i + 1
        End If
    Next c
    
    ConnectDB
    'MsgBox "check 0"
    displayQuote
    
    Set q = gcdb.Execute("SELECT * FROM accounts WHERE name = ""Imbalance-CAD""")
    If Not (q.EOF And q.BOF) Then
        q.MoveFirst
        rec_account_guid = q!guid
    End If
    Set q = gcdb.Execute("SELECT * FROM commodities WHERE fullname = ""Canadian Dollar""")
    If Not (q.EOF And q.BOF) Then
        q.MoveFirst
        currency_guid = q!guid
    End If
    Set q = gcdb.Execute("SELECT * FROM accounts WHERE name = ""Parental Fees""")
    If Not (q.EOF And q.BOF) Then
        q.MoveFirst
        payments_account_guid = q!guid
    End If
    Set q = gcdb.Execute("SELECT * FROM accounts WHERE name = ""Childcare Subsidy""")
    If Not (q.EOF And q.BOF) Then
        q.MoveFirst
        subsidy_payments_account_guid = q!guid
    End If
    
    Set q = gcdb.Execute("SELECT * FROM accounts WHERE name = ""Payroll Expenses""")
    If Not (q.EOF And q.BOF) Then
        q.MoveFirst
        payroll_account_guid = q!guid
    End If

    If Not RIDE Then
        check_for_client_changes
        For d = Date - 30 To Date - 7
            If Weekday(d) <> 1 And Weekday(d) <> 7 And Not isStatHoliday(d) Then
                Set q = db.Execute("SELECT * FROM attendance WHERE date = " & sqlDate(d))
                If q.EOF And q.BOF Then
                    MsgBox "Attendance entry missing for " & shortDate(d)
                End If
            End If
        Next d
    End If
    
    Set q = Nothing
    Set c = Nothing
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    mouseoutButtons
End Sub

Private Sub Form_Terminate()
    db.Close
End Sub


Private Sub graphsButn_Click()
    frmChart.Show 1
End Sub

Private Sub graphsButn_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    mouseoutButtons
    graphsButn.backcolor = &H80FF&
End Sub


Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    mouseoutButtons
End Sub

Private Sub ledgerButn_Click()
    frmLedger.Show 1
End Sub

Private Sub ledgerButn_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    mouseoutButtons
    ledgerButn.backcolor = &HFFCC99
End Sub


Private Sub ogpButn_Click()
    frmOGP.Show 1
End Sub

Private Sub ogpButn_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    mouseoutButtons
    ogpButn.backcolor = &H80FF&
End Sub

Private Sub qifButn_Click()
    frmQIF.Show 1
End Sub

Private Sub qifButn_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    mouseoutButtons
    qifButn.backcolor = &H80FF&
End Sub

Private Sub quote_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    mouseoutButtons
End Sub

Private Sub receiptButn_Click()
    frmReceipts.Show 1
End Sub


Private Sub receiptButn_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    mouseoutButtons
    receiptButn.backcolor = &HFFCC99
End Sub

Private Sub schoolcloseButn_Click()
    frmClosures.Show 1
End Sub

Private Sub schoolcloseButn_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    mouseoutButtons
    schoolcloseButn.backcolor = &HFFCC99
End Sub

Private Sub subButn_Click()
    frmSubsidization.Show 1
End Sub

Private Sub subButn_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    mouseoutButtons
    subButn.backcolor = &HFFCC99
End Sub

Private Sub thisweekButn_Click()
    dlgPayingThisWeek.Show 1
End Sub

Private Sub thisweekButn_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    mouseoutButtons
    thisweekButn.backcolor = &H6699FF
End Sub


Private Sub Timer1_Timer()
    timeout_counter = timeout_counter + 1
    If timeout_counter >= 100 Then
        ConnectDB
        tiemout_counter = 0
    End If
    

End Sub

Sub displayQuote()
    Dim quotes(11) As String
    
    quotes(0) = "Funtime Child Care Center: Satisfaction Guaranteed or Double Your Kids Back!"
    quotes(1) = "Children need the freedom and time to play. Play is not a luxury. Play is a necessity. ~Kay Redfield Jamison"
    quotes(2) = "A child seldom needs a good talking-to as much as a good listening-to. ~Robert Brault"
    quotes(3) = "While we try to teach our children all about life, Our children teach us what life is all about. ~Angela Schwindt"
    quotes(4) = "The prime purpose of being four is to enjoy being four — of secondary importance is to prepare for being five. ~Jim Trelease"
    quotes(5) = "Children make you want to start life over. ~Muhammad Ali"
    quotes(6) = "Even when freshly washed and relieved of all obvious confections, children tend to be sticky. ~Fran Lebowitz" & vbCrLf & "I hate sticky ~Mark Mercer"
    quotes(7) = "I brought children into this dark world because it needed the light that only a child can bring. ~Liz Armbruster"
    quotes(8) = "A child can ask questions that a wise man cannot answer."
    quotes(9) = "Children are like sponges.  They absorb all your strength and leave you limp, but give them a squeeze and you get everything back"
    quotes(10) = "A person 's a person, no matter how small. ~Dr. Seuss"
    quotes(11) = "Children are great imitators.  So give them something great to imitate."
        
    quote.Caption = quotes(CLng(Rnd * 11))
End Sub

Sub mouseoutButtons()
    Dim c As Long
    For c = 0 To 30
        'MsgBox c & "  " & (listButtons(c) Is Nothing)
        If Not listButtons(c) Is Nothing Then
            listButtons(c).backcolor = vbWhite
        End If
    Next c
End Sub

Sub updateDataHealth()
    
    Dim q As ADODB.Recordset
    Dim tempdate As Date
    Dim d As Long
    Dim rec_account_guid  As String
    Dim days(5) As Long
    Dim avgdays As Double
    Dim index As Long
    
    Set q = db.Execute("SELECT * FROM Attendance WHERE date >= " & sqlDate(Date - 6) & " ORDER BY date ASC")
    tempdate = Date - 7
    With q
        If Not (.EOF And .BOF) Then .MoveFirst
        Do Until .EOF
            If !Date > tempdate + 1 And Weekday(!Date) <> 1 And Weekday(!Date) <> 7 Then  ' tempdate + 1 is missing
                tempdate = !Date
                Exit Do
            End If
            If !Date = tempdate + 1 Then tempdate = !Date
            .MoveNext
        Loop
    End With

    If tempdate = Date Then
        ShapeAtt.backcolor = vbGreen
        LabAtt.ToolTipText = "Attendance is up to date!"
    ElseIf tempdate < Date Then
        d = DateDiff("d", tempdate, Date)
        If d = 1 Then ShapeAtt.backcolor = &HFFFF&
        If d = 2 Then ShapeAtt.backcolor = &HCCFF&
        If d = 3 Then ShapeAtt.backcolor = &H99FF&
        If d = 4 Then ShapeAtt.backcolor = &H66FF&
        If d = 5 Then ShapeAtt.backcolor = &H33FF&
        If d >= 6 Then ShapeAtt.backcolor = &HFF&
        LabAtt.ToolTipText = "Attendance records from " & shortDate(tempdate) & " are not present"
    End If
    
    
    
    Set q = gcdb.Execute("SELECT * FROM accounts WHERE name = ""Imbalance-CAD""")
    If Not (q.EOF And q.BOF) Then
        q.MoveFirst
        rec_account_guid = q!guid
    End If
    
    'Most Recent Income
    Set q = gcdb.Execute("SELECT transactions.post_date as Date FROM splits INNER JOIN transactions ON splits.tx_guid = transactions.guid WHERE account_guid != '" & rec_account_guid & "' AND value_num < 0 ORDER BY post_date DESC LIMIT 5")
    index = 1
    With q
        If Not (.EOF And .BOF) Then .MoveFirst
        Do Until .EOF
            If index > 5 Then Exit Do
            days(index) = DateDiff("d", !Date, Date)
            'MsgBox !Date & "  " & days(index)
            index = index + 1
            .MoveNext
        Loop
    End With
    avgdays = (days(1) + days(2) + days(3) + days(4) + days(5)) / 5#
    If avgdays >= 21 Then ShapeRec.backcolor = &HFF&: LabRec.ToolTipText = "Receipts are EXTREMELY behind!"
    If avgdays < 21 Then ShapeRec.backcolor = &H33FF&: LabRec.ToolTipText = "Receipts are very behind!"
    If avgdays < 17 Then ShapeRec.backcolor = &H66FF&: LabRec.ToolTipText = "Receipts are behind!"
    If avgdays < 14 Then ShapeRec.backcolor = &H99FF&: LabRec.ToolTipText = "Receipts are behind"
    If avgdays < 10 Then ShapeRec.backcolor = &HCCFF&: LabRec.ToolTipText = "Receipts are behind"
    If avgdays < 7 Then ShapeRec.backcolor = &HFFFF&: LabRec.ToolTipText = "Receipts need to be brought up to date"
    If avgdays < 3 Then ShapeRec.backcolor = &HFF00&: LabRec.ToolTipText = "Receipts are up to date!"
    
    
    'Most Recent Expenses
    'Set q = gcdb.Execute("SELECT transactions.post_date as Date FROM splits INNER JOIN transactions ON splits.tx_guid = transactions.guid WHERE account_guid != '" & rec_account_guid & "' AND value_num > 0 ORDER BY post_date DESC LIMIT 5") THIS INCLUDED PAYROLL WHICH IS UPDATED WEEKLY.
    Set q = gcdb.Execute("SELECT transactions.post_date as Date FROM splits INNER JOIN transactions ON splits.tx_guid = transactions.guid WHERE account_guid != '" & rec_account_guid & "' AND account_guid != '" & payroll_account_guid & "' AND value_num > 0 ORDER BY post_date DESC LIMIT 5")
    index = 1
    With q
        If Not (.EOF And .BOF) Then .MoveFirst
        Do Until .EOF
            If index > 5 Then Exit Do
            days(index) = DateDiff("d", !Date, Date)
            'MsgBox !Date & "  " & days(index)
            index = index + 1
            .MoveNext
        Loop
    End With
    avgdays = (days(1) + days(2) + days(3) + days(4) + days(5)) / 5#
    If avgdays >= 45 Then ShapeExp.backcolor = &HFF&: LabExp.ToolTipText = "Expenses are EXTREMELY behind!"
    If avgdays < 45 Then ShapeExp.backcolor = &H33FF&: LabExp.ToolTipText = "Expenses are very behind!"
    If avgdays < 38 Then ShapeExp.backcolor = &H66FF&: LabExp.ToolTipText = "Expenses are behind!"
    If avgdays < 30 Then ShapeExp.backcolor = &H99FF&: LabExp.ToolTipText = "Expenses are behind"
    If avgdays < 22 Then ShapeExp.backcolor = &HCCFF&: LabExp.ToolTipText = "Expenses are behind"
    If avgdays < 15 Then ShapeExp.backcolor = &HFFFF&: LabExp.ToolTipText = "Expenses need to be brought up to date"
    If avgdays < 7 Then ShapeExp.backcolor = &HFF00&: LabExp.ToolTipText = "Expenses are up to date!"
    
    Set q = Nothing
End Sub
