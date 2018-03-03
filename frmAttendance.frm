VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAttendance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attendance"
   ClientHeight    =   8610
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   8160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton printButn 
      Caption         =   "Print"
      Height          =   372
      Left            =   7560
      TabIndex        =   21
      Top             =   240
      Width           =   612
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   6240
      Top             =   240
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   312
      Index           =   3
      Left            =   1560
      TabIndex        =   11
      Top             =   1560
      Visible         =   0   'False
      Width           =   6252
      Begin VB.ComboBox cboYear 
         Height          =   288
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   0
         Width           =   1692
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   312
      Index           =   2
      Left            =   1560
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   6252
      Begin VB.ComboBox cboYearMo 
         Height          =   288
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   0
         Width           =   1212
      End
      Begin VB.ComboBox cboMonth 
         Height          =   288
         ItemData        =   "frmAttendance.frx":0000
         Left            =   0
         List            =   "frmAttendance.frx":0028
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   0
         Width           =   2292
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   312
      Index           =   1
      Left            =   1560
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   6252
      Begin VB.CheckBox chkUnpaid 
         Caption         =   "Unpaid"
         Height          =   252
         Left            =   4800
         TabIndex        =   22
         Top             =   0
         Width           =   852
      End
      Begin MSComCtl2.DTPicker dpFrom 
         Height          =   280
         Left            =   600
         TabIndex        =   14
         Top             =   0
         Width           =   1812
         _ExtentX        =   3201
         _ExtentY        =   503
         _Version        =   393216
         CalendarTitleBackColor=   16764006
         CustomFormat    =   "MMM d, yyyy"
         Format          =   75169795
         CurrentDate     =   42536
      End
      Begin MSComCtl2.DTPicker dpTo 
         Height          =   280
         Left            =   2880
         TabIndex        =   16
         Top             =   0
         Width           =   1812
         _ExtentX        =   3201
         _ExtentY        =   503
         _Version        =   393216
         CalendarTitleBackColor=   16764006
         CustomFormat    =   "MMM d, yyyy"
         Format          =   75169795
         CurrentDate     =   42536
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   252
         Left            =   2520
         TabIndex        =   17
         Top             =   0
         Width           =   492
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "From"
         Height          =   252
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   612
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   312
      Index           =   0
      Left            =   1800
      TabIndex        =   7
      Top             =   240
      Width           =   6252
      Begin VB.CheckBox chkUnpaidCl 
         Caption         =   "Unpaid"
         Height          =   252
         Left            =   2520
         TabIndex        =   23
         Top             =   0
         Width           =   852
      End
      Begin VB.ComboBox cboClient 
         Height          =   288
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   0
         Width           =   2292
      End
   End
   Begin VB.OptionButton viewOptions 
      Caption         =   "Yearly"
      Height          =   252
      Index           =   5
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   0
      Width           =   1092
   End
   Begin VB.OptionButton viewOptions 
      Caption         =   "Monthly"
      Height          =   252
      Index           =   4
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   0
      Width           =   1092
   End
   Begin VB.CommandButton enterAttendButn 
      Caption         =   "Enter Attendance"
      Height          =   492
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1812
   End
   Begin VB.TextBox Text1 
      Height          =   2772
      Left            =   1920
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "frmAttendance.frx":008E
      Top             =   3240
      Visible         =   0   'False
      Width           =   4092
   End
   Begin MSComctlLib.ListView ListView 
      Height          =   7932
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   8052
      _ExtentX        =   14208
      _ExtentY        =   13996
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Client"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Attended"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Sign In"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Sign Out"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Paid"
         Object.Width           =   1058
      EndProperty
   End
   Begin VB.OptionButton viewOptions 
      Caption         =   "Daily"
      Height          =   252
      Index           =   3
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   1092
   End
   Begin VB.OptionButton viewOptions 
      Caption         =   "By Fee Class"
      Height          =   252
      Index           =   2
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1092
   End
   Begin VB.OptionButton viewOptions 
      Caption         =   "By Date"
      Height          =   252
      Index           =   1
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   972
   End
   Begin VB.OptionButton viewOptions 
      Caption         =   "By Client"
      Height          =   252
      Index           =   0
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Value           =   -1  'True
      Width           =   1212
   End
End
Attribute VB_Name = "frmAttendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private col_width(8) As Long
Private col_label(8) As String
Private selected_view As Byte

Private Sub cboClient_Change()
    cboClient_Click
End Sub

Private Sub cboClient_Click()
    ListView.ListItems.Clear
    fillByClient
End Sub


Private Sub chkUnpaid_Click()
    dpTo_Click
End Sub

Private Sub chkUnpaidCl_Click()
    ListView.ListItems.Clear
    fillByClient
End Sub


Private Sub dpFrom_Change()
    dpFrom_Click
End Sub

Private Sub dpFrom_Click()
    ListView.ListItems.Clear
    If viewOptions(1) Then
        fillByDate
    Else
        fillByFeeClass
    End If
End Sub

Private Sub dpTo_Change()
    dpTo_Click
End Sub

Private Sub dpTo_Click()
    ListView.ListItems.Clear
    If viewOptions(1) Then
        fillByDate
    Else
        fillByFeeClass
    End If
End Sub

Private Sub enterAttendButn_Click()
    frmAttendanceEntry.Show 1
    viewOptions_Click CLng(selected_view)
End Sub


Private Sub Form_Load()
    Dim y As Long
    Dim cl As ADODB.Recordset
    Dim ddiff As Long
    For y = 2016 To Year(Date)
        cboYear.AddItem y
        cboYearMo.AddItem y
    Next y
    dpTo = Date
    dpFrom = Date '- 7
    ddiff = Weekday(dpFrom.value) - 2
    dpFrom = dpFrom.value - ddiff
    
    For y = 1 To 3 'reusing y variable just as an int
        Frame(y).Top = Frame(0).Top
        Frame(y).Left = Frame(0).Left
    Next y
    
    For y = 1 To 7 'reusing y variable just as an int again
        col_width(y) = ListView.ColumnHeaders(y).width
        col_label(y) = ListView.ColumnHeaders(y).Text
    Next y

    
    Set cl = db.Execute("SELECT * FROM clients ORDER BY last, first ASC")
    With cl
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do Until .EOF
                cboClient.AddItem "(" & Trim(str(!idclient)) & ") " & !Last & ", " & !First
                .MoveNext
            Loop
        End If
    End With
    Set cl = Nothing
End Sub

Private Sub ListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If ListView.SortKey = ColumnHeader.index - 1 Then
        If ListView.SortOrder = lvwAscending Then
            ListView.SortOrder = lvwDescending
        Else
            ListView.SortOrder = lvwAscending
        End If
    Else
        ListView.SortKey = ColumnHeader.index - 1
    End If
    ListView.Sorted = True
End Sub

Private Sub printButn_Click()
    'formPrint Me, 50, 50
    printText "Attendance", 50, 500, 15000, "Arial", 16, True, 0
    If selected_view < 2 Then
        printListView ListView, 60, 50, 1500, 1.65, True
    Else
        'add other options here
        
        printListView ListView, 60, 50, 1500, 1.5, True ' fee class
    End If
    Printer.EndDoc
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    
    viewOptions(1).value = True
    viewOptions_Click 1
    
End Sub

Private Sub viewOptions_Click(index As Integer)
    Dim f As Long
    Dim i As Long
    
    ListView.ListItems.Clear
    selected_view = index
    
    For f = 0 To 3
        Frame(f).Visible = False
    Next f
    i = index
    If i >= 2 Then i = i - 1 ' this is the mechanism that allows the use of frame(1) for button (1) and (2)
    If i <= 3 Then Frame(i).Visible = True ' this is the mechanism that makes sure you dont refer to a frame that doesn't exist
    
    If index = 0 And cboClient.ListIndex >= 0 Then 'client
        fillByClient
    ElseIf index = 1 And dpFrom <= dpTo Then 'date
        chkUnpaid.Visible = True
        fillByDate
    ElseIf index = 2 And dpFrom <= dpTo Then 'fee class
        chkUnpaid.Visible = False
        fillByFeeClass
    ElseIf index = 3 And cboMonth.ListIndex > -1 And cboYear.ListIndex > -1 Then   'day
        fillByDays
    ElseIf index = 4 And cboYear.ListIndex > -1 Then   'month
        'fillByMonth
    ElseIf index = 5 Then    'year
        'fillByYear
    End If
    
    If index = 1 And dpFrom < dpTo Then 'date
        
    End If
    
    If index = 2 And cboMonth.ListIndex >= 0 And cboYearMo.ListIndex >= 0 Then 'month
        
    End If
    
    If index = 3 And cboYear.ListIndex >= 0 Then 'year
        
    End If
    
End Sub

Function getClientID(ByVal s As String) As Long
    If Not s = "" Then getClientID = val(MiD$(Left$(s, InStr(1, s, ")") - 1), 2))
End Function


Sub fillByClient()
    Dim clientID As Long
    Dim att As ADODB.Recordset
    Dim li As ListItem
    Dim y As Long ' way to go! now I'm using y here too!!
    Dim unpaid As String
    
    For y = 1 To 7
        ListView.ColumnHeaders(y).width = col_width(y)
        ListView.ColumnHeaders(y).Text = col_label(y)
    Next y
    
    clientID = getClientID(cboClient.Text)
    
    If chkUnpaidCl Then
        unpaid = " AND paid=0"
    Else
        unpaid = ""
    End If
    
    
    Set att = db.Execute("SELECT * FROM attendance WHERE idClient = " & clientID & unpaid & " ORDER BY date DESC, idClient ASC")
    With att
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do Until .EOF
                Set li = ListView.ListItems.Add(, , !idAttendance)
                li.SubItems(1) = Trim(MiD$(cboClient.Text, InStr(cboClient.Text, ")") + 1)) '!Last & ", " & !First
                li.SubItems(2) = ansiDate(!Date)
                If !attended = 1 Then
                    li.SubItems(3) = "Attended"
                    li.SubItems(4) = Format(!signin, "hh:nn")
                    li.SubItems(5) = Format(!signout, "hh:nn")
                Else
                    li.SubItems(3) = "Did Not Attend"
                End If
                If !paid = 1 Then li.SubItems(6) = "Pd"
                
                .MoveNext
            Loop
        End If
    End With
    
    Set li = Nothing
    Set att = Nothing
End Sub


Sub fillByDate()
    Dim att As ADODB.Recordset
    Dim li As ListItem
    Dim y As Long ' way to go! now I'm using y here too!!
    Dim unpaid As String
    
    For y = 1 To 7
        ListView.ColumnHeaders(y).width = col_width(y)
        ListView.ColumnHeaders(y).Text = col_label(y)
    Next y
    
    If chkUnpaid Then
        unpaid = " AND paid=0"
    Else
        unpaid = ""
    End If
    
    Set att = db.Execute("SELECT attendance.*, clients.first, clients.last FROM attendance INNER JOIN clients ON (attendance.idClient = clients.idClient) WHERE Date >= " & sqlDate(dpFrom.value) & " AND Date <= " & sqlDate(dpTo.value) & unpaid & " ORDER BY last ASC, date DESC, idClient ASC")
    With att
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do Until .EOF
                Set li = ListView.ListItems.Add(, , !idAttendance)
                li.SubItems(1) = !Last & ", " & !First
                li.SubItems(2) = ansiDate(!Date)
                If !attended = 1 Then
                    li.SubItems(3) = "Attended"
                    li.SubItems(4) = Format(!signin, "hh:nn")
                    li.SubItems(5) = Format(!signout, "hh:nn")
                Else
                    li.SubItems(3) = "Did Not Attend"
                End If
                If !paid = 1 Then li.SubItems(6) = "Pd"
                
                .MoveNext
            Loop
        End If
    End With
    
    Set li = Nothing
    Set att = Nothing
End Sub

Sub fillByFeeClass()
    Dim att As ADODB.Recordset
    Dim li As ListItem
    ListView.ColumnHeaders(2).width = 3500
    ListView.ColumnHeaders(2).Text = "Fee Class"
    ListView.ColumnHeaders(3).width = 1000
    ListView.ColumnHeaders(3).Text = "Attended"
    ListView.ColumnHeaders(4).width = 1000
    ListView.ColumnHeaders(4).Text = "Atten %"
    ListView.ColumnHeaders(5).width = 1000
    ListView.ColumnHeaders(5).Text = "Paid"
    ListView.ColumnHeaders(6).width = 1000
    ListView.ColumnHeaders(6).Text = "Paid %"
    ListView.ColumnHeaders(7).width = 0
    ListView.ColumnHeaders(7).Text = ""
    
    'Set att = db.Execute("SELECT DISTINCT feeclass FROM attendance WHERE idClient = " & clientID & " ORDER BY date DESC")
    Set att = db.Execute("SELECT clients.feeClassID, fee_classes.description, SUM(attendance.attended) as Pres, SUM(1 - attendance.attended) as Abs, SUM(attendance.paid) as Pd, SUM(1 - attendance.paid) as Owing FROM attendance INNER JOIN clients ON (attendance.idClient = clients.idClient) INNER JOIN fee_classes ON (clients.feeClassID = fee_classes.idFeeClasses) WHERE Date >= " & sqlDate(dpFrom.value) & " AND Date <= " & sqlDate(dpTo.value) & " GROUP BY clients.feeClassID")
    
    With att
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do Until .EOF
                'MsgBox !feeClassID
                Set li = ListView.ListItems.Add(, , !feeClassID)
                li.SubItems(1) = !Description
                li.SubItems(2) = !Pres
                li.SubItems(3) = Format(!Pres / CDbl(!Pres + !Abs), "0%")
                li.SubItems(4) = !Pd
                li.SubItems(5) = Format(!Pd / (!Pd + !owing), "0%")
                
                .MoveNext
            Loop
        End If
    End With
    
    Set li = Nothing
    Set att = Nothing
End Sub


Sub fillByDays()
    Dim att As ADODB.Recordset
    Dim li As ListItem
'days, total stats per day, %capacity, %IC capacity, %PS capacity, %AS capacity
'daily - rows = days, total stats per day, %capacity, %IC capacity, %PS capacity, %AS capacity

'monthly - rows = months, total stats per day, %capacity, %IC capacity, %PS capacity, %AS capacity, final row is full Year

'yearly is same as the last two but just one line for each year - no controls needed

    ListView.ColumnHeaders(2).width = 3500
    ListView.ColumnHeaders(2).Text = "Fee Class"
    ListView.ColumnHeaders(3).width = 1000
    ListView.ColumnHeaders(3).Text = "Attended"
    ListView.ColumnHeaders(4).width = 1000
    ListView.ColumnHeaders(4).Text = "Atten %"
    ListView.ColumnHeaders(5).width = 1000
    ListView.ColumnHeaders(5).Text = "Paid"
    ListView.ColumnHeaders(6).width = 1000
    ListView.ColumnHeaders(6).Text = "Paid %"
    ListView.ColumnHeaders(7).width = 0
    ListView.ColumnHeaders(7).Text = ""
    
    'Set att = db.Execute("SELECT DISTINCT feeclass FROM attendance WHERE idClient = " & clientID & " ORDER BY date DESC")
    Set att = db.Execute("SELECT clients.feeClassID, fee_classes.description, SUM(attendance.attended) as Pres, SUM(1 - attendance.attended) as Abs, SUM(attendance.paid) as Pd, SUM(1 - attendance.paid) as Owing FROM attendance INNER JOIN clients ON (attendance.idClient = clients.idClient) INNER JOIN fee_classes ON (clients.feeClassID = fee_classes.idFeeClasses) WHERE Date >= " & sqlDate(dpFrom.value) & " AND Date <= " & sqlDate(dpTo.value) & " GROUP BY clients.feeClassID")
    
    With att
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do Until .EOF
                'MsgBox !feeClassID
                Set li = ListView.ListItems.Add(, , !feeClassID)
                li.SubItems(1) = !Description
                li.SubItems(2) = !Pres
                li.SubItems(3) = Format(!Pres / CDbl(!Pres + !Abs), "0%")
                li.SubItems(4) = !Pd
                li.SubItems(5) = Format(!Pd / (!Pd + !owing), "0%")
                
                .MoveNext
            Loop
        End If
    End With
    
    Set li = Nothing
    Set att = Nothing
End Sub

