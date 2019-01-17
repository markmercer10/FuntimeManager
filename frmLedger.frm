VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmLedger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client Ledger"
   ClientHeight    =   12855
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12855
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      ItemData        =   "frmLedger.frx":0000
      Left            =   2400
      List            =   "frmLedger.frx":0007
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   120
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListViewFiltered 
      Height          =   1695
      Left            =   0
      TabIndex        =   19
      Top             =   840
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2990
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   8454016
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Debit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Credit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Balance"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   495
      Left            =   2040
      TabIndex        =   18
      Top             =   5880
      Visible         =   0   'False
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton printButn 
      Height          =   495
      Left            =   8040
      Picture         =   "frmLedger.frx":0016
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox tooltip 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7560
      TabIndex        =   15
      Top             =   2040
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Frame FrameAdd 
      Caption         =   "Add Adjustment - "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   1800
      TabIndex        =   4
      Top             =   4680
      Visible         =   0   'False
      Width           =   6375
      Begin VB.TextBox client_id 
         Height          =   495
         Left            =   360
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtNote 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   4105
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   360
         TabIndex        =   13
         Top             =   1800
         Width           =   5175
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   4105
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   11
         Text            =   "0"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton okButn 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4200
         TabIndex        =   5
         Top             =   2880
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtPicker 
         Height          =   375
         Left            =   3480
         TabIndex        =   6
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMM d, yyyy"
         Format          =   56229891
         CurrentDate     =   42718
      End
      Begin VB.Label Label3 
         Caption         =   "Note :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Date :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   9
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Amount :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Cancel 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   6000
         TabIndex        =   7
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.CheckBox summaryButn 
      Caption         =   "Summary"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3480
      Top             =   1440
   End
   Begin MSComctlLib.ListView ListView 
      Height          =   12255
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   21616
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Debit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Credit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Balance"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView SummaryListView 
      Height          =   12855
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   22675
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Reg. Balance"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Subsidy Balance"
         Object.Width           =   3175
      EndProperty
   End
   Begin VB.ComboBox cboClient 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      ItemData        =   "frmLedger.frx":2EED
      Left            =   0
      List            =   "frmLedger.frx":2EEF
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton adjustButn 
      Caption         =   "Add Adjustment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   10
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton receiptButn 
      Caption         =   "Annual Receipt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   17
      Top             =   120
      Width           =   2055
   End
   Begin VB.Menu mnuRC 
      Caption         =   "RightClick"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "frmLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim last_pointed As Long
Dim last_highlited As Long
Dim print_annual_receipt As Boolean

Private Sub adjustButn_Click()
    dtPicker.value = Date
    txtAmount.Text = 0
    txtNote.Text = ""
    FrameAdd.Visible = True
End Sub

Private Sub Cancel_Click()
    FrameAdd.Visible = False
End Sub

Private Sub cboClient_Change()
    cboClient_Click
End Sub

Private Sub cboClient_Click()
    If cboClient.ListIndex >= 0 Then updateListview
    print_annual_receipt = False
End Sub

Private Sub Form_Load()
    Dim cl As ADODB.Recordset
    
    SummaryListView.Top = 0
    
    Set cl = db.Execute("SELECT * FROM clients ORDER BY last, first ASC")
    With cl
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do Until .EOF
                cboClient.AddItem "(" & Trim(str(!idClient)) & ")  -  " & !Last & ", " & !First
                .MoveNext
            Loop
        End If
    End With
    If cboClient.Tag <> "" Then
        cboClient.ListIndex = val(cboClient.Tag)
        cboClient.Tag = ""
    Else
        cboClient.ListIndex = -1
    End If
    Set cl = Nothing
    
    For y = 2015 To year(Date)
        cboYear.AddItem y
    Next y
    cboYear.ListIndex = 0

    ListViewFiltered.Left = ListView.Left
    ListViewFiltered.Top = ListView.Top
    ListViewFiltered.width = ListView.width
    ListViewFiltered.height = ListView.height
    
End Sub

Function getClientID(ByVal s As String) As Long
    If Len(s) > 4 Then
        getClientID = val(MiD$(Left$(s, InStr(1, s, ")") - 1), 2))
    Else
        getClientID = 0
    End If
End Function

Sub updateListview()
    Dim cl As ADODB.Recordset
    Dim pm As ADODB.Recordset
    Dim sc As ADODB.Recordset
    Dim ad As ADODB.Recordset
    Dim li As ListItem
    Dim fli As ListItem
    Dim d As Date
    Dim tempdate As Date
    Dim nextbill As Date
    Dim enddate As Date
    Dim bal As Double
    Dim i As Long
    Dim charges As Double
    Dim school_age As Boolean
    Dim kinder_age As Boolean
    Dim adjustNote As String
    Dim pmnote As String
    Dim daysperweek As Byte
    Dim day As Byte
    Dim si As ListSubItem
    
    ListView.ListItems.Clear
    bal = 0
    ListView.ColumnHeaders(2).width = 4000
    ListView.ColumnHeaders(3).width = 1440
    ListView.ColumnHeaders(4).Text = "Credit"
    ListView.ColumnHeaders(5).width = 1440
    
    If getClientID(cboClient.Text) = 29 Or getClientID(cboClient.Text) = 30 Then  'AMELIA and CATHERINE
        Set li = ListView.ListItems.Add(, , "Default")
        li.SubItems(4) = "0.00"
        Exit Sub
    End If
    
    Set cl = db.Execute("SELECT * FROM clients WHERE idClient=" & getClientID(cboClient.Text))
    Set pm = db.Execute("SELECT * FROM payments WHERE idClient=" & getClientID(cboClient.Text) & " ORDER BY Date ASC")
    Set ad = db.Execute("SELECT * FROM adjustments WHERE idClient=" & getClientID(cboClient.Text) & " ORDER BY Date ASC")
    'Clipboard.SetText "SELECT * FROM payments WHERE idClient=" & getClientID(cboClient.Text) & " ORDER BY Date ASC"
    With cl
        If Not (.EOF And .BOF) Then
            .MoveFirst
            FrameAdd.Caption = "Add Adjustment - " & !First & "  " & !Last
            client_id = !idClient
            
            'If IsNull(!enddate) Then
                enddate = getLatestEnrolledDate(client_id)
                'Always go to today because if the client is not active u still have to check for payments after the end date
            'Else
            '    enddate = !enddate
            'End If
            
            
            If Not (pm.EOF And pm.BOF) Then pm.MoveFirst
            If Not (ad.EOF And ad.BOF) Then ad.MoveFirst
            
            nextbill = nearestFriday(!startdate + 4)
            For d = !startdate To Date
                school_age = schoolAge(!DOB, d)
                
                If li Is Nothing Then
                    Set li = ListView.ListItems.Add(, , "") ' ADD OPENING BALANCE
                    li.SubItems(1) = "Opening Balance"
                    li.SubItems(3) = "0.00"
                End If
                If d = nextbill And d <= enddate Then
                    charges = 0
                    Set fc = db.Execute("SELECT * FROM fee_classes WHERE idFeeClasses = " & getFeeClassAtDate(!idClient, d))
                    If fc.EOF And fc.BOF Then
                        MsgBox !First & " " & !Last & " did not have a valid fee class as of " & d
                        ListView.ListItems.Clear
                        Exit Sub
                    Else
                        daysperweek = fc!days_per_week
                        For tempdate = nextbill - 6 To nextbill
                            day = Weekday(tempdate)
                            If day <> 1 And day <> 7 Then ' it's not the weekend
                                If ((day = 2 And fc!M = 1) Or (day = 3 And fc!T = 1) Or (day = 4 And fc!W = 1) Or (day = 5 And fc!h = 1) Or (day = 6 And fc!f = 1)) And tempdate >= !startdate And getActiveAtDate(client_id, tempdate) Then 'if billed for today
                                    If school_age Then
                                        If isStatHoliday(tempdate) Then ' parents pay minimum $30 for a stat holiday regardless of regular fees and regardless of attendance.
                                            If getFeesAtDate(!idClient, tempdate) / daysperweek <= 30 Then
                                                charges = charges + 30
                                            Else
                                                charges = charges + getFeesAtDate(!idClient, tempdate) / daysperweek
                                            End If
                                        Else
                                            Set sc = db.Execute("SELECT * FROM school_closures WHERE date = " & sqlDate(tempdate))
                                            If sc.EOF And sc.BOF Then
                                                charges = charges + getFeesAtDate(!idClient, tempdate) / daysperweek
                                            Else
                                                If sc!Type = "KC" Then
                                                    kinder_age = kindergartenAge(!DOB, tempdate)
                                                    If kinder_age Then
                                                        charges = charges + 30 ' static value ' KINDERSTART DAY, DAY OFF JUST FOR KINDERGARTEN
                                                    Else
                                                        charges = charges + getFeesAtDate(!idClient, tempdate) / daysperweek
                                                    End If
                                                Else
                                                    charges = charges + 30 ' static value ' GENERAL SCHOOL CLOSURE
                                                End If
                                            End If
                                        End If
                                    Else
                                        charges = charges + getFeesAtDate(!idClient, tempdate) / daysperweek
                                    End If
                                End If
                            End If
                            'If nextbill = CDate("11/11/2016") Then MsgBox charges
                        Next tempdate
                    End If
                    If charges > 0 Then
                        Set li = ListView.ListItems.Add(, , ansiDate(d)) ' ADD CHARGE
                        li.SubItems(3) = " "
                        li.SubItems(2) = Format(charges, "0.00")
                        li.SubItems(1) = "Child Care Fees (" & Format(nextbill - 4, "mmm d") & " to " & Format(nextbill, "mmm d") & ")"
                    End If
                    nextbill = nextbill + 7
                End If
                If Not ad.EOF Then
                    Do While d >= ad!Date
                        Set li = ListView.ListItems.Add(, , ansiDate(d)) ' ADD ADJUSTMENT
                        If ad!amount > 0 Then
                            li.SubItems(3) = " "
                            li.SubItems(2) = Format(ad!amount, "0.00")
                        Else
                            li.SubItems(3) = Format(-ad!amount, "0.00")
                            li.SubItems(2) = " "
                        End If
                        adjustNote = "" & ad!note
                        If adjustNote = "" Then
                            li.SubItems(1) = "Adjustment"
                        Else
                            li.SubItems(1) = "Adjustment - " & Trim(adjustNote)
                        End If
                        li.Tag = ad!idAdjustment
                        ad.MoveNext
                        If ad.EOF Then Exit Do
                    Loop
                End If
                If Not pm.EOF Then
                    Do While d >= pm!Date
                        Set li = ListView.ListItems.Add(, , ansiDate(d)) ' ADD PAYMENT
                        li.SubItems(3) = Format(pm!amount, "0.00")
                        li.SubItems(2) = " "
                        li.SubItems(1) = "Payment"
                        pmnote = "" & Replace("" & pm!details, vbCrLf, " | ")
                        If pmnote <> "" Then li.SubItems(1) = li.SubItems(1) & " - " & pmnote
                        li.ToolTipText = ansiDate(pm!fromDate) & "  to  " & ansiDate(pm!toDate)
                        pm.MoveNext
                        If pm.EOF Then Exit Do
                    Loop 'End If
                End If
            Next d
            bal = 0
            For i = 1 To ListView.ListItems.count
                Set li = ListView.ListItems(i)
                bal = bal + val(li.SubItems(2))
                bal = bal - val(li.SubItems(3))
                li.SubItems(4) = Format(bal, "0.00")
                If bal > !fees * !payperiod Then li.ListSubItems(3).forecolor = vbRed
            Next i
        End If
    End With
    
    If cboYear.ListIndex <> 0 Then 'a year is selected
        ListViewFiltered.ListItems.Clear
        For Each li In ListView.ListItems
            If li.Text <> "" Then
                If year(CDate(li.Text)) = val(cboYear.Text) Then
                    Set fli = ListViewFiltered.ListItems.Add(, , li.Text)
                    For Each si In li.ListSubItems
                        fli.ListSubItems.Add si.index, si.key, si.Text
                    Next si
                End If
            End If
        Next li
        ListViewFiltered.Visible = True
    Else
        ListViewFiltered.Visible = False
    End If
    
    Set ad = Nothing
    Set sc = Nothing
    Set cl = Nothing
    Set pm = Nothing
    Set li = Nothing
End Sub

Private Sub ListView_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Long
    Dim highlite_from As Date
    Dim highlite_to As Date
    Dim do_highlite As Boolean
    Dim tempdate As Date
    
    do_highlite = False
    For i = 1 To ListView.ListItems.count
        If y >= ListView.ListItems(i).Top And y <= ListView.ListItems(i).Top + ListView.ListItems(i).height Then
            If ListView.ListItems(i).ToolTipText = "" Then
                tooltip.Visible = False
                highlite_from = 2 ' bogus dates
                highlite_to = 1
            Else
                tooltip.Text = ListView.ListItems(i).ToolTipText
                tooltip.Top = ListView.Top + y + 190
                tooltip.Visible = True
                highlite_from = CDate(MiD$(tooltip.Text, 1, InStr(1, tooltip.Text, "to") - 1))
                highlite_to = CDate(MiD$(tooltip.Text, InStr(1, tooltip.Text, "to") + 2))
                If last_pointed <> i Then
                    ListView.ListItems(i).forecolor = &H9900CC
                    ListView.ListItems(i).ListSubItems(1).forecolor = &H9900CC
                    ListView.ListItems(i).ListSubItems(3).forecolor = &H9900CC
                    ListView.ListItems(i).ListSubItems(4).forecolor = &H9900CC
                    last_highlited = i
                End If
            End If
            If last_pointed <> i Then
                do_highlite = True
                last_pointed = i
            End If
            Exit For
        End If
    Next i
    
    If do_highlite Then
        
        For i = 2 To ListView.ListItems.count
            If ListView.ListItems(i).Text <> "" Then
                If highlite_to = 1 Then
                    'ListView.ListItems(i).forecolor = vbBlack
                Else
                    If ListView.ListItems(i).forecolor <> vbBlack And i <> last_highlited Then
                        ListView.ListItems(i).forecolor = vbBlack
                        ListView.ListItems(i).ListSubItems(1).forecolor = vbBlack
                        ListView.ListItems(i).ListSubItems(2).forecolor = vbBlack
                        ListView.ListItems(i).ListSubItems(3).forecolor = vbBlack
                        ListView.ListItems(i).ListSubItems(4).forecolor = vbBlack
                    End If
                    tempdate = CDate(ListView.ListItems(i).Text)
                    If tempdate >= highlite_from And tempdate <= highlite_to Then
                        If ListView.ListItems(i).SubItems(1) <> "Payment" Then
                            ListView.ListItems(i).forecolor = vbRed
                            ListView.ListItems(i).ListSubItems(1).forecolor = vbRed
                            ListView.ListItems(i).ListSubItems(2).forecolor = vbRed
                            ListView.ListItems(i).ListSubItems(4).forecolor = vbRed
                        End If
                    End If
                End If
            End If
        Next i
        ListView.Refresh
    End If
    
End Sub

Private Sub ListView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        If Left$(ListView.SelectedItem.SubItems(1), 10) = "Adjustment" Then
            Me.PopupMenu mnuRC
        End If
    End If
End Sub

Private Sub mnuDelete_Click()
    If MsgBox("Are you sure you want to delete " & ListView.SelectedItem.SubItems(1) & " " & ListView.SelectedItem.SubItems(2) & " " & ListView.SelectedItem.SubItems(3), vbYesNo) = vbYes Then
        db.Execute "DELETE FROM adjustments WHERE idAdjustment=" & ListView.SelectedItem.Tag
        updateListview
    End If
End Sub

Private Sub okButn_Click()
    db.Execute "INSERT INTO adjustments SET idClient=" & client_id.Text & ",date=" & sqlDate(dtPicker.value) & ",amount=" & txtAmount & ",note=""" & txtNote & """"
    FrameAdd.Visible = False
    updateListview

End Sub

Private Sub printButn_Click()
    If print_annual_receipt Then
        Dim cl As ADODB.Recordset
        Set cl = db.Execute("SELECT * FROM clients WHERE idClient = " & getClientID(cboClient.Text))
        Dim c As ADODB.Recordset
        Dim parentname As String
        Set c = getParents(cl!idClient)
        If Not (c.EOF And c.BOF) Then
            c.MoveFirst
            parentname = "Parent: " & c!name
        End If
        Printer.PaintPicture frmMain.Image1.Picture, 1100, 1, 1200, 1200
        printText "Funtime Child Care Center", 2500, 1, 10000, "Arial", 24, True, 0
        printText "17 Brown's Place, Whitbourne, NL, A0B 3K0", 2500, 500, 10000, "Arial", 12, True, 0
        printText "Phone: 709-759-2202      Fax : 709-759-3208", 2500, 800, 10000, "Arial", 12, True, 0
        printText "Official Receipt for Income Tax Purposes", 3000, 1250, 10000, "Arial", 14, True, 0
        printText "Received from " & parentname & " for childcare (" & Trim(MiD$(cboClient.Text, InStr(1, cboClient.Text, "-") + 2)) & ")", 1500, 1750, 10000, "Arial", 12, True, 0
        Set cl = Nothing
        printListView ListView, 55, 1500, 2550, 1, True
    Else
        If summaryButn.value = 1 Then
            printText "Account Summary as of " & shortDate(Date), 1500, 1, 10000, "Arial", 16, True, 0
            printListView SummaryListView, 65, 1500, 850, 1, True
        Else
            printText cboClient.Text & " as of " & shortDate(Date), 1500, 1, 10000, "Arial", 16, True, 0
            printListView ListView, 65, 1500, 850, 1, True
        End If
    End If
    Printer.EndDoc
End Sub

Private Sub receiptButn_Click()
    Dim yr As Long
    Dim index As Long
    Dim li As ListItem
    Dim Total As Double
    If month(Date) <= 3 Then
        yr = val(InputBox("Enter Year", "Year", year(Date) - 1))
    Else
        yr = val(InputBox("Enter Year", "Year", year(Date)))
    End If
    Total = 0
    For index = ListView.ListItems.count To 1 Step -1
        If Left$(ListView.ListItems(index).SubItems(1), 7) <> "Payment" Then
            ListView.ListItems.Remove (index)
        ElseIf year(CDate(ListView.ListItems(index).Text)) <> yr Then
            ListView.ListItems.Remove (index)
        ElseIf InStr(1, ListView.ListItems(index).SubItems(1), "Subsidy Automatic") > 0 Then
            ListView.ListItems.Remove (index)
        Else
            Total = Total + val(ListView.ListItems(index).SubItems(3))
            ListView.ListItems(index).SubItems(1) = Left$(Replace(ListView.ListItems(index).SubItems(1), vbCrLf, " "), 50)
        End If
    Next index
    ListView.ColumnHeaders(2).width = 6000
    ListView.ColumnHeaders(3).width = 0
    ListView.ColumnHeaders(4).Text = "Amount"
    ListView.ColumnHeaders(5).width = 0
    Set li = ListView.ListItems.Add(, , "")
    li.SubItems(1) = ""
    li.SubItems(2) = ""
    li.SubItems(3) = ""
    li.SubItems(4) = ""
    Set li = ListView.ListItems.Add(, , "")
    li.SubItems(1) = "Total"
    li.SubItems(3) = Format(Total, "0.00")
    print_annual_receipt = True
    Set li = Nothing
End Sub

Private Sub summaryButn_Click()
    Dim li As ListItem
    If summaryButn.value = 1 Then
        ListView.Visible = False
        ProgressBar.Visible = True
        summaryButn.Enabled = False
        SummaryListView.ListItems.Clear
        
        Dim i As Long
        Dim client As Long
        Dim bal As Double
        Dim bal_sub As Double
        Dim show_client As Boolean
        For i = 0 To cboClient.ListCount - 1
            cboClient.ListIndex = i
            DoEvents
            client = getClientID(cboClient.Text)
            ProgressBar.value = Int((i / (cboClient.ListCount - 1)) * 100)
            DoEvents
            
            If ListView.ListItems.count <= 1 Then
                'brand new client with no records yet
                show_client = False
            ElseIf getActiveAtDate(client, Date) = 0 And year(CDate(ListView.ListItems(ListView.ListItems.count).Text)) < year(Date) Then
                If val(ListView.ListItems(ListView.ListItems.count).SubItems(4)) = 0 Then
                    'old client, and paid off.
                    show_client = False
                Else
                    'deadbeat client who won't pay.
                    show_client = True
                End If
            Else
                'active or recently active client
                show_client = True
            End If
            
            If show_client Then
                Set li = SummaryListView.ListItems.Add(, , Trim(MiD$(cboClient.Text, 8)))
                If ListView.ListItems.count > 0 Then
                    If getSubsidizedAtDate(client, Date) Then
                        li.SubItems(2) = ListView.ListItems(ListView.ListItems.count).SubItems(4)
                    Else
                        li.SubItems(1) = ListView.ListItems(ListView.ListItems.count).SubItems(4)
                    End If
                Else
                    li.SubItems(1) = "0.00"
                End If
            End If
        Next i
        bal = 0
        bal_sub = 0
        For i = 1 To SummaryListView.ListItems.count
            Set li = SummaryListView.ListItems(i)
            bal = bal + val(li.SubItems(1))
            bal_sub = bal_sub + val(li.SubItems(2))
        Next i
        Set li = SummaryListView.ListItems.Add(, , "")
        Set li = SummaryListView.ListItems.Add(, , "Total Receivables")
        li.SubItems(1) = Format(bal, "$0.00")
        li.SubItems(2) = Format(bal_sub, "$0.00")
        SummaryListView.Visible = True
        summaryButn.Enabled = True
        ProgressBar.Visible = False
    Else
        ListView.Visible = True
        SummaryListView.Visible = False
    End If
    Set li = Nothing
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    print_annual_receipt = False
    If cboClient.Tag <> "" Then
        cboClient.ListIndex = val(cboClient.Tag)
        cboClient.Tag = ""
    End If
End Sub
