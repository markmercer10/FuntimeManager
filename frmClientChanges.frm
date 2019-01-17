VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmClientChanges 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Changes to Clients"
   ClientHeight    =   12870
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12870
   ScaleWidth      =   15030
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame frmEdit 
      Caption         =   "Edit Entry"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   4680
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CommandButton cancelButn 
         Caption         =   "Cancel"
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
         Left            =   1920
         TabIndex        =   25
         Top             =   5160
         Width           =   1695
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
         Height          =   495
         Left            =   3720
         TabIndex        =   24
         Top             =   5160
         Width           =   1695
      End
      Begin VB.TextBox txtAuth 
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
         Height          =   375
         Left            =   2280
         TabIndex        =   23
         Text            =   "0"
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox txtParentCont 
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
         Left            =   2280
         TabIndex        =   22
         Text            =   "0"
         Top             =   3840
         Width           =   1455
      End
      Begin VB.TextBox txtFees 
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
         Left            =   2280
         TabIndex        =   21
         Text            =   "0"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CheckBox chkActive 
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
         TabIndex        =   20
         Top             =   4320
         Width           =   855
      End
      Begin VB.CheckBox chkSubs 
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
         TabIndex        =   19
         Top             =   2880
         Width           =   855
      End
      Begin VB.ComboBox room 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "frmClientChanges.frx":0000
         Left            =   2280
         List            =   "frmClientChanges.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2400
         Width           =   1455
      End
      Begin VB.ComboBox pp 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "frmClientChanges.frx":0004
         Left            =   2280
         List            =   "frmClientChanges.frx":0014
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1920
         Width           =   1455
      End
      Begin VB.ComboBox cboFeeClasses 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   960
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker DTPicker 
         Height          =   375
         Left            =   2280
         TabIndex        =   15
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
         Format          =   135528451
         CurrentDate     =   42687
      End
      Begin MSComCtl2.DTPicker StartDTPicker 
         Height          =   375
         Left            =   2280
         TabIndex        =   26
         Top             =   4680
         Visible         =   0   'False
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
         Format          =   135528451
         CurrentDate     =   42687
      End
      Begin MSComCtl2.DTPicker EndDTPicker 
         Height          =   375
         Left            =   2280
         TabIndex        =   28
         Top             =   4800
         Visible         =   0   'False
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
         Format          =   135528451
         CurrentDate     =   42687
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Start Date"
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
         Left            =   120
         TabIndex        =   27
         Top             =   4680
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "End Date"
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
         Left            =   120
         TabIndex        =   29
         Top             =   4920
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Active"
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
         Left            =   120
         TabIndex        =   14
         Top             =   4320
         Width           =   1935
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Parent Contrib."
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
         Left            =   120
         TabIndex        =   13
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Authorization #"
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
         Left            =   120
         TabIndex        =   12
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Subsidized"
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
         Left            =   120
         TabIndex        =   11
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Room"
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
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Pay Period"
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
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Fees"
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
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Fee Class"
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
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Date"
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
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.ListBox ListFeeClasses 
      Height          =   1425
      Left            =   5280
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton printButn 
      BackColor       =   &H00FFBB66&
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
   Begin VB.ListBox ListClients 
      Height          =   1425
      Left            =   3120
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   13200
      Top             =   0
   End
   Begin VB.ComboBox cboClients 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   0
      Width           =   5175
   End
   Begin MSComctlLib.ListView ListView 
      Height          =   12375
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   21828
      SortKey         =   1
      View            =   3
      Sorted          =   -1  'True
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
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fee Class"
         Object.Width           =   7938
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Fees"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Pay Period"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Room"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Subsidized"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Authorization #"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Parental Cont."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "StartDate"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "EndDate"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "Enrolled"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Menu mnuRC 
      Caption         =   "RightClick"
      Visible         =   0   'False
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "frmClientChanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private startdate As Date
Private editID As Long


Private Sub cancelButn_Click()
    frmEdit.Visible = False
    cboClients.Enabled = True
    fillList
End Sub

Private Sub cboClients_Change()
    fillList
End Sub

Private Sub cboClients_Click()
    fillList
End Sub


Private Sub cboFeeClasses_Change()
    txtFees = val(MiD$(cboFeeClasses, InStr(1, cboFeeClasses, "$") + 1))
End Sub

Private Sub cboFeeClasses_Click()
    cboFeeClasses_Change
End Sub

Private Sub DTPicker_Change()
    If dtPicker.value < startdate Then
        MsgBox "You cannot set a client change to before the client's start date"
        dtPicker.value = startdate
    End If
End Sub

Private Sub Form_Load()
    Dim fc As ADODB.Recordset
    Set fc = db.Execute("SELECT * FROM fee_classes")
    With fc
        If Not (.BOF And .EOF) Then
            .MoveFirst
            Do Until .EOF
                cboFeeClasses.AddItem !Description & " - $" & !charge, !idFeeClasses - 1
                .MoveNext
            Loop
        End If
    End With
    Set fc = Nothing
End Sub

Private Sub ListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView.SortKey = ColumnHeader.index - 1
    fillList
End Sub

Private Sub ListView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        If ListView.SelectedItem.index = 1 Or ListView.SelectedItem.index = ListView.ListItems.count Then
            mnuDelete.Enabled = False
        Else
            mnuDelete.Enabled = True
        End If
        Me.PopupMenu mnuRC
    End If
End Sub

Private Sub mnuDelete_Click()
    editID = val(ListView.SelectedItem.Text)
    If MsgBox("Are you sure you want to delete this record?", vbYesNo) = vbYes Then
        db.Execute ("DELETE FROM client_changes WHERE idChange = " & editID)
        fillList
    End If
End Sub

Private Sub mnuEdit_Click()
    Dim q As ADODB.Recordset
    editID = val(ListView.SelectedItem.Text)
    cboClients.Enabled = False
    ListView.Enabled = False
    printButn.Enabled = False
    Set q = db.Execute("SELECT * FROM client_changes WHERE idChange = " & editID)
    With q
        If Not (.EOF And .BOF) Then
            .MoveFirst
            dtPicker.value = ansiDate(!Date)
            cboFeeClasses.ListIndex = !feeClassID - 1
            txtFees = Format(!fees, "0.00")
            comboSelectItem pp, !payperiod
            comboSelectItem room, !room
            chkSubs.value = !subsidized
            txtAuth = "" & !authorizationNumber
            txtParentCont = Format(!parentalContribution, "0.00")
            StartDTPicker.value = !startdate
            'If (IsNull(!enddate)) Then
            '    EndDTPicker.value = 0
            '    EndDTPicker.Visible = False
            'Else
            '    EndDTPicker.value = !enddate
            'End If
            chkActive.value = !active
        End If
    End With
    frmEdit.Visible = True
    
    If ListView.SelectedItem.index = 1 Then
        dtPicker.Enabled = False
        chkActive.Enabled = False
    Else
        dtPicker.Enabled = True
        chkActive.Enabled = True
    End If
    
    Set q = Nothing
End Sub

Private Sub okButn_Click()
    Dim li As ListItem
    Dim latestdate As Date
    
    
    sql = "UPDATE client_changes SET "
    sql = sql & "date=" & sqlDate(dtPicker.value) & ","
    sql = sql & "feeClassID=" & cboFeeClasses.ListIndex + 1 & ","
    sql = sql & "fees=" & txtFees & ","
    sql = sql & "payperiod=" & pp.Text & ","
    sql = sql & "room=""" & room.Text & ""","
    sql = sql & "subsidized=" & chkSubs.value & ","
    If CBool(chkSubs.value) Then
        sql = sql & "authorizationNumber=""" & txtAuth.Text & ""","
        sql = sql & "parentalContribution=" & val(txtParentCont.Text) & ","
    End If
    sql = sql & "startDate=" & sqlDate(StartDTPicker.value) & ","
    'If EndDTPicker.Visible Then
    '    sql = sql & "endDate=" & sqlDate(EndDTPicker.value) & ","
    'Else
    '    sql = sql & "endDate=NULL,"
    'End If
    sql = sql & "active=" & chkActive.value
    sql = sql & " WHERE idChange = " & editID
    
    db.Execute sql
    
    latestdate = CDate("01/01/2000")
    For Each li In ListView.ListItems
        If li <> ListView.SelectedItem Then
            If CDate(li.SubItems(1)) > latestdate Then latestdate = CDate(li.SubItems(1))
        End If
    Next li
    If dtPicker.value > latestdate Then ' this is the last entry so update the clients table to match.
        
        'NOT DONE!!!
        'This setup doesn't account for if you take the latest client change and save it as an earlier one so that a different one ends up being last.
        'in that case I have to have this part LOOK UP the new latest entry and update the clients table with that.
        'But this way is good enough for now It's 1:20AM Nov 15, 2016 and i'm going delirious.
        
        sql = "UPDATE clients SET "
        sql = sql & "feeClassID=" & cboFeeClasses.ListIndex + 1 & ","
        sql = sql & "fees=" & txtFees & ","
        sql = sql & "payperiod=" & pp.Text & ","
        sql = sql & "room=""" & room.Text & ""","
        sql = sql & "subsidized=" & chkSubs.value & ","
        If CBool(chkSubs.value) Then
            sql = sql & "authorizationNumber=""" & txtAuth.Text & ""","
            sql = sql & "parentalContribution=" & val(txtParentCont.Text) & ","
        End If
        sql = sql & "active=" & chkActive.value
        sql = sql & " WHERE idClient = " & ListClients.List(cboClients.ListIndex)
        
        db.Execute sql
    End If
    
    frmEdit.Visible = False
    cboClients.Enabled = True
    ListView.Enabled = True
    printButn.Enabled = True
    fillList
End Sub

Private Sub printButn_Click()
    printButn.Visible = False
    'formPrint Me, 50, 50
    printText cboCat.Text & " - " & cboYr.Text, 50, 50, 5000, "Arial", 12, 0, 0
    'printText cboYr.Text, 4000, 50, 5000, "Arial", 12, 0, 0
    printListView ListView, 65, 50, 900, 1.1, True
    Printer.EndDoc
    printButn.Visible = True
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    
    Dim y As Long
    Dim q As ADODB.Recordset
    
    Set q = db.Execute("SELECT * FROM clients ORDER BY Last, First ASC")
    With q
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do Until .EOF
                cboClients.AddItem !Last & ", " & !First
                ListClients.AddItem !idClient
                If cboClients.Tag = !idClient Then cboClients.ListIndex = cboClients.ListCount - 1
                .MoveNext
            Loop
        End If
    End With
    
    Set q = db.Execute("SELECT * FROM fee_classes ORDER BY idFeeClasses")
    With q
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do Until .EOF
                ListFeeClasses.AddItem !Description, !idFeeClasses - 1
                .MoveNext
            Loop
        End If
    End With
    
    Set q = db.Execute("SELECT * FROM rooms")
    With q
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do Until .EOF
                room.AddItem !Abbreviation
                .MoveNext
            Loop
        End If
    End With
    
    Set q = Nothing
End Sub

Sub fillList()
    Dim q As ADODB.Recordset
    Dim li As ListItem
    Dim Total As Double
    ListView.ListItems.Clear
    
    Total = 0
    ListView.Sorted = False
    Set q = db.Execute("SELECT * FROM client_changes WHERE idClient = """ & ListClients.List(cboClients.ListIndex) & """ ORDER BY date, idChange ASC")
    With q
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do Until .EOF
                Set li = ListView.ListItems.Add(, , !idChange)
                li.SubItems(1) = ansiDate(!Date)
                li.SubItems(2) = ListFeeClasses.List(!feeClassID - 1)
                li.SubItems(3) = Format(!fees, "0.00")
                li.SubItems(4) = !payperiod
                li.SubItems(5) = !room
                li.SubItems(6) = !subsidized
                li.SubItems(7) = "" & !authorizationNumber
                li.SubItems(8) = !parentalContribution
                li.SubItems(9) = "" & ansiDate(!startdate)
                'li.SubItems(10) = "" & ansiDate(!enddate)
                li.SubItems(11) = !active
                'total = total + (!value_num / 100#)
                .MoveNext
            Loop
        End If
    End With
    
    DoEvents
    startdate = CDate(ListView.ListItems(1).SubItems(1))
    'MsgBox startdate
    
    Set q = Nothing
    Set li = Nothing
End Sub


