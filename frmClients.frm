VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmClients 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clients"
   ClientHeight    =   6225
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   18330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   18330
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton ButnClientLabels 
      Caption         =   "Client Labels"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   7
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton butnChanges 
      Caption         =   "Client Changes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   6
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton listButn 
      Caption         =   "Print List of Active Clients"
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
      Left            =   10800
      TabIndex        =   5
      Top             =   120
      Width           =   3015
   End
   Begin MSComCtl2.DTPicker dpAges 
      Height          =   300
      Left            =   9120
      TabIndex        =   4
      Top             =   195
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "MMM d, yyyy"
      Format          =   135659523
      CurrentDate     =   42536
   End
   Begin VB.CheckBox chkActive 
      Caption         =   "Display only active clients"
      Height          =   252
      Left            =   4800
      TabIndex        =   2
      Top             =   240
      Value           =   1  'Checked
      Width           =   2292
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   7320
      Top             =   120
   End
   Begin VB.CommandButton addButn 
      Caption         =   "Add Client"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1212
   End
   Begin MSComctlLib.ListView ListView 
      Height          =   5532
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   18252
      _ExtentX        =   32200
      _ExtentY        =   9763
      View            =   3
      LabelWrap       =   0   'False
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
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   16
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Last Name"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "First Name"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Init"
         Object.Width           =   776
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "DOB"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Age"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Gender"
         Object.Width           =   758
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "MCP Number"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Allergies"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Fees"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Acct. Created"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Formerly End Date"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "PP"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Room"
         Object.Width           =   1429
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   14
         Text            =   "Subsidized"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   15
         Text            =   "Enrolled"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "List Ages as of"
      Height          =   255
      Left            =   7920
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmClients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub addButn_Click()
    dlgClient.Show 1
    updateListview
End Sub

Private Sub butnChanges_Click()
    Load frmClientChanges
    frmClientChanges.cboClients.Tag = ListView.SelectedItem.Text
    frmClientChanges.Show 1
End Sub

Private Sub ButnClientLabels_Click()
    frmClientLabels.Show 1
End Sub
Private Sub chkActive_Click()
    updateListview
End Sub

Private Sub dpAges_Change()
    dpAges_Click
End Sub

Private Sub dpAges_Click()
    Dim i As Long
    For i = 1 To ListView.ListItems.count
        ListView.ListItems(i).SubItems(9) = getAge(CDate(ListView.ListItems(i).SubItems(6)), dpAges.value)
    Next i
End Sub

Private Sub Form_Load()
    'Dim s As String
    'Dim c As Long
    's = ""
    'For c = 1 To 20
    '    s = s & createGUID & vbCrLf
    'Next c
    'MsgBox s
End Sub

Private Sub listButn_Click()
    Dim q As ADODB.Recordset
    Dim y As Long
    Dim c As Long
    'Dim li As ListItem
    
    'ListView.ListItems.Clear
    printText "Active Clients", 50, 50, 10000, "Arial", 22, True, 0
    y = 600
    c = 1
    Set q = db.Execute("SELECT * FROM clients WHERE active=1 ORDER BY room DESC, last, first ASC")
    With q
        If Not (.BOF And .EOF) Then
            .MoveFirst
            Do Until .EOF
                printText Format(c, "00") & ".  " & !room & "  -  " & !First & " " & !Last, 50, y, 10000, "Arial", 11, False, 0
                c = c + 1
                y = y + 250
                .MoveNext
            Loop
        End If
    End With
    Printer.EndDoc
    
    'If ListView.ListItems.count > 0 Then ListView.ListItems(1).Selected = True
    Set q = Nothing
    'Set li = Nothing

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

Private Sub ListView_DblClick()
    If ListView.SelectedItem >= 0 Then
        dlgClient.ID = ListView.SelectedItem.Text
        dlgClient.Show 1
        updateListview
    End If
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    dpAges = Date
    updateListview
    ListView.SortKey = 2
    ListView.SortKey = 1
    ListView.Sorted = True
End Sub

Sub updateListview()
    Dim q As ADODB.Recordset
    Dim li As ListItem
    
    ListView.ListItems.Clear
    If chkActive.value = 1 Then
        Set q = db.Execute("SELECT * FROM clients WHERE active=1")
    Else
        Set q = db.Execute("SELECT * FROM clients")
    End If
    With q
        If Not (.BOF And .EOF) Then
            .MoveFirst
            Do Until .EOF
                Set li = ListView.ListItems.Add(, , !idClient)
                li.SubItems(1) = !Last
                li.SubItems(2) = !First
                li.SubItems(3) = "" & !initial
                li.SubItems(4) = ansiDate(!DOB) 'shortDate(!DOB)
                li.SubItems(5) = getAge(!DOB, dpAges.value)
                li.SubItems(6) = !gender
                li.SubItems(7) = "" & !MCP
                li.SubItems(8) = "" & !allergies
                li.SubItems(9) = !fees
                If !fees <= 0 Then li.ListSubItems(9).forecolor = vbRed: li.ListSubItems(9).bold = True
                li.SubItems(10) = ansiDate(!startdate)
                'li.SubItems(11) = ansiDate(!enddate)
                li.SubItems(12) = !payperiod
                li.SubItems(13) = !room
                If !subsidized Then li.SubItems(14) = Chr(149)
                If !active Then li.SubItems(15) = Chr(149)
                .MoveNext
            Loop
        End If
    End With
    
    If ListView.ListItems.count > 0 Then ListView.ListItems(1).selected = True
    Set q = Nothing
    Set li = Nothing
End Sub
