VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExpenses 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Expenses"
   ClientHeight    =   12870
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12870
   ScaleWidth      =   9405
   StartUpPosition =   3  'Windows Default
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
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   1215
   End
   Begin VB.ListBox ListCat 
      Height          =   1425
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   6120
      Top             =   0
   End
   Begin VB.ComboBox cboYr 
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
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   0
      Width           =   1455
   End
   Begin VB.ComboBox cboCat 
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
      Width           =   4215
   End
   Begin MSComctlLib.ListView ListView 
      Height          =   12375
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   9375
      _ExtentX        =   16536
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Amount"
         Object.Width           =   2646
      EndProperty
   End
End
Attribute VB_Name = "frmExpenses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboCat_Change()
    fillList
End Sub

Private Sub cboCat_Click()
    fillList
End Sub

Private Sub ListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView.SortKey = ColumnHeader.index - 1
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
    
    Dim Y As Long
    Dim q As ADODB.Recordset
    
    Set q = gcdb.Execute("SELECT * FROM accounts WHERE account_type = ""EXPENSE""")
    With q
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do Until .EOF
                cboCat.AddItem !name
                ListCat.AddItem !guid
                .MoveNext
            Loop
        End If
    End With
    Set q = Nothing
    
    For Y = 2016 To year(Date)
        cboYr.AddItem Y
        cboYr.ListIndex = cboYr.ListCount - 1
    Next Y
End Sub

Sub fillList()
    Dim q As ADODB.Recordset
    Dim li As ListItem
    Dim total As Double
    ListView.ListItems.Clear
    
    total = 0
    ListView.Sorted = True
    Set q = gcdb.Execute("SELECT * FROM splits INNER JOIN transactions ON splits.tx_guid = transactions.guid WHERE splits.account_guid = """ & ListCat.List(cboCat.ListIndex) & """")
    With q
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do Until .EOF
                Set li = ListView.ListItems.Add(, , !guid)
                li.SubItems(1) = ansiDate(!post_date)
                li.SubItems(2) = !Description
                li.SubItems(3) = Format(!value_num / 100, "0.00")
                total = total + (!value_num / 100#)
                .MoveNext
            Loop
        End If
    End With
    
    DoEvents
    ListView.Sorted = False
    Set li = ListView.ListItems.Add(, , "")
    Set li = ListView.ListItems.Add(, , "Total")
    li.SubItems(2) = "Total"
    li.SubItems(3) = total
    
    Set q = Nothing
    Set li = Nothing
End Sub
