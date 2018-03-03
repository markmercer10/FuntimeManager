VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReceipts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payments"
   ClientHeight    =   9855
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9855
   ScaleWidth      =   10755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ledgerButn 
      Caption         =   "View Ledger"
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
      Height          =   492
      Left            =   9120
      TabIndex        =   5
      Top             =   0
      Width           =   1575
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
      Height          =   492
      Left            =   1680
      TabIndex        =   4
      Top             =   0
      Width           =   1212
   End
   Begin VB.ComboBox cboClient 
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
      Height          =   390
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   90
      Width           =   3495
   End
   Begin VB.CheckBox chkFilterClient 
      Caption         =   "Filter by Client"
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
      Left            =   3720
      TabIndex        =   2
      Top             =   90
      Width           =   1815
   End
   Begin VB.CommandButton addButn 
      Caption         =   "New Payment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3120
      Top             =   120
   End
   Begin MSComctlLib.ListView ListView 
      Height          =   9375
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   16536
      SortKey         =   6
      View            =   3
      SortOrder       =   -1  'True
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
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Client"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Receipt Date"
         Object.Width           =   3352
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "From Date"
         Object.Width           =   3352
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "To Date"
         Object.Width           =   3352
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amount"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ansiDate"
         Object.Width           =   0
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
Attribute VB_Name = "frmReceipts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function getClientID(ByVal s As String) As Long
    If Len(s) > 4 Then
        getClientID = val(MiD$(Left$(s, InStr(1, s, ")") - 1), 2))
    Else
        getClientID = 0
    End If
End Function

Private Sub addButn_Click()
    dlgReceipt.Show 1
    updateListview
End Sub


Private Sub cboClient_Change()
    cboClient_Click
End Sub

Private Sub cboClient_Click()
    If chkFilterClient.value = 1 And cboClient.ListIndex >= 0 Then updateListview
End Sub

Private Sub chkFilterClient_Click()
    cboClient.Enabled = CBool(chkFilterClient)
    If chkFilterClient.value = 1 And cboClient.ListIndex >= 0 Then updateListview
    If chkFilterClient.value = 0 Then updateListview
    ledgerButn.Enabled = -CBool(chkFilterClient.value)
End Sub



Private Sub Form_Load()
    Dim cl As ADODB.Recordset
    
    Set cl = db.Execute("SELECT * FROM clients ORDER BY last, first ASC")
    With cl
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do Until .EOF
                cboClient.AddItem "(" & Trim(str(!idClient)) & ") " & !Last & ", " & !First
                .MoveNext
            Loop
        End If
    End With
    cboClient.ListIndex = 0
    Set cl = Nothing

End Sub

Private Sub ledgerButn_Click()
    Load frmLedger
    frmLedger.cboClient.Tag = cboClient.ListIndex
    frmLedger.Show 1
End Sub

Private Sub ListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If ListView.SortKey = ColumnHeader.Index - 1 Then
        If ListView.SortOrder = lvwAscending Then
            ListView.SortOrder = lvwDescending
        Else
            ListView.SortOrder = lvwAscending
        End If
    Else
        ListView.SortKey = ColumnHeader.Index - 1
    End If
    ListView.Sorted = True
End Sub

Private Sub ListView_DblClick()
    dlgReceipt.ID = ListView.SelectedItem.Text
    dlgReceipt.Show 1
    updateListview
End Sub

Private Sub ListView_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 2 Then
        'ListView.SelectedItem = ListView.ListItems(listview.r
        Me.PopupMenu mnuRC
        'ListView.SetFocus
    End If
End Sub

Private Sub mnuDelete_Click()
    Dim ID As String
    ID = ListView.SelectedItem.Text
    If MsgBox("You should delete receipts only if you absolutely must!" & vbCrLf & "Are you sure you want to delete the following record?" & vbCrLf & vbCrLf & ID & vbCrLf & ListView.SelectedItem.SubItems(1) & vbCrLf & ListView.SelectedItem.SubItems(2) & vbCrLf & ListView.SelectedItem.SubItems(5), vbYesNo) = vbYes Then
        db.Execute "DELETE FROM payments WHERE guid=""" & ID & """;"
        gcdb.Execute "DELETE FROM transactions WHERE guid=""" & ID & """;"
        gcdb.Execute "DELETE FROM splits WHERE tx_guid=""" & ID & """;"
        updateListview
    End If
End Sub

Private Sub mnuEdit_Click()
    ListView_DblClick
End Sub

Private Sub printButn_Click()
    If chkFilterClient = 1 Then
        printText "Receipts - " & cboClient.Text, 50, 500, 15000, "Arial", 16, True, 0
    Else
        printText "Receipts ", 50, 500, 15000, "Arial", 16, True, 0
    End If
    printListView ListView, 60, 50, 1500, 1.15, True
    Printer.EndDoc
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    updateListview
End Sub
Sub updateListview()
    Dim q As ADODB.Recordset
    Dim cl As ADODB.Recordset
    Dim li As ListItem
    
    ListView.ListItems.Clear
    ListView.Sorted = False
    If chkFilterClient.value = 1 Then
        Set q = db.Execute("SELECT * FROM payments WHERE idClient=" & getClientID(cboClient.Text))
    Else
        Set q = db.Execute("SELECT * FROM payments")
    End If
    With q
        If Not (.BOF And .EOF) Then
            .MoveFirst
            Do Until .EOF
                Set cl = db.Execute("SELECT * FROM clients WHERE idClient=" & !idClient)
                If Not (cl.EOF And cl.BOF) Then
                    cl.MoveFirst
                        Set li = ListView.ListItems.Add(, , !guid)
                        li.SubItems(1) = cl!Last & ", " & cl!First
                        li.SubItems(2) = shortDate(!Date)
                        li.SubItems(3) = shortDate(!fromDate)
                        li.SubItems(4) = shortDate(!toDate)
                        li.SubItems(5) = Format(!amount, "0.00")
                        li.SubItems(6) = ansiDate(!Date)
                        .MoveNext
                End If
            Loop
        End If
    End With
    ListView.Sorted = True
    If ListView.ListItems.count > 0 Then ListView.ListItems(1).selected = True
    
    Set q = Nothing
    Set cl = Nothing
    Set li = Nothing
End Sub
