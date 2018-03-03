VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmQIF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Import QIF"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   15345
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame resultsFrame 
      Height          =   8295
      Left            =   5400
      TabIndex        =   4
      Top             =   -2400
      Visible         =   0   'False
      Width           =   15255
      Begin VB.ListBox ListCat 
         Height          =   840
         Left            =   7680
         TabIndex        =   9
         Top             =   5400
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.ComboBox cboAccount 
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
         Left            =   7680
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   4920
         Visible         =   0   'False
         Width           =   2070
      End
      Begin VB.CommandButton importButn 
         Caption         =   "Import"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2880
         TabIndex        =   7
         Top             =   7800
         Width           =   2655
      End
      Begin VB.CommandButton filtersButn 
         Caption         =   "Manage Filters"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         TabIndex        =   6
         Top             =   7800
         Width           =   2655
      End
      Begin MSComctlLib.ListView ListView 
         Height          =   7575
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   13361
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Memo  -  Payee"
            Object.Width           =   8290
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Amount"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Account"
            Object.Width           =   3598
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Duplicates"
            Object.Width           =   8820
         EndProperty
      End
   End
   Begin VB.TextBox filepath 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      TabIndex        =   3
      Top             =   7800
      Width           =   12615
   End
   Begin VB.FileListBox FileList 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7650
      Left            =   8040
      Pattern         =   "*.QIF"
      TabIndex        =   2
      Top             =   120
      Width           =   7215
   End
   Begin VB.DirListBox DirList 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7500
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7815
   End
   Begin VB.CommandButton readButn 
      Caption         =   "Read File"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   12840
      TabIndex        =   0
      Top             =   7800
      Width           =   2415
   End
End
Attribute VB_Name = "frmQIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private moused As Long
Private droppeddown As Boolean

Private Sub cboAccount_Click()
    If droppeddown Then
        ListView.ListItems(moused).SubItems(3) = cboAccount.Text
        ListView.ListItems(moused).Tag = ListCat.List(cboAccount.ListIndex)
        ListView.ListItems(moused).forecolor = vbBlack
    End If
    droppeddown = False
End Sub

Private Sub cboAccount_DropDown()
    droppeddown = True
End Sub

Private Sub DirList_Click()
    FileList.path = DirList.List(DirList.ListIndex)
    filepath = ""
End Sub

Private Sub FileList_Click()
    filepath = FileList.path & "\" & FileList.List(FileList.ListIndex)
End Sub

Private Sub filepath_Change()
    If filepath = "" Then
        readButn.Enabled = False
    Else
        readButn.Enabled = True
    End If
End Sub

Private Sub filtersButn_Click()
    frmQIFfilters.Show 1
    ListView.ListItems.Clear
    resultsFrame.Visible = False
End Sub

Private Sub Form_Load()
    Me.width = 15435
    resultsFrame.Left = 0
    resultsFrame.Top = 0
    DirList.path = App.path
    FileList.path = App.path
    droppeddown = False
    
    Dim y As Long
    Dim q As ADODB.Recordset
    
    Set q = gcdb.Execute("SELECT * FROM accounts WHERE account_type = ""EXPENSE"" ORDER BY name ASC")
    With q
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do Until .EOF
                cboAccount.AddItem !name
                ListCat.AddItem !guid
                .MoveNext
            Loop
        End If
    End With
    Set q = Nothing
End Sub

Private Sub importButn_Click()
    Dim lvindex As Long
    Dim count As Long
    
    For lvindex = 1 To ListView.ListItems.count
        If ListView.ListItems(lvindex).Checked And ListView.ListItems(lvindex).SubItems(3) = "" Then
            MsgBox ListView.ListItems(lvindex).Text & " - " & ListView.ListItems(lvindex).SubItems(1) & vbCrLf & "This item has no account selected!"
            ListView.ListItems(lvindex).selected = True
            Exit Sub
        End If
        If ListView.ListItems(lvindex).Checked And Left$(ListView.ListItems(lvindex).SubItems(4), 9) = "DUPLICATE" Then
            If MsgBox(ListView.ListItems(lvindex).Text & " - " & ListView.ListItems(lvindex).SubItems(1) & vbCrLf & "This item has a matching entry already in the database!  Are you sure you want to import it", vbYesNo) = vbNo Then
                ListView.ListItems(lvindex).selected = True
                Exit Sub
            End If
        End If
    Next lvindex


    importButn.Enabled = False
    createGUID_sequential True
    count = 0
    For lvindex = 1 To ListView.ListItems.count
        If ListView.ListItems(lvindex).Checked Then
            create_gnc_imported_expense createGUID_sequential, ListView.ListItems(lvindex).Tag, ListView.ListItems(lvindex).SubItems(1), ListView.ListItems(lvindex).SubItems(2), CDate(ListView.ListItems(lvindex).Text)
            count = count + 1
        End If
    Next lvindex
    resultsFrame.Visible = False
    MsgBox "Import Complete.  " & count & " records created."
    importButn.Enabled = True
End Sub

Private Sub ListView_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked Then
        If Item.SubItems(3) = "" Then
            Item.forecolor = vbRed
        End If
    Else
        Item.forecolor = vbBlack
    End If
End Sub

Private Sub ListView_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Index As Long
    Dim i As Long
    droppeddown = False
    For Index = 1 To ListView.ListItems.count
        If y > ListView.ListItems(Index).Top And y <= ListView.ListItems(Index).Top + ListView.ListItems(Index).height Then
            If ListView.ListItems(Index).Checked Then
                cboAccount.Top = ListView.ListItems(Index).Top + ListView.Top - 15
                cboAccount.Visible = True
                'cboaccount.ListIndex = the index of the same as whats in the listview cell
                For i = 0 To cboAccount.ListCount - 1
                    If cboAccount.List(i) = ListView.ListItems(Index).SubItems(3) Then
                        cboAccount.ListIndex = i
                        Exit For
                    End If
                Next i
            Else
                cboAccount.Visible = False
            End If
            moused = Index
            Exit For
        End If
        If Index = ListView.ListItems.count Then cboAccount.Visible = False
    Next Index
End Sub

Private Sub readButn_Click()
    'this version can read a qif file that has ONLY line feeds at the end of every line.  The entire file comes in as one string and must be tokenized
    Dim ST As StringTokenizer
    Dim path As String
    Dim wholefile As String
    Dim fileline As String
    Dim tempdate As String
    Dim d As Date
    Dim T As Double
    Dim P As String
    Dim M As String
    Dim found As Boolean
    Dim li As ListItem
    Dim qifFilters As ADODB.Recordset
    Dim account_guid As String
    Dim account_query As ADODB.Recordset
    Dim P_temp As String
    Dim possible_duplicates As String
    Dim possible_duplicates_count As Byte
    Dim possible_duplicates_total As Double
    
        'D - Date
        'T - Amount
        'P - Payee
        'M - Memo
        
        '!Type:Bank
        '!Type:CCard
        
        'might not need to include withdrawals.  but if i do I have to create an account for cash in gnucash
        '-WITHDRAWAL
        '-ABM Withdrawal
        
        'create rules for these
        '-Service Charge
        '-Overdrawn Handling Chg.
        '-POS Purchase (various payees)
        '-Bill Payment (various payees)
        'maybe Loan Repayments
    
    Set qifFilters = db.Execute("SELECT * FROM qif_filters")
    
    path = filepath 'App.path & "\pcbanking-1.qif"
    Open path For Input As #1
        Input #1, wholefile
    Close #1
            
    Set ST = New StringTokenizer
    ST.Init wholefile, vbLf
        Do Until Not ST.hasMoreTokens
            fileline = ST.nextToken
            If Left$(fileline, 1) = "D" Then
                tempdate = MiD$(fileline, 2)
                tempdate = MonthName(val(Left$(tempdate, InStr(1, tempdate, "/") - 1))) & " " & MiD$(tempdate, InStr(1, tempdate, "/") + 1)
                d = CDate(Replace(tempdate, "/", ","))
            End If
            If Left$(fileline, 1) = "T" Then T = val(MiD$(fileline, 2))
            If Left$(fileline, 1) = "P" Then P = MiD$(fileline, 2)
            If Left$(fileline, 1) = "M" Then M = MiD$(fileline, 2)
            If Left$(fileline, 1) = "^" Then
                Set li = ListView.ListItems.Add(, , ansiDate(d))
                li.SubItems(1) = M & " - " & P
                li.SubItems(2) = -T
                li.SubItems(3) = "" ' account
                li.SubItems(4) = "" ' duplicates
                P = Trim(UCase(P))
                M = Trim(UCase(M))
                account_guid = ""
                With qifFilters
                    If Not (.EOF And .BOF) Then .MoveFirst
                    Do Until .EOF
                        
                        If !M = "" Or M = UCase(!M) Then
                            'MsgBox M & "   " & !M
                            'MsgBox P & "   " & !P
                            If !P = "" Or InStr(1, P, UCase(!P)) > 0 Then
                                li.SubItems(3) = get_gnc_account_name(!account)
                                li.Tag = !account
                                li.Checked = True
                                account_guid = !account
                                P_temp = UCase(!P)
                                Exit Do
                            End If
                        End If
                        .MoveNext
                    Loop
                End With
                
                'if account_guid <> "" then search for duplicate entries
                If account_guid <> "" Then
                    'Set account_query = gcdb.Execute("SELECT splits.account_guid, splits.value_num / splits.value_denom as value, transactions.post_date FROM splits INNER JOIN transactions ON splits.tx_guid = transactions.guid WHERE transactions.post_date = " & sqlDate(D) & " AND splits.account_guid = """ & account_guid & """ AND value <= " & -T & " ORDER BY value DESC")
                    Set account_query = gcdb.Execute("SELECT (transactions.description) AS description, (splits.value_num / splits.value_denom) AS value FROM splits INNER JOIN transactions ON splits.tx_guid = transactions.guid WHERE DATE(transactions.post_date) = " & sqlDate(d) & " AND splits.account_guid = """ & account_guid & """ AND (splits.value_num / splits.value_denom) <= " & -T & " AND UCASE(transactions.description) LIKE ""%" & P_temp & "%"" ORDER BY value DESC")
                    'If P_temp = "ESSO" Then
                    '    Clipboard.SetText "SELECT transactions.description, (splits.value_num / splits.value_denom) AS value FROM splits INNER JOIN transactions ON splits.tx_guid = transactions.guid WHERE transactions.post_date = " & sqlDate(D) & " AND splits.account_guid = """ & account_guid & """ AND (splits.value_num / splits.value_denom) <= " & -T & " AND UCASE(transactions.description) LIKE ""%" & P_temp & "%"" ORDER BY value DESC"
                    '    MsgBox "SELECT transactions.description, (splits.value_num / splits.value_denom) AS value FROM splits INNER JOIN transactions ON splits.tx_guid = transactions.guid WHERE transactions.post_date = " & sqlDate(D) & " AND splits.account_guid = """ & account_guid & """ AND (splits.value_num / splits.value_denom) <= " & -T & " AND UCASE(transactions.description) LIKE ""%" & P_temp & "%"" ORDER BY value DESC"
                    'End If
                    
                    With account_query
                        If Not (.EOF And .BOF) Then
                            .MoveFirst
                            possible_duplicates = ""
                            possible_duplicates_count = 0
                            possible_duplicates_total = -T
                            found = False
                            Do Until .EOF
                                If !value = -T Then
                                    li.SubItems(4) = "DUPLICATE " & !Description & " - " & Format(!value, "0.00")
                                    li.Checked = False
                                    found = True
                                    Exit Do
                                ElseIf !value = possible_duplicates_total Then
                                    li.SubItems(4) = "DUPLICATES " & possible_duplicates & Format(!value, "0.00") & ", "
                                    li.Checked = False
                                    found = True
                                    Exit Do
                                Else
                                    possible_duplicates = possible_duplicates & Format(!value, "0.00") & ", "
                                    possible_duplicates_total = possible_duplicates_total - !value
                                    possible_duplicates_count = possible_duplicates_count + 1
                                End If
                                .MoveNext
                            Loop
                            If Not found Then
                                li.SubItems(4) = "??? " & possible_duplicates
                                li.ListSubItems(4).forecolor = vbRed
                            End If
                        End If
                    End With
                End If
            End If
                
        Loop
    
    Set li = Nothing
    
    Set qifFilters = Nothing
    Set account_query = Nothing
    resultsFrame.Visible = True
End Sub

Private Sub readButn_Click_old()
    'this version can read a qif file that has carriage returns at the end of every line.  Without them the entire file comes in as one string and must be tokenized
    Dim path As String
    Dim fileline As String
    Dim tempdate As String
    Dim d As Date
    Dim T As Double
    Dim P As String
    Dim M As String
    Dim found As Boolean
    Dim li As ListItem
    Dim qifFilters As ADODB.Recordset
    Dim account_guid As String
    Dim account_query As ADODB.Recordset
    Dim P_temp As String
    Dim possible_duplicates As String
    Dim possible_duplicates_count As Byte
    Dim possible_duplicates_total As Double
    
        'D - Date
        'T - Amount
        'P - Payee
        'M - Memo
        
        '!Type:Bank
        '!Type:CCard
        
        'might not need to include withdrawals.  but if i do I have to create an account for cash in gnucash
        '-WITHDRAWAL
        '-ABM Withdrawal
        
        'create rules for these
        '-Service Charge
        '-Overdrawn Handling Chg.
        '-POS Purchase (various payees)
        '-Bill Payment (various payees)
        'maybe Loan Repayments
    
    Set qifFilters = db.Execute("SELECT * FROM qif_filters")
    
    path = filepath 'App.path & "\pcbanking-1.qif"
    Open path For Input As #1
        Do
            Input #1, fileline
            If Left$(fileline, 1) = "D" Then
                tempdate = MiD$(fileline, 2)
                tempdate = MonthName(val(Left$(tempdate, InStr(1, tempdate, "/") - 1))) & " " & MiD$(tempdate, InStr(1, tempdate, "/") + 1)
                d = CDate(Replace(tempdate, "/", ","))
            End If
            If Left$(fileline, 1) = "T" Then T = val(MiD$(fileline, 2))
            If Left$(fileline, 1) = "P" Then P = MiD$(fileline, 2)
            If Left$(fileline, 1) = "M" Then M = MiD$(fileline, 2)
            If Left$(fileline, 1) = "^" Then
                Set li = ListView.ListItems.Add(, , ansiDate(d))
                li.SubItems(1) = M & " - " & P
                li.SubItems(2) = -T
                li.SubItems(3) = "" ' account
                li.SubItems(4) = "" ' duplicates
                P = UCase(P)
                M = UCase(M)
                account_guid = ""
                With qifFilters
                    If Not (.EOF And .BOF) Then .MoveFirst
                    Do Until .EOF
                        If !M = "" Or M = UCase(!M) Then
                            If !P = "" Or InStr(1, P, UCase(!P)) > 0 Then
                                li.SubItems(3) = get_gnc_account_name(!account)
                                li.Tag = !account
                                li.Checked = True
                                account_guid = !account
                                P_temp = UCase(!P)
                                Exit Do
                            End If
                        End If
                        .MoveNext
                    Loop
                End With
                
                'if account_guid <> "" then search for duplicate entries
                If account_guid <> "" Then
                    'Set account_query = gcdb.Execute("SELECT splits.account_guid, splits.value_num / splits.value_denom as value, transactions.post_date FROM splits INNER JOIN transactions ON splits.tx_guid = transactions.guid WHERE transactions.post_date = " & sqlDate(D) & " AND splits.account_guid = """ & account_guid & """ AND value <= " & -T & " ORDER BY value DESC")
                    Set account_query = gcdb.Execute("SELECT (transactions.description) AS description, (splits.value_num / splits.value_denom) AS value FROM splits INNER JOIN transactions ON splits.tx_guid = transactions.guid WHERE DATE(transactions.post_date) = " & sqlDate(d) & " AND splits.account_guid = """ & account_guid & """ AND (splits.value_num / splits.value_denom) <= " & -T & " AND UCASE(transactions.description) LIKE ""%" & P_temp & "%"" ORDER BY value DESC")
                    'If P_temp = "ESSO" Then
                    '    Clipboard.SetText "SELECT transactions.description, (splits.value_num / splits.value_denom) AS value FROM splits INNER JOIN transactions ON splits.tx_guid = transactions.guid WHERE transactions.post_date = " & sqlDate(D) & " AND splits.account_guid = """ & account_guid & """ AND (splits.value_num / splits.value_denom) <= " & -T & " AND UCASE(transactions.description) LIKE ""%" & P_temp & "%"" ORDER BY value DESC"
                    '    MsgBox "SELECT transactions.description, (splits.value_num / splits.value_denom) AS value FROM splits INNER JOIN transactions ON splits.tx_guid = transactions.guid WHERE transactions.post_date = " & sqlDate(D) & " AND splits.account_guid = """ & account_guid & """ AND (splits.value_num / splits.value_denom) <= " & -T & " AND UCASE(transactions.description) LIKE ""%" & P_temp & "%"" ORDER BY value DESC"
                    'End If
                    
                    With account_query
                        If Not (.EOF And .BOF) Then
                            .MoveFirst
                            possible_duplicates = ""
                            possible_duplicates_count = 0
                            possible_duplicates_total = -T
                            found = False
                            Do Until .EOF
                                If !value = -T Then
                                    li.SubItems(4) = "DUPLICATE " & !Description & " - " & Format(!value, "0.00")
                                    li.Checked = False
                                    found = True
                                    Exit Do
                                ElseIf !value = possible_duplicates_total Then
                                    li.SubItems(4) = "DUPLICATES " & possible_duplicates & Format(!value, "0.00") & ", "
                                    li.Checked = False
                                    found = True
                                    Exit Do
                                Else
                                    possible_duplicates = possible_duplicates & Format(!value, "0.00") & ", "
                                    possible_duplicates_total = possible_duplicates_total - !value
                                    possible_duplicates_count = possible_duplicates_count + 1
                                End If
                                .MoveNext
                            Loop
                            If Not found Then
                                li.SubItems(4) = "??? " & possible_duplicates
                                li.ListSubItems(4).forecolor = vbRed
                            End If
                        End If
                    End With
                End If
            End If
                
        Loop Until EOF(1)
    Close #1
    
    Set li = Nothing
    
    Set qifFilters = Nothing
    Set account_query = Nothing
    resultsFrame.Visible = True
End Sub

