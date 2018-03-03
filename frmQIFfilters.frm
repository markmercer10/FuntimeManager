VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmQIFfilters 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage QIF Filters"
   ClientHeight    =   10320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10320
   ScaleWidth      =   10665
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameAdd 
      Caption         =   "Add Filter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   2400
      TabIndex        =   7
      Top             =   3720
      Visible         =   0   'False
      Width           =   5415
      Begin VB.ListBox ListCat 
         Height          =   840
         Left            =   1560
         TabIndex        =   12
         Top             =   2040
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox M 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1440
         TabIndex        =   0
         Top             =   480
         Width           =   3735
      End
      Begin VB.TextBox P 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1440
         TabIndex        =   1
         Top             =   1080
         Width           =   3735
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
         ItemData        =   "frmQIFfilters.frx":0000
         Left            =   1440
         List            =   "frmQIFfilters.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1680
         Width           =   2535
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
         Left            =   3840
         TabIndex        =   3
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Account :"
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
         TabIndex        =   11
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Memo :"
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
         TabIndex        =   10
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Payee :"
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
         TabIndex        =   9
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
         Left            =   5040
         TabIndex        =   4
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.CommandButton deleteButn 
      Caption         =   "Delete"
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
      Left            =   1920
      TabIndex        =   6
      Top             =   0
      Width           =   1935
   End
   Begin VB.CommandButton addButn 
      Caption         =   "Add"
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
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   1935
   End
   Begin MSComctlLib.ListView ListView 
      Height          =   9855
      Left            =   0
      TabIndex        =   8
      Top             =   480
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   17383
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
         Text            =   "Memo"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Payee"
         Object.Width           =   8820
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Account"
         Object.Width           =   3598
      EndProperty
   End
End
Attribute VB_Name = "frmQIFfilters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub addButn_Click()
    M = ""
    P = ""
    cboAccount.ListIndex = 0
    FrameAdd.Visible = True
    M.SetFocus
End Sub

Private Sub Cancel_Click()
    FrameAdd.Visible = False
End Sub

Private Sub deleteButn_Click()
    If Not (ListView.SelectedItem Is Nothing) Then
        Set q = db.Execute("DELETE FROM qif_filters WHERE id = " & ListView.SelectedItem.Text)
        updateListview
    Else
        MsgBox "No item selected"
    End If
End Sub

Private Sub Form_Load()
    
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
    
    updateListview

End Sub

Private Sub okButn_Click()
    db.Execute "INSERT INTO qif_filters SET M=""" & M & """,P=""" & P & """,Account=""" & ListCat.List(cboAccount.ListIndex) & """"
    FrameAdd.Visible = False
    updateListview

End Sub

Sub updateListview()
    Dim q As ADODB.Recordset
    Dim li As ListItem
    Set q = db.Execute("SELECT * FROM qif_filters ORDER BY M ASC")
    ListView.ListItems.Clear
    With q
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do Until .EOF
                Set li = ListView.ListItems.Add(, , !idQIF)
                li.SubItems(1) = !M
                li.SubItems(2) = !P
                li.SubItems(3) = get_gnc_account_name(!account)
                .MoveNext
            Loop
        End If
    End With
    Set q = Nothing
    Set li = Nothing
End Sub

