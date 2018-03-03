VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmClosures 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "School Closures"
   ClientHeight    =   12735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12735
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameAdd 
      Caption         =   "Add Closure Day"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   1440
      TabIndex        =   3
      Top             =   3840
      Visible         =   0   'False
      Width           =   3855
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
         Left            =   2280
         TabIndex        =   6
         Top             =   1680
         Width           =   1335
      End
      Begin VB.ComboBox cboType 
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
         ItemData        =   "frmClosures.frx":0000
         Left            =   2280
         List            =   "frmClosures.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1080
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtPicker 
         Height          =   375
         Left            =   1560
         TabIndex        =   4
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
         CalendarTitleBackColor=   49152
         CustomFormat    =   "MMM d, yyyy"
         Format          =   109314051
         CurrentDate     =   42718
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
         Left            =   3480
         TabIndex        =   9
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Type :"
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
         TabIndex        =   8
         Top             =   1080
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
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.CommandButton delButn 
      Caption         =   "Delete"
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
      Left            =   1440
      TabIndex        =   2
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton addButn 
      Caption         =   "Add"
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
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView 
      Height          =   12255
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   6420
      _ExtentX        =   11324
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
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "id"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmClosures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub addButn_Click()
    dtPicker.value = Date
    cboType.ListIndex = 0
    FrameAdd.Visible = True
End Sub

Private Sub Cancel_Click()
    FrameAdd.Visible = False
End Sub

Private Sub delButn_Click()
    If Not (ListView.SelectedItem Is Nothing) Then
        Set q = db.Execute("DELETE FROM school_closures WHERE id = " & ListView.SelectedItem.Text)
        updateListview
    Else
        MsgBox "No item selected"
    End If
End Sub

Private Sub Form_Load()
    updateListview
End Sub

Private Sub okButn_Click()
    db.Execute "INSERT INTO school_closures SET date=" & sqlDate(dtPicker.value) & ",type=""" & cboType.Text & """"
    FrameAdd.Visible = False
    updateListview
End Sub

Sub updateListview()
    Dim q As ADODB.Recordset
    Dim li As ListItem
    Set q = db.Execute("SELECT * FROM school_closures ORDER BY date DESC")
    ListView.ListItems.Clear
    With q
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do Until .EOF
                Set li = ListView.ListItems.Add(, , !ID)
                li.SubItems(1) = shortDate(!Date)
                li.SubItems(2) = !Type
                .MoveNext
            Loop
        End If
    End With
    Set q = Nothing
    Set li = Nothing
End Sub
