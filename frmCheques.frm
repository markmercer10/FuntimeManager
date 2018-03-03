VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCheques 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Printable Cheque List"
   ClientHeight    =   9465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9465
   ScaleWidth      =   8700
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   3
      Top             =   8760
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   4
      Top             =   8760
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   9
      Top             =   8760
      Width           =   1095
   End
   Begin VB.OptionButton opAll 
      Caption         =   "List All Cheques"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   8
      Top             =   300
      Width           =   1575
   End
   Begin VB.OptionButton opThisYear 
      Caption         =   "This Year Only"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   7
      Top             =   60
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Print List"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   6
      Top             =   8760
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   6240
      Top             =   120
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print Selected Cheques"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   8760
      Width           =   1455
   End
   Begin VB.CommandButton okButn 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      TabIndex        =   0
      Top             =   8760
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView 
      Height          =   8025
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   14155
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
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
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Chq Num"
         Object.Width           =   1765
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   2294
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Pay To"
         Object.Width           =   3706
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Amount"
         Object.Width           =   1765
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Memo"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Void"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.Line Line2 
      X1              =   7080
      X2              =   7080
      Y1              =   8760
      Y2              =   9360
   End
   Begin VB.Line Line3 
      X1              =   60
      X2              =   8640
      Y1              =   8640
      Y2              =   8640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   60
      X2              =   8640
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      Caption         =   "Cheques"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   8535
   End
End
Attribute VB_Name = "frmCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub updateListview()
    Dim i As ListItem
    Dim n As Long
    Dim Add As Boolean
    ListView.ListItems.Clear
    ListView.Sorted = False
    DoEvents
    Dim cheques As ADODB.Recordset
    Set cheques = db.Execute("SELECT * FROM cheques ORDER BY chqNumber DESC")
    With cheques
        If Not .EOF Or Not .BOF Then
        .MoveFirst
        Do Until .EOF
            Add = True
            If (opThisYear.value And year(!Date) <> year(Date)) Then Add = False
            If Add Then
                Set i = ListView.ListItems.Add(, , !chqNumber)
                i.SubItems(1) = shortDate(!Date)
                i.SubItems(2) = !payto
                i.SubItems(3) = Format(!amount, "0.00")
                i.SubItems(4) = !Memo
                If !Void Then
                    i.SubItems(5) = "void"
                Else
                    i.SubItems(5) = ""
                End If
            End If
            .MoveNext
        Loop
        End If
    End With
    'DoEvents
    'ListView.SortKey = 0
    'ListView.Sorted = True
    'DoEvents
    Set i = Nothing
    ListView.SetFocus
End Sub

Private Sub cancelButn_Click()
Unload Me
End Sub

Private Sub Command1_Click()
    dlgCheque.newCheque = True
    dlgCheque.num = getLastChqNum + 1
    dlgCheque.Show 1
    updateListview
End Sub

Private Sub Command2_Click()
    'EDIT
    If ListView.ListItems.count = 0 Then Exit Sub
    dlgCheque.newCheque = False
    
    dlgCheque.num = val(ListView.SelectedItem.Text)
    dlgCheque.dt = CDate(ListView.SelectedItem.SubItems(1))
    dlgCheque.PTTOO = ListView.SelectedItem.SubItems(2)
    dlgCheque.AMT = val(ListView.SelectedItem.SubItems(3))
    dlgCheque.mMO = ListView.SelectedItem.SubItems(4)
    'dlgCheque.Vd = -CLng(ListView.SelectedItem.SubItems(5))
    If ListView.SelectedItem.SubItems(5) = "void" Then
        dlgCheque.Vd = True
    Else
        dlgCheque.Vd = False
    End If
        
    dlgCheque.Show 1
    updateListview
End Sub


Private Sub Command3_Click()
    Dim yTab As Byte
    Dim yTabMax As Long
    Dim stublistindex As Long
    Dim Total As Double
    okButn.Enabled = False
    
    MsgBox "Please load cheques into " & Printer.DeviceName
    
    'Exit Sub
    yTab = 1
    yTabMax = Int(val(InputBox("Enter number of cheques on each page")))
    If yTabMax < 0 Then yTabMax = -yTabMax
    If yTabMax < 1 Then yTabMax = 1
    If yTabMax > 3 Then yTabMax = 3
    
    If yTabMax = 1 Then
        Printer.Orientation = 2
        yTab = 4
    Else
        Printer.Orientation = 1
    End If
    
    Total = 0
    For stublistindex = ListView.ListItems.count To 1 Step -1
        If ListView.ListItems(stublistindex).Checked Then
            'printCheque_old val(ListView.ListItems(stublistindex)), yTab
            printCheque val(ListView.ListItems(stublistindex)), yTab, yTabMax
            Total = Total + val(ListView.ListItems(stublistindex).ListSubItems(3))
            yTab = yTab + 1
            If yTab > yTabMax Then yTab = 1: Printer.EndDoc
            
            If yTabMax = 1 Then
                Printer.Orientation = 2
                yTab = 4
            End If
        End If
    Next stublistindex
    Printer.EndDoc
    MsgBox "The total of all cheques being printed is $" & Total
    'dlgSupplierEnvelopes.Show 1
    
    updateListview
    okButn.Enabled = True

End Sub

Private Sub Command4_Click()
    Dim i As Long
    For i = ListView.ListItems.count To 1 Step -1
        'If i < ListView.ListItems.count Then Exit For
        If Not ListView.ListItems(i).Checked Then
            ListView.ListItems.Remove i
            'i = i - 1
            DoEvents
        End If
    Next i
    printListView ListView, 60, 100, 1000, 1.35, True
    Printer.EndDoc
    updateListview
End Sub

Private Sub filter_Click()
updateListview
End Sub

Private Sub Command5_Click()
    'DELETE
    If ListView.ListItems.count = 0 Then Exit Sub
    Dim n As Long
    n = val(ListView.SelectedItem.Text)
    If MsgBox("Are you sure you want to delete cheque number " & n & "?", vbYesNo) = vbYes Then
        'delete the gnc record
        Dim ch As ADODB.Recordset
        Set ch = db.Execute("SELECT * FROM cheques WHERE chqNumber = " & n)
        If ch.EOF And ch.BOF Then delete_gnc_transaction (ch!guid)
        Set ch = Nothing
        
        'delete the cheque record
        db.Execute "DELETE FROM cheques WHERE chqNumber = " & n
    End If
    
    updateListview
End Sub

Private Sub ListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Dim i As Long
    If ColumnHeader.Index = 1 Then
        For i = 1 To ListView.ListItems.count
            With ListView.ListItems(i)
                If .Checked Then
                    .Checked = False
                Else
                    .Checked = True
                End If
            End With
        Next i
    End If
End Sub


Private Sub okButn_Click()
Unload Me
End Sub


Private Sub opAll_Click()
    updateListview
End Sub

Private Sub opThisYear_Click()
    updateListview
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    updateListview
End Sub


