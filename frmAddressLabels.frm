VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAddressLabels 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Address Labels"
   ClientHeight    =   8796
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   5952
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8796
   ScaleWidth      =   5952
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox collate 
      Caption         =   "Collate"
      Height          =   255
      Left            =   4800
      TabIndex        =   36
      Top             =   4920
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.OptionButton startcell 
      Caption         =   "30"
      Height          =   255
      Index           =   29
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   4560
      Width           =   375
   End
   Begin VB.OptionButton startcell 
      Caption         =   "29"
      Height          =   255
      Index           =   28
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   4320
      Width           =   375
   End
   Begin VB.OptionButton startcell 
      Caption         =   "28"
      Height          =   255
      Index           =   27
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   4080
      Width           =   375
   End
   Begin VB.OptionButton startcell 
      Caption         =   "27"
      Height          =   255
      Index           =   26
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   3840
      Width           =   375
   End
   Begin VB.OptionButton startcell 
      Caption         =   "26"
      Height          =   255
      Index           =   25
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3600
      Width           =   375
   End
   Begin VB.OptionButton startcell 
      Caption         =   "25"
      Height          =   255
      Index           =   24
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3360
      Width           =   375
   End
   Begin VB.OptionButton startcell 
      Caption         =   "24"
      Height          =   255
      Index           =   23
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3120
      Width           =   375
   End
   Begin VB.OptionButton startcell 
      Caption         =   "23"
      Height          =   255
      Index           =   22
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   2880
      Width           =   375
   End
   Begin VB.OptionButton startcell 
      Caption         =   "22"
      Height          =   255
      Index           =   21
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2640
      Width           =   375
   End
   Begin VB.OptionButton startcell 
      Caption         =   "21"
      Height          =   255
      Index           =   20
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2400
      Width           =   375
   End
   Begin VB.OptionButton startcell 
      Caption         =   "20"
      Height          =   255
      Index           =   19
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4560
      Width           =   375
   End
   Begin VB.OptionButton startcell 
      Caption         =   "19"
      Height          =   255
      Index           =   18
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4320
      Width           =   375
   End
   Begin VB.OptionButton startcell 
      Caption         =   "18"
      Height          =   255
      Index           =   17
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4080
      Width           =   375
   End
   Begin VB.OptionButton startcell 
      Caption         =   "17"
      Height          =   255
      Index           =   16
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3840
      Width           =   375
   End
   Begin VB.OptionButton startcell 
      Caption         =   "16"
      Height          =   255
      Index           =   15
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3600
      Width           =   375
   End
   Begin VB.OptionButton startcell 
      Caption         =   "15"
      Height          =   255
      Index           =   14
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3360
      Width           =   375
   End
   Begin VB.OptionButton startcell 
      Caption         =   "14"
      Height          =   255
      Index           =   13
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3120
      Width           =   375
   End
   Begin VB.OptionButton startcell 
      Caption         =   "13"
      Height          =   255
      Index           =   12
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2880
      Width           =   375
   End
   Begin VB.OptionButton startcell 
      Caption         =   "12"
      Height          =   255
      Index           =   11
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2640
      Width           =   375
   End
   Begin VB.OptionButton startcell 
      Caption         =   "11"
      Height          =   255
      Index           =   10
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2400
      Width           =   375
   End
   Begin VB.OptionButton startcell 
      Caption         =   "10"
      Height          =   255
      Index           =   9
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4560
      Width           =   375
   End
   Begin VB.OptionButton startcell 
      Caption         =   "9"
      Height          =   255
      Index           =   8
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4320
      Width           =   375
   End
   Begin VB.OptionButton startcell 
      Caption         =   "8"
      Height          =   255
      Index           =   7
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4080
      Width           =   375
   End
   Begin VB.OptionButton startcell 
      Caption         =   "7"
      Height          =   255
      Index           =   6
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3840
      Width           =   375
   End
   Begin VB.OptionButton startcell 
      Caption         =   "6"
      Height          =   255
      Index           =   5
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3600
      Width           =   375
   End
   Begin VB.OptionButton startcell 
      Caption         =   "5"
      Height          =   255
      Index           =   4
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3360
      Width           =   375
   End
   Begin VB.OptionButton startcell 
      Caption         =   "4"
      Height          =   255
      Index           =   3
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3120
      Width           =   375
   End
   Begin VB.OptionButton startcell 
      Caption         =   "3"
      Height          =   255
      Index           =   2
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   375
   End
   Begin VB.OptionButton startcell 
      Caption         =   "2"
      Height          =   255
      Index           =   1
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2640
      Width           =   375
   End
   Begin VB.OptionButton startcell 
      Caption         =   "1"
      Height          =   255
      Index           =   0
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2400
      Value           =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton CloseButn 
      Caption         =   "Close"
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   8280
      Width           =   1095
   End
   Begin VB.CommandButton sortButn 
      Caption         =   "Alphabetical"
      Height          =   495
      Left            =   4800
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton clearButn 
      Caption         =   "Uncheck all"
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton printButn 
      Caption         =   "Print"
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListCustomers 
      Height          =   8775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8276
      _ExtentY        =   15473
      SortKey         =   1
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Acct#"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Customer Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Purchases"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   4800
      X2              =   5880
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "Printing Start Cell"
      Height          =   375
      Left            =   4800
      TabIndex        =   35
      Top             =   1920
      Width           =   855
   End
End
Attribute VB_Name = "frmAddressLabels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strtcell As Long
Private Sub clearButn_Click()
    Dim i As Long
    For i = 1 To ListCustomers.ListItems.count
        ListCustomers.ListItems(i).Checked = False
    Next i
End Sub

Private Sub closebutn_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim L As ListItem
    Dim r As Recordset
    Dim cName As String
    Dim purch As Double
    
    strtcell = 1
    For i = 0 To 9
        Set r = dbm.sqlQuery("SELECT * FROM Customers WHERE CustomerName = """ & Mid$(frmBestCustomers.cName(i), 1, InStr(1, frmBestCustomers.cName(i), " - (")) & """")
        r.MoveFirst
        Set L = ListCustomers.ListItems.Add(, , r!accountnumber)
        L.SubItems(1) = r!customername
        L.SubItems(2) = frmBestCustomers.Purchases(i)
        L.Checked = True
    Next i
    With frmBestCustomers.List1
        For i = 0 To .ListCount
            cName = .List(i)
            cName = Mid$(cName, InStr(1, cName, ". - ") + 4)
            cName = Mid$(cName, 1, InStr(1, cName, " - ("))
            cName = Replace(cName, """", """""")
            purch = custSales(i + 11)
            'MsgBox cName & "  " & purch
            'MsgBox Trim(Format(custSales(i + 10), "0.00"))
            'MsgBox Mid$(cName, InStr(1, cName, "  -  $") + 6)
            Set r = dbm.sqlQuery("SELECT * FROM Customers WHERE CustomerName = """ & cName & """")
            If Not (r.EOF And r.BOF) Then
                r.MoveFirst
                Set L = ListCustomers.ListItems.Add(, , r!accountnumber)
                L.SubItems(1) = r!customername
                L.SubItems(2) = Format(purch, "0.00")
                If purch > 0 Then L.Checked = True
            End If
        Next i
    End With
    
    Set r = Nothing
    Set L = Nothing
End Sub


Private Sub printButn_Click()
    Dim r As Recordset
    Dim i As Long
    Dim count As Long
    Dim cell_x As Long
    Dim cell_y As Long
    Dim space_x As Double
    Dim space_y As Double
    Dim addressleft As Double
    
    store_current_printer
    selectPrinter sync!PrimaryPrinter
    
    Printer.ScaleMode = vbInches
    Printer.Font = "Arial"
    Printer.fontsize = 10
    Printer.FontBold = False
    
    space_x = 2.75
    space_y = 1
    
    count = strtcell
    
    With ListCustomers
        For i = 1 To .ListItems.count
            If .ListItems(i).Checked Then
                Set r = dbm.sqlSimpleSelect("Customers", "accountnumber", val(.ListItems(i)))
                
                cell_y = ((count - 1) Mod 10) + 1
                cell_x = (((((count - 1) \ 10) + 1) - 1) Mod 3) + 1
                
                addressleft = space_x * cell_x - 2.5
                Printer.CurrentY = space_y * cell_y - 0.45
                Printer.CurrentX = addressleft
                Printer.Print r!customername
                
                Printer.CurrentX = addressleft
                Printer.Print r!Address
                
                Printer.CurrentX = addressleft
                Printer.Print r!city & ", " & r!province
                
                Printer.CurrentX = addressleft
                Printer.Print left$(r!postalcode, 3) & " " & right$(r!postalcode, 3)
                
                If count Mod 30 = 0 Then
                    If collate.value = 1 Then
                        Printer.NewPage
                    Else
                        Printer.EndDoc
                    End If
                End If
                count = count + 1
            End If
        Next i
    End With
    
    Printer.EndDoc
    
    Printer.ScaleMode = vbTwips
    Set r = Nothing

    recall_stored_printer

End Sub


Private Sub sortButn_Click()
    ListCustomers.sorted = True
End Sub


Private Sub startcell_Click(index As Integer)
    strtcell = index + 1
End Sub


