VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form dlgCheque 
   BackColor       =   &H00E9D6D3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Printable Cheque"
   ClientHeight    =   3915
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   11805
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox chqNum 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   490
      Left            =   9120
      TabIndex        =   3
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton autoMemo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E9D6D3&
      Caption         =   "Auto"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3120
      Width           =   855
   End
   Begin MSComctlLib.ListView memoList 
      Height          =   1215
      Left            =   2760
      TabIndex        =   21
      Top             =   3120
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2143
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Text"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Score"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox TxtSubTot 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9840
      TabIndex        =   2
      Text            =   "0.00"
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox TxtHST 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9840
      TabIndex        =   1
      Text            =   "0.00"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2040
      Top             =   360
   End
   Begin VB.TextBox txtVd 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   90
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   2400
      TabIndex        =   17
      Text            =   "VOID"
      Top             =   1080
      Visible         =   0   'False
      Width           =   7095
   End
   Begin VB.CheckBox CheckVd 
      Appearance      =   0  'Flat
      BackColor       =   &H00E9D6D3&
      Caption         =   "Void"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox TxtMemo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1440
      TabIndex        =   7
      Top             =   3120
      Width           =   5055
   End
   Begin VB.CommandButton CancelButn 
      Appearance      =   0  'Flat
      BackColor       =   &H00E9D6D3&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton SaveButn 
      Appearance      =   0  'Flat
      BackColor       =   &H00E9D6D3&
      Caption         =   "Save && Close"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox TxtCents 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9960
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "xx"
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox TxtEnglishNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFE2DD&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1920
      Width           =   8415
   End
   Begin VB.TextBox TxtAMT 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9840
      TabIndex        =   0
      Text            =   "0.00"
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox TxtPayTo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1440
      TabIndex        =   5
      Top             =   1440
      Width           =   6975
   End
   Begin MSComCtl2.DTPicker DTDate 
      Height          =   405
      Left            =   5640
      TabIndex        =   4
      Top             =   840
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   714
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
      CalendarBackColor=   16772829
      CalendarTitleBackColor=   16744544
      CalendarTitleForeColor=   -2147483634
      CalendarTrailingForeColor=   8421504
      CustomFormat    =   "MMMM dd, yyy"
      Format          =   158924803
      CurrentDate     =   38337
   End
   Begin VB.ComboBox ComboCategory 
      Appearance      =   0  'Flat
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
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2400
      Width           =   5055
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   24
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pay To"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   480
      TabIndex        =   23
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque Number:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   22
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Memo"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   20
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "HST"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9120
      TabIndex        =   19
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Subtotal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   18
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   16
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "/100"
      Height          =   255
      Left            =   10200
      TabIndex        =   13
      Top             =   2085
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9600
      TabIndex        =   12
      Top             =   1440
      Width           =   255
   End
End
Attribute VB_Name = "dlgCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim keys(200) As String
Dim EXP_GUID As String

Dim allowChange As Boolean

Public newCheque As Boolean
Public payableExists As Boolean
Public payableGUID As String

Public num As Long
Public dt As Date
Public PTTOO As String
Public AMT As Double
Public mMO As String
Public Vd As Boolean

Function getIndexFromKey(key As String) As Long
    Dim i As Long
    getIndexFromKey = 26
    For i = 0 To 26
        If keys(i) = key Then
            getIndexFromKey = i
        End If
    Next i
End Function

Private Sub cancelButn_Click()
    Unload Me
End Sub

Private Sub CheckVd_Click()
    If CheckVd.value = 1 Then
        txtVd.Visible = True
    Else
        txtVd.Visible = False
    End If
End Sub

Private Sub chqNum_Change()
    checkSaveButn
End Sub

Private Sub chqNum_KeyUp(KeyCode As Integer, Shift As Integer)
    checkSaveButn
End Sub

Private Sub ComboCategory_Click()
    checkSaveButn
End Sub


Private Sub ComboCategory_DropDown()
    allowChange = False
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim found As Boolean
    Dim keyindex As Long
    
    allowChange = True
    autoMemo.Enabled = False
    
    If newCheque Then
        Me.Caption = "New Printable Cheque"
        DTDate.value = Date
        'chqNum = num
    Else
        Me.Caption = "Edit Printable Cheque"
        chqNum.Locked = True
        chqNum.backcolor = TxtEnglishNumber.backcolor
    End If
    
    'For i = 0 To frmCheques.RecipList.ListCount - 1
    '    Select_Recipient.AddItem frmCheques.RecipList.List(i)
    'Next i

    keyindex = 1
    EXP_GUID = get_gnc_account_guid("Expenses")
    Dim exp As ADODB.Recordset
    Set exp = gcdb.Execute("SELECT * FROM accounts WHERE parent_guid = """ & EXP_GUID & """ ORDER BY name ASC")
    
    With exp
        If Not (.EOF And .BOF) Then .MoveFirst
        Do Until .EOF
            ComboCategory.AddItem !name
            keys(keyindex) = !guid
            keyindex = keyindex + 1
            
            Dim exp2 As ADODB.Recordset
            Set exp2 = gcdb.Execute("SELECT * FROM accounts WHERE parent_guid = """ & !guid & """ ORDER BY name ASC")
            
            With exp2
                If Not (.EOF And .BOF) Then .MoveFirst
                Do Until .EOF
                    ComboCategory.AddItem " -- " & !name
                    keys(keyindex) = !guid
                    keyindex = keyindex + 1
                    
                    .MoveNext
                Loop
            End With

            
            .MoveNext
        Loop
    End With
End Sub

Private Sub SaveButn_Click()
    Dim sql As String
    
    If newCheque Then
        sql = "INSERT INTO cheques (chqNumber, date, payTo, memo, amount, void) VALUES ("
        sql = sql & chqNum & ","
        sql = sql & sqlDate(DTDate.value) & ","
        sql = sql & """" & TxtPayTo & ""","
        sql = sql & """" & TxtMemo & ""","
        sql = sql & TxtAMT & ","
        sql = sql & CheckVd.value & ")"
        db.Execute sql
        
        create_gnc_transaction createGUID, keys(ComboCategory.ListIndex), chqNum, TxtPayTo & " " & TxtMemo, val(TxtAMT), DTDate.value
    Else
        sql = "UPDATE cheques SET "
        sql = sql & "date=" & sqlDate(DTDate.value) & ","
        sql = sql & "payTo=""" & TxtPayTo & ""","
        sql = sql & "memo=""" & TxtMemo & ""","
        sql = sql & "amount=" & TxtAMT & ","
        sql = sql & "void=" & CheckVd.value
        sql = sql & " WHERE chqNumber=" & chqNum
        db.Execute sql
        
        update_gnc_transaction payableGUID, keys(ComboCategory.ListIndex), chqNum, TxtPayTo & " " & TxtMemo, val(TxtAMT), DTDate.value
    End If
    
    Unload Me
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    
    If Not newCheque Then
        chqNum = num
        CheckVd.value = -CLng(Vd) ' use cBool(checkvd.value) to convert back
        DTDate.value = dt
        TxtPayTo = PTTOO
        TxtAMT = Format(AMT, "$0.00")
        TxtMemo = mMO
        Dim cheque As ADODB.Recordset
        Set cheque = db.Execute("SELECT * FROM cheques WHERE chqNumber = " & num)
        
        With cheque
            If .EOF And .BOF Then
                payableExists = False
                'BillingDate.value = dt
                TxtHST = "-"
                TxtSubTot = "-"
            Else
                .MoveFirst
                payableExists = True
                DTDate.value = !Date
                'TxtHST = Format(!HST, "0.00")
                'TxtSubTot = Format(!SubTot, "0.00")
                TxtPayTo = !payto
                TxtAMT = Format(!amount, "0.00")
                TxtMemo = !Memo
                CheckVd.value = !Void
            
                Dim tx As ADODB.Recordset
                Set tx = gcdb.Execute("SELECT transactions.guid, account_guid FROM transactions INNER JOIN splits ON (transactions.guid = splits.tx_guid) WHERE num = " & num & " AND DATE_FORMAT(enter_date, '%Y-%m-%d') = DATE_FORMAT(" & sqlDate(!Date) & ", '%Y-%m-%d') AND account_guid != """ & rec_account_guid & """;")
                With tx
                    If Not (.EOF And .BOF) Then
                        .MoveFirst
                        SaveButn.Tag = !guid
                        Dim i As Long
                        For i = 0 To ComboCategory.ListCount - 1
                            If keys(i) = !account_guid Then
                                ComboCategory.ListIndex = i
                            End If
                        Next i
                    End If
                End With
            End If
        End With
    End If
    
    
    TxtAMT.SetFocus
    TxtAMT.SelStart = 0
    TxtAMT.SelLength = Len(TxtAMT)
    
    Set cheque = Nothing
    Set tx = Nothing
End Sub

Private Sub TxtAMT_Change()
    TxtEnglishNumber = englishNumber(val(TxtAMT))
    If val(TxtAMT) = 0 Then
        TxtCents = "xx"
    Else
        If InStr(1, TxtAMT, ".") = 0 Then
            TxtCents = "xx"
        Else
            TxtCents = Right$(Format(val(TxtAMT), "0.00"), 2)
        End If
    End If
    
TxtSubTot.Text = Format(val(TxtAMT.Text) / (1# + HST_RATE), "0.00")
TxtHST.Text = Format(val(TxtAMT.Text) - val(TxtSubTot.Text), "0.00")

checkSaveButn
End Sub

Private Sub txtHST_Change()
TxtSubTot.Text = Format(val(TxtAMT.Text) - val(TxtHST.Text), "0.00")
End Sub

Sub checkSaveButn()
    
    If Trim(TxtAMT.Text) <> "" And val(TxtAMT.Text) > 0 And ComboCategory.ListIndex > -1 And val(chqNum.Text) > 0 Then
        If Not newCheque Then
            SaveButn.Enabled = True
        Else
            'check if chqNum is already in the DB.
            Dim ch As ADODB.Recordset
            Set ch = db.Execute("SELECT * FROM cheques WHERE chqNumber = " & chqNum)
            If (ch.EOF And ch.BOF) Then
                SaveButn.Enabled = True
                chqNum.forecolor = vbBlack
                chqNum.ToolTipText = ""
            Else
                SaveButn.Enabled = False
                chqNum.forecolor = vbRed
                chqNum.ToolTipText = "Duplicate cheque number!"
            End If
        End If
    Else
        SaveButn.Enabled = False
    End If
End Sub
