VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmChart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Income / Expenses Chart"
   ClientHeight    =   12510
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   18765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12510
   ScaleWidth      =   18765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton printButn 
      Height          =   495
      Left            =   10440
      Picture         =   "frmChart.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   11520
      Width           =   615
   End
   Begin VB.ComboBox cboYear 
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
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   12120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   9495
      Left            =   5280
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   7935
      Begin MSChart20Lib.MSChart PieChart 
         Height          =   9255
         Left            =   120
         OleObjectBlob   =   "frmChart.frx":2ED7
         TabIndex        =   8
         Top             =   120
         Width           =   7695
      End
   End
   Begin VB.CheckBox chkExpCats 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Show Expense Categories"
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
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   12000
      Width           =   3615
   End
   Begin VB.CheckBox chkExp 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Include Expenses"
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
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   11520
      Value           =   1  'Checked
      Width           =   3615
   End
   Begin VB.CheckBox chkIncCats 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Show Income Categories"
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
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   12000
      Width           =   3615
   End
   Begin VB.CheckBox chkInc 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Include Income"
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
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   11520
      Value           =   1  'Checked
      Width           =   3615
   End
   Begin VB.OptionButton optYearly 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Yearly"
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
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   11520
      Width           =   1095
   End
   Begin VB.OptionButton optMonthly 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Monthly"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   11520
      Value           =   -1  'True
      Width           =   1095
   End
   Begin MSChart20Lib.MSChart Chart 
      Height          =   11415
      Left            =   0
      OleObjectBlob   =   "frmChart.frx":682D
      TabIndex        =   0
      Top             =   0
      Width           =   18735
   End
End
Attribute VB_Name = "frmChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ARRAY_SIZE = 1024
Dim account_to_column() As String
Dim ROOT_ACCOUNT_GUID As String
Dim INC_GUID As String
Dim EXP_GUID As String




Private Sub cboYear_Change()
    cboYear_Click
End Sub

Private Sub cboYear_Click()
    draw_chart
End Sub

Private Sub Chart_PointSelected(Series As Integer, DataPoint As Integer, MouseFlags As Integer, Cancel As Integer)
    Chart.row = DataPoint
    MsgBox "Show button for pie chart for " & Chart.RowLabel
End Sub

Private Sub chkExp_Click()
    If chkExp.value = 1 Then
        chkExpCats.Enabled = True
    Else
        chkExpCats.Enabled = False
    End If
    draw_chart
End Sub

Private Sub chkExpCats_Click()
    draw_chart
End Sub

Private Sub chkInc_Click()
    If chkInc.value = 1 Then
        chkIncCats.Enabled = True
    Else
        chkIncCats.Enabled = False
    End If
    draw_chart
End Sub

Private Sub chkIncCats_Click()
    draw_chart
End Sub

Private Sub Form_Load()
    For y = 2016 To year(Date)
        cboYear.AddItem y
    Next y
    cboYear.ListIndex = cboYear.ListCount - 1
    
    Chart.Legend.Location.LocationType = VtChLocationTypeBottom
    PieChart.Legend.Location.LocationType = VtChLocationTypeBottom
    
    draw_chart
End Sub

Sub draw_chart()
    Dim inc_index As Long
    Dim exp_index As Long
    Dim temp_index As Long
    Dim account_count As Long
    Dim exp_account As Long
    Dim account As Long
    Dim column_count As Long
    Dim q As ADODB.Recordset
    Dim intervals() As Date
    Dim i As Long
    Dim c As Long
    Dim item_guid As String
    Dim parent_guid As String
    
    Dim temp As String
    
    ReDim account_to_column(ARRAY_SIZE, 2) As String
    Chart.RowCount = 0
    Chart.ColumnCount = 0
    
    cboYear.Visible = Not CBool(optYearly.value)
    
    ROOT_ACCOUNT_GUID = get_gnc_account_guid("Root Account")
    INC_GUID = get_gnc_account_guid("Income")
    EXP_GUID = get_gnc_account_guid("Expenses")
    inc_index = hash_guid(INC_GUID)
    exp_index = hash_guid(EXP_GUID)
    
    
    
    account = 1
    column_count = 0
    If chkInc.value = 1 Then
        account_to_column(inc_index, 1) = INC_GUID
        account_to_column(inc_index, 2) = account ' income is column 1
        account = account + 1
        column_count = column_count + 1
    End If
    
    Set q = gcdb.Execute("SELECT * FROM accounts WHERE parent_guid = """ & INC_GUID & """ ORDER BY name")
    If Not (q.EOF And q.BOF) Then
        q.MoveFirst
        Do Until q.EOF
            temp_index = hash_guid(q!guid)
            account_to_column(temp_index, 1) = q!guid
            If chkIncCats.value = 1 And chkInc.value = 1 Then
                account_to_column(temp_index, 2) = account ' if we are showing other income accounts number them following 1
                column_count = column_count + 1
            Else
                account_to_column(temp_index, 2) = account_to_column(inc_index, 2) ' if we are not showing other income accounts refer them to 1
            End If
            account = account + 1
            q.MoveNext
        Loop
    End If
    
    If chkExp.value = 1 Then ' if showing expenses
        account_to_column(exp_index, 1) = EXP_GUID
        account_to_column(exp_index, 2) = column_count + 1
        account = account + 1
        column_count = column_count + 1
    End If
    
    Set q = gcdb.Execute("SELECT * FROM accounts WHERE parent_guid = """ & EXP_GUID & """ ORDER BY name")
    If Not (q.EOF And q.BOF) Then
        q.MoveFirst
        Do Until q.EOF
            temp_index = hash_guid(q!guid)
            account_to_column(temp_index, 1) = q!guid
            If chkExpCats.value = 1 And chkExp.value = 1 Then
                column_count = column_count + 1
                account_to_column(temp_index, 2) = column_count ' if we are showing other expense accounts number them following the base expense account
            Else
                account_to_column(account, 2) = account_to_column(exp_index, 2) ' if we are not showing other expense accounts refer them to the base expense account
            End If
            account = account + 1
            q.MoveNext
        Loop
    End If

    With Chart
        Const variance = 150
        Const min = 40
        
        .ColumnCount = column_count
        If optMonthly Then
            .RowCount = 12
            ReDim intervals(0 To 12) As Date
            For i = 0 To 11
                'intervals(i) = CDate(year(Date) & "-" & Format(i + 1, "00") & "-01")
                intervals(i) = CDate(cboYear.Text & "-" & Format(i + 1, "00") & "-01")
                Chart.row = i + 1
                Chart.RowLabel = MonthName(i + 1)
            Next i
            intervals(12) = CDate(year(Date) + 1 & "-01-01")
        Else
            .RowCount = year(Date) - year(EPOCH) + 1
            ReDim intervals(0 To year(Date) - year(EPOCH) + 1) As Date
            For i = 0 To year(Date) - year(EPOCH) + 1
                intervals(i) = CDate((year(EPOCH) + i) & "-01-01")
                If i > 0 Then
                    Chart.row = i
                    Chart.RowLabel = str$(year(EPOCH) + i - 1)
                End If
            Next i
        End If
        
        Rnd (-1)
        Randomize 107
        For account = ARRAY_SIZE To 1 Step -1
            If account_to_column(account, 2) <> "" Then
                .column = val(account_to_column(account, 2))
                .ColumnLabel = get_gnc_account_name(account_to_column(account, 1))
                If get_gnc_account_parent(account_to_column(account, 1)) = INC_GUID Then ' THIS IS AN INCOME CATEGORY      (red)
                    .Plot.SeriesCollection(val(account_to_column(account, 2))).Pen.VtColor.Set 255, Int(Rnd * variance) + min + 40, Int(Rnd * variance) + min
                Else                                                                     ' THIS IS AN EXPENSE CATEGORY     (blue)
                    .Plot.SeriesCollection(val(account_to_column(account, 2))).Pen.VtColor.Set Int(Rnd * 128), Int(Rnd * variance) + min + 50, Int(Rnd * variance) + min + 50
                    '.Plot.SeriesCollection(val(account_to_column(account, 2))).Pen.VtColor.Set Int(Rnd * variance) + min, Int(Rnd * variance) + min + 50, 255
                End If
            End If
        Next account
        If chkInc.value = 1 Then
            'If chkIncCats.value = 0 Then
            .column = val(account_to_column(inc_index, 2))
            .ColumnLabel = "Total Income"
        End If
        If chkInc.value = 1 Then .Plot.SeriesCollection(val(account_to_column(inc_index, 2))).Pen.VtColor.Set 255, 0, 0
        If chkExp.value = 1 Then
            'If chkExpCats.value = 0 Then
            .column = val(account_to_column(exp_index, 2))
            .ColumnLabel = "Total Expenses"
        End If
        If chkExp.value = 1 Then .Plot.SeriesCollection(val(account_to_column(exp_index, 2))).Pen.VtColor.Set 0, 0, 255
    
        
        
        For i = 1 To UBound(intervals)
            Chart.row = i
            If chkInc.value = 1 Then Chart.column = account_to_column(inc_index, 2): Chart.Data = 0
            If chkExp.value = 1 Then Chart.column = account_to_column(exp_index, 2): Chart.Data = 0
            
            Set q = gcdb.Execute("SELECT account_guid, (SUM(value_num) / value_denom) AS amount FROM gnucash.splits INNER JOIN gnucash.transactions ON transactions.guid = splits.tx_guid WHERE post_date >= " & sqlDate(intervals(i - 1)) & " AND post_date < " & sqlDate(intervals(i)) & " GROUP BY account_guid;")
            If Not (q.EOF And q.BOF) Then
                q.MoveFirst
                Do Until q.EOF
                    'temp = q!Description
                    item_guid = get_column(q!account_guid)
                    parent_guid = get_gnc_account_parent(item_guid)
                    If parent_guid = INC_GUID And chkInc.value = 1 Then
                        If chkIncCats.value = 1 Then
                            Chart.column = account_to_column(hash_guid(item_guid), 2)
                            Chart.Data = -q!amount
                            Chart.column = account_to_column(inc_index, 2)
                            Chart.Data = Chart.Data - q!amount
                        Else
                            'Chart.column = account_to_column(hash_guid(parent_guid), 2)
                            Chart.column = account_to_column(inc_index, 2)
                            Chart.Data = Chart.Data - q!amount
                        End If
                    ElseIf parent_guid = EXP_GUID And chkExp.value = 1 Then
                        If chkExpCats.value = 1 Then
                            Chart.column = account_to_column(hash_guid(item_guid), 2)
                            Chart.Data = q!amount
                            Chart.column = account_to_column(exp_index, 2)
                            Chart.Data = Chart.Data + q!amount
                        Else
                            'Chart.column = account_to_column(hash_guid(parent_guid), 2)
                            Chart.column = account_to_column(exp_index, 2)
                            Chart.Data = Chart.Data + q!amount
                        End If
                    End If
                    'If q!amount = -5290 Then MsgBox q!account_guid
                    'MsgBox q!account_guid & " --> " & Chart.column
                    q.MoveNext
                Loop
            End If
        Next i
        
    End With
    
    Set q = Nothing
End Sub

Public Function hash_guid(ByVal guid As String) As Long
    Dim bump As Long
    bump = 1
    hash_guid = (CLng("&H" & MiD$(guid, 1, 4)) / 128 + CLng("&H" & MiD$(guid, 9, 4)) / 256 + CLng("&H" & MiD$(guid, 21, 4)) / 512 + CLng("&H" & MiD$(guid, 29, 4)) / 1024) Mod ARRAY_SIZE
    Do
        If account_to_column(hash_guid, 1) = guid Or account_to_column(hash_guid, 1) = "" Then
            'do nothing, we have the right index
            Exit Do
        Else
            hash_guid = (hash_guid + bump) Mod ARRAY_SIZE
            bump = bump + 2
        End If
    Loop
End Function

Function get_column(ByVal guid As String) As String
    'returns the GUID of the column account
    Dim temp As String
    If guid = INC_GUID Or guid = EXP_GUID Then
        get_column = ""
    ElseIf guid = ROOT_ACCOUNT_GUID Then
        get_column = ""
    Else
        'If top_level Then
            temp = guid
        '    Do
        '        temp = get_gnc_account_parent(temp)
        '    Loop Until temp = INC_GUID Or temp = EXP_GUID Or temp = ROOT_ACCOUNT_GUID
        '    If temp = ROOT_ACCOUNT_GUID Then
        '        get_column = ""
        '    Else
        '        get_column = temp
        '    End If
        'Else
            Do
                get_column = temp
                temp = get_gnc_account_parent(temp)
            Loop Until temp = INC_GUID Or temp = EXP_GUID Or temp = ROOT_ACCOUNT_GUID Or temp = ""
        'End If
    End If
End Function

Private Sub optMonthly_Click()
    draw_chart
End Sub

Private Sub optYearly_Click()
    draw_chart
End Sub

Private Sub printButn_Click()
    If optMonthly.value = True Then
        printText "Income / Expenses " & cboYear.Text, 500, 1, 10000, "Arial", 16, True, 0
    Else
        printText "Yearly Income / Expenses ", 500, 1, 10000, "Arial", 16, True, 0
    End If
    printLineGraph Printer, Chart, Printer.width * 0.05, 1000, Printer.width * 0.85, Chart.height / Chart.width * Printer.width
    Dim r As Long
    Dim x As Long
    Dim y As Long
    Dim h As Long
    Dim offset As Long
    Dim linecolor As Long
    offset = 9000
    h = 350
    Printer.Line (500, offset - 150)-(11000, offset - 150), vbBlack
    For r = 1 To Chart.ColumnCount
        Chart.column = r
        x = 500
        If r >= CLng((Chart.ColumnCount / 2) + 1) Then x = 5500
        y = ((r - 1) Mod (CLng(Chart.ColumnCount / 2 + 0.1))) * h + offset
        linecolor = CLng(Chart.Plot.SeriesCollection(r).Pen.VtColor.Red) + CLng(Chart.Plot.SeriesCollection(r).Pen.VtColor.Green) * 256 + CLng(Chart.Plot.SeriesCollection(r).Pen.VtColor.Blue) * 65536
        Printer.Line (x, y)-(x + 280, y + 160), linecolor, BF
        Printer.Line (x, y)-(x + 280, y + 160), vbBlack, B
        printText Chart.ColumnLabel, x + 450, y - 40, 6000, "Arial", 12, False, 0
    Next r
    Printer.EndDoc
End Sub
