VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBudget 
   Caption         =   "Budget"
   ClientHeight    =   10830
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   18435
   LinkTopic       =   "Form1"
   ScaleHeight     =   10830
   ScaleWidth      =   18435
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboMonth 
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
      ItemData        =   "frmBudget.frx":0000
      Left            =   120
      List            =   "frmBudget.frx":002B
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   0
      Width           =   2895
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   10455
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   18375
      _ExtentX        =   32411
      _ExtentY        =   18441
      _Version        =   393216
      Rows            =   40
      Cols            =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmBudget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim budget_expense_categories(20, 2) As String
Dim budget_income_categories(10, 2) As String


Private Sub Form_Load()
    Dim r As Long
    Dim q As ADODB.Recordset
    Dim tot As Double
    Dim offset As Byte
    'Grid.CellBackColor = vbRed
    
    cboMonth.ListIndex = 0
    offset = 7
    
    init_categories
    For r = 0 To 4
        Grid.ColWidth(r) = 2000
    Next r
    
    Grid.col = 0
    Grid.row = 0:    Grid.CellFontBold = True
    Grid.row = 5:    Grid.CellFontBold = True
    Grid.row = 7:    Grid.CellFontBold = True
    Grid.row = 20:    Grid.CellFontBold = True
    Grid.row = 0
    For r = 1 To 4
        Grid.col = r:    Grid.CellFontBold = True
    Next r
    
    Grid.TextMatrix(0, 0) = "Income"
    
    For r = 1 To 4
        Grid.TextMatrix(r, 0) = budget_income_categories(r, 1)
    Next r
    Grid.TextMatrix(5, 0) = "Total Income"

    Grid.TextMatrix(7, 0) = "Expenses"
    For r = 1 To 12
        Grid.TextMatrix(r + 7, 0) = budget_expense_categories(r, 1)
    Next r
    Grid.TextMatrix(20, 0) = "Total Expenses"
    Grid.TextMatrix(0, 1) = "Budgeted"
    Grid.TextMatrix(0, 2) = "Recorded"
    Grid.TextMatrix(0, 3) = "Difference"
    Grid.TextMatrix(0, 4) = "Moo"
    
    Set q = db.Execute("SELECT * FROM budget")
    q.MoveFirst
    tot = 0
    For r = 1 To 4
        Grid.TextMatrix(r, 1) = q.Fields(Replace(Grid.TextMatrix(r, 0), " ", "_")).value
        tot = tot + q.Fields(Replace(Grid.TextMatrix(r, 0), " ", "_")).value
    Next r
    Grid.TextMatrix(5, 1) = tot
    tot = 0
    For r = 1 To 12
        Grid.TextMatrix(r + offset, 1) = q.Fields(Replace(Grid.TextMatrix(r + offset, 0), " ", "_")).value
        tot = tot + q.Fields(Replace(Grid.TextMatrix(r + offset, 0), " ", "_")).value
    Next r
    Grid.TextMatrix(20, 1) = tot
    
    tot = 0
    For r = 1 To 4
        'MsgBox "SELECT SUM(value_num) as value FROM splits INNER JOIN transactions ON splits.guid = transactions.guid WHERE value_num > 0 AND account_guid = '" & budget_income_categories(r, 2) & "' GROUP BY account_guid"
        'Clipboard.SetText "SELECT DISTINCT SUM(value_num) as value FROM splits INNER JOIN transactions ON splits.guid = transactions.guid WHERE splits.value_num > 0 AND splits.account_guid = '" & budget_income_categories(r, 2) & "'", vbCFText
        'Set q = gcdb.Execute(SELECT SUM(value_num) as value FROM splits INNER JOIN transactions ON splits.tx_guid = transactions.guid WHERE splits.account_guid = 'd07db1f904884662a1238a455774ca8c')
        Set q = gcdb.Execute("SELECT SUM(value_num / -value_denom) as value FROM splits INNER JOIN transactions ON splits.tx_guid = transactions.guid WHERE splits.account_guid = '" & budget_income_categories(r, 2) & "'")
        If Not (q.EOF And q.BOF) Then
            q.MoveFirst
            Grid.TextMatrix(r, 2) = val("" & q!value)
        End If
        Grid.TextMatrix(r, 2) = val("" & q!value) 'budget_income_categories(r, 2)
        tot = tot + val(Grid.TextMatrix(r, 2))
    Next r
    Grid.TextMatrix(5, 2) = tot
    
    tot = 0
    For r = 8 To 19
        'MsgBox budget_expense_categories(r - 7, 2)
        Set q = gcdb.Execute("SELECT SUM(value_num / value_denom) AS value FROM splits INNER JOIN transactions ON splits.tx_guid = transactions.guid WHERE INSTR(""" & MiD$(budget_expense_categories(r - 7, 2), 3) & """, splits.account_guid) > 0")
        If Not (q.EOF And q.BOF) Then
            q.MoveFirst
            Grid.TextMatrix(r, 2) = val("" & q!value)
        End If
        Grid.TextMatrix(r, 2) = val("" & q!value)
        tot = tot + val(Grid.TextMatrix(r, 2))
    Next r
    Grid.TextMatrix(20, 2) = tot
    
    Set q = Nothing
End Sub

Sub init_categories()
    'second column will contain a list of account guid's for items that fall into that category
    budget_income_categories(1, 1) = "Parental Fees": budget_income_categories(1, 2) = "Parental Fees"
    budget_income_categories(2, 1) = "Child Care Subsidy": budget_income_categories(2, 2) = "Child Care Subsidy"
    budget_income_categories(3, 1) = "OGP funding": budget_income_categories(3, 2) = "OGP"
    budget_income_categories(4, 1) = "Other Income": budget_income_categories(4, 2) = "Other Income"
    
    budget_expense_categories(1, 1) = "Loan Payments": budget_expense_categories(1, 2) = ", Loan Repayments,"
    budget_expense_categories(2, 1) = "Insurance": budget_expense_categories(2, 2) = ", Insurance, Auto, Auto Insurance, Liability Insurance, Workers Comp,"
    budget_expense_categories(3, 1) = "Property Tax": budget_expense_categories(3, 2) = ", Property,"
    budget_expense_categories(4, 1) = "Maintenance": budget_expense_categories(4, 2) = ", Repair and Maintenance, Maintenance & Repairs, Computer Repairs, Equipment Repairs,"
    budget_expense_categories(5, 1) = "Equipment": budget_expense_categories(5, 2) = ", Equipment,"
    budget_expense_categories(6, 1) = "Power": budget_expense_categories(6, 2) = ", Electric, Utilities,"
    budget_expense_categories(7, 1) = "Phone/Internet": budget_expense_categories(7, 2) = ", Phone, Internet, Cell Phone,"
    budget_expense_categories(8, 1) = "Snow Clearing": budget_expense_categories(8, 2) = ", Outside Services,"
    budget_expense_categories(9, 1) = "Travel": budget_expense_categories(9, 2) = ", Travel, Entertainment, Meals, Gas,"
    budget_expense_categories(10, 1) = "Payroll": budget_expense_categories(10, 2) = ", Payroll Expenses, Remittance,"
    budget_expense_categories(11, 1) = "Food/Supplies": budget_expense_categories(11, 2) = ", Supplies, Groceries, Office Supplies, Postage and Delivery,"
    budget_expense_categories(12, 1) = "Miscellaneous": budget_expense_categories(12, 2) = "<Automatically Everything Else>"
    
    Dim q As ADODB.Recordset
    Dim i As Long
    
    budget_expense_categories(12, 2) = ""
    Set q = gcdb.Execute("SELECT * FROM accounts WHERE account_type = ""EXPENSE""")
    With q
    If Not (.EOF And .BOF) Then .MoveFirst
    Do Until .EOF
        For i = 1 To 11
            If InStr(1, budget_expense_categories(i, 2), ", " & Trim(!name) & ",") > 0 Then
                budget_expense_categories(i, 2) = Replace(budget_expense_categories(i, 2), ", " & !name & ",", ", " & !guid & ",")
                Exit For
            End If
        Next i
        If i = 12 Then
            'MsgBox !name
            budget_expense_categories(i, 2) = budget_expense_categories(i, 2) & !guid & ","
        End If
        .MoveNext
    Loop
    End With
        'For i = 1 To 12
        '    MsgBox budget_expense_categories(i, 2)
        'Next i
        
    Set q = gcdb.Execute("SELECT * FROM accounts WHERE account_type = ""INCOME""")
    With q
    If Not (.EOF And .BOF) Then .MoveFirst
    Do Until .EOF
        For i = 1 To 4
            If InStr(1, !name, budget_income_categories(i, 2)) > 0 Then
                budget_income_categories(i, 2) = Replace(budget_income_categories(i, 2), !name, !guid)
                'MsgBox budget_income_categories(i, 2)
                Exit For
            End If
        Next i
        .MoveNext
    Loop
    End With
    
    Set q = Nothing
End Sub

