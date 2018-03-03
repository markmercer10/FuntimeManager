VERSION 5.00
Begin VB.Form dlgPayingThisWeek 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paying This Week"
   ClientHeight    =   9300
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9300
   ScaleWidth      =   4830
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox textBox 
      Height          =   9255
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   4095
   End
   Begin VB.CommandButton printButn 
      Caption         =   "Print"
      Height          =   615
      Left            =   4080
      TabIndex        =   0
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "dlgPayingThisWeek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
    
    Dim q As ADODB.Recordset
    Dim p As ADODB.Recordset
    Dim tempdate As Date
    Dim thisfriday As Date
    Dim total As Double
    
    thisfriday = nearestFriday(Date + 3)
    total = 0
    
    Set q = db.Execute("SELECT * FROM Clients WHERE active = 1 AND subsidized = 0")
    With q
        If Not (.EOF And .BOF) Then .MoveFirst
        Do Until .EOF
            If !payperiod = 1 Then
                textBox = textBox & !First & " " & !Last & " - $" & !fees * !payperiod & vbCrLf
                total = total + !fees * !payperiod
            Else
                Set p = db.Execute("SELECT * FROM Payments WHERE idClient = " & q!idClient & " ORDER BY date DESC LIMIT 1")
                If Not (p.EOF And p.BOF) Then
                    p.MoveFirst
                    tempdate = nearestFriday(p!Date)
                End If
                
                If tempdate + !payperiod * 7 = thisfriday Or tempdate + (!payperiod * 7) * 2 = thisfriday Then
                    textBox = textBox & !First & " " & !Last & " - $" & !fees * !payperiod & vbCrLf
                    total = total + !fees * !payperiod
                End If
            End If
            .MoveNext
        Loop
    End With

    textBox = textBox & vbCrLf & "Total: $" & total
        
End Sub

Private Sub printButn_Click()
    printButn.Visible = False
    printText Me.Caption & " - " & ansiDate(Date), 50, 50, 10000, "Arial", 13, True, vbLeftJustify
    formPrint Me, 50, 500
    Printer.EndDoc
    printButn.Visible = True
End Sub
