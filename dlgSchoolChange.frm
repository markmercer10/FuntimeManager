VERSION 5.00
Begin VB.Form dlgSchoolChange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Summer Switch"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   6675
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton okButn 
      Caption         =   "OK"
      Height          =   495
      Left            =   5280
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton skipButn 
      Caption         =   "Skip"
      Height          =   495
      Left            =   4080
      TabIndex        =   10
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ComboBox cboRoom 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      ItemData        =   "dlgSchoolChange.frx":0000
      Left            =   1200
      List            =   "dlgSchoolChange.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2520
      Width           =   1212
   End
   Begin VB.TextBox txtFees 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   2520
      Width           =   855
   End
   Begin VB.ComboBox cboFeeClass 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1920
      Width           =   5295
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1440
      Top             =   1320
   End
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   2040
      Top             =   1320
   End
   Begin VB.Label message 
      Caption         =   "School's out for summer!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   6375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Effective Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   1380
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Room:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   2580
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Fee Class:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1980
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Fees:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   2580
      Width           =   1695
   End
End
Attribute VB_Name = "dlgSchoolChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboFeeClass_Change()
    cboFeeClass_Click
End Sub

Private Sub cboFeeClass_Click()
    txtFees = val(MiD$(cboFeeClass, InStr(1, cboFeeClass, "$") + 1))
End Sub

Private Sub Form_Load()
    Dim fc As ADODB.Recordset
    Set fc = db.Execute("SELECT * FROM fee_classes ORDER BY idFeeClasses ASC")
    With fc
        If Not (.BOF And .EOF) Then
            .MoveFirst
            Do Until .EOF
                cboFeeClass.AddItem !Description & " - $" & !charge, !idFeeClasses - 1
                .MoveNext
            Loop
        End If
    End With
    
    Set fc = Nothing
End Sub

Private Sub okButn_Click()
    Dim sql As String
    
    sql = "UPDATE clients SET "
    sql = sql & "feeClassID=" & cboFeeClass.ListIndex + 1 & ","
    sql = sql & "fees=" & txtFees & ","
    sql = sql & "room=""" & cboRoom.Text & """"
    sql = sql & " WHERE idClient = " & Timer1.Tag
    'MsgBox sql
    db.Execute sql
    
    'add record to client_changes table
    upsertClientChange dpEffective.value, _
    Timer1.Tag, _
    cboFeeClass.ListIndex + 1, _
    txtFees, _
    getPayperiodAtDate(val(Timer1.Tag), dpEffective.value), _
    cboRoom.Text, _
    getSubsidizedAtDate(val(Timer1.Tag), dpEffective.value), _
    getAuthorizationNumberAtDate(val(Timer1.Tag), dpEffective.value), _
    getParentContributionAtDate(val(Timer1.Tag), dpEffective.value), _
    getStartDateAtDate(val(Timer1.Tag), dpEffective.value), _
    getActiveAtDate(val(Timer1.Tag), dpEffective.value)
    
    If Abs(DateDiff("d", LDOS, Date)) < 10 Then
        LDOS = dpEffective.value
    Else
        FDOS = dpEffective.value
    End If
    
    Set q = Nothing
    Unload Me
End Sub

Private Sub skipButn_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Dim tempdate As Date
    Dim diff As Long
    Dim i As Byte
    Dim q As ADODB.Recordset
    Set q = db.Execute("SELECT * FROM clients WHERE idClient = " & Timer1.Tag)
    If Not (q.EOF And q.BOF) Then
        q.MoveFirst
        cboFeeClass.ListIndex = val(cboFeeClass.Tag) - 1
        'I THINK THIS METHOD WAS STILL BEING WORKED ON WHEN LAST I LEFT IT.
        For i = 0 To 3 'Make a global for number of rooms
            If cboRoom.List(i) = q!room Then cboRoom.ListIndex = i
        Next i
        If Abs(DateDiff("d", LDOS, Date)) < 15 Then
            dpEffective.value = LDOS
            message = message & vbCrLf & q!First & " " & q!Last & " should be switching to a different fee class.  See the suggestion below.  Please select the new fee class and other options from the below form."
        Else
            dpEffective.value = FDOS
            message = "Back To School!" & vbCrLf & q!First & " " & q!Last & " should be switching to a different fee class.  See the suggestion below.  Please select the new fee class and other options from the below form."
        End If
    End If
    
    Set q = Nothing
End Sub

Private Sub Timer2_Timer()
    Timer2.Enabled = False
    If Me.Visible Then
        dpEffective.SetFocus
    End If
End Sub

