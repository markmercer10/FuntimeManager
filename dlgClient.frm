VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form dlgClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New Client"
   ClientHeight    =   8265
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   5760
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame subsFrame 
      Caption         =   "Subsidization Info"
      Height          =   1215
      Left            =   2880
      TabIndex        =   39
      Top             =   6390
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox txtParentalContrib 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   4105
            SubFormatType   =   2
         EndProperty
         Height          =   288
         Left            =   1440
         TabIndex        =   42
         Text            =   "0.00"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtAuthNumber 
         Height          =   288
         Left            =   1320
         TabIndex        =   40
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label19 
         Caption         =   "Parental Contrib."
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "Authorization #"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CheckBox chkDropIn 
      Caption         =   "Drop In"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1560
      TabIndex        =   38
      Top             =   6840
      Width           =   1455
   End
   Begin VB.TextBox ID 
      Height          =   288
      Left            =   120
      TabIndex        =   35
      Top             =   6840
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.ComboBox cboPP 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "dlgClient.frx":0000
      Left            =   1560
      List            =   "dlgClient.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   6030
      Width           =   855
   End
   Begin VB.TextBox txtEmergency 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      TabIndex        =   6
      Top             =   2760
      Width           =   2532
   End
   Begin VB.TextBox txtPhone 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      TabIndex        =   3
      Top             =   1680
      Width           =   2532
   End
   Begin VB.CheckBox chkSubsidized 
      Caption         =   "Subsidized"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1560
      TabIndex        =   14
      Top             =   6480
      Width           =   1452
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   7200
   End
   Begin VB.TextBox txtLast 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      TabIndex        =   1
      Top             =   960
      Width           =   2532
   End
   Begin VB.CommandButton cancelButn 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2640
      TabIndex        =   17
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton saveButn 
      Caption         =   "Save"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4200
      TabIndex        =   16
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CheckBox chkActive 
      Caption         =   "Active"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1560
      TabIndex        =   15
      Top             =   7200
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox cboRoom 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "dlgClient.frx":0020
      Left            =   1560
      List            =   "dlgClient.frx":0022
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   4700
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker dpEnd 
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   5580
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MMM d, yyyy"
      Format          =   137428995
      CurrentDate     =   42531
   End
   Begin MSComCtl2.DTPicker dpStart 
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   5100
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MMM d, yyyy"
      Format          =   137428995
      CurrentDate     =   42531
   End
   Begin VB.TextBox txtFees 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      TabIndex        =   10
      Top             =   4320
      Width           =   2052
   End
   Begin VB.ComboBox cboFeeClass 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3960
      Width           =   3492
   End
   Begin VB.ComboBox cboGender 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "dlgClient.frx":0024
      Left            =   1560
      List            =   "dlgClient.frx":002E
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   3600
      Width           =   855
   End
   Begin MSComCtl2.DTPicker dpDOB 
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   3120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MMM d, yyyy"
      Format          =   158138371
      CurrentDate     =   42530
   End
   Begin VB.TextBox txtParent2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      TabIndex        =   5
      Top             =   2400
      Width           =   2532
   End
   Begin VB.TextBox txtParent1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      TabIndex        =   4
      Top             =   2040
      Width           =   2532
   End
   Begin VB.TextBox txtInitial 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   2
      Top             =   1320
      Width           =   852
   End
   Begin VB.TextBox txtFirst 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      TabIndex        =   0
      Top             =   600
      Width           =   2532
   End
   Begin MSComCtl2.DTPicker dpEffective 
      Height          =   375
      Left            =   1560
      TabIndex        =   36
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      CalendarForeColor=   255
      CalendarTitleBackColor=   65535
      CalendarTitleForeColor=   255
      CustomFormat    =   "MMM d, yyyy"
      Format          =   158138371
      CurrentDate     =   42531
   End
   Begin VB.Label Label17 
      BackColor       =   &H000000FF&
      Caption         =   "Changes Effective                                             Ensure correct date!"
      Height          =   375
      Left            =   120
      TabIndex        =   37
      Top             =   120
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label Label16 
      Caption         =   "Weeks"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   34
      Top             =   6080
      Width           =   855
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "Pay Period"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -480
      TabIndex        =   32
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "Emergency Contact"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -480
      TabIndex        =   31
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Phone Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -480
      TabIndex        =   30
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Room"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -480
      TabIndex        =   29
      Top             =   4720
      Width           =   1935
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "End Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -480
      TabIndex        =   28
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Start Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -480
      TabIndex        =   27
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Fees"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -480
      TabIndex        =   26
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Fee Class"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -480
      TabIndex        =   25
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -480
      TabIndex        =   24
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "DOB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -480
      TabIndex        =   23
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Parent 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -480
      TabIndex        =   22
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Parent 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -480
      TabIndex        =   21
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Initial"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -480
      TabIndex        =   20
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -480
      TabIndex        =   19
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "First Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -480
      TabIndex        =   18
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "dlgClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim originalStartDate As Date
Dim changeDateReminder As Boolean

Private Sub cancelButn_Click()
    Unload Me
End Sub

Private Sub cboFeeClass_Change()
    txtFees = val(MiD$(cboFeeClass, InStr(1, cboFeeClass, "$") + 1))
    check_feeclass_age
    checkSaveEnabled
End Sub

Private Sub cboFeeClass_Click()
    cboFeeClass_Change
    checkSaveEnabled
End Sub

Private Sub cboRoom_Change()
    checkSaveEnabled
End Sub

Private Sub cboRoom_Click()
    checkSaveEnabled
End Sub

Private Sub chkActive_Click()
    'If CBool(chkActive) Then chkDropIn.value = 0
    If chkActive.value = 0 Then
        dpEnd.Visible = 1
    Else
        dpEnd.Visible = 0
    End If
End Sub

Private Sub chkDropIn_Click()
    'If CBool(chkDropIn) Then chkActive.value = 0
End Sub

Private Sub chkSubsidized_Click()
    If chkSubsidized.value = 1 Then
        subsFrame.Visible = True
        cboPP.ListIndex = 3
    Else
        subsFrame.Visible = False
    End If
End Sub



Private Sub dpEffective_Change()
    changeDateReminder = False
End Sub

Private Sub dpEffective_Click()
    changeDateReminder = False
End Sub

Private Sub dpEffective_KeyDown(KeyCode As Integer, Shift As Integer)
    changeDateReminder = False
End Sub

Private Sub dpEffective_Validate(Cancel As Boolean)
    If dpEffective.value < dpStart.value Then dpEffective.value = dpStart.value
    If dpEffective.value > dpEnd.value Then dpEffective.value = dpEnd.value
End Sub

Private Sub dpEnd_Change()
    dpEffective.value = dpEnd.value
End Sub

Private Sub Form_Load()
    changeDateReminder = True
End Sub

Private Sub SaveButn_Click()
    Dim sql As String
    Dim effectiveDate As Date
    
    If changeDateReminder Then
        changeDateReminder = False
        MsgBox "Have you checked the date that these changes are being applied as?  Please ensure the correct date is chosen."
        Exit Sub
    End If
    
    If ID = "" Then ' NEW CLIENT!!!
        effectiveDate = dpStart.value
        
        sql = "INSERT INTO clients ("
        
        sql = sql & "first,"
        sql = sql & "last,"
        sql = sql & "initial,"
        sql = sql & "phone,"
        sql = sql & "parent1,"
        sql = sql & "parent2,"
        sql = sql & "emergency,"
        sql = sql & "DOB,"
        sql = sql & "gender,"
        sql = sql & "feeClassID,"
        sql = sql & "fees,"
        sql = sql & "startDate,"
        sql = sql & "payperiod,"
        sql = sql & "room,"
        sql = sql & "subsidized,"
        If CBool(chkSubsidized.value) Then
            sql = sql & "authorizationNumber,"
            sql = sql & "parentalContribution,"
        End If
        sql = sql & "active"
        
        sql = sql & ") VALUES ("
        
        sql = sql & """" & txtFirst & ""","
        sql = sql & """" & txtLast & ""","
        sql = sql & """" & txtInitial & ""","
        sql = sql & """" & txtPhone & ""","
        sql = sql & """" & txtParent1 & ""","
        sql = sql & """" & txtParent2 & ""","
        sql = sql & """" & txtEmergency & ""","
        sql = sql & sqlDate(dpDOB.value) & ","
        sql = sql & """" & cboGender.Text & ""","
        sql = sql & cboFeeClass.ListIndex + 1 & ","
        sql = sql & val(txtFees) & ","
        sql = sql & sqlDate(dpStart.value) & ","
        sql = sql & cboPP.Text & ","
        sql = sql & """" & cboRoom.Text & ""","
        sql = sql & chkSubsidized.value & ","
        If CBool(chkSubsidized.value) Then
            sql = sql & """" & txtAuthNumber.Text & ""","
            sql = sql & txtParentalContrib.Text & ","
        End If
        sql = sql & 1 'chkActive.value
        
        sql = sql & ")"
        
        'Clipboard.SetText sql
        'MsgBox sql
        db.Execute sql
        
    Else               ' EDITING CLIENT!
        effectiveDate = dpEffective.value
        
        sql = "UPDATE clients SET "
        
        sql = sql & "first=""" & txtFirst & ""","
        sql = sql & "last=""" & txtLast & ""","
        sql = sql & "initial=""" & txtInitial & ""","
        sql = sql & "phone=""" & txtPhone & ""","
        sql = sql & "parent1=""" & txtParent1 & ""","
        sql = sql & "parent2=""" & txtParent2 & ""","
        sql = sql & "emergency=""" & txtEmergency & ""","
        sql = sql & "DOB=" & sqlDate(dpDOB.value) & ","
        sql = sql & "gender=""" & cboGender.Text & ""","
        sql = sql & "feeClassID=" & cboFeeClass.ListIndex + 1 & ","
        sql = sql & "fees=" & txtFees & ","
        sql = sql & "startDate=" & sqlDate(dpStart.value) & ","
        sql = sql & "payperiod=" & cboPP.Text & ","
        sql = sql & "room=""" & cboRoom.Text & ""","
        sql = sql & "subsidized=" & chkSubsidized.value & ","
        If CBool(chkSubsidized.value) Then
            sql = sql & "authorizationNumber=""" & txtAuthNumber.Text & ""","
            sql = sql & "parentalContribution=" & val(txtParentalContrib.Text) & ","
        End If
        If chkActive.value = 0 Then
            sql = sql & "enddate=" & sqlDate(dpEnd.value) & ","
        End If
        sql = sql & "active=" & chkActive.value
        
        sql = sql & " WHERE idClient = " & ID.Text
        
        db.Execute sql
    
    End If
    DoEvents
    
    'add record to client_changes table
    Dim q As ADODB.Recordset
    Set q = db.Execute("SELECT * FROM clients WHERE first = """ & txtFirst & """ AND last = """ & txtLast & """ AND DOB = " & sqlDate(dpDOB.value))
    With q
        If Not (.EOF And .BOF) Then
            .MoveFirst
            sql = "INSERT INTO client_changes (date, idClient, feeClassID, fees, payperiod, room, subsidized, authorizationNumber, parentalContribution, active) VALUES ("
            sql = sql & sqlDate(effectiveDate) & ","
            sql = sql & !idClient & ","
            sql = sql & !feeClassID & ","
            sql = sql & !fees & ","
            sql = sql & !payperiod & ","
            sql = sql & """" & !room & ""","
            sql = sql & !subsidized & ","
            sql = sql & """" & !authorizationNumber & ""","
            sql = sql & !parentalContribution & ","
            sql = sql & !active & ")"
            db.Execute sql
            .MoveNext
        End If
    End With
    
    If ID <> "" Then ' EDITING CLIENT
        If originalStartDate <> dpStart Then
            sql = "UPDATE client_changes SET date=" & sqlDate(dpStart.value) & " WHERE idClient = " & ID & " AND date = " & sqlDate(originalStartDate)
            db.Execute sql
        End If
    End If
    
    Set q = Nothing
    
    Unload Me
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Dim fc As ADODB.Recordset
    Dim rm As ADODB.Recordset
    dpStart = Date
    dpEnd = Date
    dpEffective = Date
    dpDOB = CDate(MonthName(month(Date)) & " " & day(Date) & ", " & year(Date) - 2)
    cboPP.ListIndex = 0
    
    Set fc = db.Execute("SELECT * FROM fee_classes")
    With fc
        If Not (.BOF And .EOF) Then
            .MoveFirst
            Do Until .EOF
                cboFeeClass.AddItem !Description & " - $" & !charge, !idFeeClasses - 1
                .MoveNext
            Loop
        End If
    End With
    
    Set rm = db.Execute("SELECT * FROM rooms")
    With rm
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do Until .EOF
                cboRoom.AddItem !Abbreviation '!name
                .MoveNext
            Loop
        End If
    End With
    
    'IF EDITING FILL THE FORM
    If ID <> "" Then
        changeDateReminder = False
        Me.Caption = "Editing Client"
        Label17.Visible = True
        dpEffective.Visible = True
        Dim cl As ADODB.Recordset
        Set cl = db.Execute("SELECT * FROM clients WHERE idClient=" & ID)
        With cl
            If Not (.EOF And .BOF) Then
                .MoveFirst
                txtFirst = !First
                txtLast = !Last
                txtInitial = !initial
                txtPhone = !phone
                txtParent1 = !parent1
                txtParent2 = !parent2
                txtEmergency = !emergency
                dpDOB.value = !DOB
                comboSelectItem cboGender, !gender
                cboFeeClass.ListIndex = !feeClassID - 1
                txtFees = !fees
                dpStart.value = !startdate
                originalStartDate = !startdate
                If IsNull(!enddate) Then
                    dpEnd.value = Date
                Else
                    dpEnd.value = !enddate
                End If
                comboSelectItem cboPP, !payperiod
                comboSelectItem cboRoom, !room
                chkSubsidized.value = !subsidized
                If !subsidized Then
                    txtAuthNumber = "" & !authorizationNumber
                    txtParentalContrib = Format(!parentalContribution, "0.00")
                End If
                chkActive.value = !active
                chkActive.Visible = True
            End If
        End With
        dpEffective.value = Date
    Else
    'NEW ENTRY
    
    End If
    Set cl = Nothing
    Set fc = Nothing
End Sub

Sub check_feeclass_age()
    Dim min As Long
    Dim max As Long
    Dim age As Long
    Dim q As ADODB.Recordset
    Set q = db.Execute("SELECT * FROM fee_classes WHERE idFeeClasses = """ & cboFeeClass.ListIndex + 1 & """")
    With q
        If Not (.EOF And .BOF) Then
            .MoveFirst
            min = !min_age
            max = !max_age
        End If
    End With
    Set q = Nothing
    
    age = getAgeM(dpDOB.value, Now)
    If age < min Or age > max Then
        MsgBox "The selected fee class is meant for ages " & min & " to " & max & " months but the age of this child is " & age & " months"
    End If
End Sub

Private Sub txtFees_Change()
    checkSaveEnabled
End Sub

Private Sub txtFirst_Change()
    checkSaveEnabled
End Sub

Private Sub txtInitial_Change()
    txtInitial = UCase$(txtInitial)
End Sub

Private Sub checkSaveEnabled()
    If txtFirst <> "" And txtLast <> "" And cboFeeClass.ListIndex <> -1 And txtFees <> "" And cboRoom.ListIndex <> -1 Then
        saveButn.Enabled = True
    Else
        saveButn.Enabled = False
    End If
End Sub

Private Sub txtLast_Change()
    checkSaveEnabled
End Sub
