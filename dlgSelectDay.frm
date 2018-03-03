VERSION 5.00
Begin VB.Form dlgSelectDay 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Statutory Holiday"
   ClientHeight    =   795
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   795
   ScaleWidth      =   4515
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   120
   End
   Begin VB.CommandButton okButn 
      Caption         =   "Select"
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
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.ComboBox cboDay 
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
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Label labMo 
      Alignment       =   1  'Right Justify
      Caption         =   "Label1"
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
      Left            =   0
      TabIndex        =   1
      Top             =   300
      Width           =   1815
   End
End
Attribute VB_Name = "dlgSelectDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboDay_Click()
    If cboDay.ListIndex > -1 Then okButn.Enabled = True
End Sub

Private Sub okButn_Click()
    frmSubsidization.statDay = cboDay.Text
    Unload Me
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Dim i As Long
    labMo = MonthName(val(labMo.Tag))
    For i = 1 To daysInMonth(CDate(Year(Date) & "-" & Format(val(labMo.Tag), "00") & "-01"))
        cboDay.AddItem i
    Next i
End Sub
