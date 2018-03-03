VERSION 5.00
Begin VB.Form dlgSetTimes 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cancelButn 
      BackColor       =   &H00C0C0FF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   0
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   5400
      Top             =   2040
   End
   Begin VB.CommandButton MiD 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "\/"
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton MiU 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "/\"
      Height          =   495
      Left            =   3240
      TabIndex        =   7
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton MoD 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "\/"
      Height          =   495
      Left            =   3240
      TabIndex        =   21
      Top             =   4560
      Width           =   495
   End
   Begin VB.CommandButton MoU 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "/\"
      Height          =   495
      Left            =   3240
      TabIndex        =   20
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton MoDD 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "\/"
      Height          =   495
      Left            =   2760
      TabIndex        =   19
      Top             =   4560
      Width           =   495
   End
   Begin VB.CommandButton MoUU 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "/\"
      Height          =   495
      Left            =   2760
      TabIndex        =   18
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox Mo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   2760
      TabIndex        =   17
      Text            =   "00"
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton HoD 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "\/"
      Height          =   495
      Left            =   1680
      TabIndex        =   15
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton HoU 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "/\"
      Height          =   495
      Left            =   1680
      TabIndex        =   14
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox Ho 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   1680
      TabIndex        =   13
      Text            =   "5"
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton MiDD 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "\/"
      Height          =   495
      Left            =   2760
      TabIndex        =   10
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton MiUU 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "/\"
      Height          =   495
      Left            =   2760
      TabIndex        =   9
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Mi 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   2760
      TabIndex        =   6
      Text            =   "00"
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton HiD 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "\/"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton HiU 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "/\"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Hi 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   1680
      TabIndex        =   1
      Text            =   "8"
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton setButn 
      Caption         =   "Set"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   0
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Shape shpBorder 
      Height          =   615
      Left            =   5280
      Top             =   600
      Width           =   615
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   5760
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Sign Out"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   23
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label AMPMo 
      Alignment       =   2  'Center
      Caption         =   "PM"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   22
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   16
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Sign In"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   12
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label AMPMi 
      Alignment       =   2  'Center
      Caption         =   "AM"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   11
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   5
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label TitleBar 
      BackColor       =   &H00FFCC88&
      Caption         =   "Client - Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "dlgSetTimes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intime As Date
Dim outtime As Date
Dim mintime As Date
Dim maxtime As Date
Dim clientIndex As Long



Private Sub cancelButn_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    shpBorder.Top = 0
    shpBorder.Left = 0
    shpBorder.width = Me.width
    shpBorder.height = Me.height
    mintime = CDate("06:00:00AM")
    maxtime = CDate("07:00:00PM")
End Sub

Sub convertTimes()
    intime = CDate(Hi & ":" & Mi & ":00" & AMPMi)
    outtime = CDate(Ho & ":" & Mo & ":00" & AMPMo)
    'If intime < mintime Then intime = mintime
    'If intime > maxtime Then intime = maxtime
    'If outtime < mintime Then outtime = mintime
    'If outtime > maxtime Then outtime = maxtime
    'MsgBox intime
End Sub

Sub setTimes()
    'MsgBox intime & "  " & outtime & "  " & maxtime
    If intime < mintime Then intime = mintime
    If intime > maxtime Then intime = maxtime
    If outtime < mintime Then outtime = mintime
    If outtime > maxtime Then outtime = maxtime
    If outtime < intime Then outtime = intime
    
    If Hour(intime) Mod 12 = 0 Then
        Hi = 12
    Else
        Hi = Hour(intime) Mod 12
    End If
    Mi = Format(Minute(intime), "00")
    AMPMi = Format(intime, "AMPM")
    
    If Hour(outtime) Mod 12 = 0 Then
        Ho = 12
    Else
        Ho = Hour(outtime) Mod 12
    End If
    Mo = Format(Minute(outtime), "00")
    AMPMo = Format(outtime, "AMPM")
End Sub

Private Sub HiD_Click()
    convertTimes
    intime = TimeSerial(Hour(intime) + 24 - 1, Minute(intime), 0) - 1
    setTimes
End Sub

Private Sub HiU_Click()
    convertTimes
    intime = TimeSerial(Hour(intime) + 1, Minute(intime), 0)
    setTimes
End Sub

Private Sub MiD_Click()
    convertTimes
    intime = TimeSerial(Hour(intime), Minute(intime) - 1, 0)
    setTimes
End Sub

Private Sub MiDD_Click()
    convertTimes
    intime = TimeSerial(Hour(intime) + 24, Minute(intime) - 10, 0) - 1
    setTimes
End Sub

Private Sub MiU_Click()
    convertTimes
    intime = TimeSerial(Hour(intime) + 24, Minute(intime) + 1, 0) - 1
    setTimes
End Sub

Private Sub MiUU_Click()
    convertTimes
    intime = TimeSerial(Hour(intime), Minute(intime) + 10, 0)
    setTimes
End Sub

Private Sub HoD_Click()
    convertTimes
    outtime = TimeSerial(Hour(outtime) + 24 - 1, Minute(outtime), 0) - 1
    setTimes
End Sub

Private Sub HoU_Click()
    convertTimes
    outtime = TimeSerial(Hour(outtime) + 1, Minute(outtime), 0)
    setTimes
End Sub

Private Sub MoD_Click()
    convertTimes
    outtime = TimeSerial(Hour(outtime), Minute(outtime) - 1, 0)
    setTimes
End Sub

Private Sub MoDD_Click()
    convertTimes
    outtime = TimeSerial(Hour(outtime) + 24, Minute(outtime) - 10, 0) - 1
    setTimes
End Sub

Private Sub MoU_Click()
    convertTimes
    outtime = TimeSerial(Hour(outtime) + 24, Minute(outtime) + 1, 0) - 1
    setTimes
End Sub

Private Sub MoUU_Click()
    convertTimes
    outtime = TimeSerial(Hour(outtime), Minute(outtime) + 10, 0)
    setTimes
End Sub

Private Sub setButn_Click()
    convertTimes
    frmAttendanceEntry.signin(clientIndex) = intime
    frmAttendanceEntry.signout(clientIndex) = outtime
    Unload Me
End Sub

Private Sub Timer1_Timer()
    Dim temptime As Date
    Timer1.Enabled = False
    clientIndex = frmAttendanceEntry.selected ' .ActiveControl.index
    'MsgBox frmAttendanceEntry.signin(clientIndex) & "   " & clientIndex
    temptime = frmAttendanceEntry.signin(clientIndex)
    intime = CDate(Hour(temptime) & ":" & Minute(temptime) & ":00 " & Format(temptime, "AMPM"))
    temptime = frmAttendanceEntry.signout(clientIndex)
    outtime = CDate(Hour(temptime) & ":" & Minute(temptime) & ":00 " & Format(temptime, "AMPM"))
    TitleBar.Caption = " " & frmAttendanceEntry.labClient(clientIndex) & " - " & shortDate(frmAttendanceEntry.MonthView.value)
    setTimes
End Sub
