VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCalendar 
   Caption         =   "Calendar"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   5025
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkFullYear 
      Caption         =   "Full Year"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   2040
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.TextBox textYear 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4080
      TabIndex        =   3
      Text            =   "2005"
      Top             =   1680
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3000
      Top             =   1680
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   735
      Left            =   3000
      TabIndex        =   2
      Top             =   2400
      Width           =   1815
   End
   Begin MSComCtl2.MonthView MonthView 
      Height          =   2370
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   167247873
      CurrentDate     =   38413
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Print Year :"
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Birthday Calendar               Printer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public printYear As Long
Function daysInMonth(ByVal mnth As Byte, ByVal yr As Long) As Byte
    'Dim temp As Date
    'temp = CDate(MonthName(mnth) & " 15, " & yr) + 31
    'temp = CDate(month(temp) & " 1, " & Year(temp))
    'temp = temp - 1
    'daysInMonth = Day(temp)
    daysInMonth = DateDiff("d", CDate(MonthName(mnth) & " 15, " & yr), CDate(MonthName(month(CDate(MonthName(mnth) & " 15, " & yr) + 31)) & " 15, " & Year(CDate(MonthName(mnth) & " 15, " & yr) + 31)))
End Function

Sub printMonth(ByVal month As Byte, ByVal x As Long, ByVal y As Long)
    Dim blockWidth As Long
    Dim blockHeight As Long
    Dim weekHeight As Long
    Dim Shift As Long
    Dim dy As Byte
    Dim i As Byte
    Dim j As Byte
    Dim k As Byte
    Dim name As String
    Dim birthdays As ADODB.Recordset
    
    MonthView.value = startOfMonth(month)
    DoEvents
    Printer.FontItalic = True
    printText MonthName(month), x, y, 1000, "Times New Roman", 20, True, 0
    Printer.FontItalic = False

    blockWidth = 1600
    blockHeight = 1100
    weekHeight = 300
    Shift = 800
    
    For i = 0 To 7
        Printer.Line (x + i * blockWidth, y + Shift - weekHeight)-(x + i * blockWidth, y + Shift + (6 * blockHeight))
    Next i
    For i = 0 To 6
        Printer.Line (x, y + Shift + (i * blockHeight))-(x + 7 * blockWidth, y + Shift + (i * blockHeight))
    Next i
    
    'Week Labels
    Printer.Line (x, y + Shift - weekHeight)-(x + 7 * blockWidth, y + Shift - weekHeight)
    For i = 0 To 6
        printText WeekdayName((i + 1), False), x + 50 + i * blockWidth, y + Shift - weekHeight, 1000, "verdana", 10, False, 0
    Next i
    
    i = MonthView.DayOfWeek
    j = 0
    For dy = 1 To daysInMonth(month, printYear)
        printText dy, x + (i - 1) * blockWidth + 50, y + Shift + j * blockHeight + 40, 1000, "Verdana", 10, False, 0
        Printer.fontsize = 8
        Set birthdays = db.Execute("SELECT * FROM clients WHERE month(dob) = " & month & " AND day(dob) = " & dy)
        If birthdays.EOF And birthdays.BOF Then
            printText dy, x + (i - 1) * blockWidth + 50, y + Shift + j * blockHeight + 40, 1000, "Verdana", 10, False, 0
        Else
            k = 1
            birthdays.MoveFirst
            Do Until birthdays.EOF
                name = birthdays!First & " " & birthdays!Last
                printText name, x + (i - 1) * blockWidth + 50, y + Shift + j * blockHeight + 40 + k * Printer.TextHeight("M"), 1000, "Verdana", 8, False, 0
                k = k + 1
                birthdays.MoveNext
            Loop
            
        End If
        i = i + 1
        If i > 7 Then
            i = 1
            j = j + 1
        End If
    Next dy
    
    Set birthdays = Nothing
End Sub

Sub printText(ByVal s As String, ByVal x As Long, ByVal y As Long, ByVal textwidth As Long, ByVal fontname As String, ByVal fontsize As Byte, ByVal bold As Boolean, ByVal justify As Byte)
    
    Printer.forecolor = 0
    Printer.Font = fontname
    Printer.FontBold = bold
    Printer.fontsize = fontsize
        
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print s
    
End Sub

Function startOfMonth(ByVal month As Byte) As Date
    startOfMonth = CDate(MonthName(month) & " 1," & printYear)
End Function


Private Sub Command1_Click()
    Dim mo As Byte
    Dim pos As Byte
    Dim pos1 As Long
    Dim pos2 As Long
    Dim start As Long
    
    Printer.ScaleMode = vbTwips
    printText printYear & " Calendar", 4500, 50, 1000, "Times New Roman", 20, True, 0
    
    pos1 = 150
    pos2 = 7600
    pos = 1
    If chkFullYear = 1 Then
        start = 1
    Else
        start = month(MonthView.value)
    End If
    
    If start Mod 2 = 0 Then start = start - 1
    For mo = start To 12
        If pos = 1 Then
            printMonth mo, 150, pos1
            pos = 2
        Else
            printMonth mo, 150, pos2
            pos = 1
            Printer.NewPage
        End If
    Next mo
    Printer.EndDoc
End Sub

Private Sub Form_Load()
    printYear = Year(Date + 60)
    MonthView.value = Date
    textYear = printYear
End Sub

Private Sub MonthView_DateClick(ByVal DateClicked As Date)
    Timer1.Enabled = True
End Sub

Private Sub textYear_Change()
    printYear = val(textYear)
End Sub

Private Sub Timer1_Timer()
    Command1.SetFocus
    Timer1.Enabled = False
End Sub


Private Sub textyear_Validate(Cancel As Boolean)
    If val(textYear) < 0 Or val(textYear) - Int(val(textYear)) <> 0 Then 'positive integer
        Cancel = True
        MsgBox "Invalid Year"
    End If
End Sub


