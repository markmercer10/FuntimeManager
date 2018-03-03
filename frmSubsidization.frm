VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSubsidization 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Subsidization Forms"
   ClientHeight    =   10440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   18945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10440
   ScaleWidth      =   18945
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox editCell 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      TabIndex        =   28
      Top             =   3120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton calcButn 
      BackColor       =   &H0088DD88&
      Caption         =   "Calculate"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton loadButn 
      BackColor       =   &H00FFCC99&
      Caption         =   "Load"
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
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton saveButn 
      BackColor       =   &H00DD66DD&
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
      Height          =   735
      Left            =   16200
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   720
      Width           =   1215
   End
   Begin VB.ListBox clientList 
      Height          =   2790
      Left            =   360
      TabIndex        =   20
      Top             =   2400
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Frame frameModAttendance 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4815
      Left            =   3000
      TabIndex        =   10
      Top             =   3600
      Visible         =   0   'False
      Width           =   495
      Begin VB.CommandButton modButn 
         BackColor       =   &H00FFCC99&
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton modButn 
         BackColor       =   &H00FFCC99&
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   480
         Width           =   495
      End
      Begin VB.CommandButton modButn 
         BackColor       =   &H00FFCC99&
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton modButn 
         BackColor       =   &H00FFCC99&
         Caption         =   "IC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton modButn 
         BackColor       =   &H00FFCC99&
         Caption         =   "AS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1920
         Width           =   495
      End
      Begin VB.CommandButton modButn 
         BackColor       =   &H00FFCC99&
         Caption         =   "H"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2400
         Width           =   495
      End
      Begin VB.CommandButton modButn 
         BackColor       =   &H00FFCC99&
         Caption         =   "SH"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2880
         Width           =   495
      End
      Begin VB.CommandButton modButn 
         BackColor       =   &H00FFCC99&
         Caption         =   "SC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3360
         Width           =   495
      End
      Begin VB.CommandButton modButn 
         BackColor       =   &H00FFCC99&
         Caption         =   "KC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3840
         Width           =   495
      End
      Begin VB.CommandButton modButn 
         BackColor       =   &H00FFCC99&
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   4320
         Width           =   495
      End
   End
   Begin VB.CommandButton kcButn 
      BackColor       =   &H008888EE&
      Caption         =   "Add Kinder Care Day"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton scButn 
      BackColor       =   &H008888EE&
      Caption         =   "Add School Closure Day"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton statButn 
      BackColor       =   &H008888EE&
      Caption         =   "Add Stat Holiday"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   600
      Width           =   1575
   End
   Begin VB.ComboBox cboSaved 
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
      Left            =   4320
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   960
      Width           =   2535
   End
   Begin VB.CommandButton prntButn 
      BackColor       =   &H00DD66DD&
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   17520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
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
      ItemData        =   "frmSubsidization.frx":0000
      Left            =   240
      List            =   "frmSubsidization.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   2415
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
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid FlexGrid 
      Height          =   8775
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   18855
      _ExtentX        =   33258
      _ExtentY        =   15478
      _Version        =   393216
      Rows            =   10
      Cols            =   44
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFCCFF&
      Caption         =   " Finalize"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   15960
      TabIndex        =   24
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackColor       =   &H00CCCCFF&
      Caption         =   " Modify"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8520
      TabIndex        =   22
      Top             =   120
      Width           =   7335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00CCFFCC&
      Caption         =   " Calculate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "THIS WILL NOT FIT ONTO THE MONITOR AT THE DAYCARE!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   19200
      TabIndex        =   7
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFE0C0&
      Caption         =   "Saved Submissions:"
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
      Left            =   4320
      TabIndex        =   4
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFE0C0&
      Caption         =   " Load"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4200
      TabIndex        =   23
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmSubsidization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim periodStart As Date
Dim periodEnd As Date
Dim LOADED As Boolean
Public statDay As Byte

Sub printSubsidy()
    Dim img As StdPicture
    Printer.Orientation = vbPRORLandscape
    Dim total_width As Long
    Dim total_height As Long
    Dim TPI As Long ' twips per inch
    Dim block_height As Long
    Dim block_width As Long
    Dim margin_left As Long
    Dim margin_right As Long
    Dim grid_top As Long
    Dim grid_bottom As Long
    Dim lft As Long
    Dim columns(0 To 13) As Long
    Dim last_index As Long
    Dim current_index As Long
    Dim i As Long
    Dim j As Long
    Dim P As Byte
    Dim pages As Byte
    Dim temp As Double
    
    
    total_width = Printer.width
    total_height = Printer.height
    TPI = total_height / 8.5 ' should be 1440 for standard printers
    margin_left = 500
    margin_right = 800
    last_index = FlexGrid.rows - 3
    
    
    
    
    'FIRST FORM
    block_height = 220
    grid_top = 3800
    grid_bottom = grid_top + 27 * block_height
    columns(0) = margin_left
    columns(1) = columns(0) + 3200
    block_width = (total_width - columns(1) - margin_right) / 31
    
    
    ' DRAW THE FORM
    Set img = frmMain.ImageForms.ListImages(1).Picture
    Printer.PaintPicture img, 300, 1000, Printer.width * 0.6, Printer.height * 0.07
    
    For i = 0 To 32
        If i <= 1 Then
            Printer.Line (columns(i), grid_top)-(columns(i), grid_bottom), 0
        Else
            Printer.Line (columns(1) + (i - 1) * block_width, grid_top)-(columns(1) + (i - 1) * block_width, grid_bottom), 0
        End If
        If i = 0 Then
            printText "Child's Name", columns(0), grid_top, columns(1) - columns(0), "Arial", 8, True, 2
        ElseIf i <= 31 Then
            printText i, columns(1) + (i - 1) * block_width, grid_top, block_width, "Arial", 8, True, 2
        End If
    Next i
    For j = 0 To 27
        Printer.Line (columns(0), grid_top + j * block_height)-(total_width - margin_right, grid_top + j * block_height), 0
        If j > 0 And j < 27 Then
            printText j, columns(0) + 50, grid_top + j * block_height, 400, "Arial", 8, True, 0
        End If
    Next j
    
    printText "Child Care Center:____________________________________", margin_left, 2400, 10000, "Arial", 10, False, 0
    printText "Date:_________________________", margin_left, 3000, 10000, "Arial", 10, False, 0
    printText "Month of Attendance:__________________________", total_width / 2, 2400, 10000, "Arial", 10, False, 0
    
    printText "Please use the following", margin_left + 50, grid_bottom + 30, 10000, "Arial", 8, True, 0
    printText "I/C = Infant Care", columns(1), grid_bottom + 30, 10000, "Arial", 8, True, 0
    printText "P = Present Full Day", columns(1) + block_width * 7, grid_bottom + 30, 10000, "Arial", 8, True, 0
    printText "H = Present Half Day", columns(1) + block_width * 14, grid_bottom + 30, 10000, "Arial", 8, True, 0
    printText "AS = Afterschool Program", columns(1) + block_width * 20, grid_bottom + 30, 10000, "Arial", 8, True, 0
    printText "KC = Kindercare Day", columns(1) + block_width * 26, grid_bottom + 30, 10000, "Arial", 8, True, 0
    printText "SH = Statutory Holiday", columns(1), grid_bottom + 220, 10000, "Arial", 8, True, 0
    printText "A = Absent Day", columns(1) + block_width * 7, grid_bottom + 220, 10000, "Arial", 8, True, 0
    printText "SC = School Closure", columns(1) + block_width * 14, grid_bottom + 220, 10000, "Arial", 8, True, 0
    
    
    ' FILL THE FORM
    For i = 1 To FlexGrid.rows - 3
        For j = 0 To 31
            If j = 0 Then
                printText FlexGrid.TextMatrix(i, j), columns(j) + 300, grid_top + i * block_height + 20, 5000, "Arial", 8, False, 0
            Else
                lft = columns(1) + (j - 1) * block_width
                printText FlexGrid.TextMatrix(i, j), lft, grid_top + i * block_height + 20, block_width, "Arial", 8, False, 2
            End If
        Next j
    Next i
    
    printText "Funtime Child Care Center", margin_left + 2000, 2400, 10000, "Arial", 8, False, 0
    printText shortDate(Date), margin_left + 1000, 3000, 10000, "Arial", 8, False, 0
    printText cboMonth.Text, total_width / 2 + 3000, 2400, 10000, "Arial", 8, False, 0
    
    
    Printer.EndDoc
    
    
    
    
    
    
    
    
    'Exit Sub ' development only
    
    'SECOND FORM(S)
    block_height = 475
    grid_top = 3750
    columns(0) = margin_left
    columns(1) = columns(0) + 2800
    columns(2) = columns(1) + 2150
    columns(3) = columns(2) + 2700
    columns(4) = columns(3) + 550
    columns(5) = columns(4) + 50
    columns(6) = columns(5) + 550
    columns(7) = columns(6) + 500
    columns(8) = columns(7) + 850
    columns(9) = columns(8) + 450
    columns(10) = columns(9) + 700
    columns(11) = columns(10) + 900
    columns(12) = columns(11) + 900
    Set img = frmMain.ImageForms.ListImages(2).Picture
    
    If (last_index) / 8# > Int((FlexGrid.rows - 3) / 8) Then
        pages = Int(last_index / 8) + 1
    Else
        pages = Int(last_index / 8)
    End If
    
    For P = 1 To pages
        Printer.Orientation = vbPRORLandscape
        Printer.PaintPicture img, -100, 500, Printer.width, Printer.height * 0.9
        
        For i = 1 To 8
            current_index = i + (P - 1) * 8
            If current_index > last_index Then Exit For
            
            For j = 0 To 1
                printText FlexGrid.TextMatrix(current_index, j + 32), columns(j), grid_top + i * block_height, 3000, "Arial", 10, False, 0
            Next j
            printText FlexGrid.TextMatrix(current_index, 0), columns(2), grid_top + i * block_height, 3000, "Arial", 10, False, 0
            printText FlexGrid.TextMatrix(current_index, 34), columns(3), grid_top + i * block_height, 3000, "Arial", 7, False, 0
            printText FlexGrid.TextMatrix(current_index, 35), columns(4), grid_top + i * block_height, 3000, "Arial", 7, False, 0
            For j = 5 To 12
                printText FlexGrid.TextMatrix(current_index, j + 31), columns(j), grid_top + i * block_height, 800, "Arial", 9, False, 1
            Next j
        Next i
        For j = 5 To 12
            temp = 0
            For i = 1 + (P - 1) * 8 To 8 + (P - 1) * 8
                If i > last_index Then Exit For
                temp = temp + val(FlexGrid.TextMatrix(i, j + 31))
            Next i
            If j <= 7 Then
                printText temp, columns(j), grid_top + 9.25 * block_height, 800, "Arial", 9, False, 1
            Else
                printText Format(temp, "0.00"), columns(j), grid_top + 9.25 * block_height, 800, "Arial", 8, False, 1
            End If
            If j = 11 Then
                printText Format(temp, "0.00"), columns(11), 9350, 1200, "Arial", 12, True, 1
            End If
        Next j
        
        
        printText "Funtime Child Care Center", 1500, 1450, 5000, "Arial", 10, False, 0
        printText cboMonth.Text, 1500, 1800, 5000, "Arial", 10, False, 0
        printText Right$(cboYear.Text, 2), 3600, 1800, 5000, "Arial", 10, False, 0
        printText "709-759-2202", 1700, 2150, 5000, "Arial", 10, False, 0
        
        printText "Funtime Child Care Center", 6000, 1450, 5000, "Arial", 10, False, 0
        printText "P.O. Box 149", 6600, 1800, 5000, "Arial", 10, False, 0
        printText "Whitbourne, NL, A0B 3K0", 6600, 2150, 5000, "Arial", 10, False, 0
        
        printText P, 11600, 1450, 5000, "Arial", 10, False, 0
        printText pages, 12200, 1450, 5000, "Arial", 10, False, 0
        
        Printer.EndDoc
    Next P
    
    Printer.Orientation = vbPRORPortrait
End Sub


Private Sub calcButn_Click()
    Dim Mo As Byte
    Dim yr As Long
    Dim subs As ADODB.Recordset
    
    If CDate(cboMonth.Text & " 15, " & cboYear.Text) > EPOCH Then
        Mo = MonthNumber(cboMonth.Text)
        yr = val(cboYear.Text)
        Set subs = db.Execute("SELECT * FROM subsidy WHERE year = " & yr & " AND month = " & Mo & " LIMIT 1")
        If (subs.EOF And subs.BOF) Then 'if there is no saved entry for this month/year then
            FlexGrid.backcolor = &HDDFFDD ' green
            fillData
            SaveButn.Enabled = True
            LOADED = False
            cboMonth.Tag = cboMonth.Text
            cboYear.Tag = cboYear.Text
        Else
            If MsgBox("A record already exists for that month! You can load it by selecting from the 'load' dropdown." & vbCrLf & "Do you want to recalculate it?", vbYesNo) = vbYes Then
                FlexGrid.backcolor = &HFFEEDD ' blue
                fillData
                SaveButn.Enabled = True
                LOADED = True
                cboMonth.Tag = cboMonth.Text
                cboYear.Tag = cboYear.Text
            End If
        End If
    Else
        SaveButn.Enabled = False
        initFlexgrid
    End If
    
    'MsgBox clientList.List(0) & " - " & FlexGrid.TextMatrix(0, 0)
    'MsgBox clientList.List(1) & " - " & FlexGrid.TextMatrix(1, 0)
    'MsgBox clientList.List(clientList.ListCount - 1) & " - " & FlexGrid.TextMatrix(clientList.ListCount - 1, 0)
    'MsgBox clientList.List(clientList.ListCount) & " - " & FlexGrid.TextMatrix(clientList.ListCount, 0)
    
    Set subs = Nothing
End Sub

Private Sub cboMonth_Change()
    cboMonth_Click
End Sub

Private Sub cboMonth_Click()
    If cboMonth.ListIndex <> -1 And cboYear.ListIndex <> -1 Then
        periodStart = CDate(cboMonth.Text & " 1, " & cboYear.Text)
        periodEnd = CDate(cboMonth.Text & " " & daysInMonth(periodStart) & ", " & cboYear.Text)
    End If
    SaveButn.Enabled = False
End Sub

Private Sub cboSaved_Click()
    If cboSaved.ListIndex > -1 Then loadButn.Enabled = True
End Sub

Private Sub cboYear_Change()
    cboYear_Click
End Sub

Private Sub cboYear_Click()
    If cboMonth.ListIndex <> -1 And cboYear.ListIndex <> -1 Then
        periodStart = CDate(cboMonth.Text & " 1, " & cboYear.Text)
        periodEnd = CDate(cboMonth.Text & " " & daysInMonth(periodStart) & ", " & cboYear.Text)
    End If
    SaveButn.Enabled = False
End Sub




Private Sub editCell_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        FlexGrid.Text = editCell.Text
        editCell.Visible = False
        FlexGrid.CellBackColor = &HCCCCFF
        If FlexGrid.col = 39 Or FlexGrid.col = 41 Then
            FlexGrid.TextMatrix(FlexGrid.row, 42) = Format(FlexGrid.TextMatrix(FlexGrid.row, 39) - FlexGrid.TextMatrix(FlexGrid.row, 41), "0.00")
        End If
        If FlexGrid.col >= 34 Then
            column_tallys
        End If
    'Else
        'MsgBox KeyCode
    End If

End Sub

Private Sub FlexGrid_Click()
    'editCell.Visible = False
End Sub

Private Sub FlexGrid_DblClick()
    If FlexGrid.row <= FlexGrid.rows - 3 Then
        If FlexGrid.ColSel <= 31 Then
            frameModAttendance.Left = FlexGrid.CellLeft + FlexGrid.Left - 60
            frameModAttendance.Top = FlexGrid.CellTop + FlexGrid.Top
            frameModAttendance.Visible = True
        Else
            editCell.width = FlexGrid.CellWidth + 60
            editCell.height = FlexGrid.CellHeight + 90
            editCell.Left = FlexGrid.CellLeft + FlexGrid.Left - 30
            editCell.Top = FlexGrid.CellTop + FlexGrid.Top - 45
            editCell.Text = FlexGrid.Text
            editCell.Visible = True
            editCell.SetFocus
            editCell.SelStart = 0
            editCell.SelLength = Len(editCell.Text)
        End If
    End If
End Sub

Private Sub FlexGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If FlexGrid.row <= FlexGrid.rows - 3 Then
        If KeyCode = 46 Then
            FlexGrid.Text = ""
        ElseIf KeyCode = 13 Then
            If FlexGrid.ColSel > 31 Then
                editCell.width = FlexGrid.CellWidth + 60
                editCell.height = FlexGrid.CellHeight + 90
                editCell.Left = FlexGrid.CellLeft + FlexGrid.Left - 30
                editCell.Top = FlexGrid.CellTop + FlexGrid.Top - 45
                editCell.Text = FlexGrid.Text
                editCell.Visible = True
                editCell.SetFocus
                editCell.SelStart = 0
                editCell.SelLength = Len(editCell.Text)
            End If
        End If
    End If
End Sub

Private Sub FlexGrid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    editCell.Visible = False
End Sub

Private Sub FlexGrid_Scroll()
    If FlexGrid.ColSel <= 31 Then
        frameModAttendance.Left = FlexGrid.CellLeft + FlexGrid.Left - 60
        frameModAttendance.Top = FlexGrid.CellTop + FlexGrid.Top
    End If
End Sub

Private Sub Form_Load()
    Dim q As ADODB.Recordset
    
    'MsgBox getFeesAtDate(37, CDate("Aug 1, 2016"))
    cboMonth.ListIndex = month(Date) - 1 ' set it to this month
    If month(Date) > 1 Then cboMonth.ListIndex = cboMonth.ListIndex - 1 ' if it's not january set it to last month
    For y = 2016 To year(Date)
        cboYear.AddItem y
    Next y
    cboYear.ListIndex = cboYear.ListCount - 1
    
    Set q = db.Execute("SELECT * FROM subsidy ORDER BY year DESC, month DESC")
    With q
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do Until .EOF
                cboSaved.AddItem MonthName(!month) & " " & !year
                .MoveNext
            Loop
        End If
    End With
    
    
    initFlexgrid
    FlexGrid.ScrollTrack = True
    Set q = Nothing
End Sub

Sub fillData()
    Dim cl As ADODB.Recordset
    Dim atn As ADODB.Recordset
    Dim sc As ADODB.Recordset
    Dim fcid As Long
    Dim i As Long
    Dim d As Date
    Dim clientFrom As Date
    Dim clientTo As Date
    Dim daysAttended As Byte
    Dim daysAbsent As Byte
    Dim daysStat As Byte
    Dim totalCost As Double
    Dim c As Byte
    Dim r As Byte
    Dim temp As Double
    Dim subs_this_month As Boolean
    Dim atn_date_error As Boolean
    Dim daysperweek As Byte
    Dim wkday As Byte
    
    If Me.Visible Then initFlexgrid
    
    Set cl = db.Execute("SELECT * FROM Clients WHERE (active = 1 OR endDate >= " & sqlDate(periodStart) & ") AND startDate <= " & sqlDate(periodEnd) & " ORDER BY Last, First ASC")
    'Set cl = db.Execute("SELECT * FROM Clients WHERE subsidized = 1 AND (active = 1 OR endDate >= " & sqlDate(periodStart) & ") AND startDate <= " & sqlDate(periodEnd) & " ORDER BY Last, First ASC")
    'Clipboard.SetText "SELECT * FROM Clients WHERE subsidized = 1 AND (active = 1 OR endDate >= " & sqlDate(periodStart) & ") AND startDate <= " & sqlDate(periodEnd) & " ORDER BY Last, First ASC"
    'MsgBox "SELECT * FROM Clients WHERE subsidized = 1 AND (active = 1 OR endDate >= " & sqlDate(periodStart) & ") AND startDate <= " & sqlDate(periodEnd) & " ORDER BY Last, First ASC"
    i = 1

    With cl
        If Not (.EOF And .BOF) Then
            .MoveFirst
            
            Do Until .EOF
                'MsgBox cl.RecordCount & " Begin   " & .EOF
                subs_this_month = False
                
                Set atn = db.Execute("SELECT * FROM client_changes WHERE idClient = " & !idClient & " AND subsidized = 1")
                If atn.EOF And atn.BOF Then
                    'do nothing, this child was NEVER subsidized!
                Else
                    clientFrom = periodStart
                    clientTo = periodEnd
                    If clientFrom < !startdate Then clientFrom = !startdate
                    If clientTo > !enddate Then clientTo = !enddate
                    
                    If !subsidized = 1 Then
                        If getSubsidizedAtDate(!idClient, clientFrom) = 0 Then
                            For d = clientFrom To clientTo
                                If getSubsidizedAtDate(!idClient, d) = 1 Then
                                    clientFrom = d
                                    subs_this_month = True
                                    Exit For
                                End If
                            Next d
                        Else
                            subs_this_month = True
                        End If
                    End If
                    
                    If !subsidized = 0 Then
                        If getSubsidizedAtDate(!idClient, clientTo) = 0 Then
                            For d = clientTo To clientFrom Step -1
                                If getSubsidizedAtDate(!idClient, d) = 1 Then
                                    clientTo = d
                                    subs_this_month = True
                                    Exit For
                                End If
                            Next d
                        End If
                    End If
                End If
                
                
                If subs_this_month Then
                    FlexGrid.rows = i + 1
                    FlexGrid.row = i
                    FlexGrid.TextMatrix(i, 0) = !Last & ", " & !First
                    clientList.AddItem !idClient, i
                    FlexGrid.TextMatrix(i, 32) = !parent1
                    FlexGrid.TextMatrix(i, 33) = "" & !authorizationNumber
                    FlexGrid.TextMatrix(i, 34) = Format(clientFrom, "mmm d")
                    FlexGrid.TextMatrix(i, 35) = Format(clientTo, "mmm d")
                    FlexGrid.TextMatrix(i, 36) = "0"
                    FlexGrid.TextMatrix(i, 37) = "0"
                    FlexGrid.TextMatrix(i, 38) = "0"
                    FlexGrid.TextMatrix(i, 40) = "N/A"
                    FlexGrid.TextMatrix(i, 43) = "N/A"
                    
                    daysAttended = 0
                    daysAbsent = 0
                    daysStat = 0
                    totalCost = 0
                    atn_date_error = False
                    
                    Set atn = db.Execute("SELECT * FROM attendance WHERE idClient = " & !idClient & " AND date >= " & sqlDate(periodStart) & " AND date <= " & sqlDate(periodEnd) & "ORDER BY date ASC")
                    If Not (atn.EOF And atn.BOF) Then
                        atn.MoveFirst
                        For c = 1 To daysInMonth(periodStart) ' column (c)
                            d = DateSerial(val(cboYear.Text), cboMonth.ListIndex + 1, c)
                            fcid = getFeeClassAtDate(!idClient, d)
                            Set fc = db.Execute("SELECT * FROM fee_classes WHERE idFeeClasses = " & fcid)
                            daysperweek = fc!days_per_week
                            wkday = Weekday(d)
                            
                            If Not atn.EOF Then ' this moves the data pointer ahead to the record for the current date (d)
                                Do While d > atn!Date
                                    'MsgBox "looking for " & d & " but got attendance for " & atn!Date
                                    atn.MoveNext
                                    If atn.EOF Then Exit Do 'atn_date_error = True
                                Loop
                            End If
                            
                            'If atn_date_error Then
                            '    FlexGrid.TextMatrix(i, c) = "X" ' DATE ERROR
                            'Else
                                'MsgBox d
                            If Not isWeekend(d) And d >= EPOCH Then
                                If d >= clientFrom And d <= clientTo Then
                                    If fc.Fields(weekdayToLetter(wkday)) And d >= !startdate Then 'if billed for today
                                    'If ((wkday = 2 And fc!M = 1) Or (wkday = 3 And fc!T = 1) Or (wkday = 4 And fc!W = 1) Or (wkday = 5 And fc!h = 1) Or (wkday = 6 And fc!f = 1)) And d >= !startdate Then 'if billed for today
                                        totalCost = totalCost + getFeesAtDate(!idClient, d) / daysperweek 'CALCULATE DAILY CHARGES
                                        'If !idClient = 36 Then MsgBox totalCost
                                        If isStatHoliday(d) Then
                                            FlexGrid.TextMatrix(i, c) = "SH" ' STAT HOLIDAY
                                            FlexGrid.TextMatrix(i, 38) = val(FlexGrid.TextMatrix(i, 38)) + 1
                                        ElseIf atn.EOF Then
                                            FlexGrid.TextMatrix(i, c) = "A" ' THERE ARE NO MORE ATTENDANCE RECORDS
                                            FlexGrid.TextMatrix(i, 37) = val(FlexGrid.TextMatrix(i, 37)) + 1
                                        Else
                                            If atn!attended = 0 Then
                                                FlexGrid.TextMatrix(i, c) = "A" ' CHILD MARKED AS ABSENT TODAY
                                                FlexGrid.TextMatrix(i, 37) = val(FlexGrid.TextMatrix(i, 37)) + 1
                                            Else 'ATTENDED
                                                'MsgBox atn!signin & "  " & Hour(atn!signin)
                                                If d > getLDOS(val(cboYear)) And d < getFDOS(val(cboYear)) Then
                                                    FlexGrid.TextMatrix(i, c) = "P" ' CHILD ATTENDED TODAY _ SUMMER SO JUST MARK PRESENT
                                                    FlexGrid.TextMatrix(i, 36) = val(FlexGrid.TextMatrix(i, 36)) + 1
                                                Else
                                                    If fcid = 4 Or fcid = 5 Or fcid = 7 Then
                                                        If Hour(atn!signin) > 14 Or (Hour(atn!signin) = 14 And Minute(atn!signin) > 30) Then
                                                            FlexGrid.TextMatrix(i, c) = "AS" ' CHILD ATTENDED TODAY _ ARRIVED AFTER SCHOOL (@ > 2:30PM )
                                                        Else
                                                            FlexGrid.TextMatrix(i, c) = "P" ' CHILD ATTENDED FULL DAY
                                                        End If
                                                    'ElseIf feeclassid is the social program Hour(atn!signout) < 12 Then
                                                    '    FlexGrid.TextMatrix(i, c) = "H" ' CHILD ATTENDED ONLY MORNING _ SOCIAL PROGRAM
                                                        
                                                    Else
                                                        FlexGrid.TextMatrix(i, c) = "P" ' CHILD ATTENDED AND IS IN A FEE CLASS THAT IS BILLED FOR FULL DAYS ONLY
                                                    End If
                                                    FlexGrid.TextMatrix(i, 36) = val(FlexGrid.TextMatrix(i, 36)) + 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        FlexGrid.TextMatrix(i, c) = "/" ' NOT A SCHEDULED DAY FOR A PART TIME CLIENT.
                                        'FlexGrid.TextMatrix(i, 37) = val(FlexGrid.TextMatrix(i, 37)) + 1
                                    End If
                                Else
                                    FlexGrid.TextMatrix(i, c) = "/" ' THE CHILD HAS NOT STARTED YET OR HAS ALREADY FINISHED
                                End If
                                    
                                'Turn A's Red
                                If FlexGrid.TextMatrix(i, c) = "A" Then
                                    FlexGrid.col = c
                                    FlexGrid.CellForeColor = vbRed
                                End If
                                
                                'INFANT CARE
                                If FlexGrid.TextMatrix(i, c) = "P" And getAgeM(cl!DOB, d) < 24 Then FlexGrid.TextMatrix(i, c) = "IC"
                            
                        
                                Set sc = db.Execute("SELECT * FROM school_closures WHERE date = " & sqlDate(d) & " ORDER BY type DESC LIMIT 1")
                                If Not (sc.EOF And sc.BOF) Then
                                    sc.MoveFirst
                                    'MsgBox !First & kindergartenAge(!DOB, d)
                                    'Dim j As Byte
                                    'For j = 1 To FlexGrid.rows - 2
                                        If FlexGrid.TextMatrix(i, c) = "AS" Or FlexGrid.TextMatrix(i, c) = "P" Then
                                            If isSchoolAgeClass(getFeeClassAtDate(clientList.List(i), CDate(cboMonth.Text & " " & c & ", " & cboYear.Text))) Then
                                                If sc!Type = "SC" Then
                                                    FlexGrid.TextMatrix(i, c) = "SC"
                                                ElseIf kindergartenAge(!DOB, d) Then
                                                    FlexGrid.TextMatrix(i, c) = "KC"
                                                End If
                                            End If
                                        End If
                                    'Next j
                                End If
                            
                            
                            Else
                                FlexGrid.TextMatrix(i, c) = "" ' WEEKEND
                            End If
                            
                        Next c
                    End If
                    
                    FlexGrid.TextMatrix(i, 39) = Format(totalCost, "0.00") 'total billable
                    FlexGrid.TextMatrix(i, 41) = Format(!parentalContribution, "0.00") 'Parental Contribution
                    FlexGrid.TextMatrix(i, 42) = Format(totalCost - !parentalContribution, "0.00") ' totalcost - parental contribution.
                
                    i = i + 1
                End If
                
                'MsgBox cl.RecordCount
                .MoveNext
                'MsgBox cl.RecordCount & " End   " & .EOF
                
                
            Loop
        End If
    End With
    
    FlexGrid.rows = FlexGrid.rows + 2
    'For c = 36 To 43
    '    temp = 0
    '    For r = 1 To FlexGrid.rows - 3
    '        temp = temp + val(FlexGrid.TextMatrix(r, c))
    '    Next r
    '    If c <= 38 Then
    '        FlexGrid.TextMatrix(FlexGrid.rows - 1, c) = Int(temp)
    '    Else
    '        FlexGrid.TextMatrix(FlexGrid.rows - 1, c) = Format(temp, "0.00")
    '    End If
    '
    'Next c
    tallys
    
    
    Set cl = Nothing
    Set atn = Nothing
    Set sc = Nothing
End Sub

Sub loadData(ByVal Mo As Byte, ByVal yr As Long)
    Dim subs As ADODB.Recordset
    Dim subs_ent As ADODB.Recordset
    Dim cl As ADODB.Recordset
    Dim subs_id As Long
    Dim days As String
    Dim st As StringTokenizer
    Dim tok As String
    Dim j As Byte
    
    'Dim fc As Long
    'Dim i As Long
    'Dim d As Date
    'Dim clientFrom As Date
    'Dim clientTo As Date
    'Dim daysAttended As Byte
    'Dim daysAbsent As Byte
    'Dim daysStat As Byte
    'Dim totalCost As Double
    'Dim c As Byte
    'Dim r As Byte
    'Dim temp As Double
    'Dim subs_this_month As Boolean
    'Dim atn_date_error As Boolean
    
    If Me.Visible Then initFlexgrid
    
    Set subs = db.Execute("SELECT * FROM subsidy WHERE year = " & yr & " AND month = " & Mo & " LIMIT 1")
    If Not (subs.EOF And subs.BOF) Then
        subs.MoveFirst
        subs_id = subs!idSubsidy
    End If
    Set subs_ent = db.Execute("SELECT * FROM subsidy_entries WHERE idSubsidy = " & subs_id)
    
    i = 1
    With subs_ent
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do Until .EOF
                FlexGrid.rows = i + 1
                FlexGrid.row = i
                clientList.AddItem !idClient
                
                Set cl = db.Execute("SELECT * FROM clients WHERE idClient = " & !idClient)
                FlexGrid.TextMatrix(i, 0) = cl!Last & ", " & cl!First
                
                j = 1
                days = !day_codes
                Set st = New StringTokenizer
                st.Init days, ","
                Do Until Not st.hasMoreTokens
                    tok = st.nextToken
                    FlexGrid.TextMatrix(i, j) = tok
                    j = j + 1
                Loop
                
                FlexGrid.TextMatrix(i, 32) = !Parent
                FlexGrid.TextMatrix(i, 33) = "" & !auth
                FlexGrid.TextMatrix(i, 34) = !From
                FlexGrid.TextMatrix(i, 35) = !To
                FlexGrid.TextMatrix(i, 36) = !attended
                FlexGrid.TextMatrix(i, 37) = !absent
                FlexGrid.TextMatrix(i, 38) = !STAT
                FlexGrid.TextMatrix(i, 39) = Format(!total_cost, "0.00")
                FlexGrid.TextMatrix(i, 40) = "N/A"
                FlexGrid.TextMatrix(i, 41) = Format(!parental, "0.00")
                FlexGrid.TextMatrix(i, 42) = Format(!pay, "0.00")
                FlexGrid.TextMatrix(i, 43) = "N/A"
                
                i = i + 1
                .MoveNext
            Loop
        End If
    End With
    
    FlexGrid.rows = FlexGrid.rows + 2
    For c = 36 To 43
        temp = 0
        For r = 1 To FlexGrid.rows - 3
            temp = temp + val(FlexGrid.TextMatrix(r, c))
        Next r
        If c <= 38 Then
            FlexGrid.TextMatrix(FlexGrid.rows - 1, c) = Int(temp)
        Else
            FlexGrid.TextMatrix(FlexGrid.rows - 1, c) = Format(temp, "0.00")
        End If
        
    Next c
    
    
    Set cl = Nothing
    Set subs = Nothing
    Set subs_ent = Nothing
    Set st = Nothing
End Sub

Sub tallys()
    Dim P As Long
    Dim a As Long
    Dim sh As Long
    Dim temp As Long
    Dim r As Long
    Dim c As Long
    Dim d As Date
    Dim cl As ADODB.Recordset
    
    For r = 1 To FlexGrid.rows - 3
        a = 0
        P = 0
        sh = 0
        totalCost = 0
        Set cl = db.Execute("SELECT * FROM Clients WHERE CONCAT(Last,', ',First) = """ & FlexGrid.TextMatrix(r, 0) & """ ORDER BY startDate DESC ")
        If cl.EOF And cl.BOF Then Exit For
        'Clipboard.SetText ()
        'MsgBox cl!Last & ", " & cl!First
        
        For c = 1 To 31
            d = DateSerial(val(cboYear.Text), cboMonth.ListIndex + 1, c)
            If FlexGrid.TextMatrix(r, c) = "A" Then a = a + 1
            If FlexGrid.TextMatrix(r, c) = "S" Then a = a + 1 'sick
            If FlexGrid.TextMatrix(r, c) = "P" Then P = P + 1
            If FlexGrid.TextMatrix(r, c) = "H" Then P = P + 1
            If FlexGrid.TextMatrix(r, c) = "AS" Then P = P + 1
            If FlexGrid.TextMatrix(r, c) = "IC" Then P = P + 1
            If FlexGrid.TextMatrix(r, c) = "SC" Then P = P + 1
            If FlexGrid.TextMatrix(r, c) = "KC" Then P = P + 1
            If FlexGrid.TextMatrix(r, c) = "SH" Then sh = sh + 1
            
            
            If FlexGrid.TextMatrix(r, c) = "/" Or Trim(FlexGrid.TextMatrix(r, c) = "") Then 'weekend, not scheduled, or before start or after end
                'DO NOTHING
            'ElseIf FlexGrid.TextMatrix(r, c) = "A" Then
                'essentially we bill for absent days too so just let this case fall under 'else'
            ElseIf FlexGrid.TextMatrix(r, c) = "SC" Or FlexGrid.TextMatrix(r, c) = "KC" Then
                totalCost = totalCost + 30 'school closure or kindercare (billed at $30 per day)
            ElseIf FlexGrid.TextMatrix(r, c) = "SH" Then
                
                'MsgBox getFeesAtDate(cl!idClient, d) / 5
                If getFeesAtDate(cl!idClient, d) / feeClassDaysPerWeek(getFeeClassAtDate(cl!idClient, d)) < 30 Then
                    totalCost = totalCost + 30 'STAT HOLIDAY (school age billed at $30 per day)
                Else
                    totalCost = totalCost + getFeesAtDate(cl!idClient, d) / feeClassDaysPerWeek(getFeeClassAtDate(cl!idClient, d)) 'CALCULATE DAILY CHARGES
                End If
                
                
            Else
                totalCost = totalCost + getFeesAtDate(cl!idClient, d) / feeClassDaysPerWeek(getFeeClassAtDate(cl!idClient, d)) 'CALCULATE DAILY CHARGES
            End If
            
        Next c
        FlexGrid.TextMatrix(r, 36) = P
        FlexGrid.TextMatrix(r, 37) = a
        FlexGrid.TextMatrix(r, 38) = sh
        
        
        'MsgBox totalCost
        
        FlexGrid.TextMatrix(r, 39) = Format(totalCost, "0.00") 'total billable
        FlexGrid.TextMatrix(r, 41) = Format(cl!parentalContribution, "0.00") 'Parental Contribution
        FlexGrid.TextMatrix(r, 42) = Format(totalCost - cl!parentalContribution, "0.00") ' totalcost - parental contribution.
    
    Next r

    column_tallys
    
    Set cl = Nothing
End Sub

Sub column_tallys()
    Dim temp As Double
    Dim r As Long
    Dim c As Long
    

    For c = 36 To 43
        temp = 0
        For r = 1 To FlexGrid.rows - 3
            temp = temp + val(FlexGrid.TextMatrix(r, c))
        Next r
        If c <= 38 Then
            FlexGrid.TextMatrix(FlexGrid.rows - 1, c) = Int(temp)
        Else
            FlexGrid.TextMatrix(FlexGrid.rows - 1, c) = Format(temp, "0.00")
        End If
    Next c

End Sub

Sub initFlexgrid()
    Dim c As Long
    Dim y As Long
    
    FlexGrid.Clear
    clientList.Clear
    clientList.AddItem "Client ID's", 0
    
    FlexGrid.ColWidth(0) = 2000
    FlexGrid.TextMatrix(0, 0) = "Child's Name"
    For c = 1 To 31
        FlexGrid.ColWidth(c) = 300
        FlexGrid.TextMatrix(0, c) = c
    Next c
    For c = 34 To 43
        FlexGrid.ColWidth(c) = 800
    Next c
    FlexGrid.ColWidth(32) = 2000
    FlexGrid.TextMatrix(0, 32) = "Parent"
    FlexGrid.ColWidth(33) = 1400
    FlexGrid.TextMatrix(0, 33) = "Authorization #"
    FlexGrid.TextMatrix(0, 34) = "From"
    FlexGrid.TextMatrix(0, 35) = "To"
    FlexGrid.TextMatrix(0, 36) = "Attended"
    FlexGrid.TextMatrix(0, 37) = "Absent"
    FlexGrid.TextMatrix(0, 38) = "Stat"
    FlexGrid.TextMatrix(0, 39) = "Cost"
    FlexGrid.TextMatrix(0, 40) = "Trans"
    FlexGrid.TextMatrix(0, 41) = "Parental"
    FlexGrid.TextMatrix(0, 42) = "Pay"
    FlexGrid.TextMatrix(0, 43) = "Trans"

End Sub

Private Sub kcButn_Click()
    dlgSelectDay.labMo.Tag = cboMonth.ListIndex + 1
    dlgSelectDay.Caption = "Kinder Care Day"
    dlgSelectDay.Show 1
    If statDay > 0 Then
        'MsgBox statDay & "   " & CDate(Year(Date) & "-" & cboMonth.ListIndex + 1 & "-" & Format(statDay, "00"))
        Dim i As Byte
        For i = 1 To FlexGrid.rows - 2
            'If FlexGrid.TextMatrix(i, statDay) <> "" And FlexGrid.TextMatrix(i, statDay) <> "/" Then
            If FlexGrid.TextMatrix(i, statDay) = "AS" Or FlexGrid.TextMatrix(i, statDay) = "P" Then
                If isSchoolAgeClass(getFeeClassAtDate(clientList.List(i), CDate(cboMonth.Text & " " & statDay & ", " & cboYear.Text))) Then
                    FlexGrid.TextMatrix(i, statDay) = "KC"
                End If
            End If
        Next i
    End If
    tallys
End Sub

Private Sub loadButn_Click()
    Dim Mo As Byte
    Dim yr As Long
    'MsgBox clientList.List(1) & " " & FlexGrid.TextMatrix(1, 0) ' this shows that the flexgrid index and the clientlist index is the same.
    
    Mo = MonthNumber(MiD$(cboSaved.Text, 1, InStr(1, cboSaved.Text, " ")))
    yr = val(MiD$(cboSaved.Text, InStr(1, cboSaved.Text, " ")))
    cboMonth.ListIndex = Mo - 1
    cboYear.ListIndex = yr - 2016
    DoEvents
    cboMonth.Tag = cboMonth.Text ' or some other source of month and year
    cboYear.Tag = cboYear.Text
    
    loadData Mo, yr
    
    'if loaded then
    FlexGrid.backcolor = &HFFEEDD ' blue
    SaveButn.Enabled = True
    LOADED = True
End Sub

Private Sub modButn_Click(Index As Integer)
    FlexGrid.Text = modButn(Index).Caption
    FlexGrid.CellBackColor = &HCCCCFF
    frameModAttendance.Visible = False
    tallys
End Sub

Private Sub prntButn_Click()
    printSubsidy
End Sub

Private Sub SaveButn_Click()
    Dim sql As String
    Dim r As Long
    Dim c As Byte
    Dim q As ADODB.Recordset
    Dim subsidy_id As Long
    Dim day_codes As String
    Dim Mo As Byte
    Dim yr As Long
    Dim subs As ADODB.Recordset
    Dim subs_id As Long
    Dim clients_list() As Long
    Dim found As Boolean
    Dim ID As String
    Dim idq As ADODB.Recordset
    Dim entid As Long
    Dim record_count As Long
    Dim guid As String
    
    SaveButn.Enabled = False

    If LOADED Then 'Update the loaded entry
        'to update subsidy entries
        '-make a list of all entries currently in DB for that subsidy ID.  store in clients_list array
        '-step through list of items on screen
        '    -update records that exist and remove them from db list
        '        -update subsidy_entries table
        '        -update payments table
        '        -update gcdb.transactions table
        '        -update gcdb.splits table
        '    -add new items for records that don't exist
        '        -insert into  subsidy_entries table
        '        -insert into  payments table
        '        -insert into  gcdb.transactions table
        '        -insert into  gcdb.splits table
        '-step through the db list and any items that remain don't appear on the screen so delete them.
        '    -delete from subsidy_entries table
        '    -delete from payments table
        '    -delete from gcdb.transactions table
        '    -delete from gcdb.splits table
            
        Mo = MonthNumber(cboMonth.Tag)
        yr = val(cboYear.Tag)
        Set subs = db.Execute("SELECT * FROM subsidy WHERE year = " & yr & " AND month = " & Mo & " LIMIT 1")
        If Not (subs.EOF And subs.BOF) Then
            subs.MoveFirst
            subs_id = subs!idSubsidy
        End If
        db.Execute "UPDATE subsidy SET Total = " & FlexGrid.TextMatrix(FlexGrid.rows - 1, 42) & " WHERE idSubsidy = " & subs_id
        
        '-make a list of all entries currently in DB for that subsidy ID
        Set q = db.Execute("SELECT * FROM subsidy_entries WHERE idSubsidy = " & subs_id & ";")
        With q
            If Not (.EOF And .BOF) Then
                record_count = 0
                .MoveFirst
                Do Until .EOF
                    record_count = record_count + 1
                    .MoveNext
                Loop
                ReDim clients_list(record_count) As Long
                record_count = 0
                .MoveFirst
                Do Until .EOF
                    record_count = record_count + 1
                    clients_list(record_count) = !idClient
                    .MoveNext
                Loop
            End If
        End With
        
        
        '-step through list of items on screen
        For r = 1 To clientList.ListCount - 1
            Set q = db.Execute("SELECT * FROM subsidy_entries WHERE idSubsidy = " & subs_id & " AND idClient = " & clientList.List(r))
            With q
                If Not (.EOF And .BOF) Then
        '-update records that exist and remove them from db list
                
        '        -update subsidy_entries table
                    sql = "UPDATE subsidy_entries SET "
                    day_codes = ""
                    For c = 1 To 31
                        If c > 1 Then day_codes = day_codes & ","
                        day_codes = day_codes & FlexGrid.TextMatrix(r, c)
                    Next c
                    sql = sql & "day_codes=""" & day_codes & ""","
                    sql = sql & "parent=""" & FlexGrid.TextMatrix(r, 32) & ""","
                    sql = sql & "auth=""" & FlexGrid.TextMatrix(r, 33) & ""","
                    sql = sql & "subsidy_entries.from=""" & FlexGrid.TextMatrix(r, 34) & ""","
                    sql = sql & "subsidy_entries.to=""" & FlexGrid.TextMatrix(r, 35) & ""","
                    sql = sql & "attended=" & val(FlexGrid.TextMatrix(r, 36)) & ","
                    sql = sql & "absent=" & val(FlexGrid.TextMatrix(r, 37)) & ","
                    sql = sql & "stat=" & val(FlexGrid.TextMatrix(r, 38)) & ","
                    sql = sql & "total_cost=" & val(FlexGrid.TextMatrix(r, 39)) & ","
                    sql = sql & "total_trans=" & val(FlexGrid.TextMatrix(r, 40)) & ","
                    sql = sql & "parental=" & val(FlexGrid.TextMatrix(r, 41)) & ","
                    sql = sql & "pay=" & val(FlexGrid.TextMatrix(r, 42)) & ","
                    sql = sql & "trans=" & val(FlexGrid.TextMatrix(r, 43)) & " "
                    sql = sql & "WHERE idEntry = " & !idEntry & ";"
                    db.Execute sql
                    
                    
        '        -update a payment in the receipts table
                    Set idq = db.Execute("SELECT * FROM payments WHERE idClient=" & clients_list(r) & " AND fromdate=" & sqlDate(CDate(FlexGrid.TextMatrix(r, 34) & ", " & cboYear.Text)) & " AND todate=" & sqlDate(CDate(FlexGrid.TextMatrix(r, 35) & ", " & cboYear.Text)) & " LIMIT 1")
                    If Not (idq.EOF And idq.BOF) Then
                        .MoveFirst
                        guid = idq!guid
                        sql = "UPDATE payments SET "
                        sql = sql & "receivedFrom=" & """Subsidy (" & !idEntry & ")"","
                        sql = sql & "date=" & sqlDate(Date) & ","
                        sql = sql & "amount=" & FlexGrid.TextMatrix(r, 42) & ","
                        sql = sql & "attendance=" & """"","
                        sql = sql & "details=" & """Subsidy Automatic Entry"""
                        sql = sql & " WHERE guid = """ & guid & """"
                        'sql = sql & " WHERE idClient=" & clients_list(r) & " AND fromdate=" & sqlDate(CDate(FlexGrid.TextMatrix(r, 34) & ", " & cboYear.Text)) & " AND todate=" & sqlDate(CDate(FlexGrid.TextMatrix(r, 35) & ", " & cboYear.Text)) & " LIMIT 1"
                        db.Execute sql
                    
        '        -update gcdb.transactions table
        '        -update gcdb.splits table
                        update_gnc_receipt guid, FlexGrid.TextMatrix(r, 0) & " -- " & shortDate(CDate(FlexGrid.TextMatrix(r, 34) & ", " & cboYear.Text)) & " - " & shortDate(CDate(FlexGrid.TextMatrix(r, 35) & ", " & cboYear.Text)), val(FlexGrid.TextMatrix(r, 42)), Date
                    End If
                    'remove from the list of entries
                    For c = 1 To record_count
                        If clients_list(c) = !idClient Then
                            clients_list(c) = 0
                            Exit For
                        End If
                    Next c
                Else
        '    -add new items for records that don't exist
                    
                    'if the item on screen isn't 0 then add it to the database.
                    If clientList.List(r) <> 0 Then
        '        -insert into  subsidy_entries table
                        sql = "INSERT INTO subsidy_entries (idSubsidy,idClient,day_codes,parent,auth,subsidy_entries.from,subsidy_entries.to,attended,absent,stat,total_cost,total_trans,parental,pay,trans) VALUES ("
                        day_codes = ""
                        For c = 1 To 31
                            If c > 1 Then day_codes = day_codes & ","
                            day_codes = day_codes & FlexGrid.TextMatrix(r, c)
                        Next c
                        sql = sql & subs_id & ","
                        sql = sql & clientList.List(r) & ","
                        sql = sql & """" & day_codes & ""","
                        For c = 32 To 35
                            sql = sql & """" & FlexGrid.TextMatrix(r, c) & ""","
                        Next c
                        For c = 36 To 42
                            sql = sql & val(FlexGrid.TextMatrix(r, c)) & ","
                        Next c
                        sql = sql & val(FlexGrid.TextMatrix(r, 43)) & ");"
                        'MsgBox sql
                        
                        db.Execute sql
                        db.Execute "SET @subentid = LAST_INSERT_ID();"
                        Set idq = db.Execute("SELECT @subentid AS ent;")
                        entid = idq!ent
                            
        '        -insert into  payments table
                        'save a payment in the receipts table and in gnucash
                        guid = createGUID
                        sql = "INSERT INTO payments ("
                        sql = sql & "guid,idClient,receivedFrom,date,fromdate,todate,amount,attendance,details) VALUES ("
                        sql = sql & """" & guid & ""","
                        sql = sql & clientList.List(r) & ","
                        sql = sql & """Subsidy (" & entid & ")"","
                        sql = sql & sqlDate(Date) & ","
                        sql = sql & sqlDate(CDate(FlexGrid.TextMatrix(r, 34) & ", " & cboYear.Text)) & ","  'from date
                        sql = sql & sqlDate(CDate(FlexGrid.TextMatrix(r, 35) & ", " & cboYear.Text)) & "," 'to date
                        sql = sql & FlexGrid.TextMatrix(r, 42) & ","
                        sql = sql & """"","
                        sql = sql & """Subsidy Automatic Entry"")"
                        db.Execute sql
                        
        '        -insert into  gcdb.transactions table
        '        -insert into  gcdb.splits table
                        create_gnc_receipt guid, FlexGrid.TextMatrix(r, 0) & " -- " & shortDate(CDate(FlexGrid.TextMatrix(r, 34) & ", " & cboYear.Text)) & " - " & shortDate(CDate(FlexGrid.TextMatrix(r, 35) & ", " & cboYear.Text)), val(FlexGrid.TextMatrix(r, 42)), Date, True
                    End If
                End If
            End With
        Next r



        '-step through the db list and any items that remain don't appear on the screen so delete them.
        For c = 1 To record_count
            If clients_list(c) <> 0 Then
                Set q = db.Execute("SELECT * FROM subsidy_entries WHERE idSubsidy = " & subs_id & " AND idClient = " & clients_list(c))
                If Not (q.EOF And q.BOF) Then
                    entid = q!idEntry
        
        '    -delete from subsidy_entries table
                    db.Execute "DELETE FROM subsidy_entries WHERE idEntry=" & entid
        
        '    -delete from payments table
                    Set idq = db.Execute("SELECT * FROM payments WHERE idClient=" & clients_list(c) & " AND fromdate=" & sqlDate(CDate(q!From & ", " & cboYear.Text)) & " AND todate=" & sqlDate(CDate(q!To & ", " & cboYear.Text)) & " LIMIT 1")
                    If Not (idq.EOF And idq.BOF) Then
                        idq.MoveFirst
                        guid = idq!guid
                    
                        'db.Execute "DELETE FROM payments WHERE idClient=" & clients_list(c) & " AND fromdate=" & sqlDate(CDate(q!From & ", " & cboYear.Text)) & " AND todate=" & sqlDate(CDate(q!to & ", " & cboYear.Text)) & " LIMIT 1"
                        db.Execute "DELETE FROM payments WHERE guid=""" & guid & """"
                    
        '    -delete from gcdb.transactions table
        '    -delete from gcdb.splits table
                        delete_gnc_transaction guid
                    End If
                End If
            End If
        Next c
        
    Else ' Create a new subsidy entry
        db.Execute "INSERT INTO subsidy (Month, Year, Total) VALUES (" & MonthNumber(cboMonth.Text) & ", " & cboYear.Text & ", " & FlexGrid.TextMatrix(FlexGrid.rows - 1, 42) & ")"
        db.Execute "SET @subid = LAST_INSERT_ID();"
        For r = 1 To FlexGrid.rows - 3
            sql = "INSERT INTO subsidy_entries (idSubsidy,idClient,day_codes,parent,auth,subsidy_entries.from,subsidy_entries.to,attended,absent,stat,total_cost,total_trans,parental,pay,trans) VALUES ("
            day_codes = ""
            For c = 1 To 31
                If c > 1 Then day_codes = day_codes & ","
                day_codes = day_codes & FlexGrid.TextMatrix(r, c)
            Next c
            sql = sql & "@subid," '1 & "," 'subsidy_id & ","
            sql = sql & clientList.List(r) & ","
            sql = sql & """" & day_codes & ""","
            For c = 32 To 35
                sql = sql & """" & FlexGrid.TextMatrix(r, c) & ""","
            Next c
            For c = 36 To 42
                sql = sql & val(FlexGrid.TextMatrix(r, c)) & ","
            Next c
            sql = sql & val(FlexGrid.TextMatrix(r, 43)) & ");"
            'MsgBox sql
            db.Execute sql
            db.Execute "SET @subentid = LAST_INSERT_ID();"
            Set idq = db.Execute("SELECT @subentid AS ent;")
            entid = idq!ent
            
            'save a payment in the receipts table and in gnucash
            guid = createGUID
            sql = "INSERT INTO payments ("
            sql = sql & "guid,idClient,receivedFrom,date,fromdate,todate,amount,attendance,details) VALUES ("
            sql = sql & """" & guid & ""","
            sql = sql & clientList.List(r) & ","
            sql = sql & """Subsidy (" & entid & ")"","
            sql = sql & sqlDate(Date) & ","
            sql = sql & sqlDate(CDate(FlexGrid.TextMatrix(r, 34) & ", " & cboYear.Text)) & ","  'from date
            sql = sql & sqlDate(CDate(FlexGrid.TextMatrix(r, 35) & ", " & cboYear.Text)) & "," 'to date
            sql = sql & FlexGrid.TextMatrix(r, 42) & ","
            sql = sql & """"","
            sql = sql & """Subsidy Automatic Entry"")"
            db.Execute sql
            
            db.Execute ("UPDATE attendance SET paid = 1 WHERE idClient = " & clientList.List(r) & " AND date >= " & sqlDate(CDate(FlexGrid.TextMatrix(r, 34) & ", " & cboYear.Text)) & " AND date <= " & sqlDate(CDate(FlexGrid.TextMatrix(r, 35) & ", " & cboYear.Text)))
            
            create_gnc_receipt guid, FlexGrid.TextMatrix(r, 0) & " -- " & shortDate(CDate(FlexGrid.TextMatrix(r, 34) & ", " & cboYear.Text)) & " - " & shortDate(CDate(FlexGrid.TextMatrix(r, 35) & ", " & cboYear.Text)), val(FlexGrid.TextMatrix(r, 42)), Date, True
        Next r
    End If
    
    FlexGrid.backcolor = &HFFDDFF ' pink
    Set q = Nothing
    Set idq = Nothing
    Set subs = Nothing
    
End Sub

Private Sub scButn_Click()
    dlgSelectDay.labMo.Tag = cboMonth.ListIndex + 1
    dlgSelectDay.Caption = "School Closure???"
    dlgSelectDay.Show 1
    If statDay > 0 Then
        'MsgBox statDay & "   " & CDate(Year(Date) & "-" & cboMonth.ListIndex + 1 & "-" & Format(statDay, "00"))
        Dim i As Byte
        For i = 1 To FlexGrid.rows - 2
            'If FlexGrid.TextMatrix(i, statDay) <> "" And FlexGrid.TextMatrix(i, statDay) <> "/" Then
            If FlexGrid.TextMatrix(i, statDay) = "AS" Or FlexGrid.TextMatrix(i, statDay) = "P" Then
                If isSchoolAgeClass(getFeeClassAtDate(clientList.List(i), CDate(cboMonth.Text & " " & statDay & ", " & cboYear.Text))) Then
                    FlexGrid.TextMatrix(i, statDay) = "SC"
                End If
            End If
        Next i
    End If
    tallys
End Sub

Private Sub statButn_Click()
    dlgSelectDay.labMo.Tag = cboMonth.ListIndex + 1
    dlgSelectDay.Caption = "Statutory Holiday"
    dlgSelectDay.Show 1
    If statDay > 0 Then
        'MsgBox statDay & "   " & CDate(Year(Date) & "-" & cboMonth.ListIndex + 1 & "-" & Format(statDay, "00"))
        Dim i As Byte
        For i = 1 To FlexGrid.rows - 2
            If FlexGrid.TextMatrix(i, statDay) <> "" And FlexGrid.TextMatrix(i, statDay) <> "/" Then
                FlexGrid.TextMatrix(i, statDay) = "SH"
                FlexGrid.row = i
                FlexGrid.col = statDay
                FlexGrid.CellBackColor = &HCCCCFF
            End If
        Next i
    End If
    tallys
End Sub
