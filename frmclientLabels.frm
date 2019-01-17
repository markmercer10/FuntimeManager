VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClientLabels 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client Labels"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   6075
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkInactive 
      Caption         =   "Show Inactive"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   37
      Top             =   120
      Width           =   1215
   End
   Begin VB.CheckBox collate 
      Caption         =   "Collate"
      Height          =   255
      Left            =   4800
      TabIndex        =   36
      Top             =   7560
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton startcell 
      Caption         =   "30"
      Height          =   255
      Index           =   29
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   5760
      Width           =   560
   End
   Begin VB.OptionButton startcell 
      Caption         =   "29"
      Height          =   255
      Index           =   28
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   5520
      Width           =   560
   End
   Begin VB.OptionButton startcell 
      Caption         =   "28"
      Height          =   255
      Index           =   27
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   5280
      Width           =   560
   End
   Begin VB.OptionButton startcell 
      Caption         =   "27"
      Height          =   255
      Index           =   26
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5040
      Width           =   560
   End
   Begin VB.OptionButton startcell 
      Caption         =   "26"
      Height          =   255
      Index           =   25
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   4800
      Width           =   560
   End
   Begin VB.OptionButton startcell 
      Caption         =   "25"
      Height          =   255
      Index           =   24
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   4560
      Width           =   560
   End
   Begin VB.OptionButton startcell 
      Caption         =   "24"
      Height          =   255
      Index           =   23
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4320
      Width           =   560
   End
   Begin VB.OptionButton startcell 
      Caption         =   "23"
      Height          =   255
      Index           =   22
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4080
      Width           =   560
   End
   Begin VB.OptionButton startcell 
      Caption         =   "22"
      Height          =   255
      Index           =   21
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3840
      Width           =   560
   End
   Begin VB.OptionButton startcell 
      Caption         =   "21"
      Height          =   255
      Index           =   20
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3600
      Width           =   560
   End
   Begin VB.OptionButton startcell 
      Caption         =   "20"
      Height          =   255
      Index           =   19
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3360
      Width           =   560
   End
   Begin VB.OptionButton startcell 
      Caption         =   "19"
      Height          =   255
      Index           =   18
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3120
      Width           =   560
   End
   Begin VB.OptionButton startcell 
      Caption         =   "18"
      Height          =   255
      Index           =   17
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2880
      Width           =   560
   End
   Begin VB.OptionButton startcell 
      Caption         =   "17"
      Height          =   255
      Index           =   16
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2640
      Width           =   560
   End
   Begin VB.OptionButton startcell 
      Caption         =   "16"
      Height          =   255
      Index           =   15
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2400
      Width           =   560
   End
   Begin VB.OptionButton startcell 
      Caption         =   "15"
      Height          =   255
      Index           =   14
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5760
      Width           =   560
   End
   Begin VB.OptionButton startcell 
      Caption         =   "14"
      Height          =   255
      Index           =   13
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5520
      Width           =   560
   End
   Begin VB.OptionButton startcell 
      Caption         =   "13"
      Height          =   255
      Index           =   12
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5280
      Width           =   560
   End
   Begin VB.OptionButton startcell 
      Caption         =   "12"
      Height          =   255
      Index           =   11
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5040
      Width           =   560
   End
   Begin VB.OptionButton startcell 
      Caption         =   "11"
      Height          =   255
      Index           =   10
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4800
      Width           =   560
   End
   Begin VB.OptionButton startcell 
      Caption         =   "10"
      Height          =   255
      Index           =   9
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4560
      Width           =   560
   End
   Begin VB.OptionButton startcell 
      Caption         =   "9"
      Height          =   255
      Index           =   8
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4320
      Width           =   560
   End
   Begin VB.OptionButton startcell 
      Caption         =   "8"
      Height          =   255
      Index           =   7
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4080
      Width           =   560
   End
   Begin VB.OptionButton startcell 
      Caption         =   "7"
      Height          =   255
      Index           =   6
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3840
      Width           =   560
   End
   Begin VB.OptionButton startcell 
      Caption         =   "6"
      Height          =   255
      Index           =   5
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3600
      Width           =   560
   End
   Begin VB.OptionButton startcell 
      Caption         =   "5"
      Height          =   255
      Index           =   4
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3360
      Width           =   560
   End
   Begin VB.OptionButton startcell 
      Caption         =   "4"
      Height          =   255
      Index           =   3
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3120
      Width           =   560
   End
   Begin VB.OptionButton startcell 
      Caption         =   "3"
      Height          =   255
      Index           =   2
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   560
   End
   Begin VB.OptionButton startcell 
      Caption         =   "2"
      Height          =   255
      Index           =   1
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2640
      Width           =   560
   End
   Begin VB.OptionButton startcell 
      Caption         =   "1"
      Height          =   255
      Index           =   0
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2400
      Value           =   -1  'True
      Width           =   560
   End
   Begin VB.CommandButton CloseButn 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton sortButn 
      Caption         =   "Alphabetical"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton clearButn 
      Caption         =   "Uncheck All"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton printButn 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   6120
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListClients 
      Height          =   8775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   15478
      SortKey         =   1
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   565
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Client Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Room"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Printing Start Cell"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   35
      Top             =   1800
      Width           =   1095
   End
End
Attribute VB_Name = "frmClientLabels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strtcell As Long

Private Sub chkInactive_Click()
    populate
End Sub

Private Sub clearButn_Click()
    Dim i As Long
    For i = 1 To ListClients.ListItems.count
        ListClients.ListItems(i).Checked = False
    Next i
End Sub

Private Sub closebutn_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    populate
End Sub

Sub populate()
    Dim L As ListItem
    Dim r As Recordset
    Dim sql As String
    
    If CBool(chkInactive) Then
        sql = "SELECT * FROM Clients ORDER BY room, first, last ASC"
    Else
        sql = "SELECT * FROM Clients WHERE active >= 1 ORDER BY room, first, last ASC"
    End If
    
    ListClients.ListItems.Clear
    Set r = db.Execute(sql)
    With r
        If Not (.EOF And .BOF) Then .MoveFirst
        Do Until .EOF
            Set L = ListClients.ListItems.Add(, , !idClient)
            L.SubItems(1) = !First & " " & !Last
            L.SubItems(2) = !room
            L.Checked = True
            .MoveNext
        Loop
    End With
    
    Set r = Nothing
    Set L = Nothing
End Sub

Private Sub printButn_Click()
    Dim r As Recordset
    Dim i As Long
    Dim count As Long
    Dim textLeft As Double
    Dim cell_x As Long
    Dim cell_y As Long
    Dim space_x As Double
    Dim space_y As Double
    Dim offset_x As Double
    Dim offset_y As Double
    Dim current_x As Double
    Dim current_y As Double
    Dim label_width As Double
    
    Printer.ScaleMode = vbInches
    
    space_x = 4
    space_y = 0.6667
    offset_x = 0.5
    offset_y = 0.5
    label_width = 3.35
    
    count = strtcell + 1
    
    With ListClients
        For i = 1 To .ListItems.count
            If .ListItems(i).Checked Then
                Set r = db.Execute("SELECT * FROM clients WHERE idClient = " & .ListItems(i).Text)
                
                cell_y = ((count - 1) Mod 15)
                cell_x = (((count - 1) \ 15) Mod 2)
                current_x = space_x * cell_x + offset_x
                current_y = space_y * cell_y + offset_y
                
                'Printer.Line (current_x, current_y)-(current_x + label_width, current_y)' alignment during setup only
                                
                textLeft = current_x + 0.05
                With r
                    printText "Name:", textLeft, current_y + 0.05, 4, "Arial", 7, False, 0
                    printText !First & " " & !Last, textLeft + 0.3, current_y + 0.03, 4, "Arial", 9, True, 0
                    Dim c As ADODB.Recordset
                    Set c = getParents(!idClient)
                    If Not (c.EOF And c.BOF) Then
                        c.MoveFirst
                        printText "Parent: " & c!name & " " & getBestContactInfo(c!idContact), textLeft, current_y + 0.2, 4, "Arial", 7, False, 0
                        c.MoveNext
                        If Not c.EOF Then printText "Parent: " & c!name & " " & getBestContactInfo(c!idContact), textLeft, current_y + 0.35, 4, "Arial", 7, False, 0
                    End If
                    
                    Set c = getEmergency(!idClient)
                    If Not (c.EOF And c.BOF) Then
                        c.MoveFirst
                        printText "Emerg: " & c!name & " " & getBestContactInfo(c!idContact), textLeft, current_y + 0.5, 4, "Arial", 7, False, 0
                    End If
                    
                    printText "DoB: " & !DOB, current_x, current_y + 0.05, label_width, "Arial", 7, False, 1
                    printText "MCP: " & Int(!MCP), current_x, current_y + 0.2, label_width, "Arial", 7, False, 1
                    printText "Allergies: " & Left(!allergies, 30), current_x, current_y + 0.35, label_width, "Arial", 7, False, 1
                    
                    Set c = getDoctor(!idClient)
                    If Not (c.EOF And c.BOF) Then
                        c.MoveFirst
                        printText "Dr: " & c!name & " " & getBestContactInfo(c!idContact), current_x, current_y + 0.5, label_width, "Arial", 7, False, 1
                    End If
                End With
                
                If count Mod 30 = 0 Then
                '    If collate.value = 1 Then
                '        Printer.NewPage
                '    Else
                        Printer.EndDoc
                '    End If
                End If
                count = count + 1
            End If
        Next i
    End With
    
    Printer.EndDoc
    
    Printer.ScaleMode = vbTwips
    Set r = Nothing

End Sub


Private Sub sortButn_Click()
    ListClients.Sorted = True
End Sub


Private Sub startcell_Click(index As Integer)
    strtcell = index + 1
End Sub


