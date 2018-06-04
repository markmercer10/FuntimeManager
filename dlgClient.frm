VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form dlgClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New Client"
   ClientHeight    =   9975
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9975
   ScaleWidth      =   5760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton butnReactivate 
      BackColor       =   &H00FFBB88&
      Caption         =   "Re-Activate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   6360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer Animation2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5280
      Top             =   720
   End
   Begin VB.Frame AddContactsFrame 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5655
      Left            =   0
      TabIndex        =   46
      Top             =   2160
      Visible         =   0   'False
      Width           =   775
      Begin VB.Frame AddingContactOptions 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3975
         Left            =   0
         TabIndex        =   67
         Top             =   1080
         Visible         =   0   'False
         Width           =   3060
         Begin VB.CommandButton AddingButnBack 
            BackColor       =   &H00FFBB88&
            Caption         =   "Back"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   77
            Top             =   3120
            Width           =   1215
         End
         Begin VB.OptionButton ContactType 
            BackColor       =   &H00FFFF80&
            Caption         =   "Other"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   75
            Top             =   1920
            Width           =   2535
         End
         Begin VB.CheckBox AddingCanSignOut 
            BackColor       =   &H00FFFF00&
            Caption         =   "This person can sign the child out"
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
            Left            =   720
            TabIndex        =   74
            Top             =   2640
            Width           =   4335
         End
         Begin VB.CommandButton AddingButnSave 
            BackColor       =   &H00FFBB88&
            Caption         =   "Add"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4200
            Style           =   1  'Graphical
            TabIndex        =   73
            Top             =   3120
            Width           =   1215
         End
         Begin VB.OptionButton ContactType 
            BackColor       =   &H00FFFF80&
            Caption         =   "Doctor"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   2
            Left            =   3720
            Picture         =   "dlgClient.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   72
            Top             =   840
            Width           =   1575
         End
         Begin VB.OptionButton ContactType 
            BackColor       =   &H00FFFF80&
            Caption         =   "Emergency"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   1
            Left            =   2160
            Picture         =   "dlgClient.frx":31C7
            Style           =   1  'Graphical
            TabIndex        =   71
            Top             =   840
            Width           =   1575
         End
         Begin VB.OptionButton ContactType 
            BackColor       =   &H00FFFF80&
            Caption         =   "Parent / Guardian"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   0
            Left            =   480
            Picture         =   "dlgClient.frx":657D
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label AddingNameAs 
            BackStyle       =   0  'Transparent
            Caption         =   "<Name> as:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   960
            TabIndex        =   69
            Top             =   480
            Width           =   4455
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "Adding Contact"
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
            Left            =   720
            TabIndex        =   68
            Top             =   120
            Width           =   1815
         End
      End
      Begin VB.Frame EditContactFrame 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   4935
         Left            =   0
         TabIndex        =   55
         Top             =   480
         Visible         =   0   'False
         Width           =   4530
         Begin VB.CommandButton butnCancelSaveContact 
            BackColor       =   &H00FFBB88&
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   2640
            Width           =   1215
         End
         Begin VB.CommandButton butnSaveContact 
            BackColor       =   &H00FFBB88&
            Caption         =   "Save"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4200
            Style           =   1  'Graphical
            TabIndex        =   65
            Top             =   2640
            Width           =   1215
         End
         Begin VB.TextBox contEmail 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1800
            TabIndex        =   63
            Top             =   2040
            Width           =   3615
         End
         Begin VB.TextBox contLand 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1800
            TabIndex        =   61
            Top             =   1560
            Width           =   3615
         End
         Begin VB.TextBox contCell 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1800
            TabIndex        =   59
            Top             =   1080
            Width           =   3615
         End
         Begin VB.TextBox contName 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1800
            TabIndex        =   57
            Top             =   600
            Width           =   3615
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Email Address"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -480
            TabIndex        =   64
            Top             =   2040
            Width           =   2175
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Land Line"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -480
            TabIndex        =   62
            Top             =   1560
            Width           =   2175
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cell Number"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -480
            TabIndex        =   60
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Full Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -480
            TabIndex        =   58
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label LabEditContactInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "Create New Contact"
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
            Left            =   360
            TabIndex        =   56
            Top             =   120
            Width           =   3000
         End
      End
      Begin VB.CommandButton butnCreateContact 
         BackColor       =   &H00FFBB88&
         Caption         =   "Create"
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
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton butnCancelContact 
         BackColor       =   &H00FFBB88&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CommandButton butnInnerAddContact 
         BackColor       =   &H00FFBB88&
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton butnUpdateContact 
         BackColor       =   &H00FFBB88&
         Caption         =   "Upd."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox ContactFilter 
         BackColor       =   &H00FFFFEE&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   840
         TabIndex        =   49
         Top             =   480
         Width           =   2655
      End
      Begin MSComctlLib.ListView LV_contacts 
         Height          =   3420
         Left            =   0
         TabIndex        =   50
         Top             =   1080
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   6033
         View            =   3
         SortOrder       =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777198
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cell"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Land"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Email"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   3920
         X2              =   3800
         Y1              =   800
         Y2              =   680
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Height          =   255
         Left            =   3600
         Shape           =   3  'Circle
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Add Contact"
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
         Left            =   120
         TabIndex        =   47
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.Timer Animation 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4800
      Top             =   720
   End
   Begin VB.Frame ContactsFrame 
      BackColor       =   &H00D0D0D0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2175
      Left            =   0
      TabIndex        =   40
      Top             =   7080
      Width           =   5775
      Begin VB.ListBox contactIDs 
         Height          =   255
         Left            =   120
         TabIndex        =   76
         Top             =   1320
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox ClientContacts 
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
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   360
         Width           =   2655
      End
      Begin VB.CommandButton butnRemoveContact 
         BackColor       =   &H00FFBB88&
         Caption         =   "Remove"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton butnAddContact 
         BackColor       =   &H00FFBB88&
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label LabCanSignOut 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   78
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Contacts"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   45
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label contactInfo 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   2760
         TabIndex        =   44
         Top             =   360
         Width           =   4995
      End
   End
   Begin VB.Frame subsFrame 
      Caption         =   "Subsidization Info"
      Height          =   1215
      Left            =   2880
      TabIndex        =   35
      Top             =   5670
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
         TabIndex        =   38
         Text            =   "0.00"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtAuthNumber 
         Height          =   288
         Left            =   1320
         TabIndex        =   36
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label19 
         Caption         =   "Parental Contrib."
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "Authorization #"
         Height          =   255
         Left            =   120
         TabIndex        =   37
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
      TabIndex        =   34
      Top             =   6120
      Width           =   1455
   End
   Begin VB.TextBox ID 
      Height          =   288
      Left            =   120
      TabIndex        =   31
      Top             =   6120
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
      ItemData        =   "dlgClient.frx":9866
      Left            =   1560
      List            =   "dlgClient.frx":9876
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   5310
      Width           =   855
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
      TabIndex        =   12
      Top             =   5760
      Width           =   1452
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   4320
      Top             =   720
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
      BackColor       =   &H008080FF&
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
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9360
      Width           =   1455
   End
   Begin VB.CommandButton saveButn 
      BackColor       =   &H0080FF80&
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
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   9360
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
      TabIndex        =   13
      Top             =   6480
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
      ItemData        =   "dlgClient.frx":9886
      Left            =   1560
      List            =   "dlgClient.frx":9888
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   3975
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker dpEnd 
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   4860
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
      Format          =   124256259
      CurrentDate     =   42531
   End
   Begin MSComCtl2.DTPicker dpStart 
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   4380
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
      Format          =   124256259
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
      TabIndex        =   8
      Top             =   3600
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
      TabIndex        =   7
      Top             =   3240
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
      ItemData        =   "dlgClient.frx":988A
      Left            =   1560
      List            =   "dlgClient.frx":9894
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2160
      Width           =   855
   End
   Begin MSComCtl2.DTPicker dpDOB 
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   1680
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
      Format          =   124256259
      CurrentDate     =   42530
   End
   Begin VB.TextBox txtAllergies 
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
      Top             =   2880
      Width           =   4095
   End
   Begin VB.TextBox txtMCP 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   4105
         SubFormatType   =   1
      EndProperty
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
      MaxLength       =   12
      TabIndex        =   3
      Top             =   2520
      Width           =   1695
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
      TabIndex        =   32
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
      Format          =   107413507
      CurrentDate     =   42531
   End
   Begin VB.Label Label17 
      BackColor       =   &H000000FF&
      Caption         =   "Changes Effective                                             Ensure correct date!"
      Height          =   375
      Left            =   120
      TabIndex        =   33
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
      TabIndex        =   30
      Top             =   5355
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
      TabIndex        =   28
      Top             =   5400
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
      TabIndex        =   27
      Top             =   4005
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
      TabIndex        =   26
      Top             =   4920
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
      TabIndex        =   25
      Top             =   4440
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
      TabIndex        =   24
      Top             =   3600
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
      TabIndex        =   23
      Top             =   3240
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
      TabIndex        =   22
      Top             =   2160
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
      TabIndex        =   21
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Allergies"
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
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "MCP #"
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
      Top             =   2520
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
      TabIndex        =   18
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
      TabIndex        =   17
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
      TabIndex        =   16
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "dlgClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim originalStartDate As Date
Dim changeDateReminder As Boolean
Dim contactsLineNum As Long
Dim animate_direction As Byte
Dim animate_interval As Long
Dim animate_progress As Double

Private Sub AddingButnBack_Click()
    AddingContactOptions.Visible = False
End Sub

Private Sub AddingButnSave_Click()
    butnCreateContact.Visible = True
    AddingContactOptions.Visible = False
    
    'Delete old record
    db.Execute "DELETE FROM client_contacts WHERE idClient=" & ID.Text & " AND idContact=" & LV_contacts.ListItems(contactsLineNum).Text
    
    'Create DB Record
    db.Execute "INSERT INTO client_contacts (idClient, idContact, type, canSignOut) VALUES (" & ID.Text & "," & LV_contacts.ListItems(contactsLineNum).Text & ",""" & AddingContactOptions.Tag & """," & AddingCanSignOut.value & ")"
    
    butnCancelContact_Click
    UpdateClientContacts
End Sub

Private Sub Animation_Timer()
    Dim draw_progress As Double
    animate_progress = animate_progress + (Animation.interval / animate_interval) * CDbl(animate_direction * 2 - 1)
    draw_progress = EasingFunction(animate_progress)
    If animate_progress >= 1 And animate_direction = 1 Then
        animate_progress = 1
        Animation.Enabled = False
    End If
    If animate_progress <= 0.01 And animate_direction = 0 Then
        animate_progress = 0
        Animation.Enabled = False
        AddContactsFrame.Visible = False
        ContactFilter = ""
    End If
    AddContactsFrame.width = butnAddContact.width + (Me.width - butnAddContact.width) * draw_progress
    AddContactsFrame.Top = ContactsFrame.Top + butnAddContact.Top + -(ContactsFrame.Top + butnAddContact.Top) * draw_progress
    AddContactsFrame.height = butnAddContact.height + (Me.height - butnAddContact.height) * draw_progress
    DoEvents
End Sub

Private Sub Animation2_Timer()
    Dim draw_progress As Double
    animate_progress = animate_progress + (Animation.interval / animate_interval) * CDbl(animate_direction * 2 - 1)
    draw_progress = EasingFunction(animate_progress)
    If animate_progress >= 1 And animate_direction = 1 Then
        animate_progress = 1
        Animation2.Enabled = False
    End If
    If animate_progress <= 0 And animate_direction = 0 Then
        animate_progress = 0
        Animation2.Enabled = False
        EditContactFrame.Visible = False
        contName = ""
        contCell = ""
        contLand = ""
        contEmail = ""
    End If
    EditContactFrame.height = 5000 * draw_progress
    DoEvents
End Sub

Private Sub butnAddContact_Click()
    AddContactsFrame.width = butnAddContact.width
    AddContactsFrame.Visible = True
    Animate 1, 1000
End Sub

Private Sub butnCancelContact_Click()
    Animate 0, 1000
End Sub

Private Sub butnCancelSaveContact_Click()
    Animate2 0, 1000
    contName = ""
    contCell = ""
    contLand = ""
    contEmail = ""
End Sub

Private Sub butnCreateContact_Click()
    LabEditContactInfo = "Create New Contact"
    EditContactFrame.Tag = "Add"
    EditContactFrame.height = 0
    EditContactFrame.Visible = True
    Animate2 1, 1000
End Sub

Private Sub butnInnerAddContact_Click()
    Dim c As Byte
    butnCreateContact.Visible = False
    AddingContactOptions.Visible = True
    AddingButnSave.Enabled = False
    AddingNameAs = LV_contacts.ListItems(contactsLineNum).SubItems(1) & " as:"
    AddingCanSignOut = 0
    For c = 0 To 3
        ContactType(c).value = False
    Next c
    
End Sub

Private Sub butnReactivate_Click()
    If MsgBox("Are you sure you want to reactivate " & txtFirst & " " & txtLast & "?", vbYesNo) = vbYes Then
        Dim cl As ADODB.Recordset
        Set cl = db.Execute("SELECT * FROM clients WHERE idClient = " & ID.Text)
        With cl
            If Not (.EOF And .BOF) Then
                .MoveFirst
                insertClientChange dpEffective.value, !idClient, !feeClassID, !fees, !payperiod, !room, !subsidized, "" & !authorizationNumber, !parentalContribution, dpEffective.value, Null, 1
            End If
        End With
        Set cl = Nothing
        
        db.Execute "UPDATE clients SET active=1, startDate=" & sqlDate(dpEffective.value) & ", endDate=NULL WHERE idClient = " & ID.Text
        MsgBox "Account reactivated. Window will be closed."
        Unload Me
    End If
End Sub

Private Sub butnRemoveContact_Click()
    If MsgBox("Are you sure you wish to remove " & ClientContacts.Text & " As a contact for " & txtFirst & " " & txtLast & "?", vbYesNo) = vbYes Then
        db.Execute ("DELETE FROM client_contacts WHERE idContact = " & contactInfo.Tag & " AND idClient = " & ID.Text)
        UpdateClientContacts
    End If
End Sub

Private Sub butnSaveContact_Click()
    If EditContactFrame.Tag = "Add" Then
        db.Execute "INSERT INTO contacts (name, cell, land, email) VALUES (""" & contName & """,""" & contCell & """,""" & contLand & """,""" & contEmail & """)"
    Else
        db.Execute "UPDATE contacts SET name=""" & contName & """, cell=""" & contCell & """, land=""" & contLand & """, email=""" & contEmail & """ WHERE idContact = " & LV_contacts.ListItems(contactsLineNum).Text
    End If
    Animate2 0, 1000
End Sub

Private Sub butnUpdateContact_Click()
    Dim q As ADODB.Recordset
    
    LabEditContactInfo = "Update Contact Info"
    EditContactFrame.Tag = "Edit"
    EditContactFrame.height = 0
    EditContactFrame.Visible = True
    Animate2 1, 1000
    
    Set q = db.Execute("SELECT * FROM contacts WHERE idContact = " & LV_contacts.ListItems(contactsLineNum).Text)
    If Not (q.EOF And q.BOF) Then
        q.MoveFirst
        contName = q!name
        contCell = q!cell
        contLand = q!land
        contEmail = q!email
    End If
End Sub

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

Private Sub ClientContacts_Click()
    Dim cName As String
    Dim cType As String
    Dim cCell As String
    Dim cPhone As String
    Dim cEmail As String
    Dim q As ADODB.Recordset
    Set q = db.Execute("SELECT * FROM client_contacts INNER JOIN contacts ON (contacts.idContact = client_contacts.idContact) WHERE id=" & contactIDs.List(ClientContacts.ListIndex))
    With q
        If Not (.EOF And .BOF) Then
            .MoveFirst
            butnRemoveContact.Enabled = True
            Do Until .EOF
                cName = !name
                cType = !Type
                cCell = !cell
                cPhone = !land
                cEmail = !email
                If cType = "Parent" Then cType = "Parent/Guardian"
                contactInfo = "Type: " & cType & vbCrLf & "Name: " & cName & vbCrLf & "Cell: " & cCell & vbCrLf & "Land: " & cPhone & vbCrLf & "Email: " & cEmail
                contactInfo.Tag = !idContact
                If -CBool(!canSignOut) Then
                    LabCanSignOut = "Can Sign Out"
                Else
                    LabCanSignOut = ""
                End If
                .MoveNext
            Loop
        Else
            butnRemoveContact.Enabled = False
        End If
    End With
End Sub

Private Sub ContactFilter_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim q As ADODB.Recordset
    Dim r As ADODB.Recordset
    
    Dim li As ListItem
    If Len(ContactFilter) > 0 Then
        Set q = db.Execute("SELECT * FROM contacts WHERE name LIKE ""%" & ContactFilter & "%"" OR email LIKE ""%" & ContactFilter & "%""")
        LV_contacts.ListItems.Clear
        With q
            If Not (.EOF And .BOF) Then
                .MoveFirst
                Do Until .EOF
                    Set r = db.Execute("SELECT * FROM client_contacts WHERE idContact = " & !idContact & " AND idClient = " & ID.Text)
                    ' this lists only contacts that aren't already linked to this client
                    If r.EOF And r.BOF Then
                        Set li = LV_contacts.ListItems.Add(, , !idContact)
                        li.SubItems(1) = !name
                        li.SubItems(2) = truncatePhoneNumber(!cell)
                        li.SubItems(3) = truncatePhoneNumber(!cell)
                        li.SubItems(4) = !email
                    End If
                    .MoveNext
                Loop
            End If
        End With
    Else
        LV_contacts.ListItems.Clear
    End If
End Sub

Private Sub ContactType_Click(index As Integer)
    Dim T As String
    If index = 0 Then T = "Parent"
    If index = 1 Then T = "Emergency"
    If index = 2 Then T = "Doctor"
    If index = 3 Then T = "Other"
    AddingContactOptions.Tag = T
    AddingButnSave.Enabled = True
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
    AddContactsFrame.Top = butnAddContact.Top
    AddContactsFrame.height = butnAddContact.height
    AddContactsFrame.width = butnAddContact.width
    AddingContactOptions.width = Me.width
    EditContactFrame.width = Me.width
    changeDateReminder = True
End Sub

Private Sub LV_contacts_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lineheight As Long
    Dim lineNum As Long
    lineheight = 110
    
    If LV_contacts.ListItems.count > 0 Then
        lineheight = LV_contacts.ListItems(1).height
        butnInnerAddContact.height = lineheight
        butnUpdateContact.height = lineheight
    End If
    
    lineNum = Int((y - 90) / lineheight)
    
    If lineNum > 0 And lineNum <= LV_contacts.ListItems.count Then
        butnInnerAddContact.Tag = lineNum
        butnInnerAddContact.Top = LV_contacts.Top + 90 + lineNum * lineheight
        butnInnerAddContact.Visible = True
        butnUpdateContact.Top = LV_contacts.Top + 90 + lineNum * lineheight
        butnUpdateContact.Visible = True
        contactsLineNum = lineNum
    Else
        butnInnerAddContact.Visible = False
        butnUpdateContact.Visible = False
    End If
    
End Sub

Private Sub SaveButn_Click()
    Dim sql As String
    Dim effectiveDate As Date
    
    If ID = "" Then ' NEW CLIENT!!!
        Dim search As ADODB.Recordset
        Set search = db.Execute("SELECT * FROM clients WHERE first LIKE ""%" & txtFirst & "%"" AND last LIKE ""%" & txtLast & "%""")
        With search
            If Not (.EOF And .BOF) Then
                .MoveFirst
                If MsgBox("There is already a client in the database named " & !First & " " & !Last & "!  If this is the same child you are creating, to avoid duplicate accounts please click 'No' and then 'Cancel' and use the account that already exists." & vbCrLf & "Do you still wish to create this account?", vbExclamation + vbYesNo) = vbNo Then Exit Sub
            End If
        End With
        effectiveDate = dpStart.value
        sql = "INSERT INTO clients ("
        sql = sql & "first,"
        sql = sql & "last,"
        sql = sql & "initial,"
        sql = sql & "MCP,"
        sql = sql & "allergies,"
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
        sql = sql & """" & txtMCP & ""","
        sql = sql & """" & txtAllergies & ""","
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
        If changeDateReminder Then
            changeDateReminder = False
            MsgBox "Have you checked the date that these changes are being applied as?  Please ensure the correct date is chosen."
            Exit Sub
        End If
        
        effectiveDate = dpEffective.value
        
        sql = "UPDATE clients SET "
        sql = sql & "first=""" & txtFirst & ""","
        sql = sql & "last=""" & txtLast & ""","
        sql = sql & "initial=""" & txtInitial & ""","
        sql = sql & "MCP=""" & txtMCP & ""","
        sql = sql & "allergies=""" & txtAllergies & ""","
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
            insertClientChange dpEffective.value, !idClient, !feeClassID, !fees, !payperiod, !room, !subsidized, "" & !authorizationNumber, !parentalContribution, !startDate, !endDate, !active
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
                txtMCP = "" & !MCP
                txtAllergies = "" & !allergies
                dpDOB.value = !DOB
                comboSelectItem cboGender, !gender
                cboFeeClass.ListIndex = !feeClassID - 1
                txtFees = !fees
                dpStart.value = !startDate
                originalStartDate = !startDate
                If IsNull(!endDate) Then
                    dpEnd.value = Date
                Else
                    dpEnd.value = !endDate
                End If
                comboSelectItem cboPP, !payperiod
                comboSelectItem cboRoom, !room
                chkSubsidized.value = !subsidized
                If !subsidized Then
                    txtAuthNumber = "" & !authorizationNumber
                    txtParentalContrib = Format(!parentalContribution, "0.00")
                End If
                chkActive.value = !active
                If Not CBool(!active) Then
                    butnReactivate.Visible = True
                    Dim ctrl As Control
                    For Each ctrl In Me.Controls
                        If Not (ctrl.name = "cancelButn" Or ctrl.name = "butnReactivate" Or ctrl.name = "dpEffective") Then
                            On Error Resume Next
                            ctrl.Enabled = False
                        End If
                    Next ctrl
                End If
                chkActive.Visible = True
            End If
        End With
        dpEffective.value = Date
        
        UpdateClientContacts
    Else
        'NEW ENTRY
        ContactsFrame.Visible = False
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
        SaveButn.Enabled = True
    Else
        SaveButn.Enabled = False
    End If
End Sub

Private Sub txtLast_Change()
    checkSaveEnabled
End Sub

Private Function EasingFunction(progress As Double) As Double 'progress is a double value between 0 and 1
    'EasingFunction = progress 'linear
    'EasingFunction = Sin(progress * PI / 2) 'ease out
    'EasingFunction = (Sin((progress - 1) * PI / 2#) + 1) ^ 2# 'ease in
    EasingFunction = (Sin(Sin((progress * 2 - 1) * PI / 2#) * PI / 2#) + 1) / 2# 'ease in and out
End Function

Private Sub Animate(direction As Byte, interval As Long)
    animate_direction = direction
    animate_interval = interval
    animate_progress = 1 - direction
    Animation.Enabled = True
End Sub

Private Sub Animate2(direction As Byte, interval As Long)
    animate_direction = direction
    animate_interval = interval
    animate_progress = 1 - direction
    Animation2.Enabled = True
End Sub

Sub UpdateClientContacts()
    Dim q As ADODB.Recordset
    Set q = db.Execute("SELECT name, id FROM client_contacts INNER JOIN contacts ON (contacts.idContact = client_contacts.idContact) WHERE idClient=" & ID.Text)
    ClientContacts.Clear
    contactIDs.Clear
    With q
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do Until .EOF
                ClientContacts.AddItem !name
                contactIDs.AddItem !ID
                .MoveNext
            Loop
        Else
            butnRemoveContact.Enabled = False
        End If
    End With
End Sub
