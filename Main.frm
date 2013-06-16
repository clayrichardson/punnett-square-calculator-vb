VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Punnett Square Calculator"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   9225
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   6943.874
   ScaleMode       =   0  'User
   ScaleWidth      =   9255.049
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   2760
      Top             =   -120
   End
   Begin MSComDlg.CommonDialog Cdl 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   7920
      Top             =   3120
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1320
      Top             =   360
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   4080
      Top             =   3120
   End
   Begin VB.Frame Frame1 
      Caption         =   "Number of Traits:"
      Height          =   1095
      Left            =   120
      TabIndex        =   41
      Top             =   120
      Width           =   1695
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H0000C0C0&
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Enviro"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "First Parent:"
      Height          =   3015
      Left            =   1920
      TabIndex        =   21
      Top             =   120
      Width           =   7215
      Begin VB.Frame Frame3 
         Caption         =   "Genotypes:"
         Height          =   2655
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   6975
         Begin VB.Frame Frame4 
            Caption         =   "First Allele:"
            Height          =   2295
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   2055
            Begin VB.ListBox List 
               Appearance      =   0  'Flat
               Height          =   1785
               Index           =   1
               ItemData        =   "Main.frx":0000
               Left            =   1320
               List            =   "Main.frx":0002
               TabIndex        =   38
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox Text2 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               MaxLength       =   2
               TabIndex        =   37
               Top             =   240
               Width           =   1095
            End
            Begin VB.CommandButton Command1 
               BackColor       =   &H0000C000&
               Caption         =   "Change"
               BeginProperty Font 
                  Name            =   "Enviro"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   35
               Top             =   840
               Width           =   1095
            End
            Begin VB.CommandButton Command15 
               BackColor       =   &H000080FF&
               Caption         =   "/\"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Enviro"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   39
               Top             =   1200
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.CommandButton Command16 
               BackColor       =   &H00FF0000&
               Caption         =   "\/"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Enviro"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   840
               Style           =   1  'Graphical
               TabIndex        =   36
               Top             =   1200
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.CommandButton Command2 
               BackColor       =   &H000000C0&
               Caption         =   "Remove"
               BeginProperty Font 
                  Name            =   "Enviro"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               MaskColor       =   &H00FFFFFF&
               Style           =   1  'Graphical
               TabIndex        =   40
               Top             =   1560
               Width           =   1095
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Second Allele:"
            Height          =   2295
            Left            =   2280
            TabIndex        =   27
            Top             =   240
            Width           =   2055
            Begin VB.ListBox List 
               Appearance      =   0  'Flat
               Height          =   1785
               Index           =   2
               ItemData        =   "Main.frx":0004
               Left            =   120
               List            =   "Main.frx":0006
               TabIndex        =   32
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox Text3 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   840
               MaxLength       =   2
               TabIndex        =   31
               Top             =   240
               Width           =   1095
            End
            Begin VB.CommandButton Command3 
               BackColor       =   &H0000C000&
               Caption         =   "Change"
               BeginProperty Font 
                  Name            =   "Enviro"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   840
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   840
               Width           =   1095
            End
            Begin VB.CommandButton Command17 
               BackColor       =   &H000080FF&
               Caption         =   "/\"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Enviro"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   840
               Style           =   1  'Graphical
               TabIndex        =   30
               Top             =   1200
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.CommandButton Command18 
               BackColor       =   &H00FF0000&
               Caption         =   "\/"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Enviro"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1560
               Style           =   1  'Graphical
               TabIndex        =   29
               Top             =   1200
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.CommandButton Command4 
               BackColor       =   &H000000C0&
               Caption         =   "Remove"
               BeginProperty Font 
                  Name            =   "Enviro"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   840
               Style           =   1  'Graphical
               TabIndex        =   33
               Top             =   1560
               Width           =   1095
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Preview:"
            Height          =   2295
            Left            =   4440
            TabIndex        =   23
            Top             =   240
            Width           =   2415
            Begin VB.TextBox Text4 
               Height          =   1875
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   26
               Top             =   240
               Width           =   735
            End
            Begin VB.CommandButton Command5 
               BackColor       =   &H00C0C000&
               Caption         =   "Refresh"
               BeginProperty Font 
                  Name            =   "Enviro"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   960
               Style           =   1  'Graphical
               TabIndex        =   25
               Top             =   600
               Width           =   1335
            End
            Begin VB.CommandButton Command6 
               BackColor       =   &H000040C0&
               Caption         =   "Calculate"
               BeginProperty Font 
                  Name            =   "Enviro"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   960
               Style           =   1  'Graphical
               TabIndex        =   24
               Top             =   1440
               Width           =   1335
            End
         End
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Second Parent:"
      Height          =   3015
      Left            =   1920
      TabIndex        =   1
      Top             =   3240
      Width           =   7215
      Begin VB.Frame Frame8 
         Caption         =   "Genotypes:"
         Height          =   2655
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   6975
         Begin VB.Frame Frame9 
            Caption         =   "Preview:"
            Height          =   2295
            Left            =   4440
            TabIndex        =   17
            Top             =   240
            Width           =   2415
            Begin VB.CommandButton Command7 
               BackColor       =   &H000040C0&
               Caption         =   "Calculate"
               BeginProperty Font 
                  Name            =   "Enviro"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   960
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   1440
               Width           =   1335
            End
            Begin VB.CommandButton Command8 
               BackColor       =   &H00C0C000&
               Caption         =   "Refresh"
               BeginProperty Font 
                  Name            =   "Enviro"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   960
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   600
               Width           =   1335
            End
            Begin VB.TextBox Text5 
               Height          =   1875
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   18
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Second Allele:"
            Height          =   2295
            Left            =   2280
            TabIndex        =   10
            Top             =   240
            Width           =   2055
            Begin VB.CommandButton Command10 
               BackColor       =   &H0000C000&
               Caption         =   "Change"
               BeginProperty Font 
                  Name            =   "Enviro"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   840
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   840
               Width           =   1095
            End
            Begin VB.ListBox List 
               Appearance      =   0  'Flat
               Height          =   1785
               Index           =   4
               ItemData        =   "Main.frx":0008
               Left            =   120
               List            =   "Main.frx":000A
               TabIndex        =   12
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox Text7 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   840
               MaxLength       =   2
               TabIndex        =   11
               Top             =   240
               Width           =   1095
            End
            Begin VB.CommandButton Command21 
               BackColor       =   &H000080FF&
               Caption         =   "/\"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Enviro"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   840
               Style           =   1  'Graphical
               TabIndex        =   14
               Top             =   1200
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.CommandButton Command22 
               BackColor       =   &H00FF0000&
               Caption         =   "\/"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Enviro"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1560
               Style           =   1  'Graphical
               TabIndex        =   15
               Top             =   1200
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.CommandButton Command9 
               BackColor       =   &H000000C0&
               Caption         =   "Remove"
               BeginProperty Font 
                  Name            =   "Enviro"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   840
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   1560
               Width           =   1095
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "First Allele:"
            Height          =   2295
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   2055
            Begin VB.ListBox List 
               Appearance      =   0  'Flat
               Height          =   1785
               Index           =   3
               ItemData        =   "Main.frx":000C
               Left            =   1320
               List            =   "Main.frx":000E
               TabIndex        =   8
               Top             =   240
               Width           =   615
            End
            Begin VB.CommandButton Command12 
               BackColor       =   &H0000C000&
               Caption         =   "Change"
               BeginProperty Font 
                  Name            =   "Enviro"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   840
               Width           =   1095
            End
            Begin VB.TextBox Text6 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               MaxLength       =   2
               TabIndex        =   4
               Top             =   240
               Width           =   1095
            End
            Begin VB.CommandButton Command19 
               BackColor       =   &H000080FF&
               Caption         =   "/\"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Enviro"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   7
               Top             =   1200
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.CommandButton Command20 
               BackColor       =   &H00FF0000&
               Caption         =   "\/"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Enviro"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   840
               Style           =   1  'Graphical
               TabIndex        =   6
               Top             =   1200
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.CommandButton Command11 
               BackColor       =   &H000000C0&
               Caption         =   "Remove"
               BeginProperty Font 
                  Name            =   "Enviro"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               MaskColor       =   &H00FFFFFF&
               Style           =   1  'Graphical
               TabIndex        =   9
               Top             =   1560
               Width           =   1095
            End
         End
      End
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00FF00FF&
      Caption         =   "Gicantically big pink button that I have absolutely no idea why it was put here. GO AHEAD AND CLICK IT, I DARE YOU."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   1695
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   120
      TabIndex        =   44
      Top             =   6960
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   238
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label10"
      Height          =   255
      Left            =   7800
      TabIndex        =   45
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   46
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label8"
      Height          =   255
      Left            =   4080
      TabIndex        =   47
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label7"
      Height          =   255
      Left            =   2760
      TabIndex        =   48
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
      Height          =   255
      Left            =   120
      TabIndex        =   49
      Top             =   6600
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Action:"
      Height          =   255
      Left            =   120
      TabIndex        =   54
      Top             =   6360
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Item:"
      Height          =   255
      Left            =   2760
      TabIndex        =   53
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Total:"
      Height          =   255
      Left            =   4080
      TabIndex        =   52
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Data:"
      Height          =   255
      Left            =   5400
      TabIndex        =   51
      Top             =   6360
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "% Completed:"
      Height          =   255
      Left            =   7800
      TabIndex        =   50
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Menu filemnu 
      Caption         =   "&File"
      Begin VB.Menu NewMnu 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu SaveMnu 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu SaveAsMnu 
         Caption         =   "Save &As"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu LoadMnu 
         Caption         =   "&Load"
         Shortcut        =   ^L
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu CloseMnu 
         Caption         =   "&Close"
         Shortcut        =   ^C
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu ExitMnu 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu editmnu 
      Caption         =   "&Edit"
      Begin VB.Menu ClrAllMnu 
         Caption         =   "Clear Alleles"
         Shortcut        =   +{DEL}
      End
   End
   Begin VB.Menu viewmnu 
      Caption         =   "&View"
   End
   Begin VB.Menu actionsmnu 
      Caption         =   "&Actions"
      Begin VB.Menu CalMnu 
         Caption         =   "&Calculate"
         Shortcut        =   {F7}
      End
      Begin VB.Menu RevGamMnu 
         Caption         =   "&Review Gametes"
         Shortcut        =   {F6}
      End
      Begin VB.Menu PrevAllMnu 
         Caption         =   "&Preview All"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu helpmnu 
      Caption         =   "&Help"
      Begin VB.Menu RelNoteMnu 
         Caption         =   "&Release Notes"
         Shortcut        =   {F2}
      End
      Begin VB.Menu AboutMnu 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'YOU WERE ABOUT TO FINISH THE PROCESS OF SAVING AND LOADING THE LISTS AND
'SAVING THE RESULTS AND TELLING IF THE USER HAS SAVED ALREADY AND PUTTING BACK
'THE SAVE AS MENU ITEM

Option Explicit

Dim LCount1 As Integer                  'count for storing lists 1 & 2 in a variable
Dim LCount2 As Integer                  'count for storing lists 3 & 4 in a variable
Dim LCount3 As Integer                  'count for storing lists 1 & 2 in a variable
Dim LCount4 As Integer                  'count for storing lists 3 & 4 in a variable

Dim AddAlleleCounter As Integer         'counter for add allele loops
Dim diff As Integer                     'diff. from actual # of alleles & wanted #
Public Prnt As String                   'variable that stores the rows to be
                                        'printed
                                        
Dim Percent As Integer                  'percent variable for status label
Dim Lpercent As Double                  'percent variable for progress bar

Dim PreviewInProgress As Boolean        'whether preview generation is in progress
Dim PrevCounter As Integer              'preview counter

Dim CalculateInProgress As Boolean      'whether calculation is in progress

Dim EditInProgress1 As Boolean          'whether list edit is in progress
Dim EditInProgress2 As Boolean          'whether list edit is in progress
Dim EditInProgress3 As Boolean          'whether list edit is in progress
Dim EditInProgress4 As Boolean          'whether list edit is in progress
Dim EmptyListCheck As Integer           'counter to check for empty items in list


'VARIABLES FOR THE GAMETE CALCULATOR 'ENGINE'

Dim ListStoreCounter As Long            'list storage counter
Dim CoOrdinate As String                'gamete cordinate 'data' geno cordinate
Dim ItemCount As Long                   'item no. currently being translated
Dim MaxX As Long                        'max no. of columns
Dim maxY As Long                        'man no. of rows
Dim Gex As Long                         'genotype x coordinate
Dim Gey As Long                         'genotype y coordinate
Dim Gax As Long                         'gamete x coordinate
Dim Gay As Long                         'gamete y coordinate
Dim Geno1()                             'genotype parent 1 2-d array
Dim Gamete1()                           'gamete parent 1 2-d array
Dim Geno2()                             'genotype parent 2 2-d array
Dim Gamete2()                           'gamete parent 2 2-d array
Dim OuterPattern As Long                'how many times to switch from list1 to list2
Dim Pattern As Long                     'how many times to print same charac. in column
Public Traits As Long                   'how many traits there are
Dim OutCounter As Long                  'declares the outer loop counter
Dim MidCounter As Long                  'the middle loop counter
Dim InCounter As Long                   'the inner loop counter


'VARIABLES FOR THE OFFSPRING CALCULATOR 'ENGINE'

Dim OutComes As Long
Dim Offspring()
Dim Geno()
Dim GameteX As Long
Dim OffX As Long
Dim OffY As Long
Dim Temp As Variant
Dim Pcounter As Long
Dim OffPrev As Variant


'VARIABLES FOR THE SAVE/LOAD PROCESS

Dim NewFile As Boolean
Dim FileOpen As Boolean

Dim SaveList As Long
Dim ListCont As String
Dim SaveListNum As Integer

Dim SaveResponse As Integer

Private Sub AboutMnu_Click()
    
    MsgBox "This is BETA VERSION 1.0", , "BETA VERSION 1.0"
    
End Sub

Private Sub CalMnu_Click()
    
    Call CalculateGamete1(True)
    Call CalculateGamete2(True)
    
End Sub

Private Sub CloseMnu_Click()
    
    Call AskSaveYet
    
    FileOpen = False
    Call ClearAlleles
    
End Sub

Private Sub ClrAllMnu_Click()
    
    Call ClearAlleles
    
End Sub

Private Sub Command1_Click()
    
    If EditInProgress1 = False Then
    
        EditInProgress1 = True
        
        Call DisableButtons
        
        Command1.Caption = "Done"
    
        Text2.SetFocus
        Text2.SelStart = 0
        Text2.SelLength = 2
    
    ElseIf EditInProgress1 = True Then
    
        For EmptyListCheck = 0 To _
                (List(1).ListCount - 1)
            DoEvents
            
            List(1).ListIndex = _
            EmptyListCheck
            
            If List(1).Text = "" Then
                MsgBox "You must " _
                & "select " _
                & "a value for this " _
                & "allele."
                GoTo EndEditing1
            End If
            
        Next
                   
                   
        Call EnableButtons
        
        EditInProgress1 = False
        
        Command1.Caption = "Change"
        
        List(1).ListIndex = -1
        
        End If
        
EndEditing1:
        
End Sub

Private Sub Command10_Click()
    
    If EditInProgress4 = False Then
    
        EditInProgress4 = True
        
        Call DisableButtons
        
        Command10.Caption = "Done"
    
        Text7.SetFocus
        Text7.SelStart = 0
        Text7.SelLength = 2
    
    ElseIf EditInProgress4 = True Then
    
        For EmptyListCheck = 0 To _
                (List(4).ListCount - 1)
            DoEvents
            
            List(4).ListIndex = _
            EmptyListCheck
            
            If List(4).Text = "" Then
                MsgBox "You must " _
                & "select " _
                & "a value for this " _
                & "allele."
                GoTo EndEditing4
            End If
            
        Next
        
        Call EnableButtons
        
        EditInProgress4 = False
        
        Command10.Caption = "Change"
        
        List(4).ListIndex = -1
        
        End If

EndEditing4:
        
End Sub

Private Sub Command11_Click()
    
    List(3).RemoveItem List(3).ListIndex
    
End Sub

Private Sub Command12_Click()
    
    If EditInProgress3 = False Then
    
        EditInProgress3 = True
        
        Call DisableButtons
        
        Command12.Caption = "Done"
    
        Text6.SetFocus
        Text6.SelStart = 0
        Text6.SelLength = 2
    
    ElseIf EditInProgress3 = True Then
        
        For EmptyListCheck = 0 To _
                (List(3).ListCount - 1)
            DoEvents
            
            List(3).ListIndex = _
            EmptyListCheck
            
            If List(3).Text = "" Then
                MsgBox "You must " _
                & "select " _
                & "a value for this " _
                & "allele."
                GoTo EndEditing3
            End If
            
        Next
               
               
        Call EnableButtons
        
        EditInProgress3 = False
        
        Command12.Caption = "Change"
        
        List(3).ListIndex = -1
        
        End If
        
EndEditing3:
        
End Sub

Private Sub Command13_Click()
        
PreviewInProgress = True
        
Call DisableButtons
    
    If Val(Text1.Text) > 100 Then
        MsgBox "Dude, you've gotta be kiddin me..."
    End If
        
    LCount1 = List(1).ListCount           'gets the number of items in list 1
    LCount2 = List(2).ListCount           'gets the number of items in list 2
    LCount3 = List(3).ListCount           'gets the number of items in list 3
    LCount4 = List(4).ListCount           'gets the number of items in list 4
    
    If Val(Text1.Text) > LCount1 Then   'if the no. in text1 is greater than
        diff = Val(Text1.Text) - LCount1
            For AddAlleleCounter = 1 To diff     'list1's count, then add the no. of
                    DoEvents
                List(1).AddItem _
                (List(1).ListCount + 1)   'the diff. to make the list have that
                    DoEvents
                    
                Lpercent = _
                (AddAlleleCounter / _
                diff) * 100             'figures the percent to display
                Percent = Lpercent      'rounds long percent into integer
                    DoEvents
                ProgressBar1.Value = _
                Lpercent                'updates progress bar to percent
                    DoEvents
                                        'displays status in status label

Label6.Caption = "Adding item..."
Label7.Caption = AddAlleleCounter
Label8.Caption = diff
Label9.Caption = "Parent 1 " & "Allele 1"
Label10.Caption = Round(Lpercent, 10)


                    DoEvents
            Next                        'many items without adding too much
    End If                              'or not adding enough
            
    If Val(Text1.Text) > LCount2 Then   'if the no. in text1 is greater than
        diff = Val(Text1.Text) - LCount2
            For AddAlleleCounter = _
                        1 To diff       'list1's count, then add the no. of
                    DoEvents
                List(2).AddItem _
                (List(2).ListCount + 1)   'the diff. to make the list have that
                    DoEvents
                
                Lpercent = _
                (AddAlleleCounter / _
                diff) * 100             'figures the percent to display
                Percent = Lpercent      'rounds long percent into integer
                    DoEvents
                ProgressBar1.Value = _
                Lpercent                'updates progress bar to percent
                    DoEvents
                                        'displays status in status label

Label6.Caption = "Adding item..."
Label7.Caption = AddAlleleCounter
Label8.Caption = diff
Label9.Caption = "Parent 1 " & "Allele 2"
Label10.Caption = Round(Lpercent, 10)


                    DoEvents
            Next                        'many items without adding too much
    End If                              'or not adding enough
            
    If Val(Text1.Text) > LCount3 Then   'if the no. in text1 is greater than
        diff = Val(Text1.Text) - _
                                LCount3
            For AddAlleleCounter = _
            1 To diff                   'list1's count, then add the no. of
                    DoEvents
                List(3).AddItem _
                (List(3).ListCount + 1)   'the diff. to make the list have that
                    DoEvents
                    
                Lpercent = _
                (AddAlleleCounter / _
                diff) * 100             'figures the percent to display
                Percent = Lpercent      'rounds long percent into integer
                    DoEvents
                ProgressBar1.Value = _
                Lpercent                'updates progress bar to percent
                    DoEvents
                                        'displays status in status label
Label6.Caption = "Adding item..."
Label7.Caption = AddAlleleCounter
Label8.Caption = diff
Label9.Caption = "Parent 2 " & "Allele 1"
Label10.Caption = Round(Lpercent, 10)
                
                    DoEvents
            Next                        'many items without adding too much
    End If                              'or not adding enough
    
    If Val(Text1.Text) > LCount4 Then   'if the no. in text1 is greater than
        diff = Val(Text1.Text) - _
                                LCount4
            For AddAlleleCounter = _
                            1 To diff   'list1's count, then add the no. of
                    DoEvents
                List(4).AddItem _
                (List(4).ListCount + 1)   'the diff. to make the list have that
                    DoEvents
                
                Lpercent = _
                (AddAlleleCounter / _
                diff) * 100             'figures the percent to display
                Percent = Lpercent      'rounds long percent into integer
                    DoEvents
                ProgressBar1.Value = _
                Lpercent                'updates progress bar to percent
                    DoEvents
                                        'displays status in status label
Label6.Caption = "Adding item..."
Label7.Caption = AddAlleleCounter
Label8.Caption = diff
Label9.Caption = "Parent 2 " & "Allele 2"
Label10.Caption = Round(Lpercent, 10)

                    DoEvents
            Next                        'many items without adding too much
    End If                              'or not adding enough

EndAddAlleles:

ProgressBar1.Value = "0"

Label6.Caption = "Ready"
Label7.Caption = ""
Label8.Caption = ""
Label9.Caption = ""
Label10.Caption = ""

Call EnableButtons

Text1.Text = ""

PreviewInProgress = False
    
End Sub

Private Sub Command14_Click()
    
    MsgBox "Ha ha, made you click...", , "BETA VERSION 1.0"
    
End Sub

Private Sub Command2_Click()
    
    List(1).RemoveItem List(1).ListIndex
    
End Sub

Private Sub Command3_Click()
    
    If EditInProgress2 = False Then
    
        EditInProgress2 = True
        
        Call DisableButtons
        
        Command3.Caption = "Done"
    
        Text3.SetFocus
        Text3.SelStart = 0
        Text3.SelLength = 2
    
    ElseIf EditInProgress2 = True Then
    
        For EmptyListCheck = 0 To _
                (List(2).ListCount - 1)
            DoEvents
            
            List(2).ListIndex = _
            EmptyListCheck
            
            If List(2).Text = "" Then
                MsgBox "You must " _
                & "select " _
                & "a value for this " _
                & "allele."
                GoTo EndEditing2
            End If
            
        Next
               
               
        Call EnableButtons
        
        EditInProgress2 = False
        
        Command3.Caption = "Change"
        
        List(2).ListIndex = -1
        
        End If
        
EndEditing2:
        
End Sub

Private Sub Command4_Click()
    
    List(2).RemoveItem List(2).ListIndex
    
End Sub

Private Sub Command5_Click()

    Call PreviewGeno1                   'calls the subroutine to preview genotypes

End Sub

Private Sub Command6_Click()
    
    Call CalculateGamete1(True)
    
End Sub

Private Sub Command7_Click()
    
    Call CalculateGamete1(True)
    Call CalculateGamete2(True)
    
End Sub

Private Sub Command8_Click()
    
    Call PreviewGeno2
    
End Sub

Private Sub Command9_Click()
    
    List(4).RemoveItem List(4).ListIndex
    
End Sub

Private Sub ExitMnu_Click()
    
    If FileOpen = True Then
        Call AskSaveYet
    End If
    
End Sub

Private Sub Form_Load()
        
    EditInProgress1 = False             'sets the editing status to false
    EditInProgress2 = False             'sets the editing status to false
    EditInProgress3 = False             'sets the editing status to false
    EditInProgress4 = False             'sets the editing status to false
    
    FileOpen = False
    
    
    'List(1).AddItem "A"
    'List(1).AddItem "B"
    'List(1).AddItem "D"
    'List(1).AddItem "G"
    'list(1).AddItem "H"
    'list(1).AddItem "I"
    'list(1).AddItem "J"
    'list(1).AddItem "K"
    'list(1).AddItem "L"
    'list(1).AddItem "M"
    
    'List(2).AddItem "a"
    'List(2).AddItem "b"
    'List(2).AddItem "d"
    'List(2).AddItem "g"
    'list(2).AddItem "h"
    'list(2).AddItem "i"
    'list(2).AddItem "j"
    'list(2).AddItem "k"
    'list(2).AddItem "l"
    'list(2).AddItem "m"
    
    
    'List(3).AddItem "N"
    'List(3).AddItem "O"
    'List(3).AddItem "P"
    'List(3).AddItem "Q"
    'list(3).AddItem "R"
    'list(3).AddItem "S"
    'list(3).AddItem "T"
    'list(3).AddItem "U"
    'list(3).AddItem "V"
    'list(3).AddItem "W"
    
    'List(4).AddItem "n"
    'List(4).AddItem "o"
    'List(4).AddItem "p"
    'List(4).AddItem "q"
    'list(4).AddItem "r"
    'list(4).AddItem "s"
    'list(4).AddItem "t"
    'list(4).AddItem "u"
    'list(4).AddItem "v"
    'list(4).AddItem "w"
        
    
    Label6.Caption = "Ready"
    Label7.Caption = ""
    Label8.Caption = ""
    Label9.Caption = ""
    Label10.Caption = ""
    
               
'Call PreviewGeno                       'calls the subroutine to preview genotypes
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call AskSaveYet

End Sub

Private Sub List_Click(Index As Integer)
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
    
End Sub

Private Sub List_GotFocus(Index As Integer)
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
    
End Sub

Private Sub List_KeyDown(KeyCode As Integer, Shift As Integer, Index As Integer)
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
    
End Sub

Private Sub List_KeyPress(KeyAscii As Integer, Index As Integer)
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
    
End Sub

Private Sub List_KeyUp(KeyCode As Integer, Shift As Integer, Index As Integer)
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
    
End Sub

Private Sub List_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
    
End Sub

Private Sub List_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
    
End Sub

Private Sub LoadMnu_Click()

Dim ListNum As String
Dim ListData As String

If FileOpen = True Then
    Call AskSaveYet
End If

    Cdl.DialogTitle = "Save file..."
    Cdl.InitDir = App.Path
    Cdl.Filter = "Parent Genotypes (*.txt)|*.txt"
    Cdl.ShowOpen
    
    If Cdl.FileName = "" Then
        GoTo EndOfLoad
    ElseIf Cdl.FileName <> "" Then
        Open Cdl.FileName For Input As #1
    End If
    
    Call ClearAlleles
    
    Do While Not EOF(1)
        DoEvents
        
        Input #1, ListNum, ListData
        
        If ListNum = "[1]" Then
            List(1).AddItem ListData
        ElseIf ListNum = "[2]" Then
            List(2).AddItem ListData
        ElseIf ListNum = "[3]" Then
            List(3).AddItem ListData
        ElseIf ListNum = "[4]" Then
            List(4).AddItem ListData
        ElseIf Left$(ListNum, 1) <> "[" Then
            MsgBox "Make sure this file is in the correct format or it is " _
            & "not corrupted."
            Exit Do
            Close #1
        End If
    Loop
    
    Close #1
    
    
    FileCopy Cdl.FileName, App.Path & "\temp.txt"
    
    NewFile = False
    FileOpen = True
    
EndOfLoad:
    
End Sub

Private Sub NewMnu_Click()

    If FileOpen = True Then
        Call AskSaveYet
    End If
    
    Cdl.DialogTitle = "Save new file..."
    Cdl.InitDir = App.Path
    Cdl.Filter = "Parent Genotypes (*.txt)|*.txt"
    Cdl.ShowSave
    Open Cdl.FileName For Append As #1
    Close #1
    Open App.Path & "\temp\temp.txt" For Append As #1
    Close #1
    Open App.Path & "\temp\temp.txt" For Input As #1
    Close #1
    
    Call ClearAlleles
    NewFile = True
    FileOpen = True
    
End Sub

Private Sub PrevAllMnu_Click()
    
    Call PreviewGeno1
    Call PreviewGeno2
    
End Sub

Private Sub RelNoteMnu_Click()
    
    MsgBox "This is BETA VERSION 1.0", , "BETA VERSION 1.0"
    
End Sub

Private Sub RevGamMnu_Click()
    
    Call CalculateGamete1(False)
    Call CalculateGamete2(False)
    
End Sub

Private Sub SaveMnu_Click()
    
    Call SaveWork(False, NewFile)
    
End Sub


Private Sub SaveAsMnu_Click()

    Call SaveWork(True, NewFile)
    
End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'Call NumericOnly                    'calls the subroutine to show numeric only
    
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    
    'Call NumericOnly                    'calls the subroutine to show numeric only
    
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Call NumericOnly                    'calls the subroutine to show numeric only
    
End Sub

Private Sub Text2_Change()

Dim indexsel As Integer                 'declares the variable
        
    If EditInProgress1 = False Then     'if an edit is not in progress then
        List(1).Text = Text2.Text         'find what is in text box 2
    ElseIf EditInProgress1 = True Then  'else if an edit is in progress then
        indexsel = List(1).ListIndex      'store the selected item's list index in a
                                        'variable
            List(1).AddItem Text2.Text, _
                    List(1).ListIndex     'add the text in text box 2
            List(1).RemoveItem _
                    List(1).ListIndex     'remove the old item
            List(1).ListIndex = indexsel  'select the item that was added
        
    End If
        
End Sub

Private Sub Text3_Change()

Dim indexsel As Integer                 'declares the variable
        
    If EditInProgress2 = False Then     'if an edit is not in progress then
        List(2).Text = Text3.Text         'find what is in text box 2
    ElseIf EditInProgress2 = True Then  'else if an edit is in progress then
        indexsel = List(2).ListIndex      'store the selected item's list index in a
                                        'variable
            List(2).AddItem Text3.Text, _
                    List(2).ListIndex     'add the text in text box 2
            List(2).RemoveItem _
                    List(2).ListIndex     'remove the old item
            List(2).ListIndex = indexsel  'select the item that was added
        
    End If
        
End Sub

Private Sub Text6_Change()

Dim indexsel As Integer                 'declares the variable
        
    If EditInProgress3 = False Then     'if an edit is not in progress then
        List(3).Text = Text6.Text         'find what is in text box 2
    ElseIf EditInProgress3 = True Then  'else if an edit is in progress then
        indexsel = List(3).ListIndex      'store the selected item's list index in a
                                        'variable
            List(3).AddItem Text6.Text, _
                    List(3).ListIndex     'add the text in text box 2
            List(3).RemoveItem _
                    List(3).ListIndex     'remove the old item
            List(3).ListIndex = indexsel  'select the item that was added
        
    End If
        
End Sub

Private Sub Text7_Change()

Dim indexsel As Integer                 'declares the variable
        
    If EditInProgress4 = False Then     'if an edit is not in progress then
        List(4).Text = Text7.Text         'find what is in text box 2
    ElseIf EditInProgress4 = True Then  'else if an edit is in progress then
        indexsel = List(4).ListIndex      'store the selected item's list index in a
                                        'variable
            List(4).AddItem Text7.Text, _
                    List(4).ListIndex     'add the text in text box 2
            List(4).RemoveItem _
                    List(4).ListIndex     'remove the old item
            List(4).ListIndex = indexsel  'select the item that was added
        
    End If
        
End Sub

Private Sub Timer1_Timer()
    
    DoEvents
    If Val(Text1.Text) = 0 Then
        Command13.Enabled = False
    ElseIf Val(Text1.Text) > 0 Then
        Command13.Enabled = True
    End If
    
End Sub

Private Sub PreviewGeno1()
        
    PreviewInProgress = True
    
Call DisableButtons

If List(1).ListCount <> List(2).ListCount Then
    MsgBox "The number of Alleles in lists 1 and 2 must be equal."
    GoTo PreviewGeno1End
End If
    
    
'PREVIEW FOR LISTS 1 & 2
    
LCount1 = List(1).ListCount - 1           'how many items are in list 1 & 2
                                        'subtract 1 because list begins with 0

Prnt = ""                               'sets prnt to "null"


For PrevCounter = 0 To LCount1          'this loop goes through lists 1 & 2 and stores
        DoEvents                        'each item in the list with a
                                        'carriage-return/linefeed combination as
    Prnt = Prnt & List(1).List(PrevCounter) & _
    List(2).List(PrevCounter) & vbCrLf                 'variable prnt
        DoEvents
        
                Lpercent = (PrevCounter + 1) _
                / (LCount1 + 1) * 100     'figures the percent to display
                Percent = Lpercent      'rounds long percent into integer
                    DoEvents
    ProgressBar1.Value = Lpercent       'updates progress bar to percent
                    DoEvents
                                        'displays status in status label
                
Label6.Caption = "Creating preview..."
Label7.Caption = PrevCounter + 1
Label8.Caption = LCount1 + 1
Label9.Caption = "Preview window 1"
Label10.Caption = Lpercent

                    DoEvents
    
Next

Text4.Text = Prnt                       '"prints" the contents of the variable in
                                        'text box 4
                                        
PreviewGeno1End:

List(1).ListIndex = 0                     'returns to top of list
List(2).ListIndex = 0                     'returns to top of list
List(1).ListIndex = -1                    'deselects the list
List(2).ListIndex = -1                    'deselects the list

ProgressBar1.Value = "0"

Label6.Caption = "Ready"                'sets the status label to 'Ready'
Label7.Caption = ""
Label8.Caption = ""
Label9.Caption = ""
Label10.Caption = ""


Call EnableButtons

    PreviewInProgress = False
                                        
End Sub

Private Sub PreviewGeno2()

    PreviewInProgress = True

Call DisableButtons

If List(3).ListCount <> List(4).ListCount Then
    MsgBox "The number of Alleles in lists 3 and 4 must be equal."
    GoTo PreviewGeno2End
End If

'PREVIEW FOR LISTS 3 & 4

LCount2 = List(3).ListCount - 1           'how many items are in list 3 & 4
                                        'subtract 1 because list begins with 0
                                        
Prnt = ""                               'sets prnt to "null"


For PrevCounter = 0 To LCount2              'this loop goes through lists 1 & 2 and stores
        DoEvents
                                            'each item in the list with a
                                            'carriage-return/linefeed combination as
    Prnt = Prnt & List(3).List(PrevCounter) & _
    List(4).List(PrevCounter) & vbCrLf                 'variable prnt
        DoEvents
        
        Lpercent = (PrevCounter + 1) _
                / (LCount2 + 1) * 100   'figures the percent to display
                Percent = Lpercent      'rounds long percent into integer
                    DoEvents
    ProgressBar1.Value = Lpercent       'updates progress bar to percent
                    DoEvents
                                        'displays status in status label
Label6.Caption = "Creating preview..."
Label7.Caption = PrevCounter + 1
Label8.Caption = LCount2 + 1
Label9.Caption = "Preview window 2"
Label10.Caption = Lpercent


                    DoEvents
        
Next

Text5.Text = Prnt                       '"prints" the contents of the variable in
                                        'text box 4
                                        
PreviewGeno2End:

List(3).ListIndex = 0                    'returns to top of list
List(4).ListIndex = 0                    'returns to top of list
List(3).ListIndex = -1                   'deselects the list
List(4).ListIndex = -1                   'deselects the list

ProgressBar1.Value = "0"                'sets the progress bar to 0%
Label6.Caption = "Ready"                'sets the status label to 'Ready'
Label7.Caption = ""
Label8.Caption = ""
Label9.Caption = ""
Label10.Caption = ""


Call EnableButtons                      'enable buttons

    PreviewInProgress = False           'set generating preview to false
                                        
End Sub

Private Sub NumericOnly()
    
    Text1.Text = Val(Text1.Text)        'keeps user from entering non-numeric values
    
End Sub

Private Sub TextList()
    
    If EditInProgress1 = True Then
        Text2.Text = List(1).Text
        Text2.SetFocus
        Text2.SelStart = 0
        Text2.SelLength = 2
        
    ElseIf EditInProgress2 = True Then
        Text3.Text = List(2).Text
        Text3.SetFocus
        Text3.SelStart = 0
        Text3.SelLength = 2
        
    ElseIf EditInProgress3 = True Then
        Text6.Text = List(3).Text
        Text6.SetFocus
        Text6.SelStart = 0
        Text6.SelLength = 2
        
    ElseIf EditInProgress4 = True Then
        Text7.Text = List(4).Text
        Text7.SetFocus
        Text7.SelStart = 0
        Text7.SelLength = 2
        
    Else
        Text2.Text = List(1).Text             'displays the item selected in the list
        Text3.Text = List(2).Text             'in the corresponding text box
        Text6.Text = List(3).Text             'to be edited whenever a new item
        Text7.Text = List(4).Text             'is selected
    End If
    
    'need to put editing code here
    
End Sub

Private Sub DisableButtons()
    
    If EditInProgress1 = True Then      'if user is editing list 1 then
        Command1.Enabled = True         'enable the edit button next to list 1
        Command3.Enabled = False        'disable the other edit buttons
        Command10.Enabled = False       'disable the other edit buttons
        Command12.Enabled = False       'disable the other edit buttons
        
    ElseIf EditInProgress2 = True Then  'else if user is editing list then
        Command1.Enabled = False        'disable the other edit buttons
        Command3.Enabled = True         'enable the edit button next to list 2
        Command10.Enabled = False       'disable the other edit buttons
        Command12.Enabled = False       'disable the other edit buttons
        
    ElseIf EditInProgress3 = True Then  'else if user is editing list 3 then
        Command1.Enabled = False        'disable the other edit buttons
        Command3.Enabled = False        'disable the other eidt buttons
        Command10.Enabled = True        'enable the edit button next to list 3
        Command12.Enabled = False       'disable the other edit buttons
        
    ElseIf EditInProgress2 = True Then  'elseif user is editing list 4 then
        Command1.Enabled = False        'disable the other edit buttons
        Command3.Enabled = False        'disable the other edit buttons
        Command10.Enabled = False       'disable the other edit buttons
        Command12.Enabled = True        'enable the edit button next to list 4
        
    End If                              'can't have an if without an end if
    
                                        'FROM THIS LINE
'   Command1.Enabled = False   'OMITED - negates code above
    Command2.Enabled = False
'   Command3.Enabled = False   'OMITED - negates code above
    Command4.Enabled = False
    Command5.Enabled = False
    Command6.Enabled = False
    Command7.Enabled = False
    Command8.Enabled = False
    Command9.Enabled = False
'   Command10.Enabled = False  'OMITED - negates code above
    Command11.Enabled = False
'   Command12.Enabled = False  'OMITED - negates code above
    Command13.Enabled = False
    Command14.Enabled = False
    Command15.Enabled = False
    Command16.Enabled = False
    Command17.Enabled = False
    Command18.Enabled = False
    Command19.Enabled = False
    Command20.Enabled = False
    Command21.Enabled = False
    Command22.Enabled = False
    
    Text1.Enabled = False
        
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Timer4.Enabled = False
    
                                            'TO THIS LINE
                                        'IS PRETTY SELF-EXPLANATORY
    
End Sub

Private Sub EnableButtons()
    
                                        'FROM THIS LINE
    
   Command1.Enabled = True   'OMITED - timer handles enable/disable of edit buttons
    Command2.Enabled = True
   Command3.Enabled = True   'OMITED - timer handles enable/disable of edit buttons
    Command4.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True
    Command7.Enabled = True
    Command8.Enabled = True
    Command9.Enabled = True
   Command10.Enabled = True  'OMITED - timer handles enable/disable of edit buttons
    Command11.Enabled = True
   Command12.Enabled = True  'OMITED - timer handles enable/disable of edit buttons
    Command13.Enabled = True
    Command14.Enabled = True
    Command15.Enabled = True
    Command16.Enabled = True
    Command17.Enabled = True
    Command18.Enabled = True
    Command19.Enabled = True
    Command20.Enabled = True
    Command21.Enabled = True
    Command22.Enabled = True
    
    Text1.Enabled = True
    
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Timer4.Enabled = False

                                        'TO THIS LINE
                                        'IS PRETTY SELF EXPLANATORY
    
End Sub

Private Sub Timer2_Timer()
    
'    If CalculateInProgress = True Then
'        Command1.Enabled = False
'        Command3.Enabled = False
'        Command10.Enabled = False
'        Command12.Enabled = False
'        Command2.Enabled = False
'        Command4.Enabled = False
'        Command11.Enabled = False
'        Command9.Enabled = False
'        Text2.Text = ""
'        Text2.Enabled = False
'        Text3.Text = ""
'        Text3.Enabled = False
'            Text4.Enabled = False
'            Text5.Enabled = False
'        Text6.Text = ""
'        Text6.Enabled = False
'        Text7.Text = ""
'        Text7.Enabled = False
'    ElseIf CalculateInProgress = False Then
'        Text2.Enabled = True
'        Text3.Enabled = True
'            Text4.Enabled = True
'            Text5.Enabled = True
'        Text6.Enabled = True
'        Text7.Enabled = True
'    End If
    
    If PreviewInProgress = True Then    'if the preview is being generated
        Command1.Enabled = False        'then disable the 'Change' button
        Command3.Enabled = False
        Command10.Enabled = False
        Command12.Enabled = False
        
        Command2.Enabled = False        'and disable the 'Remove' buttons
        Command4.Enabled = False
        Command11.Enabled = False
        Command9.Enabled = False
        
    End If
        
    If PreviewInProgress = True _
        Or CalculateInProgress = _
                            True Then   'if the preview is being generated
        Command1.Enabled = False        'then disable the 'Change' button
        Command3.Enabled = False
        Command10.Enabled = False
        Command12.Enabled = False
        
        Command2.Enabled = False        'and disable the 'Remove' buttons
        Command4.Enabled = False
        Command11.Enabled = False
        Command9.Enabled = False
    
    ElseIf EditInProgress1 = True Then      'if user is editing list 1 then
        Command1.Enabled = True         'enable the edit button next to list 1
        Command3.Enabled = False        'disable the other edit buttons
        Command10.Enabled = False       'disable the other edit buttons
        Command12.Enabled = False       'disable the other edit buttons
        
        Command2.Enabled = False        'and disable the 'Remove' buttons
        Command4.Enabled = False
        Command11.Enabled = False
        Command9.Enabled = False
        
    ElseIf EditInProgress2 = True Then  'else if user is editing list then
        Command1.Enabled = False        'disable the other edit buttons
        Command3.Enabled = True         'enable the edit button next to list 2
        Command10.Enabled = False       'disable the other edit buttons
        Command12.Enabled = False       'disable the other edit buttons
        
        Command2.Enabled = False        'and disable the 'Remove' buttons
        Command4.Enabled = False
        Command11.Enabled = False
        Command9.Enabled = False
        
    ElseIf EditInProgress3 = True Then  'else if user is editing list 3 then
        Command1.Enabled = False        'disable the other edit buttons
        Command3.Enabled = False        'disable the other eidt buttons
        Command10.Enabled = False       'disable the other edit buttons
        Command12.Enabled = True        'enable the edit button next to list 3
        
        Command2.Enabled = False        'and disable the 'Remove' buttons
        Command4.Enabled = False
        Command11.Enabled = False
        Command9.Enabled = False
        
    ElseIf EditInProgress4 = True Then  'elseif user is editing list 4 then
        Command1.Enabled = False        'disable the other edit buttons
        Command3.Enabled = False        'disable the other edit buttons
        Command10.Enabled = True        'enable the edit button next to list 4
        Command12.Enabled = False       'disable the other edit buttons
        
        Command2.Enabled = False        'and disable the 'Remove' buttons
        Command4.Enabled = False
        Command11.Enabled = False
        Command9.Enabled = False
        
    End If
    
    If EditInProgress1 = True Then      'if user is editing list 1 then
                
        List(1).Enabled = True           'enable the list being edited
        List(2).Enabled = False           'and disable the others
        List(3).Enabled = False           'so user can't select other
        List(4).Enabled = False            'lists while editing
        
        List(2).ListIndex = -1
        List(3).ListIndex = -1
        List(4).ListIndex = -1
                
    ElseIf EditInProgress2 = True Then  'else if user is editing list then
                
        List(1).Enabled = False           'enable the list being edited
        List(2).Enabled = True           'and disable the others
        List(3).Enabled = False           'so user can't select other
        List(4).Enabled = False            'lists while editing
        
        List(1).ListIndex = -1
        List(3).ListIndex = -1
        List(4).ListIndex = -1
        
    ElseIf EditInProgress3 = True Then  'else if user is editing list 3 then
                
        List(1).Enabled = False           'enable the list being edited
        List(2).Enabled = False           'and disable the others
        List(3).Enabled = True           'so user can't select other
        List(4).Enabled = False            'lists while editing
        
        List(1).ListIndex = -1
        List(2).ListIndex = -1
        List(4).ListIndex = -1
                
    ElseIf EditInProgress4 = True Then 'elseif user is editing list 4 then
                
        List(1).Enabled = False           'enable the list being edited
        List(2).Enabled = False           'and disable the others
        List(3).Enabled = False           'so user can't select other
        List(4).Enabled = True            'lists while editing
        
        List(1).ListIndex = -1
        List(2).ListIndex = -1
        List(3).ListIndex = -1
        
    ElseIf EditInProgress1 = False _
        And EditInProgress2 = False _
        And EditInProgress3 = False _
        And EditInProgress4 = False _
                                Then
        List(1).Enabled = True
        List(2).Enabled = True
        List(3).Enabled = True
        List(4).Enabled = True
    End If
    
    If List(1).ListIndex = -1 Then        'else if nothing is selected
        Command1.Enabled = False        'then disable the 'Change' button
    ElseIf List(1).ListIndex <> -1 _
    And PreviewInProgress = False _
    And CalculateInProgress = False Then   'else if something is selected
        Command1.Enabled = True         'then enable the 'Change' button
    End If                              'can't have an if without an 'End If'
    
    If List(2).ListIndex = -1 Then        'else if nothing is selected
        Command3.Enabled = False        'then disable the 'Change' button
    ElseIf List(2).ListIndex <> -1 _
    And PreviewInProgress = False _
    And CalculateInProgress = False Then   'else if something is selected
        Command3.Enabled = True         'then enable the 'Change' button
    End If                              'can't have an if without an 'End If'
    
    If List(3).ListIndex = -1 Then        'else if nothing is selected
        Command12.Enabled = False       'then disable the 'Change' button
    ElseIf List(3).ListIndex <> -1 _
    And PreviewInProgress = False _
    And CalculateInProgress = False Then   'else if something is selected
        Command12.Enabled = True        'then enable the 'Change' button
    End If                              'can't have an if without an 'End If'
    
    If List(4).ListIndex = -1 Then        'else if nothing is selected
        Command10.Enabled = False       'then disable the 'Change' button
    ElseIf List(4).ListIndex <> -1 _
    And PreviewInProgress = False _
    And CalculateInProgress = False Then   'else if something is selected
        Command10.Enabled = True        'then enable the 'Change' button
    End If                              'can't have an if without an 'End If'
    
End Sub


Private Sub Timer3_Timer()
    
    If List(1).ListCount = 0 Then
        Command5.Enabled = False
    ElseIf List(2).ListCount = 0 Then
        Command5.Enabled = False
    ElseIf List(1).ListCount <> 0 And PreviewInProgress = False _
    And CalculateInProgress = False Then
        Command5.Enabled = True
    ElseIf List(2).ListCount <> 0 And PreviewInProgress = False _
    And CalculateInProgress = False Then
        Command5.Enabled = True
    End If
    
    
    If List(3).ListCount = 0 Then
        Command8.Enabled = False
    ElseIf List(4).ListCount = 0 Then
        Command8.Enabled = False
    ElseIf List(3).ListCount <> 0 And PreviewInProgress = False _
    And CalculateInProgress = False Then
        Command8.Enabled = True
    ElseIf List(4).ListCount <> 0 And PreviewInProgress = False _
    And CalculateInProgress = False Then
        Command8.Enabled = True
    End If
    
    
    
    If List(1).ListIndex = -1 Then        'else if nothing is selected
        Command2.Enabled = False        'then disable the 'Change' button
    ElseIf List(1).ListIndex <> -1 _
    And EditInProgress1 = False _
    And EditInProgress2 = False _
    And EditInProgress3 = False _
    And EditInProgress4 = False _
    And PreviewInProgress = False _
    And CalculateInProgress = False Then 'else if something is selected
        Command2.Enabled = True         'then enable the 'Change' button
    End If                              'can't have an if without an 'End If'
    
    
    If List(2).ListIndex = -1 Then        'else if nothing is selected
        Command4.Enabled = False        'then disable the 'Change' button
    ElseIf List(2).ListIndex <> -1 _
    And EditInProgress1 = False _
    And EditInProgress2 = False _
    And EditInProgress3 = False _
    And EditInProgress4 = False _
    And PreviewInProgress = False _
    And CalculateInProgress = False Then 'else if something is selected
        Command4.Enabled = True         'then enable the 'Change' button
    End If                              'can't have an if without an 'End If'
    
    
    If List(3).ListIndex = -1 Then        'else if nothing is selected
        Command11.Enabled = False       'then disable the 'Change' button
    ElseIf List(3).ListIndex <> -1 _
    And EditInProgress1 = False _
    And EditInProgress2 = False _
    And EditInProgress3 = False _
    And EditInProgress4 = False _
    And PreviewInProgress = False _
    And CalculateInProgress = False Then 'else if something is selected
        Command11.Enabled = True        'then enable the 'Change' button
    End If                              'can't have an if without an 'End If'
    
    
    If List(4).ListIndex = -1 Then        'else if nothing is selected
        Command9.Enabled = False        'then disable the 'Change' button
    ElseIf List(4).ListIndex <> -1 _
    And EditInProgress1 = False _
    And EditInProgress2 = False _
    And EditInProgress3 = False _
    And EditInProgress4 = False _
    And PreviewInProgress = False _
    And CalculateInProgress = False Then 'else if something is selected
        Command9.Enabled = True         'then enable the 'Change' button
    End If                              'can't have an if without an 'End If'
    
    
    If Command5.Enabled = True _
            And Command8.Enabled _
                        = True Then
        Command6.Enabled = True
        Command7.Enabled = True
        CalMnu.Enabled = True
        RevGamMnu.Enabled = True
    ElseIf Command5.Enabled = False _
            Or Command8.Enabled = _
                            False Then
        Command6.Enabled = False
        Command7.Enabled = False
        CalMnu.Enabled = False
        RevGamMnu.Enabled = False
    End If
    
End Sub

Private Sub CalculateGamete1(Continue1 As Boolean)

Call DisableButtons

If List(1).ListCount <> List(2).ListCount Or _
    List(2).ListCount <> List(3).ListCount Or _
    List(3).ListCount <> List(4).ListCount Or _
    List(1).ListCount <> List(3).ListCount Or _
    List(1).ListCount <> List(4).ListCount Or _
    List(2).ListCount <> List(4).ListCount Or _
    List(2).ListCount <> List(4).ListCount Then _

MsgBox "All alleles have to be equal."

GoTo EndCalculate1

End If


Form3.Show

CalculateInProgress = True

Call DisableButtons

ReDim Geno1(2, (List(1).ListCount - 1))


'THIS STORES LIST 1 IN A 2-D ARRAY

    
    For ListStoreCounter = 0 To _
                (List(1).ListCount - 1)
            DoEvents
        Gex = 1                         'for list 1, duh
        Gey = ListStoreCounter          'for the item number
        Geno1(Gex, Gey) = _
        List(1).List(ListStoreCounter)    'stores the list in a 2d-array
        
        
        Lpercent = _
                (ListStoreCounter / _
                (List(1).ListCount)) _
                            * 100       'figures the percent to display
                Percent = Lpercent      'rounds long percent into integer
                ProgressBar1.Value = _
                Lpercent                'updates progress bar to percent
                
                                        'displays status in status label
Label6.Caption = "Storing List... " & Gex
Label7.Caption = (ListStoreCounter + 1)
Label8.Caption = List(1).ListCount
Label9.Caption = Geno1(Gex, Gey) & " (" & Gex & "," & Gey & ")"
Label10.Caption = Round(Lpercent, 10)
        
        
        
        
        Form3.Cls
        Form3.Print Geno1(Gex, Gey) & " " & "(" & Gex; "," & Gey & ")"
    Next
        List(1).ListIndex = -1

'THIS STORES LIST 2 IN A 2-D ARRAY
    
    For ListStoreCounter = 0 To _
                (List(2).ListCount - 1)
        DoEvents
        Gex = 2                         'for list 2, duh
        Gey = ListStoreCounter          'for the item number
        Geno1(Gex, Gey) = _
        List(2).List(ListStoreCounter)    'stores the list in a 2d-array
        
        
        
        Lpercent = _
                (ListStoreCounter / _
                (List(1).ListCount)) _
                            * 100       'figures the percent to display
                Percent = Lpercent      'rounds long percent into integer
                ProgressBar1.Value = _
                Lpercent                'updates progress bar to percent
                
                                        'displays status in status label



Label6.Caption = "Storing List... " & Gex
Label7.Caption = (ListStoreCounter + 1)
Label8.Caption = List(1).ListCount
Label9.Caption = Geno1(Gex, Gey) & " ( " & Gex & " , " & Gey & " )"
Label10.Caption = Round(Lpercent, 10)






'Label1.Caption = "Storing item " & (ListStoreCounter + 1) & " of " & list(1).ListCount & _
" from List " & Gex & " as " & geno1(Gex, Gey) & " With " & "Coordinates " & "(" & Gex & _
"," & Gey & ")" & "  -  " & Percent & "%"
        
        
        
        
        Form3.Cls
        Form3.Print Geno1(Gex, Gey) & " " & "(" & Gex; "," & Gey & ")"
    Next
        List(2).ListIndex = -1
    

'THIS FINDS THE POSSIBLE GAMETES FROM PARENT 1
    
ReDim Gamete1(List(1).ListCount, _
            (2 ^ (List(1).ListCount)))    'this declares it as a
                                        '2-d array
ReDim Preserve Geno1(2, _
        (List(1).ListCount - 1))          'redeclares it but preserves
                                        'the data contained in it
    

    Traits = (List(1).ListCount)          'how many traits are in list 1
    MaxX = Traits
    maxY = (2 ^ (Traits - 1))
    OuterPattern = 2                    'outerpattern always starts at 2
    Pattern = (2 ^ (Traits - 1))        'how many times to write the same character
    Gax = 1                             'gamete1 x coordinate starts at 1
    Gay = 1                             'gamete1 y coordinate starts at 1
    Gex = 1                             'genotype list 1
    Gey = 0                             'genotype item list index
    ItemCount = 1                       'sets the item count
    
    For OutCounter = 0 To (Traits)  'this loop tells it how many columns there are
            DoEvents
        For MidCounter = 1 To OuterPattern 'how many time to chg. between list 1 & list 2
                DoEvents
            For InCounter = 1 To Pattern 'how many time to write came charac. before chg.
                    DoEvents
                
                Gamete1(Gax, Gay) = _
                Geno1(Gex, Gey)          'finds gametes and writes them in a 2-d array
                                        'beginning with (1,1)
                
                
CoOrdinate = "(" & Gax & "," & Gay & ")" & " " & Gamete1(Gax, Gay) & " " & "(" & Gex & "," _
& Gey & ")"                             'this stores the item being stored to this
                                        'variable having the format:
                                        '(GameteX,GameteY) 'data' ('list#','listindex#')
                                        'where list# is GenoX and listindex# in GenoY
                                        
                Lpercent = _
                (ItemCount / _
                (Traits * (2 ^ Traits))) _
                            * 100       'figures the percent to display
                Percent = Lpercent      'rounds long percent into integer
                ProgressBar1.Value = _
                Lpercent                'updates progress bar to percent
                
                                        'displays status in status label


Label6.Caption = "Translating parent 1 table..."
Label7.Caption = ItemCount
Label8.Caption = (Traits * (2 ^ Traits))
Label9.Caption = CoOrdinate
Label10.Caption = Round(Lpercent, 10)









'Label1.Caption = "Translating table... " & "Item no. " & ItemCount & " of " _
& (Traits * (2 ^ (Traits))) & " from " & "(" & Gex & "," & Gey & ")" _
& " to " & "(" & Gax & "," & Gay & ")" & " data: " _
& gamete1(Gax, Gay) & "  -  " & Percent & "%"
                                        
                                        
                                        
                                        
                ItemCount = ItemCount + 1
                Form3.Cls
                Form3.Print CoOrdinate
                
                                        
                    Gay = Gay + 1       'starts a new row
            Next
                If Gex = 1 Then         'if it's on list 1 then
                    Gex = Gex + 1       'change to list 2
                ElseIf Gex = 2 Then     'else if it's on list 2 then
                    Gex = Gex - 1       'change to list 1
                End If                  'can't have an if without an end if
        Next
            Gax = Gax + 1               'starts a new column
            Gey = Gey + 1               'starts a new column
            Gay = 1                     'resets the GameteY value to 1
                                        'to start a new column
            OuterPattern = _
                OuterPattern * 2        'doubles the times it chg. between lists
                
            'If OuterPattern > (Traits ^ 2) Then
            '    Exit For
            'End If
                
            Pattern = Pattern / 2       'halves the times it prints a charc. before
                                        'changing to the next list
    Next
    

'THIS PRINTS THE 2-D ARRAY IN A TEXT BOX GOING THROUGH EVERY COLUMN AND ROW

    
    Dim xCounter As Long                'declares the GameteX columns
    Dim yCounter As Long                'declares the GameteY rows
    Dim X As Long                       'declares the X counter
    Dim Y As Long                       'declares the Y counter
    Dim Fgamete As String               'gamete1, or 1 row

        DoEvents
    Form2.Text1.Text = ""
    Form2.Text2.Text = ""
    Form2.Text3.Text = ""
    Fgamete = ""
    xCounter = Traits                   'GameteX = no. of columns
    yCounter = (2 ^ Traits)             'GameteY = no. of rows

    For Y = 1 To yCounter               'this loop goes through rows
                DoEvents
        For X = 1 To xCounter           'this loop goes through columns
                DoEvents
            Fgamete = Fgamete & _
            Gamete1(X, Y)  'this gets the data and puts it in a text box
            
            Form3.Cls
            Form3.Print Fgamete
            
            
            
            Lpercent = _
                (Y / yCounter) * 100    'figures the percent to display
                Percent = Lpercent      'rounds long percent into integer
                ProgressBar1.Value = _
                Lpercent                'updates progress bar to percent
                
                                        'displays status in status label
            
            
            
            
            
Label6.Caption = "Calculating parent 1 gametes..."
Label7.Caption = Y
Label8.Caption = yCounter

Label10.Caption = Round(Lpercent, 10)
                                   
        Next
            
            Form2.Text1.Text = _
            Form2.Text1.Text & _
            Fgamete & vbCrLf            'at the end of every row, inserts a
                                        '(carriage return/line feed) combo
            Label9.Caption = Fgamete    'this prints the found gamete in status label
            Fgamete = ""
            
    Next
    
    Form2.Label1.Caption = "Results:"
    'Form2.Show 1, Form1                 'shows form 2, having form 1 as the owner form
    
'Call EnableButtons                     'enables the buttons on form 1

Label6.Caption = "Ready"                'sets status labels to ready
Label7.Caption = ""
Label8.Caption = ""
Label9.Caption = ""
Label10.Caption = ""

EndCalculate1:

If Continue1 = True Then
    Call DisableButtons
    Call CalculateGamete2(True)
    Continue1 = False
    Exit Sub
ElseIf Continue1 = False Then
    CalculateInProgress = False
    Call EnableButtons
    Exit Sub
End If

End Sub

Private Sub CalculateGamete2(Continue2 As Boolean)

If List(1).ListCount <> List(2).ListCount Or _
    List(2).ListCount <> List(3).ListCount Or _
    List(3).ListCount <> List(4).ListCount Or _
    List(1).ListCount <> List(3).ListCount Or _
    List(1).ListCount <> List(4).ListCount Or _
    List(2).ListCount <> List(4).ListCount Or _
    List(2).ListCount <> List(4).ListCount Then _

MsgBox "All alleles have to be equal."

GoTo EndCalculate2

End If


Form3.Show

CalculateInProgress = True

Call DisableButtons

ReDim Geno2(2, (List(3).ListCount - 1))


'THIS STORES LIST 3 IN A 2-D ARRAY

    
    For ListStoreCounter = 0 To _
                (List(3).ListCount - 1)
            DoEvents
        Gex = 1                         'for list 1, duh
        Gey = ListStoreCounter          'for the item number
        Geno2(Gex, Gey) = _
        List(3).List(ListStoreCounter)    'stores the list in a 2d-array
        
        
        Lpercent = _
                (ListStoreCounter / _
                (List(3).ListCount)) _
                            * 100       'figures the percent to display
                Percent = Lpercent      'rounds long percent into integer
                ProgressBar1.Value = _
                Lpercent                'updates progress bar to percent
                
                                        'displays status in status label
Label6.Caption = "Storing List... " & Gex
Label7.Caption = (ListStoreCounter + 1)
Label8.Caption = List(3).ListCount
Label9.Caption = Geno2(Gex, Gey) & " (" & Gex & "," & Gey & ")"
Label10.Caption = Round(Lpercent, 10)
        
        
        
        
        Form3.Cls
        Form3.Print Geno2(Gex, Gey) & " " & "(" & Gex; "," & Gey & ")"
    Next
        List(3).ListIndex = -1

'THIS STORES LIST 4 IN A 2-D ARRAY
    
    For ListStoreCounter = 0 To _
                (List(4).ListCount - 1)
        DoEvents
        Gex = 2                         'for list 2, duh
        Gey = ListStoreCounter                   'for the item number
        Geno2(Gex, Gey) = _
        List(4).List(ListStoreCounter)     'stores the list in a 2d-array
        
        
        
        Lpercent = _
                (ListStoreCounter / _
                (List(3).ListCount)) _
                            * 100       'figures the percent to display
                Percent = Lpercent      'rounds long percent into integer
                ProgressBar1.Value = _
                Lpercent                'updates progress bar to percent
                
                                        'displays status in status label



Label6.Caption = "Storing List... " & Gex
Label7.Caption = (ListStoreCounter + 1)
Label8.Caption = List(3).ListCount
Label9.Caption = Geno2(Gex, Gey) & " ( " & Gex & " , " & Gey & " )"
Label10.Caption = Round(Lpercent, 10)






'Label1.Caption = "Storing item " & (ListStoreCounter + 1) & " of " & list(3).ListCount & _
" from List " & Gex & " as " & geno2(Gex, Gey) & " With " & "Coordinates " & "(" & Gex & _
"," & Gey & ")" & "  -  " & Percent & "%"
        
        
        
        
        Form3.Cls
        Form3.Print Geno2(Gex, Gey) & " " & "(" & Gex; "," & Gey & ")"
    Next
        List(4).ListIndex = -1
    

'THIS FINDS THE POSSIBLE GAMETES FROM PARENT 1
    
ReDim Gamete2(List(3).ListCount, _
            (2 ^ (List(3).ListCount)))    'this declares it as a
                                        '2-d array
ReDim Preserve Geno2(2, _
        (List(3).ListCount - 1))          'redeclares it but preserves
                                        'the data contained in it
    

    Traits = (List(3).ListCount)          'how many traits are in list 1
    MaxX = Traits
    maxY = (2 ^ (Traits - 1))
    OuterPattern = 2                    'outerpattern always starts at 2
    Pattern = (2 ^ (Traits - 1))        'how many times to write the same character
    Gax = 1                             'gamete2 x coordinate starts at 1
    Gay = 1                             'gamete2 y coordinate starts at 1
    Gex = 1                             'genotype list 1
    Gey = 0                             'genotype item list index
    ItemCount = 1                       'sets the item count
    
    For OutCounter = 0 To (Traits)  'this loop tells it how many columns there are
            DoEvents
        For MidCounter = 1 To OuterPattern 'how many time to chg. between list 1 & list 2
                DoEvents
            For InCounter = 1 To Pattern 'how many time to write came charac. before chg.
                    DoEvents
                
                Gamete2(Gax, Gay) = _
                Geno2(Gex, Gey)          'finds gametes and writes them in a 2-d array
                                        'beginning with (1,1)
                
                
CoOrdinate = "(" & Gax & "," & Gay & ")" & " " & Gamete2(Gax, Gay) & " " & "(" & Gex & "," _
& Gey & ")"                             'this stores the item being stored to this
                                        'variable having the format:
                                        '(GameteX,GameteY) 'data' ('list#','listindex#')
                                        'where list# is GenoX and listindex# in GenoY
                                        
                Lpercent = _
                (ItemCount / _
                (Traits * (2 ^ Traits))) _
                            * 100       'figures the percent to display
                Percent = Lpercent      'rounds long percent into integer
                ProgressBar1.Value = _
                Lpercent                'updates progress bar to percent
                
                                        'displays status in status label


Label6.Caption = "Translating parent 2 table..."
Label7.Caption = ItemCount
Label8.Caption = (Traits * (2 ^ Traits))
Label9.Caption = CoOrdinate
Label10.Caption = Round(Lpercent, 10)









'Label1.Caption = "Translating table... " & "Item no. " & ItemCount & " of " _
& (Traits * (2 ^ (Traits))) & " from " & "(" & Gex & "," & Gey & ")" _
& " to " & "(" & Gax & "," & Gay & ")" & " data: " _
& gamete2(Gax, Gay) & "  -  " & Percent & "%"
                                        
                                        
                                        
                                        
                ItemCount = ItemCount + 1
                Form3.Cls
                Form3.Print CoOrdinate
                
                                        
                    Gay = Gay + 1       'starts a new row
            Next
                If Gex = 1 Then         'if it's on list 1 then
                    Gex = Gex + 1       'change to list 2
                ElseIf Gex = 2 Then     'else if it's on list 2 then
                    Gex = Gex - 1       'change to list 1
                End If                  'can't have an if without an end if
        Next
            Gax = Gax + 1               'starts a new column
            Gey = Gey + 1               'starts a new column
            Gay = 1                     'resets the GameteY value to 1
                                        'to start a new column
            OuterPattern = _
                OuterPattern * 2        'doubles the times it chg. between lists
                
            'If OuterPattern > (Traits ^ 2) Then
            '    Exit For
            'End If
                
            Pattern = Pattern / 2       'halves the times it prints a charc. before
                                        'changing to the next list
    Next
    

'THIS PRINTS THE 2-D ARRAY IN A TEXT BOX GOING THROUGH EVERY COLUMN AND ROW

    
    Dim xCounter As Long                'declares the GameteX columns
    Dim yCounter As Long                'declares the GameteY rows
    Dim X As Long                       'declares the X counter
    Dim Y As Long                       'declares the Y counter
    Dim Fgamete As String               'gamete2, or 1 row

        DoEvents
    Form2.Text2.Text = ""
    Fgamete = ""
    xCounter = Traits                   'GameteX = no. of columns
    yCounter = (2 ^ Traits)             'GameteY = no. of rows

    For Y = 1 To yCounter               'this loop goes through rows
                DoEvents
        For X = 1 To xCounter           'this loop goes through columns
                DoEvents
            Fgamete = Fgamete & _
            Gamete2(X, Y)               'this gets the data and puts it in a text box
            
            Form3.Cls
            Form3.Print Fgamete
            
            
            
            Lpercent = _
                (Y / yCounter) * 100    'figures the percent to display
                Percent = Lpercent      'rounds long percent into integer
                ProgressBar1.Value = _
                Lpercent                'updates progress bar to percent
                
                                        'displays status in status label
            
            
            
            
            
Label6.Caption = "Calculating parent 2 gametes..."
Label7.Caption = Y
Label8.Caption = yCounter

Label10.Caption = Round(Lpercent, 10)
                                   
        Next
            
            Form2.Text2.Text = _
            Form2.Text2.Text & _
            Fgamete & vbCrLf            'at the end of every row, inserts a
                                        '(carriage return/line feed) combo
            Label9.Caption = Fgamete
            Fgamete = ""
    Next
    
    Form2.Label1.Caption = "Results:"
    'Form2.Show 1, Form1                 'shows form 2, having form 1 as the owner form
    
'Call EnableButtons                     'enables the buttons on form 1

Label6.Caption = "Ready"
Label7.Caption = ""
Label8.Caption = ""
Label9.Caption = ""
Label10.Caption = ""

EndCalculate2:

If Continue2 = True Then
    Call DisableButtons
    Call CalculateOffspring
    Continue2 = False
    Exit Sub
ElseIf Continue2 = False Then
    CalculateInProgress = False
    Call EnableButtons
    Exit Sub
End If


End Sub

Private Sub CalculateOffspring()

Call DisableButtons
    
    OutComes = (2 ^ Traits)
    Temp = ""
    ReDim Offspring(OutComes, OutComes)
    Pcounter = 0
    
    For OffY = 1 To OutComes
        DoEvents
        For OffX = 1 To OutComes
            DoEvents
            For GameteX = 1 To Traits
                DoEvents
                Temp = Temp & Gamete1(GameteX, OffY)
                Temp = Temp & Gamete2(GameteX, OffX)
                Label9.Caption = GameteX & _
                " (" & OffX & "," & OffY & ")"
            Next
            Pcounter = Pcounter + 1
            Offspring(OffX, OffY) = Temp
            Temp = ""
            
            Lpercent = (Pcounter / (OutComes ^ 2)) * 100
            ProgressBar1.Value = Lpercent

            Label6.Caption = "Translating offspring table..."
            Label7.Caption = Pcounter
            Label8.Caption = OutComes ^ 2
            Label10.Caption = Round(Lpercent, 10)
        Next
    Next
    
    
    
    Temp = ""
    Pcounter = 0
    
    Form2.Text3.Text = ""
    For OffY = 1 To OutComes
        DoEvents
        For OffX = 1 To OutComes
            DoEvents
            Form2.Text3.Text = Form2.Text3.Text & Offspring(OffX, OffY) & "   "
            
            
            
            
            Label9.Caption = Offspring(OffX, OffY)
            Pcounter = Pcounter + 1
            Label7.Caption = Pcounter
            Lpercent = (Pcounter / (OutComes ^ 2)) * 100
            ProgressBar1.Value = Lpercent
            Label10.Caption = Round(Lpercent, 10)
            
        Next
        Form2.Text3.Text = Form2.Text3.Text & vbCrLf & vbCrLf
        
        Label6.Caption = "Calculating offspring..."
        Label8.Caption = OutComes ^ 2
    Next
    
    Form2.Command2.Enabled = True
    Form2.Label1.Caption = "Results:"
    Form2.Show 0, Form1
    
    CalculateInProgress = False
    
    Call EnableButtons

End Sub

Private Sub ClearAlleles()
    
    List(1).Clear
    List(2).Clear
    List(3).Clear
    List(4).Clear
    
End Sub

Private Sub AskSaveYet()

    If FileOpen = True Then
        SaveResponse = MsgBox("Do you " _
        & "wish to save your changes?", _
        vbYesNo)
        If SaveResponse = vbYes Then
            Call SaveWork(False, NewFile)
        End If
    End If

End Sub
Private Sub SaveWork(SaveAsNew As Boolean, Fresh As Boolean)
    
    If SaveAsNew = True Then
        Cdl.DialogTitle = "Save file as..."
        Cdl.InitDir = App.Path
        Cdl.Filter = "Parent Genotypes (*.txt)|*.txt"
        Cdl.ShowSave
        Open Cdl.FileName For Append As #1
        Close #1
    ElseIf Fresh = True Then
        Cdl.DialogTitle = "Save file..."
        Cdl.InitDir = App.Path
        Cdl.Filter = "Parent Genotypes (*.txt)|*.txt"
        Cdl.ShowSave
        Open Cdl.FileName For Append As #1
        Close #1
    End If
    
        Open Cdl.FileName For Output As #1
        
        For SaveListNum = 1 To 4
            DoEvents
            For SaveList = 0 To (List(SaveListNum).ListCount - 1)
                DoEvents
    
                Write #1, "[" & SaveListNum & "]", List(SaveListNum).List(SaveList)
            Next
        Next
    
        Close #1
    
    NewFile = False
    
End Sub

Private Sub DetermineChange()
    Dim TempInput1 As String
    Dim TempInputNum1 As String
    Dim TempInput2 As String
    Dim TempInputNum2 As String
    
    Open App.Path & "\temp.txt" For Output As #1
        For SaveListNum = 1 To 4
            DoEvents
            For SaveList = 0 To (List(SaveListNum).ListCount - 1)
                DoEvents
    
                Write #1, "[" & SaveListNum & "]", List(SaveListNum).List(SaveList)
            Next
        Next
    
        Close #1
        
    If FileOpen = True Then
        
        Open Cdl.FileName For Input As #1
        Open App.Path & "\temp.txt" For Input As #2
        
        Do While Not EOF(1) And Not EOF(2)
            Input #1, TempInputNum1, TempInput1
            Input #2, TempInputNum2, TempInput2
                        
            If TempInputNum1 <> TempInputNum2 Then
                SaveYet = False
                NewFile = False
            ElseIf TempInput1 <> TempInput2 Then
                SaveYet = False
                NewFile = False
            End If
        Loop
        
        Close #1
        Close #2
    End If
        
End Sub

Private Sub Timer4_Timer()
    
    If FileOpen = False Then
        SaveMnu.Enabled = False
        SaveAsMnu.Enabled = False
        CloseMnu.Enabled = False
        ClrAllMnu.Enabled = False
        Text1.Enabled = False
        Text2.Enabled = False
        Text3.Enabled = False
        Text4.Enabled = False
        Text5.Enabled = False
        Text6.Enabled = False
        Text7.Enabled = False
        List(1).Enabled = False
        List(2).Enabled = False
        List(3).Enabled = False
        List(4).Enabled = False
    ElseIf FileOpen = True Then
        SaveMnu.Enabled = True
        SaveAsMnu.Enabled = True
        CloseMnu.Enabled = True
        ClrAllMnu.Enabled = False
        Text1.Enabled = True
        Text2.Enabled = True
        Text3.Enabled = True
        Text4.Enabled = True
        Text5.Enabled = True
        Text6.Enabled = True
        Text7.Enabled = True
        List(1).Enabled = True
        List(2).Enabled = True
        List(3).Enabled = True
        List(4).Enabled = True
    End If
    
End Sub


Private Sub Timer5_Timer()
        If List(1).ListIndex = 0 Then
            Command15.Enabled = True
            Command16.Enabled = False
        ElseIf List(1).ListIndex = (List(1).ListCount - 1) _
                And List(1).ListIndex <> -1 Then
            Command15.Enabled = False
            Command16.Enabled = True
        ElseIf List(1).ListIndex = -1 Then
        '    Command15.Enabled = False
        '    Command16.Enabled = False
        ElseIf List(2).ListIndex = 0 Then
            Command17.Enabled = True
            Command18.Enabled = False
        ElseIf List(2).ListIndex = (List(2).ListCount - 1) _
                And List(2).ListIndex <> -1 Then
            Command17.Enabled = False
            Command18.Enabled = True
        ElseIf List(2).ListIndex = -1 Then
        '    Command17.Enabled = False
        '    Command18.Enabled = False
        ElseIf List(3).ListIndex = 0 Then
            Command19.Enabled = True
            Command20.Enabled = False
        ElseIf List(3).ListIndex = (List(3).ListCount - 1) _
                And List(3).ListIndex <> -1 Then
            Command19.Enabled = False
            Command20.Enabled = True
        ElseIf List(3).ListIndex = -1 Then
        '    Command19.Enabled = False
        '    Command20.Enabled = False
        ElseIf List(4).ListIndex = 0 Then
            Command21.Enabled = True
            Command22.Enabled = False
        ElseIf List(4).ListIndex = (List(4).ListCount - 1) _
                And List(4).ListIndex <> -1 Then
            Command21.Enabled = False
            Command22.Enabled = True
        ElseIf List(4).ListIndex = -1 Then
        '    Command21.Enabled = False
        '    Command22.Enabled = False
        End If
        
End Sub


Private Sub viewmnu_Click()
    
    MsgBox "Nothing under here yet.", , "BETA VERSION 1.0"
    MsgBox "But if you think of something, let me know...", , "BETA VERSION 1.0"
    
End Sub


