VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PunnettSquareCalculator"
   ClientHeight    =   7320
   ClientLeft      =   4980
   ClientTop       =   4905
   ClientWidth     =   9240
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   9240
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   4080
      Top             =   3120
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1320
      Top             =   360
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
      TabIndex        =   33
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Frame Frame7 
      Caption         =   "Second Parent:"
      Height          =   3015
      Left            =   1920
      TabIndex        =   18
      Top             =   3240
      Width           =   7215
      Begin VB.Frame Frame8 
         Caption         =   "Genotypes:"
         Height          =   2655
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   6975
         Begin VB.Frame Frame11 
            Caption         =   "First Allele:"
            Height          =   2295
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   2055
            Begin VB.TextBox Text6 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               MaxLength       =   2
               TabIndex        =   43
               Text            =   "Text6"
               Top             =   240
               Width           =   1095
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
               TabIndex        =   30
               Top             =   840
               Width           =   1095
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
               TabIndex        =   39
               Top             =   1200
               Width           =   375
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
               TabIndex        =   38
               Top             =   1200
               Width           =   375
            End
            Begin VB.ListBox List3 
               Appearance      =   0  'Flat
               Height          =   1785
               ItemData        =   "Form1.frx":0000
               Left            =   1320
               List            =   "Form1.frx":0002
               TabIndex        =   31
               Top             =   240
               Width           =   615
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
               TabIndex        =   29
               Top             =   1560
               Width           =   1095
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Second Allele:"
            Height          =   2295
            Left            =   2280
            TabIndex        =   24
            Top             =   240
            Width           =   2055
            Begin VB.TextBox Text7 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   840
               MaxLength       =   2
               TabIndex        =   42
               Text            =   "Text7"
               Top             =   240
               Width           =   1095
            End
            Begin VB.ListBox List4 
               Appearance      =   0  'Flat
               Height          =   1785
               ItemData        =   "Form1.frx":0004
               Left            =   120
               List            =   "Form1.frx":0006
               TabIndex        =   27
               Top             =   240
               Width           =   615
            End
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
               TabIndex        =   26
               Top             =   840
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
               TabIndex        =   40
               Top             =   1200
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
               TabIndex        =   41
               Top             =   1200
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
               TabIndex        =   25
               Top             =   1560
               Width           =   1095
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Preview:"
            Height          =   2295
            Left            =   4440
            TabIndex        =   20
            Top             =   240
            Width           =   2415
            Begin VB.TextBox Text5 
               Height          =   1875
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   23
               Top             =   240
               Width           =   735
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
               TabIndex        =   22
               Top             =   600
               Width           =   1335
            End
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
               TabIndex        =   21
               Top             =   1440
               Width           =   1335
            End
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "First Parent:"
      Height          =   3015
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   7215
      Begin VB.Frame Frame3 
         Caption         =   "Genotypes:"
         Height          =   2655
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   6975
         Begin VB.Frame Frame6 
            Caption         =   "Preview:"
            Height          =   2295
            Left            =   4440
            TabIndex        =   14
            Top             =   240
            Width           =   2415
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
               TabIndex        =   17
               Top             =   1440
               Width           =   1335
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
               TabIndex        =   16
               Top             =   600
               Width           =   1335
            End
            Begin VB.TextBox Text4 
               Height          =   1875
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   15
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Second Allele:"
            Height          =   2295
            Left            =   2280
            TabIndex        =   6
            Top             =   240
            Width           =   2055
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
               TabIndex        =   12
               Top             =   840
               Width           =   1095
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
               TabIndex        =   37
               Top             =   1200
               Width           =   375
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
               TabIndex        =   36
               Top             =   1200
               Width           =   375
            End
            Begin VB.TextBox Text3 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   840
               MaxLength       =   2
               TabIndex        =   11
               Text            =   "Text3"
               Top             =   240
               Width           =   1095
            End
            Begin VB.ListBox List2 
               Appearance      =   0  'Flat
               Height          =   1785
               ItemData        =   "Form1.frx":0008
               Left            =   120
               List            =   "Form1.frx":000A
               TabIndex        =   7
               Top             =   240
               Width           =   615
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
               TabIndex        =   13
               Top             =   1560
               Width           =   1095
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "First Allele:"
            Height          =   2295
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   2055
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
               TabIndex        =   9
               Top             =   840
               Width           =   1095
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
               TabIndex        =   35
               Top             =   1200
               Width           =   375
            End
            Begin VB.TextBox Text2 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               MaxLength       =   2
               TabIndex        =   8
               Text            =   "Text2"
               Top             =   240
               Width           =   1095
            End
            Begin VB.ListBox List1 
               Appearance      =   0  'Flat
               Height          =   1785
               ItemData        =   "Form1.frx":000C
               Left            =   1320
               List            =   "Form1.frx":000E
               TabIndex        =   5
               Top             =   240
               Width           =   615
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
               TabIndex        =   34
               Top             =   1200
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
               TabIndex        =   10
               Top             =   1560
               Width           =   1095
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Number of Traits:"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
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
         TabIndex        =   32
         Top             =   600
         Width           =   1455
      End
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
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   120
      TabIndex        =   44
      Top             =   7080
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
      Left            =   7440
      TabIndex        =   54
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label9"
      Height          =   255
      Left            =   5640
      TabIndex        =   53
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label8"
      Height          =   255
      Left            =   3720
      TabIndex        =   52
      Top             =   6720
      Width           =   1815
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label7"
      Height          =   255
      Left            =   1920
      TabIndex        =   51
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
      Height          =   255
      Left            =   120
      TabIndex        =   50
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Percent Completed:"
      Height          =   255
      Left            =   7440
      TabIndex        =   49
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Data:"
      Height          =   255
      Left            =   5640
      TabIndex        =   48
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Total:"
      Height          =   255
      Left            =   3720
      TabIndex        =   47
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Item:"
      Height          =   255
      Left            =   1920
      TabIndex        =   46
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Action:"
      Height          =   255
      Left            =   120
      TabIndex        =   45
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu new 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu save 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu saveas 
         Caption         =   "Save &As"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu load 
         Caption         =   "&Load"
         Shortcut        =   ^L
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu close 
         Caption         =   "&Close"
         Shortcut        =   ^C
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu edit 
      Caption         =   "&Edit"
      Begin VB.Menu clralle 
         Caption         =   "Clear Alleles"
      End
   End
   Begin VB.Menu view 
      Caption         =   "&View"
   End
   Begin VB.Menu actions 
      Caption         =   "&Actions"
      Begin VB.Menu cal 
         Caption         =   "Calculate"
      End
      Begin VB.Menu revgam 
         Caption         =   "Review Gametes"
      End
      Begin VB.Menu prevall 
         Caption         =   "Preview All"
      End
      Begin VB.Menu opt 
         Caption         =   "Options"
         Begin VB.Menu optopt 
            Caption         =   "Optional Options"
            Begin VB.Menu relopt 
               Caption         =   "Really Optional"
               Begin VB.Menu warn 
                  Caption         =   "Don't Say I Didn't Warn You"
                  Begin VB.Menu redbut 
                     Caption         =   "THE RED BUTTON"
                  End
               End
            End
         End
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu helpcon 
         Caption         =   "Help Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu relnote 
         Caption         =   "Release Notes"
      End
      Begin VB.Menu about 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim LCount1 As Integer                  'count for storing lists 1 & 2 in a variable
Dim LCount2 As Integer                  'count for storing lists 3 & 4 in a variable
Dim LCount3 As Integer                  'count for storing lists 1 & 2 in a variable
Dim LCount4 As Integer                  'count for storing lists 3 & 4 in a variable

Dim AddAlleleCounter As Integer
Dim diff As Integer
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

Dim ListStoreCounter As Long            'list storage counter
Dim CoOrdinate As String                'gamete cordinate 'data' geno cordinate
Dim ItemCount As Long                   'item no. currently being translated
Dim MaxX As Long                        'max no. of columns
Dim maxY As Long                        'man no. of rows
Dim Gex As Long                         'genotype x coordinate
Dim Gey As Long                         'genotype y coordinate
Dim Gax As Long                         'gamete x coordinate
Dim Gay As Long                         'gamete y coordinate
Dim Geno()                              'genotype 2-d array
Dim Gamete()                            'gamete 2-d array
Dim OuterPattern As Long                'how many times to switch from list1 to list2
Dim Pattern As Long                     'how many times to print same charac. in column
Public Traits As Long                   'how many traits there are
Dim InCounter As Long                   'the inner loop counter
Dim MidCounter As Long                  'the middle loop counter
Dim OutCounter As Long                  'declares the outer loop counter



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
                (List1.ListCount - 1)
            DoEvents
            
            List1.ListIndex = _
            EmptyListCheck
            
            If List1.Text = "" Then
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
                (List4.ListCount - 1)
            DoEvents
            
            List4.ListIndex = _
            EmptyListCheck
            
            If List4.Text = "" Then
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
        
        End If

EndEditing4:
        
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
                (List3.ListCount - 1)
            DoEvents
            
            List3.ListIndex = _
            EmptyListCheck
            
            If List3.Text = "" Then
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
        
        End If
        
EndEditing3:
        
End Sub

Private Sub Command13_Click()
        
Call DisableButtons
        
    If Val(Text1.Text) Mod 2 <> 0 Then
        MsgBox "You have to enter a number divisible by 2."
        GoTo EndAddAlleles
    End If
    
    If Val(Text1.Text) > 100 Then
        MsgBox "Dude, you've gotta be kiddin me..."
    End If
        
    LCount1 = List1.ListCount           'gets the number of items in list 1
    LCount2 = List2.ListCount           'gets the number of items in list 2
    LCount3 = List3.ListCount           'gets the number of items in list 3
    LCount4 = List4.ListCount           'gets the number of items in list 4
    
    If Val(Text1.Text) > LCount1 Then   'if the no. in text1 is greater than
        diff = Val(Text1.Text) - LCount1
            For AddAlleleCounter = 1 To diff     'list1's count, then add the no. of
                    DoEvents
                List1.AddItem _
                (List1.ListCount + 1)   'the diff. to make the list have that
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

Label6.Caption = "Adding item"
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
                List2.AddItem _
                (List2.ListCount + 1)   'the diff. to make the list have that
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

Label6.Caption = "Adding item"
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
                List3.AddItem _
                (List3.ListCount + 1)   'the diff. to make the list have that
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
Label6.Caption = "Adding item"
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
                List4.AddItem _
                (List4.ListCount + 1)   'the diff. to make the list have that
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
Label6.Caption = "Adding item"
Label7.Caption = AddAlleleCounter
Label8.Caption = diff
Label9.Caption = "Parent 2 " & "Allele 2"
Label10.Caption = Round(Lpercent, 10)

                    DoEvents
            Next                        'many items without adding too much
    End If                              'or not adding enough

EndAddAlleles:

ProgressBar1.Value = "0"
Label1.Caption = "Ready"

Call EnableButtons

Text1.Text = ""
    
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
                (List2.ListCount - 1)
            DoEvents
            
            List2.ListIndex = _
            EmptyListCheck
            
            If List2.Text = "" Then
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
        
        End If
        
EndEditing2:
        
End Sub

Private Sub Command5_Click()

    Call PreviewGeno1                   'calls the subroutine to preview genotypes

End Sub

Private Sub Command6_Click()

Form3.Show

CalculateInProgress = True

Call DisableButtons

ReDim Geno(2, (List1.ListCount - 1))


'THIS STORES LIST 1 IN A 2-D ARRAY

Debug.Print "----------"
    
    For ListStoreCounter = 0 To _
                (List1.ListCount - 1)
            DoEvents
        List1.ListIndex = _
                    ListStoreCounter
        Gex = 1                         'for list 1, duh
        Gey = ListStoreCounter          'for the item number
        Geno(Gex, Gey) = List1.Text     'stores the list in a 2d-array
        
        
        Lpercent = _
                (ListStoreCounter / _
                (List1.ListCount - 1)) _
                            * 100       'figures the percent to display
                Percent = Lpercent      'rounds long percent into integer
                ProgressBar1.Value = _
                Lpercent                'updates progress bar to percent
                
                                        'displays status in status label
Label6.Caption = "Storing List... " & Gex
Label7.Caption = (ListStoreCounter + 1)
Label8.Caption = List1.ListCount
Label9.Caption = Geno(Gex, Gey) & " ( " & Gex & " , " & Gey & " )"
Label10.Caption = Round(Lpercent, 10)
        
        
        
        
        Form3.Cls
        Debug.Print Geno(Gex, Gey) & " " & "(" & Gex; "," & Gey & ")"
    Next
        List1.ListIndex = -1

'THIS STORES LIST 2 IN A 2-D ARRAY

'Debug.Print "---"
    
    For ListStoreCounter = 0 To _
                (List2.ListCount - 1)
        List2.ListIndex = _
                ListStoreCounter
        Gex = 2                         'for list 2, duh
        Gey = ListStoreCounter                   'for the item number
        Geno(Gex, Gey) = List2.Text     'stores the list in a 2d-array
        
        
        
        Lpercent = _
                (ListStoreCounter / _
                (List1.ListCount - 1)) _
                            * 100       'figures the percent to display
                Percent = Lpercent      'rounds long percent into integer
                ProgressBar1.Value = _
                Lpercent                'updates progress bar to percent
                
                                        'displays status in status label



Label6.Caption = "Storing List... " & Gex
Label7.Caption = (ListStoreCounter + 1)
Label8.Caption = List1.ListCount
Label9.Caption = Geno(Gex, Gey) & " ( " & Gex & " , " & Gey & " )"
Label10.Caption = Round(Lpercent, 10)






'Label1.Caption = "Storing item " & (ListStoreCounter + 1) & " of " & List1.ListCount & _
" from List " & Gex & " as " & Geno(Gex, Gey) & " With " & "Coordinates " & "(" & Gex & _
"," & Gey & ")" & "  -  " & Percent & "%"
        
        
        
        
        Form3.Cls
        Debug.Print Geno(Gex, Gey) & " " & "(" & Gex; "," & Gey & ")"
    Next
        List2.ListIndex = -1
    

'THIS FINDS THE POSSIBLE GAMETES FROM PARENT 1
    
ReDim Gamete(List1.ListCount, _
            (2 ^ (List1.ListCount)))    'this declares it as a
                                        '2-d array
ReDim Preserve Geno(2, _
        (List1.ListCount - 1))          'redeclares it but preserves
                                        'the data contained in it
    

    Traits = (List1.ListCount)          'how many traits are in list 1
    MaxX = Traits
    maxY = (2 ^ (Traits - 1))
    OuterPattern = 2                    'outerpattern always starts at 2
    Pattern = (2 ^ (Traits - 1))        'how many times to write the same character
    Gax = 1                             'gamete x coordinate starts at 1
    Gay = 1                             'gamete y coordinate starts at 1
    Gex = 1                             'genotype list 1
    Gey = 0                             'genotype item list index
    ItemCount = 1                       'sets the item count
    
    For OutCounter = 0 To (Traits)  'this loop tells it how many columns there are
            DoEvents
        For MidCounter = 1 To OuterPattern 'how many time to chg. between list 1 & list 2
                DoEvents
            For InCounter = 1 To Pattern 'how many time to write came charac. before chg.
                    DoEvents
                
                Gamete(Gax, Gay) = _
                Geno(Gex, Gey)          'finds gametes and writes them in a 2-d array
                                        'beginning with (1,1)
                
                
CoOrdinate = "(" & Gax & "," & Gay & ")" & " " & Gamete(Gax, Gay) & " " & "(" & Gex & "," _
& Gey & ")"                             'this stores the item being stored to this
                                        'variable having the format:
                                        '(GameteX,GameteY) 'data' ('list#','listindex#')
                                        'where list# is GenoX and listindex# in GenoY
                                        
                 Debug.Print CoOrdinate
                                        
                Lpercent = _
                (ItemCount / _
                (Traits * (2 ^ Traits))) _
                            * 100       'figures the percent to display
                Percent = Lpercent      'rounds long percent into integer
                ProgressBar1.Value = _
                Lpercent                'updates progress bar to percent
                
                                        'displays status in status label


Label6.Caption = "Translating table..."
Label7.Caption = ItemCount
Label8.Caption = (Traits * (2 ^ Traits))
Label9.Caption = CoOrdinate
Label10.Caption = Round(Lpercent, 10)









'Label1.Caption = "Translating table... " & "Item no. " & ItemCount & " of " _
& (Traits * (2 ^ (Traits))) & " from " & "(" & Gex & "," & Gey & ")" _
& " to " & "(" & Gax & "," & Gay & ")" & " data: " _
& Gamete(Gax, Gay) & "  -  " & Percent & "%"
                                        
                                        
                                        
                                        
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
    Dim Fgamete As String               'gamete, or 1 row

    Debug.Print "---"
        DoEvents
    Form2.Text1.Text = ""
    Fgamete = ""
    xCounter = Traits                   'GameteX = no. of columns
    yCounter = (2 ^ Traits)             'GameteY = no. of rows

    For Y = 1 To yCounter               'this loop goes through rows
                DoEvents
        For X = 1 To xCounter           'this loop goes through columns
                DoEvents
            Fgamete = Fgamete & _
            Gamete(X, Y)  'this gets the data and puts it in a text box
            
            Form3.Cls
            Form3.Print Fgamete
            
            
            
            Lpercent = _
                (Y / yCounter) * 100    'figures the percent to display
                Percent = Lpercent      'rounds long percent into integer
                ProgressBar1.Value = _
                Lpercent                'updates progress bar to percent
                
                                        'displays status in status label
            
            
            
            
            
Label6.Caption = "Calculating gametes..."
Label7.Caption = Y
Label8.Caption = yCounter
Label9.Caption = Fgamete
Label10.Caption = Round(Lpercent, 10)
                                   
        Next
            
            Form2.Text1.Text = _
            Form2.Text1.Text & _
            Fgamete & vbCrLf            'at the end of every row, inserts a
                                        '(carriage return/line feed) combo
            Debug.Print Fgamete
            Fgamete = ""
    Next
   
    Form2.Show 1, Form1                 'shows form 2, having form 1 as the owner form
    
'Call EnableButtons                     'enables the buttons on form 1

Label6.Caption = "Ready"
Label7.Caption = ""
Label8.Caption = ""
Label9.Caption = ""
Label10.Caption = ""

CalculateInProgress = False

End Sub

Private Sub Command7_Click()

Call DisableButtons



Call EnableButtons

End Sub

Private Sub Command8_Click()
    
    Call PreviewGeno2
    
End Sub

Private Sub Form_Load()
        
    EditInProgress1 = False             'sets the editing status to false
    EditInProgress2 = False             'sets the editing status to false
    EditInProgress3 = False             'sets the editing status to false
    EditInProgress4 = False             'sets the editing status to false
    
    List1.AddItem "A"
    List1.AddItem "B"
    List1.AddItem "D"
    List1.AddItem "G"
    List1.AddItem "H"
    List1.AddItem "I"
    List1.AddItem "J"
    List1.AddItem "K"
    
    List2.AddItem "a"
    List2.AddItem "b"
    List2.AddItem "d"
    List2.AddItem "g"
    List2.AddItem "h"
    List2.AddItem "i"
    List2.AddItem "j"
    List2.AddItem "k"
    
            
'Call PreviewGeno                       'calls the subroutine to preview genotypes
    
End Sub

Private Sub List1_Click()
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
    
End Sub

Private Sub List1_GotFocus()
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
    
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
    
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
    
End Sub

Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
    
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
    
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
    
End Sub

Private Sub List2_Click()
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
        
End Sub

Private Sub List2_GotFocus()
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
    
End Sub

Private Sub List2_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
    
End Sub

Private Sub List2_KeyPress(KeyAscii As Integer)
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
    
End Sub

Private Sub List2_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
    
End Sub

Private Sub List2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
    
End Sub

Private Sub List2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
    
End Sub

Private Sub List3_Click()
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
    
End Sub

Private Sub List3_GotFocus()
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
    
End Sub

Private Sub List3_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
    
End Sub

Private Sub List3_KeyPress(KeyAscii As Integer)
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
    
End Sub

Private Sub List3_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
    
End Sub

Private Sub List3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
    
End Sub

Private Sub List3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
    
End Sub

Private Sub List4_Click()
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
    
End Sub

Private Sub List4_GotFocus()
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
    
End Sub

Private Sub List4_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
    
End Sub

Private Sub List4_KeyPress(KeyAscii As Integer)
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
    
End Sub

Private Sub List4_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
    
End Sub

Private Sub List4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
    
End Sub

Private Sub List4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call TextList                       'calls the subroutine to display list text
                                        'in the corresponding text box
    
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call NumericOnly                    'calls the subroutine to show numeric only
    
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    
    Call NumericOnly                    'calls the subroutine to show numeric only
    
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Call NumericOnly                    'calls the subroutine to show numeric only
    
End Sub

Private Sub Text2_Change()

Dim indexsel As Integer                 'declares the variable
        
    If EditInProgress1 = False Then     'if an edit is not in progress then
        List1.Text = Text2.Text         'find what is in text box 2
    ElseIf EditInProgress1 = True Then  'else if an edit is in progress then
        indexsel = List1.ListIndex      'store the selected item's list index in a
                                        'variable
            List1.AddItem Text2.Text, _
                    List1.ListIndex     'add the text in text box 2
            List1.RemoveItem _
                    List1.ListIndex     'remove the old item
            List1.ListIndex = indexsel  'select the item that was added
        
    End If
        
End Sub

Private Sub Text3_Change()

Dim indexsel As Integer                 'declares the variable
        
    If EditInProgress2 = False Then     'if an edit is not in progress then
        List2.Text = Text3.Text         'find what is in text box 2
    ElseIf EditInProgress2 = True Then  'else if an edit is in progress then
        indexsel = List2.ListIndex      'store the selected item's list index in a
                                        'variable
            List2.AddItem Text3.Text, _
                    List2.ListIndex     'add the text in text box 2
            List2.RemoveItem _
                    List2.ListIndex     'remove the old item
            List2.ListIndex = indexsel  'select the item that was added
        
    End If
        
End Sub

Private Sub Text6_Change()

Dim indexsel As Integer                 'declares the variable
        
    If EditInProgress3 = False Then     'if an edit is not in progress then
        List3.Text = Text6.Text         'find what is in text box 2
    ElseIf EditInProgress3 = True Then  'else if an edit is in progress then
        indexsel = List3.ListIndex      'store the selected item's list index in a
                                        'variable
            List3.AddItem Text6.Text, _
                    List3.ListIndex     'add the text in text box 2
            List3.RemoveItem _
                    List3.ListIndex     'remove the old item
            List3.ListIndex = indexsel  'select the item that was added
        
    End If
        
End Sub

Private Sub Text7_Change()

Dim indexsel As Integer                 'declares the variable
        
    If EditInProgress4 = False Then     'if an edit is not in progress then
        List4.Text = Text7.Text         'find what is in text box 2
    ElseIf EditInProgress4 = True Then  'else if an edit is in progress then
        indexsel = List4.ListIndex      'store the selected item's list index in a
                                        'variable
            List4.AddItem Text7.Text, _
                    List4.ListIndex     'add the text in text box 2
            List4.RemoveItem _
                    List4.ListIndex     'remove the old item
            List4.ListIndex = indexsel  'select the item that was added
        
    End If
        
End Sub

Private Sub Timer1_Timer()
    
    DoEvents
    Call NumericOnly                    'calls the subroutine to show numeric only
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

If List1.ListCount <> List2.ListCount Then
    MsgBox "The number of Alleles in lists 1 and 2 must be equal."
    GoTo PreviewGeno1End
End If
    
    
'PREVIEW FOR LISTS 1 & 2
    
LCount1 = List1.ListCount - 1           'how many items are in list 1 & 2
                                        'subtract 1 because list begins with 0

Prnt = ""                               'sets prnt to "null"


For PrevCounter = 0 To LCount1          'this loop goes through lists 1 & 2 and stores
        DoEvents
    List1.ListIndex = PrevCounter       'each item in the list with a
        DoEvents
    List2.ListIndex = PrevCounter       'carriage-return/linefeed combination as
        DoEvents
    Prnt = Prnt & List1.Text & _
    List2.Text & vbCrLf                 'variable prnt
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

List1.ListIndex = 0                     'returns to top of list
List2.ListIndex = 0                     'returns to top of list
List1.ListIndex = -1                    'deselects the list
List2.ListIndex = -1                    'deselects the list

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

If List3.ListCount <> List4.ListCount Then
    MsgBox "The number of Alleles in lists 3 and 4 must be equal."
    GoTo PreviewGeno2End
End If

'PREVIEW FOR LISTS 3 & 4

LCount2 = List3.ListCount - 1           'how many items are in list 3 & 4
                                        'subtract 1 because list begins with 0
                                        
Prnt = ""                               'sets prnt to "null"


For PrevCounter = 0 To LCount2              'this loop goes through lists 1 & 2 and stores
        DoEvents
    List3.ListIndex = PrevCounter           'each item in the list with a
        DoEvents
    List4.ListIndex = PrevCounter           'carriage-return/linefeed combination as
        DoEvents
    Prnt = Prnt & List3.Text & _
    List4.Text & vbCrLf                 'variable prnt
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

List3.ListIndex = 0                    'returns to top of list
List4.ListIndex = 0                    'returns to top of list
List3.ListIndex = -1                   'deselects the list
List4.ListIndex = -1                   'deselects the list

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
        Text2.Text = List1.Text
        Text2.SetFocus
        Text2.SelStart = 0
        Text2.SelLength = 2
        
    ElseIf EditInProgress2 = True Then
        Text3.Text = List2.Text
        Text3.SetFocus
        Text3.SelStart = 0
        Text3.SelLength = 2
        
    ElseIf EditInProgress3 = True Then
        Text6.Text = List3.Text
        Text6.SetFocus
        Text6.SelStart = 0
        Text6.SelLength = 2
        
    ElseIf EditInProgress4 = True Then
        Text7.Text = List4.Text
        Text7.SetFocus
        Text7.SelStart = 0
        Text7.SelLength = 2
        
    Else
        Text2.Text = List1.Text             'displays the item selected in the list
        Text3.Text = List2.Text             'in the corresponding text box
        Text6.Text = List3.Text             'to be edited whenever a new item
        Text7.Text = List4.Text             'is selected
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
    
                                        'TO THIS LINE
                                        'IS PRETTY SELF-EXPLANATORY
    
End Sub

Private Sub EnableButtons()
    
                                        'FROM THIS LINE
    
'   Command1.Enabled = True   'OMITED - timer handles enable/disable of edit buttons
    Command2.Enabled = True
'   Command3.Enabled = True   'OMITED - timer handles enable/disable of edit buttons
    Command4.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True
    Command7.Enabled = True
    Command8.Enabled = True
    Command9.Enabled = True
'   Command10.Enabled = True  'OMITED - timer handles enable/disable of edit buttons
    Command11.Enabled = True
'   Command12.Enabled = True  'OMITED - timer handles enable/disable of edit buttons
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

                                        'TO THIS LINE
                                        'IS PRETTY SELF EXPLANATORY
    
End Sub

Private Sub Timer2_Timer()
    
    If CalculateInProgress = True Then
        Command1.Enabled = False
        Command3.Enabled = False
        Command10.Enabled = False
        Command12.Enabled = False
        Text2.Text = ""
        Text2.Enabled = False
        Text3.Text = ""
        Text3.Enabled = False
            Text4.Enabled = False
            Text5.Enabled = False
        Text6.Text = ""
        Text6.Enabled = False
        Text7.Text = ""
        Text7.Enabled = False
    ElseIf CalculateInProgress = False Then
        Text2.Enabled = True
        Text3.Enabled = True
            Text4.Enabled = True
            Text5.Enabled = True
        Text6.Enabled = True
        Text7.Enabled = True
    End If
    
    If PreviewInProgress = True Then    'if the preview is being generated
        Command1.Enabled = False        'then disable the 'Change' button
        Command3.Enabled = False
        Command10.Enabled = False
        Command12.Enabled = False
    End If
        
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
        Command10.Enabled = False       'disable the other edit buttons
        Command12.Enabled = True        'enable the edit button next to list 3
        
    ElseIf EditInProgress4 = True Then  'elseif user is editing list 4 then
        Command1.Enabled = False        'disable the other edit buttons
        Command3.Enabled = False        'disable the other edit buttons
        Command10.Enabled = True        'enable the edit button next to list 4
        Command12.Enabled = False       'disable the other edit buttons
        
    End If
    
    If EditInProgress1 = True Then      'if user is editing list 1 then
                
        List1.Enabled = True           'enable the list being edited
        List2.Enabled = False           'and disable the others
        List3.Enabled = False           'so user can't select other
        List4.Enabled = False            'lists while editing
        
        List2.ListIndex = -1
        List3.ListIndex = -1
        List4.ListIndex = -1
                
    ElseIf EditInProgress2 = True Then  'else if user is editing list then
                
        List1.Enabled = False           'enable the list being edited
        List2.Enabled = True           'and disable the others
        List3.Enabled = False           'so user can't select other
        List4.Enabled = False            'lists while editing
        
        List1.ListIndex = -1
        List3.ListIndex = -1
        List4.ListIndex = -1
        
    ElseIf EditInProgress3 = True Then  'else if user is editing list 3 then
                
        List1.Enabled = False           'enable the list being edited
        List2.Enabled = False           'and disable the others
        List3.Enabled = True           'so user can't select other
        List4.Enabled = False            'lists while editing
        
        List1.ListIndex = -1
        List2.ListIndex = -1
        List4.ListIndex = -1
                
    ElseIf EditInProgress4 = True Then 'elseif user is editing list 4 then
                
        List1.Enabled = False           'enable the list being edited
        List2.Enabled = False           'and disable the others
        List3.Enabled = False           'so user can't select other
        List4.Enabled = True            'lists while editing
        
        List1.ListIndex = -1
        List2.ListIndex = -1
        List3.ListIndex = -1
        
    ElseIf EditInProgress1 = False _
        And EditInProgress2 = False _
        And EditInProgress3 = False _
        And EditInProgress4 = False _
                                Then
        List1.Enabled = True
        List2.Enabled = True
        List3.Enabled = True
        List4.Enabled = True
    End If
    
    If List1.ListIndex = -1 Then    'else if nothing is selected
        Command1.Enabled = False        'then disable the 'Change' button
    ElseIf List1.ListIndex <> -1 Then   'else if something is selected
        Command1.Enabled = True         'then enable the 'Change' button
    End If                              'can't have an if without an 'End If'
    
    If List2.ListIndex = -1 Then    'else if nothing is selected
        Command3.Enabled = False        'then disable the 'Change' button
    ElseIf List2.ListIndex <> -1 Then   'else if something is selected
        Command3.Enabled = True         'then enable the 'Change' button
    End If                              'can't have an if without an 'End If'
    
    If List3.ListIndex = -1 Then    'else if nothing is selected
        Command12.Enabled = False        'then disable the 'Change' button
    ElseIf List3.ListIndex <> -1 Then   'else if something is selected
        Command12.Enabled = True         'then enable the 'Change' button
    End If                              'can't have an if without an 'End If'
    
    If List4.ListIndex = -1 Then    'else if nothing is selected
        Command10.Enabled = False        'then disable the 'Change' button
    ElseIf List4.ListIndex <> -1 Then   'else if something is selected
        Command10.Enabled = True         'then enable the 'Change' button
    End If                              'can't have an if without an 'End If'
    
End Sub
