VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form2 
   Caption         =   "Punnett Square Calculator - Confirm gamete results"
   ClientHeight    =   12840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10215
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   12840
   ScaleWidth      =   10215
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9240
      Top             =   0
   End
   Begin MSComDlg.CommonDialog Cdl 
      Left            =   9720
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Save Results"
      Height          =   375
      Left            =   7560
      TabIndex        =   8
      Top             =   12360
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "Form2.frx":0000
      Top             =   5760
      Width           =   9735
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   5280
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Text            =   "Form2.frx":0007
      Top             =   840
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "Form2.frx":000E
      Top             =   840
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   8880
      TabIndex        =   0
      Top             =   12360
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parent 1:"
      Height          =   4935
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   4935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Parent 2:"
      Height          =   4935
      Left            =   5160
      TabIndex        =   5
      Top             =   600
      Width           =   4935
   End
   Begin VB.Frame Frame3 
      Caption         =   "Possible Offspring:"
      Height          =   6735
      Left            =   120
      TabIndex        =   6
      Top             =   5520
      Width           =   9975
   End
   Begin VB.Label Label1 
      Caption         =   "Results:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
    
    Me.Hide
    
    Command2.Enabled = False
    
End Sub

Private Sub SaveResults()
    
    Cdl.DialogTitle = "Save results..."
    Cdl.InitDir = App.Path
    Cdl.Filter = "Text File (*.txt)|*.txt"
    Cdl.ShowSave
    Open Cdl.FileName For Append As #1
    Close #1
    
    Open Cdl.FileName For Input As #1
    
    Write #1, "PARENT 1 GAMETES:"
    Write #1, ""
    Write #1, Text1.Text
    Write #1, ""
    Write #1, ""
    Write #1, "PARENT 2 GAMETES:"
    Write #1, ""
    Write #1, Text2.Text
    Write #1, ""
    Write #1, ""
    Write #1, "OFFSPRING:"
    Write #1, ""
    Write #1, "[END OF FILE]"
    
    Close #1
    
End Sub

Private Sub Form_Load()

    Text3.Enabled = False
    
End Sub

Private Sub Timer1_Timer()
    
    Text3.Enabled = True
    
End Sub
