VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Label2"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents Tim As MMTimer
Attribute Tim.VB_VarHelpID = -1
Private mCounter As Long
Private dt As Single

Private Sub Form_Load()
    Set Tim = New MMTimer
    Tim.Enabled = False
    Tim.Interval = 1
    Label1.Caption = mCounter
End Sub

Private Sub Command1_Click()
    dt = Timer
    Label2.Caption = dt
    Tim.Enabled = Not Tim.Enabled
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Tim.Enabled = False
End Sub

Private Sub Tim_Timer()
    Label1.Caption = mCounter
    Label3.Caption = Timer - dt
    mCounter = mCounter + 1
End Sub
