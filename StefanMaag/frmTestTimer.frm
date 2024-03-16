VERSION 5.00
Begin VB.Form frmTestTimer 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdStop2 
      Caption         =   "Stop"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdStart2 
      Caption         =   "Start"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      Caption         =   "Label2"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   1935
   End
End
Attribute VB_Name = "frmTestTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents Timer1 As cTimer
Attribute Timer1.VB_VarHelpID = -1
Dim WithEvents Timer2 As cTimer
Attribute Timer2.VB_VarHelpID = -1

Private Sub cmdStart_Click()
   Timer1.CreateTimer 500
End Sub

Private Sub cmdStart2_Click()
   Timer2.CreateTimer 1000
End Sub

Private Sub cmdStop_Click()
   Timer1.DestroyTimer
End Sub

Private Sub cmdStop2_Click()
   Timer2.DestroyTimer
End Sub

Private Sub Form_Load()
   Set Timer1 = New cTimer
   Set Timer2 = New cTimer
End Sub

Private Sub Timer1_Timer(ByVal ThisTime As Long)
   Label1.Caption = ThisTime
End Sub

Private Sub Timer2_Timer(ByVal ThisTime As Long)
   Label2.Caption = ThisTime

End Sub
