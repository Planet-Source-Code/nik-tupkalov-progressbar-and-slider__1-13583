VERSION 5.00
Object = "{6ADB6618-CDC8-11D4-BC6B-60F14FC10000}#5.0#0"; "PBarS.ocx"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   1680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   1680
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin PBarS.PBarY PBarY2 
      Height          =   240
      Left            =   225
      TabIndex        =   5
      Top             =   585
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   423
      Value           =   50
      Min             =   10
      Max             =   200
      BackColor       =   0
      MouseIcon       =   "frmTest.frx":0000
      MousePointer    =   99
      BackStyle       =   1
      picStep         =   20
      Style           =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4725
      Top             =   495
   End
   Begin VB.CheckBox Check4 
      Caption         =   "EnabledSlider"
      Height          =   285
      Left            =   5400
      TabIndex        =   4
      Top             =   1170
      Value           =   1  'Checked
      Width           =   1725
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Normal/Digital"
      Height          =   285
      Left            =   5400
      TabIndex        =   3
      Top             =   900
      Value           =   1  'Checked
      Width           =   1725
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Timer stop/start"
      Height          =   285
      Left            =   5400
      TabIndex        =   2
      Top             =   630
      Value           =   1  'Checked
      Width           =   1725
   End
   Begin VB.CheckBox Check1 
      Caption         =   "3D"
      Height          =   285
      Left            =   5400
      TabIndex        =   1
      Top             =   360
      Value           =   1  'Checked
      Width           =   1725
   End
   Begin PBarS.PBarY PBarY1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   503
      BackColor       =   0
      picForeColor    =   8421504
      picFillColor    =   65280
      Style           =   1
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim K As Boolean

Private Sub Check1_Click()
InitForm
End Sub

Private Sub Check2_Click()
InitForm

End Sub

Private Sub Check3_Click()
InitForm

End Sub

Private Sub Check4_Click()
InitForm

End Sub

Private Sub Form_Load()
Me.Show
InitForm
K = True
End Sub

Private Sub InitForm()
PBarY1.BackStyle = Check1.Value
Timer1.Enabled = Check2.Value
PBarY1.Style = 1 - Check3.Value
PBarY1.EnabledSlider = Check4.Value
PBarY1.picStep = PBarY2.Value
End Sub

Private Sub PBarY2_ChangeValue(NewValue As Long, OldValue As Long)
PBarY1.picStep = PBarY2.Value
End Sub

Private Sub Timer1_Timer()
With PBarY1
If K Then
    If .Value = .Max Then
    K = False: .Value = .Value - 1
    Else
    .Value = .Value + 1
    End If
Else
    If .Value = .Min Then
    K = True: .Value = .Value + 1
    Else
    .Value = .Value - 1
    End If
End If
End With
End Sub
