VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000008&
   Caption         =   "Stopwatch-Mosihur"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9735
   FillColor       =   &H0080FFFF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H8000000B&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   9735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4920
      TabIndex        =   4
      Top             =   4440
      Width           =   2655
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1200
      MaskColor       =   &H00FF00FF&
      TabIndex        =   3
      Top             =   4440
      Width           =   2535
   End
   Begin VB.TextBox txtMiniSecond 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   975
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "00"
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox txtSecond 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   855
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "00"
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox txtMinute 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Text            =   "00"
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Timer Timer3 
      Interval        =   10
      Left            =   6480
      Top             =   1800
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   3480
      Top             =   1920
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   840
      Top             =   1800
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mosihur--Stopwatch"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   5
      Top             =   240
      Width           =   5535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text2_Change()

End Sub

Private Sub Image1_Click()

End Sub

Private Sub Picture1_Click()

End Sub

Private Sub cmdReset_Click()
cmdStart.Caption = "Start"
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
txtMiniSecond.Text = "00"
txtSecond.Text = "00"
txtMinute.Text = "00"

End Sub

Private Sub cmdStart_Click()
If cmdStart.Caption = "Start" Then
cmdStart.Caption = "Stop"
Else
cmdStart.Caption = "Start"
End If
If cmdStart.Caption = "Stop" Then
Timer3.Enabled = True
Timer2.Enabled = True
Timer1.Enabled = True
Else
Timer3.Enabled = False
Timer2.Enabled = False
Timer1.Enabled = False
End If
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Timer1_Timer()
If txtMinute.Text < 59 Then
txtMinute.Text = Format(txtMinute.Text + 1, "00")
Else
txtMinute.Text = Format(0, "00")
End If
End Sub

Private Sub Timer2_Timer()
If txtSecond.Text < 59 Then
txtSecond.Text = Format(txtSecond.Text + 1, "00")
Else
txtSecond.Text = Format(0, "00")
End If
End Sub

Private Sub Timer3_Timer()
If txtMiniSecond.Text < 99 Then
txtMiniSecond.Text = Format(txtMiniSecond.Text + 1, "00")
Else
txtMiniSecond.Text = Format(0, "00")
End If
End Sub

Private Sub txtMiniSecond_Change()

End Sub
