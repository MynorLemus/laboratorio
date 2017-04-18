VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   6285
   ClientLeft      =   7575
   ClientTop       =   2670
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   7515
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   7320
      Top             =   0
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5520
      Top             =   4440
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5280
      Top             =   5520
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4080
      Top             =   5400
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4560
      Top             =   4320
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3840
      Top             =   4320
   End
   Begin VB.Image shape5 
      Height          =   1335
      Left            =   5160
      Picture         =   "Form1.frx":0000
      Top             =   0
      Width           =   2205
   End
   Begin VB.Image Shape4 
      Height          =   1335
      Left            =   0
      Picture         =   "Form1.frx":1946
      Top             =   5160
      Width           =   2205
   End
   Begin VB.Image Shape3 
      Height          =   1335
      Left            =   0
      Picture         =   "Form1.frx":35EE
      Top             =   0
      Width           =   2205
   End
   Begin VB.Image Shape2 
      Height          =   1335
      Left            =   4800
      Picture         =   "Form1.frx":4E17
      Top             =   4800
      Width           =   2205
   End
   Begin VB.Image Shape1 
      Height          =   1335
      Left            =   0
      Picture         =   "Form1.frx":67BA
      Top             =   0
      Width           =   2205
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Shape1.Visible = True
Shape2.Visible = False
Shape3.Visible = False
Shape4.Visible = False
shape5.Visible = False

x = 124
x = Shape1.Top
x = x + 50
Shape1.Top = x
x = Shape1.Left
x = x + 50
Shape1.Left = x
If Shape1.Top > 4800 Then
Shape1.Top = 4800
Timer2.Enabled = True
Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()

Shape1.Visible = False
Shape2.Visible = True
Shape3.Visible = False
Shape4.Visible = False

x = Shape2.Top
x = x - 50
Shape2.Top = x
If Shape2.Top < 0 Then
Shape2.Top = 0
x = Shape2.Left
x = x - 50
Shape2.Left = x
End If
If Shape2.Left < 0 Then
Shape2.Left = 0
Timer2.Enabled = False
Timer3.Enabled = True
End If
End Sub

Private Sub Timer3_Timer()
Shape1.Visible = False
Shape2.Visible = False
Shape3.Visible = True
Shape4.Visible = False
x = Shape3.Top
x = x + 50
Shape3.Top = x
If Shape3.Top > 5160 Then
Shape3.Top = 5160
Timer3.Enabled = False
Timer4.Enabled = True
End If
End Sub

Private Sub Timer4_Timer()
Shape1.Visible = False
Shape2.Visible = False
Shape3.Visible = False
Shape4.Visible = True
x = Shape4.Top
x = x - 50
Shape4.Top = x
x = Shape4.Left
x = x + 50
Shape4.Left = x
If Shape4.Left > 5160 Then
Shape4.Left = 5160
Timer5.Enabled = True
Timer4.Enabled = False
End If
End Sub

Private Sub Timer5_Timer()
Shape1.Visible = False
Shape2.Visible = False
Shape3.Visible = False
Shape4.Visible = False
shape5.Visible = True
x = shape5.Left
x = x - 50
shape5.Left = x
If shape5.Left < 0 Then
shape5.Left = 0
Timer5.Enabled = False
Timer6.Enabled = True
End If
End Sub

Private Sub Timer6_Timer()
y = 5
If y > 0 Then
Shape1.Left = 0
Shape1.Top = 0
Shape2.Left = 4800
Shape2.Top = 4800
Shape3.Left = 0
Shape3.Top = 0
Shape4.Left = 0
Shape4.Top = 5160
shape5.Left = 5160
shape5.Top = 0
Shape1.Visible = False
Shape2.Visible = False
Shape3.Visible = False
Shape4.Visible = False
shape5.Visible = False

Timer6.Enabled = False
Timer1.Enabled = True
End If

End Sub
