VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H8000000D&
   Caption         =   "Form3"
   ClientHeight    =   4065
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8940
   LinkTopic       =   "Form3"
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   4065
   ScaleWidth      =   8940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "Salir"
      Height          =   495
      Left            =   6240
      TabIndex        =   6
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Autos de Lujo"
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Autos Estandar"
      Height          =   495
      Left            =   6240
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Autos Semi Nuevos"
      Height          =   495
      Left            =   6240
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clientes"
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Empleados"
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   2895
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form3.Hide
Form1.Show
End Sub

Private Sub Command2_Click()
Form3.Hide
Form2.Show
End Sub

Private Sub Command3_Click()
Form3.Hide
Form4.Show
End Sub

Private Sub Command4_Click()
Form3.Hide
Form5.Show
End Sub

Private Sub Command5_Click()
Form3.Hide
Form6.Show
End Sub

Private Sub Command6_Click()
Form3.Hide
Form7.Show
End Sub

Private Sub Command8_Click()
End
End Sub
