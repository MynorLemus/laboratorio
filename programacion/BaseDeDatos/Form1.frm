VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   Caption         =   "Form1"
   ClientHeight    =   4575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   4575
   ScaleMode       =   0  'User
   ScaleWidth      =   9105
   Begin VB.TextBox Text5 
      DataField       =   "Salario"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   7440
      TabIndex        =   16
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      DataField       =   "Fecha_Nacimiento"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   4800
      TabIndex        =   15
      Top             =   1560
      Width           =   2655
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   4440
      TabIndex        =   12
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Menu"
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Siguente"
      Height          =   375
      Left            =   6240
      TabIndex        =   8
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Anterior"
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\stricker\Desktop\BaseDeDatos\Toyota.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Empleados"
      Top             =   7320
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.TextBox Text3 
      DataField       =   "Puesto"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      DataField       =   "Nombre"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      DataField       =   "Codigo"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Salario"
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
      Left            =   7440
      TabIndex        =   14
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Fecha de Nacimiento"
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
      Left            =   4920
      TabIndex        =   13
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Puesto"
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
      Left            =   3360
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Nombre"
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
      Left            =   1680
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Codigo"
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
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Empleados"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.MovePrevious
If Data1.Recordset.BOF Then
Data1.Recordset.MoveLast
End If
End Sub

Private Sub Command2_Click()
Data1.Recordset.MoveNext
If Data1.Recordset.EOF Then
Data1.Recordset.MoveFirst
End If
End Sub

Private Sub Command3_Click()
Form1.Hide
Form3.Show
End Sub

Private Sub Command4_Click()
Data1.Recordset.AddNew
End Sub

Private Sub Command5_Click()
Data1.Recordset.Update
Data1.Recordset.MoveNext
End Sub

Private Sub Command6_Click()
Data1.Recordset.Delete
Data1.Recordset.MovePrevious

End Sub

