VERSION 5.00
Begin VB.Form login 
   Caption         =   "Blue Bezarts Login"
   ClientHeight    =   7170
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   10275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton rstcmnd 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8160
      TabIndex        =   3
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton lgncmnd 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   2
      Top             =   4920
      Width           =   1935
   End
   Begin VB.TextBox pwtxt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   5040
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "Enter Admin Password"
      Top             =   3840
      Width           =   3375
   End
   Begin VB.TextBox idtxt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      ToolTipText     =   "Enter Admin ID"
      Top             =   2880
      Width           =   3375
   End
   Begin VB.Image Image2 
      Height          =   1920
      Left            =   4080
      Picture         =   "Blue Blazerts Project.frx":0000
      Top             =   240
      Width           =   2040
   End
   Begin VB.Label lblusr 
      Alignment       =   2  'Center
      Caption         =   "Admin ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label lblpw 
      Alignment       =   2  'Center
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   7545
      Left            =   -120
      Picture         =   "Blue Blazerts Project.frx":291F
      Top             =   0
      Width           =   12810
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs1 As New ADODB.Recordset
Dim flag As Boolean
Public Function connect_db() As Integer
con.Open ("log")
rs1.Open "select * from tab_login", con, adOpenDynamic, adLockOptimistic, adCmdText
End Function
Public Function close_db() As Integer
con.Close
End Function
Private Sub lgncmnd_Click()
connect_db
Do While Not rs1.EOF
If idtxt.Text = rs1(0) And pwtxt.Text = rs1(1) Then
flag = True
hotel.Show
Unload Me
End If
rs1.MoveNext
Loop
If flag = False Then
MsgBox ("Incorrect Id and Password")
End If
close_db
End Sub
Private Sub rstcmnd_Click()
idtxt.Text = ""
pwtxt.Text = ""
End Sub
