VERSION 5.00
Begin VB.Form hotel 
   Caption         =   "Blue Bezarts Booking"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   9315
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7680
      TabIndex        =   4
      Top             =   2640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cnclcmnd 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   3
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton nxtcmnd 
      Caption         =   "Next"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   2
      Top             =   4680
      Width           =   1335
   End
   Begin VB.ComboBox lctncmbo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "booking.frx":0000
      Left            =   2160
      List            =   "booking.frx":0002
      TabIndex        =   1
      Text            =   "Select Your Location"
      Top             =   2640
      Width           =   4935
   End
   Begin VB.Label htllbl 
      Alignment       =   2  'Center
      Caption         =   "Enter Location"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   0
      Top             =   840
      Width           =   4335
   End
End
Attribute VB_Name = "hotel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs2 As New ADODB.Recordset
Public Function connect_db() As Integer
con.Open ("log")
rs2.Open "select * from hot", con, adOpenDynamic, adLockOptimistic, adCmdText
End Function
Public Function close_db() As Integer
con.Close
End Function
Private Sub cnclcmnd_Click()
Dim x As Integer
x = MsgBox("Sure to Exit", vbQuestion + vbYesNo, "Exit")
If x = vbYes Then
 End
End If
End Sub
Private Sub Form_Load()
i = connect_db
While rs2.EOF <> True
lctncmbo.AddItem (rs2(0))
rs2.MoveNext
Wend
close_db
End Sub
Private Sub lctncmbo_click()
nxtcmnd.Enabled = True
End Sub
Private Sub nxtcmnd_Click()
i = lctncmbo.ListIndex
Text1.Text = Text1.Text + lctncmbo.List(i)
booking1.lctntxt.Text = Text1.Text
printpage.lctntxt1.Text = Text1.Text
booking1.Show
Unload Me
End Sub

