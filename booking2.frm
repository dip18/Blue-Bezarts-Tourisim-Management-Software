VERSION 5.00
Begin VB.Form booking2 
   Caption         =   "Blue Bezarts Booking"
   ClientHeight    =   6570
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   10875
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox text6 
      Height          =   285
      Left            =   9840
      TabIndex        =   14
      Top             =   3000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   7920
      TabIndex        =   13
      Top             =   4080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox noroomcmbo 
      Height          =   315
      ItemData        =   "booking2.frx":0000
      Left            =   5400
      List            =   "booking2.frx":0013
      TabIndex        =   12
      Text            =   "Select"
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   7920
      TabIndex        =   10
      Top             =   3000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   7920
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   7920
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox roomcmbo 
      Height          =   315
      ItemData        =   "booking2.frx":0026
      Left            =   5400
      List            =   "booking2.frx":0028
      TabIndex        =   5
      Text            =   "Select"
      Top             =   3000
      Width           =   1695
   End
   Begin VB.ComboBox pplcmbo 
      Height          =   315
      ItemData        =   "booking2.frx":002A
      Left            =   5400
      List            =   "booking2.frx":0040
      TabIndex        =   4
      Text            =   "Select"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton cstmrcmnd 
      Caption         =   "Customer Details"
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
      Height          =   735
      Left            =   4800
      TabIndex        =   3
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton backcmnd1 
      Caption         =   "Back"
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
      Left            =   9000
      TabIndex        =   2
      Top             =   5880
      Width           =   1455
   End
   Begin VB.ComboBox durtncmbo 
      Height          =   315
      ItemData        =   "booking2.frx":0057
      Left            =   5400
      List            =   "booking2.frx":0067
      TabIndex        =   1
      Text            =   "Select"
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label foodlbl 
      Alignment       =   2  'Center
      Caption         =   "Food Cost Rs. 400 Per Day Per Person"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label noroomlbl 
      Alignment       =   2  'Center
      Caption         =   "No.Of Room"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   11
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label roomlbl 
      Alignment       =   2  'Center
      Caption         =   "Room Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label durtnlbl 
      Alignment       =   2  'Center
      Caption         =   "Duration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label ppllbl 
      Alignment       =   2  'Center
      Caption         =   "No of People"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   1920
      Width           =   1695
   End
End
Attribute VB_Name = "booking2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs5 As New ADODB.Recordset
Public Function connect_db() As Integer
con.Open ("log")
rs2.Open "select * from hot", con, adOpenDynamic, adLockOptimistic, adCmdText
rs3.Open "select * from room", con, adOpenDynamic, adLockOptimistic, adCmdText
rs5.Open "select * from roomprice", con, adOpenDynamic, adLockOptimistic, adCmdText
End Function
Public Function close_db() As Integer
con.Close
End Function
Private Sub backcmnd1_Click()
booking1.Show
Unload Me
End Sub
Private Sub cstmrcmnd_Click()
k = roomcmbo.ListIndex
Text5.Text = Text5.Text + roomcmbo.List(k)
rm = Text5.Text
printpage.roomtxt1.Text = Text5.Text
a = connect_db()
Do While Not rs5.EOF
 If rs5(0) = rm Then
     text6.Text = rs5(1)
     End If
rs5.MoveNext
Loop
close_db
printpage.rmvaluetxt1.Text = text6.Text
i = pplcmbo.ListIndex
j = durtncmbo.ListIndex
l = noroomcmbo.ListIndex
Text3.Text = Text3.Text + pplcmbo.List(i)
Text4.Text = Text4.Text + durtncmbo.List(j)
Text7.Text = Text7.Text + noroomcmbo.List(l)
printpage.ppltxt1.Text = Text3.Text
printpage.durtntxt1.Text = Text4.Text
printpage.noofroomtxt1.Text = Text7.Text
customer.Show
Unload Me
End Sub
Private Sub Form_Load()
i = connect_db
While rs5.EOF <> True
roomcmbo.AddItem (rs5(0))
rs5.MoveNext
Wend
close_db
End Sub
Private Sub pplcmbo_Click()
cstmrcmnd.Enabled = True
End Sub

