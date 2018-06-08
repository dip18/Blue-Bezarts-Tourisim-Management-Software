VERSION 5.00
Begin VB.Form booking1 
   Caption         =   "Blue Bezarts Booking"
   ClientHeight    =   3240
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton backcmnd 
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
      Left            =   7440
      TabIndex        =   6
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   7800
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton avlcmnd 
      Caption         =   "Check Availability"
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
      Height          =   615
      Left            =   3840
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
   Begin VB.ComboBox htlcmbo 
      Height          =   315
      Left            =   2760
      TabIndex        =   3
      Text            =   "Select Your Hotel"
      Top             =   1440
      Width           =   4935
   End
   Begin VB.TextBox lctntxt 
      Height          =   405
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label htllbl 
      Alignment       =   2  'Center
      Caption         =   "Hotel"
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
      Left            =   600
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lctnlbl 
      Alignment       =   2  'Center
      Caption         =   "Location"
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
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "booking1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim flag As Boolean
Public Function connect_db() As Integer
con.Open ("log")
rs2.Open "select * from hot", con, adOpenDynamic, adLockOptimistic, adCmdText
rs3.Open "select * from room", con, adOpenDynamic, adLockOptimistic, adCmdText
End Function
Public Function close_db() As Integer
con.Close
End Function
Private Sub avlcmnd_Click()
i = htlcmbo.ListIndex
Text2.Text = Text2.Text + htlcmbo.List(i)
ht = Text2.Text
printpage.htltxt1.Text = Text2.Text
j = connect_db()
Do While Not rs3.EOF
    If rs3(0) = ht Then
        If rs3(1) = "nil" Or rs3(2) = "nil" Or rs3(3) = "nil" Then
            MsgBox "Available"
            flag = True
            booking2.Show
            Unload Me
            i = close_db()
            Exit Sub
        End If
    End If
rs3.MoveNext
Loop
If flag = False Then
   MsgBox ("Not Available")
End If
i = close_db()
End Sub
Private Sub backcmnd_Click()
hotel.Show
Unload Me
End Sub
Private Sub Form_Load()
i = connect_db()
While rs2.EOF <> True
    If rs2(0) = hotel.Text1.Text Then
        htlcmbo.AddItem (rs2(1))
        htlcmbo.AddItem (rs2(2))
        htlcmbo.AddItem (rs2(3))
    End If
rs2.MoveNext
Wend
close_db
End Sub
Private Sub htlcmbo_Click()
avlcmnd.Enabled = True
End Sub
