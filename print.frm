VERSION 5.00
Begin VB.Form printpage 
   Caption         =   "Blue Bezarts Booking"
   ClientHeight    =   7920
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13260
   LinkTopic       =   "Form1"
   Picture         =   "print.frx":0000
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox prcetxt1 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   10920
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   6360
      Width           =   2055
   End
   Begin VB.TextBox rmvaluetxt1 
      Height          =   285
      Left            =   7200
      TabIndex        =   29
      Top             =   4440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton prcecmnd 
      Caption         =   "Click To Get Price"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10920
      TabIndex        =   28
      Top             =   7200
      Width           =   2055
   End
   Begin VB.TextBox noofroomtxt1 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   10920
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   5040
      Width           =   2055
   End
   Begin VB.TextBox pantxt1 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   5280
      Width           =   3615
   End
   Begin VB.TextBox opttxt1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   6840
      TabIndex        =   22
      Top             =   240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cnclcmnd1 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12000
      TabIndex        =   21
      Top             =   9480
      Width           =   2055
   End
   Begin VB.CommandButton printcmnd 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   20
      Top             =   9480
      Width           =   2055
   End
   Begin VB.TextBox ppltxt1 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   10920
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox durtntxt1 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   10920
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox roomtxt1 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   10920
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   4080
      Width           =   2055
   End
   Begin VB.TextBox htltxt1 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   10920
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox lctntxt1 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   10920
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   240
      Width           =   2055
   End
   Begin VB.TextBox nametxt1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   240
      Width           =   3615
   End
   Begin VB.TextBox addtxt1 
      Appearance      =   0  'Flat
      Height          =   1335
      Left            =   2640
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   960
      Width           =   3615
   End
   Begin VB.TextBox citytxt1 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   2640
      Width           =   3615
   End
   Begin VB.TextBox statetxt1 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   3600
      Width           =   3615
   End
   Begin VB.TextBox mobiletxt1 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   4440
      Width           =   3615
   End
   Begin VB.Label noroomlbl1 
      Caption         =   "No.Of Rooms"
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
      Left            =   8880
      TabIndex        =   26
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label prcelbl1 
      Caption         =   "Total Price"
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
      Left            =   8880
      TabIndex        =   25
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Label panlbl1 
      Caption         =   "Pan Cardno."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   23
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label ppllbl1 
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
      Left            =   8880
      TabIndex        =   16
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label durtnlbl1 
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
      Left            =   8880
      TabIndex        =   15
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label roomlbl1 
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
      Left            =   8880
      TabIndex        =   14
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label htllbl1 
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
      Left            =   8880
      TabIndex        =   12
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lctnlbl1 
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
      Left            =   8880
      TabIndex        =   10
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label namelbl1 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label addlbl1 
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label citylbl1 
      Caption         =   "City"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label statelbl1 
      Caption         =   "State"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label mobilelbl1 
      Caption         =   "Mobile no."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   4560
      Width           =   1335
   End
End
Attribute VB_Name = "printpage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cnclcmnd1_Click()
Dim z As Integer
z = MsgBox("Sure to Exit", vbQuestion + vbYesNo, "Exit")
If z = vbYes Then
 End
End If
End Sub

Private Sub prcecmnd_Click()
p = Val(ppltxt1.Text)
d = Val(durtntxt1.Text)
r = Val(noofroomtxt1.Text)
rt = Val(rmvaluetxt1.Text)
tt = ((p * 400 * d) + (d * rt * r))
prcetxt1.Text = tt
End Sub
Private Sub printcmnd_Click()
printpage.PrintForm
End Sub
