VERSION 5.00
Begin VB.Form customer 
   Caption         =   "Blue Bezarts Booking"
   ClientHeight    =   9510
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   ScaleHeight     =   9510
   ScaleWidth      =   11265
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox pantxt 
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      Top             =   8520
      Width           =   3615
   End
   Begin VB.CommandButton backcmnd2 
      Caption         =   "Back"
      Height          =   375
      Left            =   13200
      TabIndex        =   20
      Top             =   10200
      Width           =   1575
   End
   Begin VB.CommandButton nxtcmnd1 
      Caption         =   "Next"
      Enabled         =   0   'False
      Height          =   375
      Left            =   11400
      TabIndex        =   19
      Top             =   10200
      Width           =   1575
   End
   Begin VB.CommandButton bookcmnd 
      Caption         =   "Book"
      Height          =   375
      Left            =   9600
      TabIndex        =   18
      Top             =   10200
      Width           =   1575
   End
   Begin VB.TextBox lnametxt 
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   2400
      Width           =   3615
   End
   Begin VB.TextBox mnametxt 
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Frame sexframe 
      Caption         =   "SEX"
      Height          =   1815
      Left            =   7080
      TabIndex        =   14
      Top             =   240
      Width           =   2655
      Begin VB.OptionButton fmaleopt 
         Caption         =   "Female"
         Height          =   375
         Left            =   600
         TabIndex        =   15
         Top             =   1080
         Width           =   855
      End
      Begin VB.OptionButton maleopt 
         Caption         =   "Male"
         Height          =   375
         Left            =   600
         TabIndex        =   9
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.TextBox mobiletxt 
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   7440
      Width           =   3615
   End
   Begin VB.TextBox statetxt 
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   6360
      Width           =   3615
   End
   Begin VB.TextBox citytxt 
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   5280
      Width           =   3615
   End
   Begin VB.TextBox addtxt 
      Height          =   1215
      Left            =   1920
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   3480
      Width           =   3615
   End
   Begin VB.TextBox fnametxt 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "a"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   0
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label panlbl 
      Alignment       =   2  'Center
      Caption         =   "Pan Card no."
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
      Left            =   240
      TabIndex        =   21
      Top             =   8640
      Width           =   1455
   End
   Begin VB.Label lnamelbl 
      Alignment       =   2  'Center
      Caption         =   "Last Name"
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
      Left            =   240
      TabIndex        =   17
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label mnamelbl 
      Alignment       =   2  'Center
      Caption         =   "Middle Name"
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
      Left            =   240
      TabIndex        =   16
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label mobilelbl 
      Alignment       =   2  'Center
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
      Left            =   240
      TabIndex        =   13
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Label statelbl 
      Alignment       =   2  'Center
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
      Left            =   240
      TabIndex        =   12
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label citylbl 
      Alignment       =   2  'Center
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
      Left            =   240
      TabIndex        =   11
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label addlbl 
      Alignment       =   2  'Center
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
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label fnamelbl 
      Alignment       =   2  'Center
      Caption         =   "First Name"
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
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim x As String
Dim y As Integer
Public Function connect_db() As Integer
con.Open ("log")
rs1.Open "select * from tab_login", con, adOpenDynamic, adLockOptimistic, adCmdText
rs2.Open "select * from hot", con, adOpenDynamic, adLockOptimistic, adCmdText
rs3.Open "select * from room", con, adOpenDynamic, adLockOptimistic, adCmdText
End Function
Public Function close_db() As Integer
con.Close
End Function
Private Sub backcmnd2_Click()
booking2.Show
Unload Me
End Sub
Private Sub bookcmnd_Click()
nxtcmnd1.Enabled = True
i = connect_db()
rs4.Open "select * from room where hotelname='" & ht & "'", con, adOpenDynamic, adLockOptimistic, adCmdText
If rs4(1) = "nil" Then
    Set cmd.ActiveConnection = con
    cmd.CommandText = "update room set room1='Booked' where hotelname='" & ht & "'"
    cmd.Execute
    MsgBox "Booked"
    Exit Sub
ElseIf rs4(2) = "nil" Then
    Set cmd.ActiveConnection = con
    cmd.CommandText = "update room set room2='Booked' where hotelname='" & ht & "'"
    cmd.Execute
    MsgBox "Booked"
    Exit Sub
ElseIf rs4(3) = "nil" Then
    Set cmd.ActiveConnection = con
    cmd.CommandText = "update room set room3='Booked' where hotelname='" & ht & "'"
    cmd.Execute
    MsgBox "Booked"
    Exit Sub
Else
    MsgBox "Not Available"
End If
i = close_db()
End Sub
Private Sub fmaleopt_Click()
If fmaleopt.Value = True Then
 printpage.opttxt1 = "Female"
 End If
End Sub
Private Sub fnametxt_KeyPress(KeyAscii As Integer)
Select Case Chr(KeyAscii)
Case "A" To "Z"
Case "a" To "z"
Case vbBack
Case Else
KeyAscii = 0
End Select
End Sub
Private Sub lnametxt_KeyPress(KeyAscii As Integer)
Select Case Chr(KeyAscii)
Case "A" To "Z"
Case "a" To "z"
Case vbBack
Case Else
KeyAscii = 0
End Select
End Sub
Private Sub maleopt_Click()
If maleopt.Value = True Then
   printpage.opttxt1 = "Male"
End If
End Sub
Private Sub mnametxt_KeyPress(KeyAscii As Integer)
Select Case Chr(KeyAscii)
Case "A" To "Z"
Case "a" To "z"
Case vbBack
Case Else
KeyAscii = 0
End Select
End Sub

Private Sub mobiletxt_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case vbKey0 To vbKey9
Case vbKeyBack
Case Else
    KeyAscii = 0
End Select
End Sub
Private Sub nxtcmnd1_Click()
x = mobiletxt.Text
y = Len(x)
printpage.htltxt1.Text = ht
printpage.nametxt1.Text = fnametxt.Text & " " & mnametxt.Text & " " & lnametxt.Text
printpage.addtxt1.Text = addtxt.Text
printpage.citytxt1.Text = citytxt.Text
printpage.statetxt1.Text = statetxt.Text
printpage.mobiletxt1.Text = mobiletxt.Text
printpage.pantxt1 = pantxt.Text
If fnametxt.Text = "" And mnametxt.Text = "" Or lnametxt.Text = "" Or citytxt.Text = "" Or addtxt.Text = "" Or statetxt.Text = "" Or mobiletxt.Text = "" Or pantxt.Text = "" Then
MsgBox ("All Field should be filled")
ElseIf y <> 10 Then
 MsgBox ("Improper Mobile no.")
Else
printpage.Show
Unload Me
End If
End Sub

