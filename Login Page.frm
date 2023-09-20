VERSION 5.00
Begin VB.Form Login1 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Login page"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17415
   FillColor       =   &H00FFFFC0&
   BeginProperty Font 
      Name            =   "Segoe Print"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFC0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9255
   ScaleLeft       =   120
   ScaleMode       =   0  'User
   ScaleTop        =   465
   ScaleWidth      =   16350
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Show Password"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   12360
      TabIndex        =   6
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Forget Password"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   470
      Left            =   10737
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3661
      Width           =   2189
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   698
      Left            =   7413
      MaskColor       =   &H00FFFFC0&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   2189
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   698
      Left            =   8400
      TabIndex        =   3
      Top             =   2760
      Width           =   3723
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   727
      Left            =   8436
      TabIndex        =   2
      Top             =   1920
      Width           =   3723
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6480
      TabIndex        =   7
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   5520
      TabIndex        =   1
      Top             =   2760
      Width           =   3600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   5490
      TabIndex        =   0
      Top             =   1830
      Width           =   3600
   End
End
Attribute VB_Name = "Login1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim C As String, S As String

Private Sub Check1_Click()
If Check1.Value = 1 Then
Text2.PasswordChar = ""
Text2.Text = Text2.Text
Check1.Caption = "Hide Password"
Text1.SetFocus

Else
Text2.PasswordChar = "*"
Check1.Caption = "Show Password"
Text2.SetFocus
End If

End Sub


Private Sub Form_Load()
Text1.Text = "PRJ2333E"
End Sub

Private Sub Text2_Change()
Text2.PasswordChar = "*"
End Sub


Private Sub Command1_Click()
If Text1.Text = "" And Text2.Text <> "" Then
MsgBox "Enter the Username"
Text1.SetFocus
End If

If Text1.Text <> "" And Text2.Text = "" Then
MsgBox "Enter the Password"
Text2.SetFocus
End If

If Text1.Text = "" And Text2.Text = "" Then
MsgBox "Enter username and password"
End If
If Text1.Text = "PRJ2333E" And Text2.Text = "PRJ2333E" Then
MsgBox "Login Successfully"
SFMG.Show
Unload Me

Else
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
Command2.Caption = "Show Password"
MsgBox "Wrong Username or Password"
End If



End Sub

Private Sub Command2_Click()


C = (InputBox("What is your school Id"))

If (C = "PR001") Then
S = (InputBox("What is your unique number"))
If (S = "P12345678") Then
MsgBox "Login Successfully"
End If
SFMG.Show
Unload Me

Else
MsgBox " Wrong Password Contact to Techinical Team"
Text1.SetFocus
End If
End Sub



Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.SetFocus
End If

End Sub

Private Sub Text1_LostFocus()
Text1.Text = UCase(Text1.Text)
End Sub



Private Sub Text2_LostFocus()
Text2.Text = UCase(Text2.Text)

End Sub
