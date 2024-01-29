VERSION 5.00
Begin VB.Form Login 
   Caption         =   "LOGIN"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17205
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9105
   ScaleWidth      =   17205
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      ForeColor       =   &H00400000&
      Height          =   9135
      Left            =   0
      Picture         =   "login.frx":0000
      ScaleHeight     =   9075
      ScaleWidth      =   17115
      TabIndex        =   0
      Top             =   0
      Width           =   17175
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H00FFFFFF&
         Height          =   7455
         Left            =   4080
         TabIndex        =   1
         Top             =   960
         Width           =   8535
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            ForeColor       =   &H80000008&
            Height          =   6135
            Left            =   480
            TabIndex        =   2
            Top             =   360
            Width           =   8655
            Begin VB.CommandButton Command1 
               Caption         =   "LOGIN"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   1440
               Picture         =   "login.frx":12955
               Style           =   1  'Graphical
               TabIndex        =   9
               Top             =   5280
               Width           =   4935
            End
            Begin VB.TextBox Text1 
               BackColor       =   &H8000000E&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00E0E0E0&
               Height          =   255
               Left            =   1320
               TabIndex        =   4
               Text            =   "TYPE YOUR USERNAME"
               Top             =   2160
               Width           =   4455
            End
            Begin VB.TextBox Text2 
               BackColor       =   &H8000000E&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00E0E0E0&
               Height          =   255
               Left            =   1320
               TabIndex        =   3
               Text            =   "TYPE YOUR PASSWORD"
               Top             =   3840
               Width           =   4335
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Login"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   21.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   735
               Left            =   2760
               TabIndex        =   8
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Username"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   375
               Left            =   600
               TabIndex        =   7
               Top             =   1680
               Width           =   1335
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00000000&
               X1              =   600
               X2              =   6960
               Y1              =   2520
               Y2              =   2520
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Password"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   375
               Left            =   600
               TabIndex        =   6
               Top             =   3240
               Width           =   1575
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00000000&
               X1              =   600
               X2              =   7080
               Y1              =   4200
               Y2              =   4200
            End
            Begin VB.Image Image3 
               Height          =   420
               Left            =   720
               Picture         =   "login.frx":131C1
               Top             =   2040
               Width           =   390
            End
            Begin VB.Image Image4 
               Height          =   585
               Left            =   600
               Picture         =   "login.frx":13605
               Top             =   3600
               Width           =   555
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Forget Password"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   495
               Left            =   5400
               TabIndex        =   5
               Top             =   4320
               Width           =   1695
            End
            Begin VB.Image Image5 
               Height          =   435
               Left            =   6720
               Picture         =   "login.frx":139E4
               Top             =   3720
               Width           =   585
            End
            Begin VB.Image Image6 
               Height          =   390
               Left            =   6720
               Picture         =   "login.frx":13F0F
               Top             =   3720
               Width           =   555
            End
         End
         Begin VB.Image Image2 
            Height          =   5805
            Left            =   8760
            Picture         =   "login.frx":14417
            Top             =   600
            Width           =   7515
         End
      End
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim u As String
Dim p As String
Dim A As Byte

Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "Enter Username and Password"
ElseIf u = Text1.Text And p = Text2.Text Then
Unload Me
SFMG.Show
Else
MsgBox "Wrong Password"
Text2.Text = ""
Text2.SetFocus
End If
End Sub

Private Sub Form_Load()
Image5.Visible = False
Image6.Visible = False
End Sub

Private Sub Image5_Click()
Image6.Visible = True
Image5.Visible = False
If (Text2.Text = "Type your password") Then
Text2.PasswordChar = ""
Else
Text2.PasswordChar = "*"
End If
'Text2.SetFocus
End Sub

Private Sub Image6_Click()
Image5.Visible = True
Image6.Visible = False
Text2.PasswordChar = ""
'Text2.SetFocus
End Sub





Private Sub Label4_Click()
Id = (InputBox("Enter Your School ID"))
If (Id = "1") Then
Y = (InputBox("Enter Your School Registration Number"))
If (Y = "2021") Then
MsgBox "Login Successfully"
SFMG.Show
Unload Me
Else
MsgBox "Wrong Details contact to technical team"
Text1.SetFocus
End If
End If

End Sub

Private Sub Label5_Click()

End Sub






Private Sub Text1_CLICK()
Text1.Text = ""
Text1.ForeColor = vbBlack
If Text2.Text = "" Then
Text2.Text = "Type Your Password"
Text2.PasswordChar = ""
Text2.ForeColor = QBColor(8)
End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Image5.Visible = True
Image6.Visible = True
On Error GoTo xxxx
Text1.Text = UCase(Text1.Text)
Abc
sql = "select * from login where Uname='" + Text1.Text + "' "
Set r = c.Execute(sql)
u = r.Fields("Uname")
p = r.Fields("Passw")
Text2.Text = ""
Text2.SetFocus
Text2.ForeColor = vbBlack
Exit Sub
xxxx:
MsgBox "WRONG ID"
Text1.Text = ""
Text1.SetFocus


End If


End Sub

Private Sub Text1_LostFocus()
Text1.Text = UCase(Text1.Text)
End Sub


Private Sub Text2_Change()
If (Text2.Text = "" Or Text2.Text = "Type your password") Then
Text2.Text = Text2.Text
Else
Text2.PasswordChar = "*"
End If
End Sub

Private Sub Text2_Click()
Text2.Text = ""
Text2.ForeColor = vbBlack
If Text1.Text = "" Then
Text1.Text = "Type Your Username"
Text1.ForeColor = QBColor(8)
End If
End Sub

Private Sub Text2_GotFocus()
If (Text2.Text = "Type your password") Then
Text2.Text = ""
Else
Text2.Text = Text2.Text
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.Text = UCase(Text2.Text)

Command1.SetFocus
End If
End Sub

