VERSION 5.00
Begin VB.Form Signup 
   Caption         =   "Form1"
   ClientHeight    =   9660
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17850
   LinkTopic       =   "Form1"
   ScaleHeight     =   9660
   ScaleWidth      =   17850
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   7095
      Left            =   4200
      TabIndex        =   0
      Top             =   1800
      Width           =   9375
      Begin VB.CommandButton Command1 
         Caption         =   "Sign Up"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         Picture         =   "Signup.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5520
         Width           =   4335
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Text            =   "Enter User id "
         Top             =   1800
         Width           =   5415
      End
      Begin VB.TextBox Text2 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Text            =   "Enter Username "
         Top             =   3000
         Width           =   5535
      End
      Begin VB.TextBox Text3 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   2400
         TabIndex        =   1
         Text            =   "Enter Password"
         Top             =   4200
         Width           =   5655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Sign"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   1320
         X2              =   8280
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line1 
         X1              =   1320
         X2              =   8280
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line3 
         X1              =   1440
         X2              =   8520
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "User ID"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   2760
         Width           =   735
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
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   1320
         Picture         =   "Signup.frx":086C
         Top             =   1800
         Width           =   510
      End
      Begin VB.Image Image3 
         Height          =   420
         Left            =   1440
         Picture         =   "Signup.frx":0CFC
         Top             =   3000
         Width           =   390
      End
      Begin VB.Image Image4 
         Height          =   585
         Left            =   1440
         Picture         =   "Signup.frx":1140
         Top             =   4080
         Width           =   555
      End
      Begin VB.Image Image5 
         Height          =   390
         Left            =   8040
         Picture         =   "Signup.frx":151F
         Top             =   4200
         Width           =   555
      End
      Begin VB.Image Image6 
         Height          =   435
         Left            =   8040
         Picture         =   "Signup.frx":1A27
         Top             =   4200
         Width           =   585
      End
   End
   Begin VB.Image Sign 
      Height          =   10380
      Left            =   0
      Picture         =   "Signup.frx":1F52
      Top             =   0
      Width           =   18450
   End
End
Attribute VB_Name = "Signup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s As String

Private Sub Form_Load()
Image5.Visible = False
Image6.Visible = False
Text1.TabIndex = 0
Abc
Set r = New ADODB.Recordset
sql = "select count(uid) from login"
Set r = c.Execute(sql)
s = "SFMS"
'Text1.Text = Mid(Text1.Text, 5) + 1
Text1.Text = s & "" & r.Fields(0) + 1
End Sub
Private Sub Text1_CLICK()
If (Text1.Text = "Enter User Id") Then
Text1.Text = ""
'Text1.ForeColor = QBColor(8)
End If
End Sub
Private Sub Text1_GotFocus()
Text1.ForeColor = vbBlack
If (Text1.Text = "Enter User Id") Then
Text1.Text = ""
Else
Text1.Text = Text1.Text
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.ForeColor = vbBlack
Text1.Text = UCase(Text1.Text)
Text2.Text = ""
Text2.SetFocus
End If
End Sub

Private Sub Text1_LostFocus()
If (Text1.Text = "") Then
Text1.Text = "Enter User Id"
Text1.ForeColor = vbBlack
End If
End Sub
Private Sub Text2_Click()
Text2.ForeColor = vbBlack
If (Text2.Text = "Enter Username") Then
Text2.Text = ""
End If
End Sub

Private Sub Text2_GotFocus()
Text2.Text = ""
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If (Text2.Text = "") Then
Text2.SetFocus
Else
Text3.ForeColor = vbBlack
Text2.Text = UCase(Text2.Text)
'Text3.PasswordChar = "*"
Text3.SetFocus
Text3.Text = ""
Image5.Visible = True
Image6.Visible = True
End If
End If
End Sub
Private Sub Text2_LostFocus()
If (Text2.Text = "") Then
MsgBox "Please Firstly Fill Username!!", vbCritical
Text2.Text = "Enter Username"
Text2.SetFocus
End If
End Sub
Private Sub Text3_Change()
If (Text3.Text = "" Or Text3.Text = "Enter Password") Then
Text3.PasswordChar = ""
Text3.Text = Text3.Text
Else
Text3.PasswordChar = "*"
End If
End Sub

Private Sub Text3_Click()
Image5.Visible = True
Image6.Visible = True
Text3.ForeColor = vbBlack
If (Text3.Text = "Enter Password") Then
Text3.Text = ""
End If
End Sub

Private Sub Text3_GotFocus()
Text3.ForeColor = vbBlack
If (Text3.Text = "Enter Password") Then
Text3.Text = ""
Else
Text3.Text = Text3.Text
End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text3.Text = UCase(Text3.Text)
Command1.SetFocus
End If
End Sub

Private Sub Text3_LostFocus()
If (Len(Text3.Text) < 8) Then
MsgBox "Password must be of 8 characters"
End If
'Text3.SetFocus
'Image5.Visible = False
'Image6.Visible = False
End Sub
Private Sub Command1_Click()
If (Text2.Text = "" Or Text2.Text = "Enter Username" Or Text3.Text = "Enter Password") Then
MsgBox "Please fill all the fields!!", vbCritical
Text2.SetFocus
Text3.Text = "Enter Password"
Else
sql = "Insert into LOGIN values('" + Text1.Text + "','" + Text2.Text + "','" + Text3.Text + "')"
Set r = c.Execute(sql)
MsgBox "New User Created", vbInformation
s = "SFMS"
Text1.Text = Mid(Text1.Text, 5) + 1
Text1.Text = s & Text1.Text
Text3.Text = "Enter Password"
Text2.Text = "Enter Username"
Text2.TabIndex = 0
Image5.Visible = False
Image6.Visible = False
If (Text2.Text = "Enter Username" Or Text3.Text = "Enter Password") Then
Text2.ForeColor = QBColor(7)
Text3.ForeColor = QBColor(7)
End If
End If
End Sub
Private Sub Image5_Click()
Image6.Visible = True
Image5.Visible = False
If (Text3.Text = "Enter Password") Then
Text3.PasswordChar = ""
Else
Text3.PasswordChar = "*"
End If
'Text2.SetFocus
End Sub

Private Sub Image6_Click()
Image5.Visible = True
Image6.Visible = False
Text3.PasswordChar = ""
'Text2.SetFocus
End Sub


