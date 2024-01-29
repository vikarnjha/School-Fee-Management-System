VERSION 5.00
Begin VB.Form CLASS_DET 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18255
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   18255
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000015&
      Height          =   7935
      Left            =   0
      Picture         =   "Class.frx":0000
      ScaleHeight     =   7875
      ScaleWidth      =   18195
      TabIndex        =   0
      Top             =   0
      Width           =   18255
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   120
         Top             =   6840
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   5775
         Left            =   600
         TabIndex        =   1
         Top             =   840
         Width           =   16815
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1695
            Left            =   11400
            TabIndex        =   11
            Top             =   360
            Width           =   4215
            Begin VB.CommandButton Command1 
               BackColor       =   &H00FFFFFF&
               Caption         =   "SEARCH"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   1320
               Picture         =   "Class.frx":17515
               Style           =   1  'Graphical
               TabIndex        =   14
               Top             =   840
               Width           =   2535
            End
            Begin VB.TextBox Text4 
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   1800
               TabIndex        =   13
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "SEARCH"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   240
               TabIndex        =   12
               Top             =   360
               Width           =   1215
            End
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            ForeColor       =   &H80000008&
            Height          =   975
            Left            =   0
            TabIndex        =   10
            Top             =   4080
            Width           =   14055
            Begin VB.Image Image7 
               Height          =   855
               Left            =   11760
               Picture         =   "Class.frx":17D38
               Top             =   0
               Width           =   2325
            End
            Begin VB.Image Image6 
               Height          =   855
               Left            =   9360
               Picture         =   "Class.frx":188EB
               Top             =   0
               Width           =   2325
            End
            Begin VB.Image Image5 
               Height          =   855
               Left            =   7080
               Picture         =   "Class.frx":194F5
               Top             =   0
               Width           =   2250
            End
            Begin VB.Image Image4 
               Height          =   870
               Left            =   4800
               Picture         =   "Class.frx":1A190
               Top             =   0
               Width           =   2250
            End
            Begin VB.Image Image3 
               Height          =   885
               Left            =   2400
               Picture         =   "Class.frx":1ACE0
               Top             =   0
               Width           =   2250
            End
            Begin VB.Image Image2 
               Height          =   870
               Left            =   240
               Picture         =   "Class.frx":1B9EB
               Top             =   0
               Width           =   2010
            End
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   7800
            TabIndex        =   8
            Top             =   1920
            Width           =   1215
         End
         Begin VB.TextBox Text2 
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
            Height          =   405
            Left            =   7800
            TabIndex        =   7
            Top             =   2400
            Width           =   1215
         End
         Begin VB.TextBox Text3 
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
            Height          =   405
            Left            =   7800
            TabIndex        =   6
            Top             =   2880
            Width           =   1215
         End
         Begin VB.Image Image8 
            Height          =   855
            Left            =   14160
            Picture         =   "Class.frx":1C349
            Top             =   4080
            Width           =   2565
         End
         Begin VB.Label Label6 
            Caption         =   "CLASS"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   615
            Left            =   6600
            TabIndex        =   18
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label16 
            Height          =   255
            Left            =   9240
            TabIndex        =   17
            Top             =   2520
            Width           =   1095
         End
         Begin VB.Image Image1 
            Height          =   1680
            Left            =   120
            Picture         =   "Class.frx":1CFAC
            Top             =   120
            Width           =   2130
         End
         Begin VB.Label Label5 
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
            Height          =   375
            Left            =   7800
            TabIndex        =   9
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "CID"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4680
            TabIndex        =   5
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "CLASS"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   4680
            TabIndex        =   4
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "SECTION"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   4680
            TabIndex        =   3
            Top             =   2520
            Width           =   975
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "CLASS STRENGTH"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   4680
            TabIndex        =   2
            Top             =   3000
            Width           =   2055
         End
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
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
         Left            =   15960
         TabIndex        =   16
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
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
         Left            =   15960
         TabIndex        =   15
         Top             =   120
         Width           =   1335
      End
   End
End
Attribute VB_Name = "CLASS_DET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo xxxx
Abc
sql = "select * from class_detail where cid='" + Text4.Text + "'"
Set r = c.Execute(sql)
Label5.Caption = r.Fields("CID")
Text1.Text = r.Fields("CLASS")
Text2.Text = r.Fields("SECTION")
Text3.Text = r.Fields(3)
Text4.Text = ""
Text4.SetFocus
Image5.Enabled = True
Image6.Enabled = True
Label5.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Exit Sub
xxxx:
MsgBox "Wrong CLASS ID"
Text4.Text = ""
Text4.SetFocus
End Sub

Private Sub Form_Load()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Label5.Enabled = False
Image5.Enabled = False
Image6.Enabled = False
Abc
sql = "select count(cno)from class_detail"
Set r = c.Execute(sql)
Label16.Caption = r.Fields(0) + 1
Label16.TabIndex = 0
Label16.Visible = False
End Sub

Private Sub Image2_Click()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text1.SetFocus
End Sub

Private Sub Image3_Click()
If Text1.Enabled = True Then
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Label5.Caption = ""
Text1.SetFocus
Else
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Label5.Caption = ""
End If
End Sub

Private Sub Image4_Click()
On Error GoTo ins
Abc
sql = "Insert into class_detail values('" + Label5.Caption + "','" + Text1.Text + "','" + Text2.Text + "'," + Text3.Text + "," + Label16.Caption + ")"
Set r = c.Execute(sql)
MsgBox "Record Saved"
Label16.Caption = Label16.Caption + 1
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Label5.Caption = ""
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False

Exit Sub
ins:
MsgBox "exist"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Label5.Caption = ""
Text1.SetFocus
End Sub

Private Sub Image5_Click()
Abc
sql = "Delete from class_detail where cid='" + Label5.Caption + "'"
Set r = c.Execute(sql)
MsgBox "Record Deleted"
Label5.Caption = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End Sub

Private Sub Image6_Click()
Abc
If Text2.Text = "A" Or Text2.Text = "B" Then
sql = "Update  class_detail set class='" + Text1.Text + "',section='" + Text2.Text + "',cstr= " + Text3.Text + " where cid='" + Label5.Caption + "'"
Set r = c.Execute(sql)
MsgBox "Record Updated"
Label5.Caption = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Else
MsgBox "Please Enter Section A OR B"
End If
End Sub

Private Sub Label10_Click()

End Sub

Private Sub Label8_Click()

End Sub

Private Sub image7_Click()
B
End Sub

Private Sub Label11_Click()

End Sub

Private Sub Label7_Click()

End Sub

Private Sub Label9_Click()

End Sub

Private Sub Image8_Click()
Class_report.Show

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Then
MsgBox "Please Enter in only in roman numeral"
Text1.Text = ""
Else
If KeyAscii = 13 Then
Text2.SetFocus
Text1.Text = UCase(Text1.Text)
End If
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.Text = UCase(Text2.Text)
If Text2.Text <> "A" And Text2.Text <> "B" Then
MsgBox "Please Enter Only Section A AND B"
Text2.Text = ""
Else
Label5.Caption = Text1.Text & "-" & Text2.Text



Text3.SetFocus

End If
End If
End Sub




Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text4.Text = UCase(Text4.Text)
Command1.SetFocus
End If
End Sub

Private Sub Timer1_Timer()
Label14.Caption = Format(Now, "yyyy-MM-dd ")
Label15.Caption = Format(Now, "hh:mm:ss AM/PM")
End Sub
