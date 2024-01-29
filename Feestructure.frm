VERSION 5.00
Begin VB.Form Feestructure 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   12270
   ClientLeft      =   3195
   ClientTop       =   405
   ClientWidth     =   20910
   FillColor       =   &H00FFFFC0&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   17.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFC0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   12270
   ScaleWidth      =   20910
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   10320
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   10815
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   19575
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   12960
         TabIndex        =   18
         Top             =   600
         Width           =   5055
         Begin VB.OptionButton Option2 
            Caption         =   "Search"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   3000
            TabIndex        =   24
            Top             =   1920
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "View all"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   3000
            TabIndex        =   23
            Top             =   1560
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
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
            Height          =   435
            Left            =   1320
            TabIndex        =   21
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox Text4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   435
            Left            =   1680
            TabIndex        =   20
            Text            =   "feetype-classname"
            Top             =   840
            Width           =   2415
         End
         Begin VB.Label Label16 
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
            Height          =   375
            Left            =   240
            TabIndex        =   19
            Top             =   840
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   0
         TabIndex        =   17
         Top             =   9000
         Width           =   16455
         Begin VB.Image Image9 
            Height          =   855
            Left            =   14400
            Picture         =   "Feestructure.frx":0000
            Top             =   600
            Width           =   2325
         End
         Begin VB.Image Image8 
            Height          =   855
            Left            =   12120
            Picture         =   "Feestructure.frx":0BB3
            Top             =   600
            Width           =   2250
         End
         Begin VB.Image Image7 
            Height          =   855
            Left            =   9720
            Picture         =   "Feestructure.frx":184E
            Top             =   600
            Width           =   2325
         End
         Begin VB.Image Image6 
            Height          =   855
            Left            =   7320
            Picture         =   "Feestructure.frx":2458
            Top             =   600
            Width           =   2325
         End
         Begin VB.Image Image5 
            Height          =   870
            Left            =   4920
            Picture         =   "Feestructure.frx":2FB5
            Top             =   600
            Width           =   2250
         End
         Begin VB.Image Image4 
            Height          =   885
            Left            =   2520
            Picture         =   "Feestructure.frx":3B05
            Top             =   600
            Width           =   2250
         End
         Begin VB.Image Image3 
            Height          =   870
            Left            =   360
            Picture         =   "Feestructure.frx":4810
            Top             =   600
            Width           =   2010
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   4815
         Left            =   2400
         TabIndex        =   13
         Top             =   4320
         Width           =   13575
         Begin VB.TextBox Text5 
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   10680
            TabIndex        =   25
            Top             =   4200
            Width           =   2295
         End
         Begin VB.ListBox List5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3420
            ItemData        =   "Feestructure.frx":516E
            Left            =   10800
            List            =   "Feestructure.frx":5170
            TabIndex        =   22
            Top             =   480
            Width           =   2415
         End
         Begin VB.ListBox List4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3420
            ItemData        =   "Feestructure.frx":5172
            Left            =   7320
            List            =   "Feestructure.frx":5174
            TabIndex        =   16
            Top             =   480
            Width           =   2415
         End
         Begin VB.ListBox List3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3420
            ItemData        =   "Feestructure.frx":5176
            Left            =   3840
            List            =   "Feestructure.frx":5178
            TabIndex        =   15
            Top             =   480
            Width           =   2415
         End
         Begin VB.ListBox List2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3420
            ItemData        =   "Feestructure.frx":517A
            Left            =   480
            List            =   "Feestructure.frx":517C
            TabIndex        =   14
            Top             =   480
            Width           =   2415
         End
         Begin VB.Label Label21 
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   375
            Left            =   10920
            TabIndex        =   30
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label Label20 
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   375
            Left            =   7680
            TabIndex        =   29
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label Label19 
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   375
            Left            =   4080
            TabIndex        =   28
            Top             =   120
            Width           =   2175
         End
         Begin VB.Label Label18 
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   375
            Left            =   840
            TabIndex        =   27
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9360
            TabIndex        =   26
            Top             =   4320
            Width           =   975
         End
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   9480
         TabIndex        =   12
         ToolTipText     =   "PLEASE ENTER ONLY NUMERIC VALUE"
         Top             =   3480
         Width           =   1815
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9480
         TabIndex        =   10
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox Text2 
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
         Left            =   9480
         TabIndex        =   9
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   9480
         TabIndex        =   7
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Image Image10 
         Height          =   855
         Left            =   16560
         Picture         =   "Feestructure.frx":517E
         Top             =   9600
         Width           =   2565
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   11
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Fee Type"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   8
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Class"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   6
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
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
         Left            =   9480
         TabIndex        =   5
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fees Structure Id"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   4
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Image Image2 
         Height          =   1680
         Left            =   0
         Picture         =   "Feestructure.frx":5DE1
         Top             =   0
         Width           =   2130
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "FEE STRUCTURE"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5760
         TabIndex        =   1
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12840
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12840
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   21960
      Left            =   120
      Picture         =   "Feestructure.frx":76EB
      Top             =   240
      Width           =   40425
   End
End
Attribute VB_Name = "Feestructure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False









Private Sub Command1_Click()
List2.Clear
List3.Clear
List4.Clear
List5.Clear
Text1.Locked = True
If Option1.Value = True Then
Option2.Value = False
On Error GoTo xxxx
Abc
sql = "select * from fee_structure where class='" + Text4.Text + "'"
Set r = c.Execute(sql)
 Do While (r.EOF = False)
List2.AddItem r.Fields("FSID")
List3.AddItem r.Fields("CLASS")
List4.AddItem r.Fields("FEETYPE")
List5.AddItem r.Fields("AMOUNT")
r.MoveNext
Loop

Image7.Enabled = False
Image8.Enabled = False
Text4.Text = ""
Text4.SetFocus
'Label2.Enabled = True
Text1.Locked = False
Text2.Locked = False
Text3.Locked = False
'List2.Visible = False
'List3.Visible = False
'List3.Visible = False
'Frame2.Visible = False
Dim Total As Integer
Dim i As Integer
For i = 0 To List5.ListCount - 1
Total = Total + Val(List5.List(i))
Next i
Text5.Text = Format(Total, "###0.00")
Exit Sub
xxxx:
MsgBox "Wrong Class"
Text4.Text = ""
Text4.SetFocus
End If
If Option2.Value = True Then
On Error GoTo xxx
Abc
sql = "select * from fee_structure where fsid='" + Text4.Text + "'"
Set r = c.Execute(sql)
Label2.Caption = r.Fields("FSID")
Text1.Text = r.Fields("CLASS")
Text2.Text = r.Fields("FEETYPE")
Text3.Text = r.Fields("AMOUNT")
Text4.Text = ""
Text4.SetFocus
Image8.Enabled = True
Image7.Enabled = True
Label2.Enabled = True
Text1.Locked = False
Text2.Locked = False
Text3.Locked = False
'List2.Visible = False
'List3.Visible = False
'List3.Visible = False
Frame2.Visible = False
Exit Sub
xxx:
MsgBox "Wrong FEE STRUCTURE ID"
Text4.Text = ""
Text4.SetFocus
End If

End Sub

Private Sub Form_Load()
Text4.Locked = True
Image5.Enabled = False
Image6.Enabled = False
List1.Visible = False
List5.Visible = False
Frame4.Visible = False
Label1.Visible = False
Label2.Visible = True
List2.Enabled = False
List3.Enabled = False
List4.Enabled = False
List5.Enabled = False
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text5.Visible = False
Label17.Visible = False
Text5.Locked = True
Abc
sql = "select distinct class from class_detail order by class asc"
Set r = c.Execute(sql)
Do While Not r.EOF
List1.AddItem r.Fields("CLASS")
r.MoveNext
Loop
Image7.Enabled = False
Image8.Enabled = False
End Sub



Private Sub Image3_Click()
List1.Visible = True
List1.SetFocus
Text1.Visible = False
List1.SetFocus
List2.Enabled = False
List3.Enabled = False
List4.Enabled = False
Text1.Locked = False
'Text2.Locked = False
Text3.Locked = False
List2.Enabled = True
List3.Enabled = True
List4.Enabled = True
Image6.Enabled = True
Image5.Enabled = True
Label18.Caption = "Fee structure id"
Label19.Caption = "Fee Type"
Label20.Caption = "Amount"
End Sub

Private Sub Image4_Click()
Text2.Text = ""
Text3.Text = ""
List2.Text = ""
List3.Text = ""
List4.Text = ""
End Sub

Private Sub Image5_Click()
Abc
On Error GoTo CCC
Dim i As Integer
For i = 0 To List2.ListCount - 1
sql = "Insert into FEE_STRUCTURE values('" + List2.List(i) + "','" + List1.Text + "','" + List3.List(i) + "'," + List4.List(i) + ")"
Set r = c.Execute(sql)
'List1.Enabled = True
'List1.SetFocus
Next
MsgBox "Record Saved"
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Label18.Caption = ""
Label19.Caption = ""
Label20.Caption = ""
List2.Clear
List3.Clear
List4.Clear
Text2.Locked = True
Text3.Locked = True
Text1.Visible = True
List1.Visible = False

Exit Sub
CCC:
MsgBox "Exist"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
List2.Clear
List3.Clear
List4.Clear
Text2.SetFocus
End Sub

Private Sub Image6_Click()
Frame4.Visible = True
Text1.Visible = True
List1.Visible = False
Label1.Visible = True
Label2.Visible = True
List5.Visible = True
Image6.Enabled = False
MsgBox "please select all fees or selected fees", vbInformation
Image5.Enabled = False
List2.Clear
List3.Clear
List4.Clear
Label18.Caption = "Fee structure id "
Label19.Caption = "Class"
Label20.Caption = "Fee type"
Label21.Caption = "Amount"
End Sub

Private Sub Image8_Click()
Abc
sql = "Delete from fee_structure where fsid='" + Label2.Caption + "'"
Set r = c.Execute(sql)
MsgBox "Record Deleted"
Image8.Enabled = False
List1.Visible = False
Text1.Visible = True
Label2.Caption = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Frame2.Visible = True
Text1.Visible = False
List1.Visible = True
Text2.Locked = True
Text3.Locked = True
Frame4.Visible = False
Text1.Visible = False
Label1.Visible = False
End Sub

Private Sub Image9_Click()
B
End Sub

Private Sub image7_Click()
Abc
sql = "Update  FEE_STRUCTURE set class='" + Text1.Text + "',FEETYPE='" + Text2.Text + "',AMOUNT= " + Text3.Text + " where FSID='" + Label2.Caption + "'"
Set r = c.Execute(sql)
MsgBox "Record Updated"
Image7.Enabled = False
List1.Visible = False
Text1.Visible = True
Label2.Caption = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Frame2.Visible = True
Text1.Visible = False
List1.Visible = True
Text2.Locked = True
Text3.Locked = True
Frame4.Visible = False
Text1.Visible = False
Label1.Visible = False
Label21.Caption = ""
End Sub

Private Sub Label11_Click()

End Sub

Private Sub Label15_Click()
End Sub

Private Sub Label6_Click()


End Sub

Private Sub Label7_Click()


End Sub

Private Sub Label8_Click()

End Sub

Private Sub Label9_Click()
End Sub

Private Sub List1_Click()
Text2.SetFocus
Text2.Locked = False
End Sub




Private Sub Option1_Click()
Text4.SetFocus
Text4.Locked = False
Frame2.Visible = True
List5.Visible = True
Text5.Visible = True
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Label2.Caption = ""
If Option1.Value = True Then
Label17.Visible = True
Else
Label17.Visible = False
End If
'Label1.Caption = False
'Label2.Caption = False
'Label3.Caption = False
'Label4.Caption = False
'Label5.Caption = False
'List1.Visible = False
'Text1.Visible = False
'Text2.Visible = False
'Text3.Visible = False

End Sub

Private Sub Option2_Click()
Text4.Locked = False
Option1.Value = False
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List5.Visible = False
Text4.SetFocus
Frame2.Visible = False
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Label2.Caption = ""
End Sub



Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Frame2.Visible = True Then
Text2.Text = UCase(Text2.Text)
 List2.AddItem Text2.Text & "-" & List1.Text
 List3.AddItem Text2.Text
 Text3.SetFocus
 ElseIf (KeyAscii = 13 And Frame2.Visible = False) Then
 Text3.SetFocus
End If
End Sub







Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text3.Text = UCase(Text3.Text)
List1.SetFocus
Text2.Locked = True
 'List4.AddItem Mid(Text3.Text, 1, 3) & ".00"
 List4.AddItem Format(Text3.Text, "###0.00")
 Text2.Text = ""
 Text3.Text = ""
 ElseIf (KeyAscii = 13 And Frame2.Visible = False) Then
 Text1.SetFocus
 'If KeyAscii <> 48 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 And KeyAscii <> 53 And KeyAscii <> 54 And KeyAscii <> 55 And KeyAscii <> 56 And KeyAscii <> 57 Then
 'MsgBox "PLEASE ENTER A NUMBER"
 'Text3.SetFocus
 'Text3.Text = ""
'End If
End If
If KeyAscii < 48 Or KeyAscii > 57 Then
If KeyAscii <> 8 Then
KeyAscii = 0
End If
End If
End Sub



Private Sub Text3_LostFocus()

If Text2.Text = "" Then
MsgBox "please enter the fee type"
List1.SetFocus
Text3.Text = ""
End If

End Sub

Private Sub Text4_Click()
If (Text4.Text = "feetype-classname") Then
Text4.Text = ""
End If
End Sub







Private Sub Text4_GotFocus()
Text4.ForeColor = vbBlack
If (Text4.Text = "feetype-classname") Then
Text4.Text = ""
Else
Text4.Text = Text4.Text
End If
End Sub



Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text4.Text = UCase(Text4.Text)
Command1.SetFocus
End If
End Sub


Private Sub Text4_LostFocus()
If (Text4.Text = "") Then
Text4.Text = "feetype-classname"
Text1.ForeColor = vbBlack
End If
End Sub

Private Sub Timer1_Timer()
Label13.Caption = Format(Now, "yyyy-MM-dd ")
Label14.Caption = Format(Now, "hh:mm:ss AM/PM")
End Sub
