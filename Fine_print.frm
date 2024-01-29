VERSION 5.00
Begin VB.Form Class_report 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   4605
   ClientLeft      =   5040
   ClientTop       =   2850
   ClientWidth     =   7920
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   7920
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3135
      Left            =   960
      TabIndex        =   0
      Top             =   720
      Width           =   6015
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   720
         TabIndex        =   2
         Top             =   1080
         Width           =   4815
         Begin VB.CommandButton Command2 
            Caption         =   "All"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2640
            TabIndex        =   5
            Top             =   1080
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H80000016&
            Caption         =   "View"
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
            Left            =   480
            TabIndex        =   4
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox Text1 
            Height          =   495
            Left            =   1200
            TabIndex        =   3
            Top             =   0
            Width           =   2295
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Class"
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
            Left            =   360
            TabIndex        =   6
            Top             =   120
            Width           =   735
         End
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Class Report"
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
         Left            =   2400
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Image Image1 
      Height          =   21960
      Left            =   -360
      Picture         =   "Fine_print.frx":0000
      Top             =   -240
      Width           =   40425
   End
End
Attribute VB_Name = "Class_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

DataEnvironment1.Command2 Text1.Text
DataReport2.Show
Text1.Text = ""
Set DataEnvironment1 = Nothing



End Sub
Private Sub Command2_Click()

DataReport1.LeftMargin = 0.25
DataReport1.RightMargin = 0.25
DataReport1.Show

End Sub

Private Sub Form_Load()
'Adodc1.Visible = False

End Sub



Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text1.Text = UCase(Text1.Text)
Command1.SetFocus
End If
End Sub

