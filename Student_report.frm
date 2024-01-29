VERSION 5.00
Begin VB.Form Student_report 
   Caption         =   "Form1"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   10065
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   3120
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   720
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   975
      Left            =   5040
      TabIndex        =   1
      Top             =   2160
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   1680
      TabIndex        =   0
      Top             =   2160
      Width           =   2535
   End
End
Attribute VB_Name = "Student_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
DataReport3.LeftMargin = 2
DataReport3.RightMargin = 2
DataReport3.Show
End Sub

Private Sub Command2_Click()
DataEnvironment1.Command4 Text1.Text
DataReport4.LeftMargin = 2
DataReport4.RightMargin = 2
DataReport4.Show
Text1.Text = ""
Set DataEnvironment1 = Nothing


End Sub
