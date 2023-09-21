VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Schoolfee 
   Caption         =   "School Fee Management System"
   ClientHeight    =   10095
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18630
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10095
   ScaleWidth      =   18630
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   270
      Left            =   3360
      TabIndex        =   2
      Top             =   11400
      Width           =   16215
      _ExtentX        =   28601
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   17880
      Top             =   9600
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Registered Since :- 2021 "
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9480
      TabIndex        =   3
      Top             =   2520
      Width           =   5175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   735
      Left            =   9960
      TabIndex        =   1
      Top             =   10800
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading...."
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   855
      Left            =   8760
      TabIndex        =   0
      Top             =   10320
      Width           =   5895
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   13095
      Left            =   0
      OLEDragMode     =   1  'Automatic
      Picture         =   "Splash_Screen.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   23775
   End
End
Attribute VB_Name = "Schoolfee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim randomNumber As Integer

Private Sub Form_Load()
    Timer1.Enabled = True
    ' Initialize the random number generator
    Randomize

    ' Generate a random number between 1 and 100
    randomNumber = Int((4 * Rnd) + 1)
End Sub


Private Sub Timer1_Timer()
ProgressBar1.Visible = True
ProgressBar1.Value = ProgressBar1.Value + 4
Label3.Visible = True
Label4.Caption = ProgressBar1.Value & "%"
If ProgressBar1.Value = ProgressBar1.Max Then
'ProgressBar1.value = 100
Timer1.Enabled = False
Login1.Show
Unload Me
End If
End Sub
