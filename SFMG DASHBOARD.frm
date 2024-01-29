VERSION 5.00
Begin VB.MDIForm SFMG 
   BackColor       =   &H8000000C&
   Caption         =   "SFM"
   ClientHeight    =   12195
   ClientLeft      =   1755
   ClientTop       =   1350
   ClientWidth     =   19620
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   Picture         =   "SFMG DASHBOARD.frx":0000
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   0
      Picture         =   "SFMG DASHBOARD.frx":0E3F
      ScaleHeight     =   1905
      ScaleWidth      =   19590
      TabIndex        =   1
      Top             =   0
      Width           =   19620
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      AutoSize        =   -1  'True
      Height          =   16260
      Left            =   0
      Picture         =   "SFMG DASHBOARD.frx":FCD5
      ScaleHeight     =   16200
      ScaleWidth      =   19560
      TabIndex        =   0
      Top             =   -4065
      Width           =   19620
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   11055
         Left            =   0
         Picture         =   "SFMG DASHBOARD.frx":1D828
         ScaleHeight     =   11025
         ScaleWidth      =   3585
         TabIndex        =   2
         Top             =   6480
         Width           =   3615
         Begin VB.PictureBox Picture12 
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   240
            Picture         =   "SFMG DASHBOARD.frx":23595
            ScaleHeight     =   735
            ScaleWidth      =   2895
            TabIndex        =   11
            Top             =   8760
            Width           =   2895
         End
         Begin VB.PictureBox Picture11 
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   240
            Picture         =   "SFMG DASHBOARD.frx":2429F
            ScaleHeight     =   735
            ScaleWidth      =   2895
            TabIndex        =   10
            Top             =   7800
            Width           =   2895
         End
         Begin VB.PictureBox Picture10 
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   240
            Picture         =   "SFMG DASHBOARD.frx":252FE
            ScaleHeight     =   735
            ScaleWidth      =   2895
            TabIndex        =   9
            Top             =   6840
            Width           =   2895
         End
         Begin VB.PictureBox Picture9 
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   240
            Picture         =   "SFMG DASHBOARD.frx":26338
            ScaleHeight     =   735
            ScaleWidth      =   2895
            TabIndex        =   8
            Top             =   5880
            Width           =   2895
         End
         Begin VB.PictureBox Picture8 
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   240
            Picture         =   "SFMG DASHBOARD.frx":2733E
            ScaleHeight     =   735
            ScaleWidth      =   2895
            TabIndex        =   7
            Top             =   5040
            Width           =   2895
         End
         Begin VB.PictureBox Picture7 
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   240
            Picture         =   "SFMG DASHBOARD.frx":283D1
            ScaleHeight     =   735
            ScaleWidth      =   2895
            TabIndex        =   6
            Top             =   4080
            Width           =   2895
         End
         Begin VB.PictureBox Picture6 
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   240
            Picture         =   "SFMG DASHBOARD.frx":294F7
            ScaleHeight     =   735
            ScaleWidth      =   2895
            TabIndex        =   5
            Top             =   3120
            Width           =   2895
         End
         Begin VB.PictureBox Picture5 
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   240
            Picture         =   "SFMG DASHBOARD.frx":2A349
            ScaleHeight     =   735
            ScaleWidth      =   2895
            TabIndex        =   4
            Top             =   2160
            Width           =   2895
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   240
            Picture         =   "SFMG DASHBOARD.frx":2B45B
            ScaleHeight     =   735
            ScaleWidth      =   2895
            TabIndex        =   3
            Top             =   1320
            Width           =   2895
         End
      End
   End
   Begin VB.Menu CLASS 
      Caption         =   "CLASS"
   End
   Begin VB.Menu STUDENT 
      Caption         =   "STUDENT"
   End
   Begin VB.Menu FEE_STRUCTURE 
      Caption         =   "FEE STRUCTURE"
   End
   Begin VB.Menu FEE_PAYMENT 
      Caption         =   "FEE PAYMENT"
   End
   Begin VB.Menu FEES_DUES 
      Caption         =   "FEE DUES"
   End
   Begin VB.Menu REPORT 
      Caption         =   "REPORT"
   End
   Begin VB.Menu PROFILE 
      Caption         =   "PROFILE"
      Begin VB.Menu NEW_PASSWORD 
         Caption         =   "NEW PASSWORD"
         Checked         =   -1  'True
      End
      Begin VB.Menu CHANGE_PASSWORD 
         Caption         =   "CHANGE  PASSWORD"
      End
      Begin VB.Menu LOGOUT 
         Caption         =   "LOGOUT"
      End
   End
End
Attribute VB_Name = "SFMG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub CHANGE_PASSWORD_Click()
Load Forget_Password
Forget_Password.Show
End Sub

Private Sub CLASS_Click()
Load CLASS_DET
CLASS_DET.Show
End Sub

Private Sub FEE_PAYMENT_Click()
Load Feepayment
Feepayment.Show
End Sub

Private Sub FEE_STRUCTURE_Click()
Load Feestructure
Feestructure.Show
End Sub

Private Sub FEES_DUES_Click()
Load FeeDues
FeeDues.Show
End Sub



Private Sub FINE_Click()

End Sub

Private Sub NEW_PASSWORD_Click()
Load Signup
Signup.Show

End Sub

Private Sub Picture10_Click()
Load Signup
Signup.Show

End Sub

Private Sub Picture4_Click()
Load CLASS_DET
CLASS_DET.Show
End Sub

Private Sub Picture5_Click()
Load Student_details
Student_details.Show
End Sub

Private Sub Picture6_Click()
Load Feestructure
Feestructure.Show
End Sub

Private Sub Picture7_Click()
Load Feepayment
Feepayment.Show
End Sub

Private Sub Picture8_Click()
Load FeeDues
FeeDues.Show
End Sub

Private Sub STUDENT_Click()
Load Student_details
Student_details.Show
End Sub
