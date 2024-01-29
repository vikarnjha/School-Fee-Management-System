VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Student_details 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Student"
   ClientHeight    =   12465
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   22920
   BeginProperty Font 
      Name            =   "Segoe Print"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   12465
   ScaleWidth      =   22920
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   11640
   End
   Begin VB.Frame Frame1 
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
      Height          =   10695
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   21975
      Begin VB.TextBox Text24 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   70
         Top             =   5040
         Width           =   1575
      End
      Begin VB.TextBox Text21 
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
         Left            =   4440
         TabIndex        =   68
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox Text20 
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
         Left            =   4440
         TabIndex        =   67
         Top             =   2400
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   2880
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   61
         Text            =   "Combo1"
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox Text23 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   60
         Top             =   5520
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Height          =   555
         Left            =   15480
         TabIndex        =   59
         Top             =   8520
         Width           =   255
      End
      Begin VB.Frame Frame3 
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
         Height          =   1695
         Left            =   12360
         TabIndex        =   55
         Top             =   240
         Width           =   5775
         Begin VB.CheckBox Check2 
            Caption         =   "Retrive"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   4320
            TabIndex        =   69
            Top             =   960
            Width           =   975
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Search"
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
            Left            =   1560
            Picture         =   "Student_details.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   58
            Top             =   960
            Width           =   2655
         End
         Begin VB.TextBox Text22 
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
            Left            =   1560
            TabIndex        =   57
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   "Search"
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
            Left            =   360
            TabIndex        =   56
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   360
         TabIndex        =   54
         Top             =   9120
         Width           =   19695
         Begin VB.Image Image11 
            Height          =   855
            Left            =   16200
            Picture         =   "Student_details.frx":075D
            Top             =   480
            Width           =   2565
         End
         Begin VB.Image Image4 
            Height          =   870
            Left            =   0
            Picture         =   "Student_details.frx":15B9
            Top             =   480
            Width           =   2010
         End
         Begin VB.Image Image8 
            Height          =   855
            Left            =   11400
            Picture         =   "Student_details.frx":1F17
            Top             =   480
            Width           =   2325
         End
         Begin VB.Image Image7 
            Height          =   855
            Left            =   9000
            Picture         =   "Student_details.frx":2A74
            Top             =   480
            Width           =   2325
         End
         Begin VB.Image Image6 
            Height          =   855
            Left            =   6600
            Picture         =   "Student_details.frx":367E
            Top             =   480
            Width           =   2250
         End
         Begin VB.Image Image5 
            Height          =   870
            Left            =   4320
            Picture         =   "Student_details.frx":4319
            Top             =   480
            Width           =   2250
         End
         Begin VB.Image Image3 
            Height          =   885
            Left            =   2040
            Picture         =   "Student_details.frx":4E69
            Top             =   480
            Width           =   2250
         End
         Begin VB.Image Image10 
            Height          =   855
            Left            =   13800
            Picture         =   "Student_details.frx":5B74
            Top             =   480
            Width           =   2325
         End
      End
      Begin MSComCtl2.MonthView MonthView2 
         Height          =   2370
         Left            =   18720
         TabIndex        =   53
         Top             =   4920
         Width           =   2910
         _ExtentX        =   5133
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   0
         BackColor       =   16777215
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         StartOfWeek     =   157155330
         CurrentDate     =   45268
      End
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2370
         Left            =   7680
         TabIndex        =   52
         Top             =   4560
         Width           =   2910
         _ExtentX        =   5133
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   0
         BackColor       =   0
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         StartOfWeek     =   157155330
         CurrentDate     =   45268
      End
      Begin VB.TextBox Text19 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   15840
         TabIndex        =   49
         Top             =   8280
         Width           =   5535
      End
      Begin VB.TextBox Text18 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   15840
         TabIndex        =   47
         Top             =   7440
         Width           =   5535
      End
      Begin VB.TextBox Text17 
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
         Left            =   15840
         TabIndex        =   45
         Top             =   6840
         Width           =   2775
      End
      Begin VB.TextBox Text16 
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
         Left            =   4440
         TabIndex        =   43
         ToolTipText     =   "Please enter only 12 numbers"
         Top             =   7080
         Width           =   2895
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Others"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   6720
         TabIndex        =   41
         Top             =   5400
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Female"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   5640
         TabIndex        =   40
         Top             =   5400
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Male"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   4440
         TabIndex        =   39
         Top             =   5400
         Width           =   855
      End
      Begin VB.TextBox Text15 
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
         Left            =   4440
         TabIndex        =   37
         Top             =   6480
         Width           =   2655
      End
      Begin VB.TextBox Text14 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   35
         Top             =   6000
         Width           =   2655
      End
      Begin VB.TextBox Text13 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   16320
         TabIndex        =   33
         ToolTipText     =   "Please enter only 10 numbers"
         Top             =   6360
         Width           =   2295
      End
      Begin VB.TextBox Text12 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   15840
         TabIndex        =   31
         Top             =   5760
         Width           =   2775
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   15840
         TabIndex        =   29
         Top             =   5280
         Width           =   2535
      End
      Begin VB.TextBox Text10 
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
         Left            =   4440
         TabIndex        =   28
         Top             =   4560
         Width           =   2655
      End
      Begin VB.TextBox Text9 
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
         Left            =   15840
         TabIndex        =   24
         Top             =   4680
         Width           =   2895
      End
      Begin VB.TextBox Text8 
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
         Left            =   15840
         TabIndex        =   22
         Top             =   4200
         Width           =   3015
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   16320
         TabIndex        =   21
         ToolTipText     =   "Please enter only 10 numbers"
         Top             =   3720
         Width           =   2895
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   15840
         TabIndex        =   18
         Top             =   3120
         Width           =   3135
      End
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
         Left            =   15840
         TabIndex        =   16
         Top             =   2520
         Width           =   3135
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
         Height          =   435
         Left            =   15840
         TabIndex        =   14
         Top             =   1920
         Width           =   2775
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
         Height          =   375
         Left            =   4920
         TabIndex        =   12
         ToolTipText     =   "Please enter only 10 numbers"
         Top             =   8280
         Width           =   2775
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
         Height          =   435
         Left            =   4440
         TabIndex        =   10
         Top             =   7680
         Width           =   4215
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
         Height          =   435
         Left            =   4440
         TabIndex        =   5
         Top             =   3960
         Width           =   4335
      End
      Begin VB.Label Label47 
         BackColor       =   &H80000016&
         Caption         =   "No dues no fee"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000016&
         Height          =   375
         Left            =   19680
         TabIndex        =   75
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000016&
         Height          =   375
         Left            =   16320
         TabIndex        =   74
         Top             =   10080
         Width           =   495
      End
      Begin VB.Label Label45 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "+91"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   15840
         TabIndex        =   73
         Top             =   6360
         Width           =   495
      End
      Begin VB.Label Label44 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "+91"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   15840
         TabIndex        =   72
         Top             =   3720
         Width           =   495
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
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
         Left            =   1200
         TabIndex        =   71
         Top             =   5040
         Width           =   975
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   66
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "+91"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   65
         Top             =   8280
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Label10"
         Height          =   495
         Left            =   3120
         TabIndex        =   64
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label40 
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
         Height          =   375
         Left            =   4440
         TabIndex        =   62
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Correspondence Address"
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
         Left            =   11880
         TabIndex        =   51
         Top             =   8520
         Width           =   3015
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Permanent Address"
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
         Left            =   11880
         TabIndex        =   50
         Top             =   7680
         Width           =   2775
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of admission"
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
         Left            =   11880
         TabIndex        =   48
         Top             =   6840
         Width           =   2415
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Aadhar Number"
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
         Left            =   1200
         TabIndex        =   46
         Top             =   7080
         Width           =   2295
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "City"
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
         Left            =   1200
         TabIndex        =   44
         Top             =   6000
         Width           =   1095
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
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
         Left            =   1200
         TabIndex        =   42
         Top             =   5520
         Width           =   855
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "State"
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
         Left            =   1200
         TabIndex        =   38
         Top             =   6480
         Width           =   1215
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Guardian mobile number"
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
         Left            =   11880
         TabIndex        =   34
         Top             =   6360
         Width           =   2895
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Student Relation with guardian"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   11880
         TabIndex        =   32
         Top             =   5640
         Width           =   2415
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Guardian Name"
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
         Left            =   11880
         TabIndex        =   30
         Top             =   5160
         Width           =   3015
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of birth"
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
         Left            =   1200
         TabIndex        =   27
         Top             =   4560
         Width           =   2175
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Mother's Occupation"
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
         Left            =   11880
         TabIndex        =   26
         Top             =   4680
         Width           =   2775
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Mother's Qualification "
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
         Left            =   11880
         TabIndex        =   25
         Top             =   4200
         Width           =   2895
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Mother's Mobile Number"
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
         Left            =   11880
         TabIndex        =   23
         Top             =   3720
         Width           =   2895
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Mother's Name"
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
         Left            =   11880
         TabIndex        =   20
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Father's Occupation"
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
         Left            =   11880
         TabIndex        =   19
         Top             =   2520
         Width           =   2775
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Father's Qualification"
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
         Left            =   11880
         TabIndex        =   17
         Top             =   1920
         Width           =   2655
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Father's Mobile Number"
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
         Left            =   1200
         TabIndex        =   15
         Top             =   8280
         Width           =   2895
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Father's Name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1200
         TabIndex        =   13
         Top             =   7680
         Width           =   2415
      End
      Begin VB.Label Label9 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   11
         Top             =   1920
         Width           =   2895
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Student Name"
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
         Left            =   1200
         TabIndex        =   9
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Roll Number"
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
         Left            =   1200
         TabIndex        =   8
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Section"
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
         Left            =   1200
         TabIndex        =   7
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label5 
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
         Left            =   1200
         TabIndex        =   6
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Registration Number"
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
         Left            =   1200
         TabIndex        =   4
         Top             =   1920
         Width           =   2655
      End
      Begin VB.Line Line1 
         X1              =   9480
         X2              =   12120
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Student Details"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9480
         TabIndex        =   3
         Top             =   480
         Width           =   2655
      End
      Begin VB.Image Image2 
         Height          =   1680
         Left            =   0
         Picture         =   "Student_details.frx":6727
         Top             =   0
         Width           =   2130
      End
   End
   Begin VB.Label Label24 
      Caption         =   "Label24"
      Height          =   495
      Left            =   10920
      TabIndex        =   36
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label2 
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
      Height          =   375
      Left            =   21000
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label1 
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
      Left            =   21000
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   21960
      Left            =   -1200
      Picture         =   "Student_details.frx":8031
      Top             =   0
      Width           =   40425
   End
End
Attribute VB_Name = "Student_details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
If Check1.Value = 1 Then
Text19.Text = Text18.Text
Text19.Locked = True
Else
Text19.Text = ""
Text19.Locked = False
End If
End Sub

Private Sub Combo1_Click()
Abc
Combo2.AddItem "A"
Combo2.AddItem "B"
End Sub

Private Sub Combo2_Click()

Dim year As String
year = Format(Now, "yyyy")
Static count As Integer
sql = "SELECT COUNT(ROLL) FROM student WHERE class='" + Combo1.Text + "' AND Section='" + Combo2.Text + "' "
Set r = c.Execute(sql)
Label10.Caption = r.Fields(0)
If Label10.Caption = 45 Then
Label40.Caption = 0
MsgBox "FULL"
Else
sql = "select count(roll)from student where class='" + Combo1.Text + "' AND Section='" + Combo2.Text + "'"
Set r = c.Execute(sql)
Label40.Caption = r.Fields(0) + 1
End If
sql = "select count(regid) from student"
Set r = c.Execute(sql)
count = r.Fields(0) + 1
Label9.Caption = year & "-" & Combo1.Text & "-" & "000" & count
A = Text3.Text
Dim selectedSection As String
selectedSection = Combo2.Text
Dim rollnumber As Integer
rollnumber = Label40.Caption
If selectedSection = "A" Then
If rollnumber >= 1 And rollnumber <= 45 Then
MsgBox "Roll No " & rollnumber & " Section A is selected", vbInformation
Else
MsgBox "class full. cannot select section a . you have to  select section b,", vbInformation
End If
ElseIf selectedSection = "B" Then
If rollnumber < 45 Then
MsgBox "Invalid selection. Roll number is less than 45. please select section A", vbCritical
Combo2.Clear
Combo2.AddItem "A"
Combo2.AddItem "B"
Combo2.SetFocus
Exit Sub
End If
MsgBox "Roll No"
rollnumber = Label40.Caption
End If
End Sub







Private Sub Command1_Click()
Text1.Locked = False
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
Text5.Locked = False
Text6.Locked = False
Text7.Locked = False
Text8.Locked = False
Text9.Locked = False
Text10.Locked = False
Text11.Locked = False
Text12.Locked = False
Text13.Locked = False
Text14.Locked = False
Text15.Locked = False
Text16.Locked = False
Text17.Locked = False
Text18.Locked = False
Text19.Locked = False
Text23.Locked = False
On Error GoTo sss
If Check2.Value = False Then
Image6.Enabled = True
Image7.Enabled = True
Image11.Enabled = False
Abc
sql = "select * from student where regid='" + Text22.Text + "' and status =1"
Set r = c.Execute(sql)
Label9.Caption = r.Fields("REGID")
Text20.Text = r.Fields("Class")
Text21.Text = r.Fields("Section")
Label40.Caption = r.Fields("Roll")
Text1.Text = r.Fields("Sname")
Text2.Text = r.Fields("Fname")
Text3.Text = r.Fields("Fmob")
Text4.Text = r.Fields("Fquali")
Text5.Text = r.Fields("Foccup")
Text6.Text = r.Fields("Mname")
Text7.Text = r.Fields("Mmob")
Text8.Text = r.Fields("Mquali")
Text9.Text = r.Fields("Moccup")
Text10.Text = r.Fields("Dob")
Text11.Text = r.Fields("Gname")
Text12.Text = r.Fields("Gsr")
Text13.Text = r.Fields("Gmob")
Text23.Text = r.Fields("Gender")
Text14.Text = r.Fields("City")
Text15.Text = r.Fields("State")
Text16.Text = r.Fields("Aadharno")
Text17.Text = r.Fields("Doa")
Text18.Text = r.Fields("Padd")
Text19.Text = r.Fields("Cadd")
Label25.Caption = r.Fields("Status")
Text24.Text = r.Fields("Age")
Exit Sub:
sss:
MsgBox "No Record Found"
Image6.Enabled = False
Image7.Enabled = False
Text22.Text = ""
Text22.SetFocus
Combo1.Locked = False
Combo2.Locked = False
Text1.Locked = False
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
Text5.Locked = False
Text6.Locked = False
Text7.Locked = False
Text8.Locked = False
Text9.Locked = False
Text10.Locked = False
Text11.Locked = False
Text12.Locked = False
Text13.Locked = False
Text14.Locked = False
Text15.Locked = False
Text16.Locked = False
Text17.Locked = False
Text18.Locked = False
Text19.Locked = False
Text23.Locked = False
End If
On Error GoTo ssss
If Check2.Value = 1 Then
Image6.Enabled = False
Image7.Enabled = False
Image11.Enabled = True
Abc
sql = "select * from student where regid='" + Text22.Text + "' and status =0"
Set r = c.Execute(sql)
Label9.Caption = r.Fields("REGID")
Text20.Text = r.Fields("Class")
Text21.Text = r.Fields("Section")
Label40.Caption = r.Fields("Roll")
Text1.Text = r.Fields("Sname")
Text2.Text = r.Fields("Fname")
Text3.Text = r.Fields("Fmob")
Text4.Text = r.Fields("Fquali")
Text5.Text = r.Fields("Foccup")
Text6.Text = r.Fields("Mname")
Text7.Text = r.Fields("Mmob")
Text8.Text = r.Fields("Mquali")
Text9.Text = r.Fields("Moccup")
Text10.Text = r.Fields("Dob")
Text11.Text = r.Fields("Gname")
Text12.Text = r.Fields("Gsr")
Text13.Text = r.Fields("Gmob")
Text23.Text = r.Fields("Gender")
Text14.Text = r.Fields("City")
Text15.Text = r.Fields("State")
Text16.Text = r.Fields("Aadharno")
Text17.Text = r.Fields("Doa")
Text18.Text = r.Fields("Padd")
Text19.Text = r.Fields("Cadd")
Label25.Caption = r.Fields("Status")
Text24.Text = r.Fields("Age")
Exit Sub:
ssss:
MsgBox "No Record Found"
Text22.Text = ""
Text22.SetFocus
Combo1.Locked = False
Combo2.Locked = False
Text1.Locked = False
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
Text5.Locked = False
Text6.Locked = False
Text7.Locked = False
Text8.Locked = False
Text9.Locked = False
Text10.Locked = False
Text11.Locked = False
Text12.Locked = False
Text13.Locked = False
Text14.Locked = False
Text15.Locked = False
Text16.Locked = False
Text17.Locked = False
Text18.Locked = False
Text19.Locked = False
Text23.Locked = False
'Text24.Locked = False

End If

End Sub



Private Sub Form_Load()
Image8.Enabled = False
Label46.Visible = False
Label47.Visible = False
'Dim year As String
'year = Format(Now, "yyyy")
'Static count As Integer
'count = count + 1
Image5.Enabled = False
Image6.Enabled = False
Image7.Enabled = False
Image11.Enabled = False
Label25.Visible = False
Label10.Visible = False
MonthView1.Visible = False
MonthView2.Visible = False
Text23.Visible = False
Abc
sql = "select distinct  class from class_detail order by class"
Set r = c.Execute(sql)
Do While Not r.EOF
Combo1.AddItem r.Fields(0)
r.MoveNext
Loop

'Label9.Caption = year & "000" & count
Combo1.Locked = True
Combo2.Locked = True
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
Text6.Locked = True
Text7.Locked = True
Text8.Locked = True
Text9.Locked = True
Text10.Locked = True
Text11.Locked = True
Text12.Locked = True
Text13.Locked = True
Text14.Locked = True
Text15.Locked = True
Text16.Locked = True
Text17.Locked = True
Text18.Locked = True
Text19.Locked = True
Text23.Locked = True
Frame3.Visible = False
Text20.Visible = False
Text21.Visible = False
Combo1.Visible = True
Combo2.Visible = True
MonthView1.Value = Date
Text10.Text = Format(DateClicked, "dd mmm yyyy")
calculateage
End Sub

Private Sub calculateage()
Dim selecteddate As Date
If IsDate(Text10.Text) Then
selecteddate = CDate(Text10.Text)
Dim age As Integer
age = DateDiff("yyyy", selecteddate, Date)
Text24.Text = CStr(age)
Dim selectedClass As String
selectedClass = Combo1.List(Combo1.ListIndex)
Select Case selectedClass
Case "KGI-I"
If age < 4 Then
MsgBox "Age is not within the acceptable for KGI-I"
Text10.Text = ""
Text24.Text = ""
Text10.SetFocus
End If
Case "KGI-II"
If age < 5 Then
MsgBox "Age is not within the acceptable for KGI-II"
Text10.Text = ""
Text24.Text = ""
Text10.SetFocus
End If
Case "I"
If age < 6 Then
MsgBox "Age is not within the acceptable for class I"
Text10.Text = ""
Text24.Text = ""
Text10.SetFocus
End If
Case "II"
If age < 7 Then
MsgBox "Age is not within the acceptable for class II"
Text10.Text = ""
Text24.Text = ""
Text10.SetFocus
End If
Case "III"
If age < 8 Then
MsgBox "Age is not within the acceptable for class III"
Text10.Text = ""
Text24.Text = ""
Text10.SetFocus
End If
Case "IV"
If age < 9 Then
MsgBox "Age is not within the acceptable for class IV"
Text10.Text = ""
Text24.Text = ""
Text10.SetFocus
End If
Case "V"
If age < 10 Then
MsgBox "Age is not within the acceptable for class V"
Text10.Text = ""
Text24.Text = ""
Text10.SetFocus
End If
Case "VI"
If age < 11 Then
MsgBox "Age is not within the acceptable for class VI"
Text10.Text = ""
Text24.Text = ""
Text10.SetFocus
End If
Case "VII"
If age < 12 Then
MsgBox "Age is not within the acceptable for class VII"
Text10.Text = ""
Text24.Text = ""
Text10.SetFocus
End If
Case "VIII"
If age < 13 Then
MsgBox "Age is not within the acceptable for class VIII"
Text10.Text = ""
Text24.Text = ""
Text10.SetFocus
End If
Case "IX"
If age < 14 Then
MsgBox "Age is not within the acceptable for class IX"
Text10.Text = ""
Text24.Text = ""
Text10.SetFocus
End If
Case "X"
If age < 15 Then
MsgBox "Age is not within the acceptable for class X"
Text10.Text = ""
Text24.Text = ""
Text10.SetFocus
End If
Case Else
MsgBox "Invalid class selected"
Text10.Text = ""
Text24.Text = ""
Text10.SetFocus
End Select
Else
Text24.Text = ""
End If
End Sub

Private Sub Image10_Click()
B
End Sub

Private Sub Image11_Click()
MsgBox " Record Retrived"
sql = "update student set status = 1 where regid='" + Text22.Text + "'"
Set r = c.Execute(sql)
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = ""
Text19.Text = ""
Text23.Text = ""
Label9.Caption = ""
Text20.Text = ""
Text21.Text = ""
Label40.Caption = ""
Text22.SetFocus
Text22.Text = blank
Image11.Enabled = False
End Sub

Private Sub Image3_Click()
Image4.Enabled = True
Frame3.Visible = False
Image8.Enabled = False
Combo1.Text = ""
Combo2.Clear
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = ""
Text19.Text = ""
Text23.Text = ""
Label9.Caption = ""
Combo2.AddItem "A"
Combo2.AddItem "B"
Check1.Value = False
Label40.Caption = ""
Text24.Text = ""
Combo1.Visible = True
Combo2.Visible = True
Text20.Visible = False
Text21.Visible = False
End Sub

Private Sub Image4_Click()
Combo1.SetFocus
Image8.Enabled = True
Image5.Enabled = True
Combo1.Locked = False
Combo2.Locked = False
Text1.Locked = False
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
Text5.Locked = False
Text6.Locked = False
Text7.Locked = False
Text8.Locked = False
Text9.Locked = False
Text10.Locked = False
Text11.Locked = False
Text12.Locked = False
Text13.Locked = False
Text14.Locked = False
Text15.Locked = False
Text16.Locked = False
Text17.Locked = False
Text18.Locked = False
Text19.Locked = False
Text23.Locked = False
Frame3.Visible = False
End Sub

Private Sub Image5_Click()
Abc
On Error GoTo CCC
If Combo1.Text = Combo1.Text And Combo2.Text = Combo2.Text And Label40.Caption = 46 Then
MsgBox "Class full"
Else
sql = "insert into student values('" + Label9.Caption + "','" + Combo1.Text + "', '" + Combo2.Text + "', " + Label40.Caption + " , '" + Text1.Text + "', '" + Text2.Text + "', '" + Text3.Text + "', '" + Text4.Text + "', '" + Text5.Text + "', '" + Text6.Text + "', '" + Text7.Text + "','" + Text8.Text + "', '" + Text9.Text + "', '" + Format(Text10.Text, "dd mmm yyyy") + "','" + Text11.Text + "', '" + Text12.Text + "', '" + Text13.Text + "','" + Text23.Text + "', '" + Text14.Text + "', '" + Text15.Text + "', " + Text16.Text + ",'" + Format(Text17.Text, "dd mmm yyyy") + "', '" + Text18.Text + "', '" + Text19.Text + "'," + Label25.Caption + "," + Text24.Text + ")"
Set r = c.Execute(sql)
Label40.Caption = Label40.Caption + 1
MsgBox "Record Saved"
sql = "insert into dues values('" + Combo1.Text + "','" + Combo2.Text + "','" + Label9.Caption + "','" + Text1.Text + "'," + Label46.Caption + "," + Label46.Caption + ",'" + Label47.Caption + "')"
Set r = c.Execute(sql)
Image5.Enabled = False
Combo1.Text = ""
Combo2.Clear


Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = ""
Text19.Text = ""
Text23.Text = ""
Label9.Caption = ""
Check1.Value = False
Label40.Caption = ""
Text24.Text = ""
Combo1.Locked = True
Combo2.Locked = True
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
Text6.Locked = True
Text7.Locked = True
Text8.Locked = True
Text9.Locked = True
Text10.Locked = True
Text11.Locked = True
Text12.Locked = True
Text13.Locked = True
Text14.Locked = True
Text15.Locked = True
Text16.Locked = True
Text17.Locked = True
Text18.Locked = True
Text19.Locked = True
Text23.Locked = True


Exit Sub
CCC:
MsgBox "Exist"
Image5.Enabled = False
Combo1.Text = ""
Combo2.Clear
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = ""
Text19.Text = ""
Text23.Text = ""
Label40.Caption = ""
Text24.Text = ""
End If
End Sub

Private Sub Image6_Click()
MsgBox "Record deleted"
sql = "update student set status = 0 where regid='" + Text22.Text + "'"
Set r = c.Execute(sql)
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = ""
Text19.Text = ""
Text23.Text = ""
Label9.Caption = ""
Text20.Text = ""
Text21.Text = ""
Label40.Caption = ""
Text24.Text = ""
Image8.Enabled = False
End Sub

Private Sub image7_Click()
Abc
sql = "Update student set class= '" + Text20.Text + "' , section= '" + Text21.Text + "', sname = '" + Text1.Text + "',fname ='" + Text2.Text + "', Fmob ='" + Text3.Text + "', Fquali ='" + Text4.Text + "', Foccup = '" + Text5.Text + "', Mname ='" + Text6.Text + "', Mmob= '" + Text7.Text + "', Mquali='" + Text8.Text + "' , Moccup = '" + Text9.Text + "',  Dob ='" + Format(Text10.Text, "dd mmm yyyy") + "' , Gname = '" + Text11.Text + "', Gsr ='" + Text12.Text + "', Gmob= '" + Text13.Text + "', Gender = '" + Text23.Text + "', City ='" + Text14.Text + "', State ='" + Text15.Text + "', Aadharno = '" + Text16.Text + "', Doa = '" + Format(Text17.Text, "dd mmm yyyy") + "', Padd ='" + Text18.Text + "', Cadd= '" + Text19.Text + "',age=" + Text24.Text + " where regid ='" + Text22.Text + "'"
Set r = c.Execute(sql)
MsgBox " Record Updated"
Image7.Enabled = False
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = ""
Text19.Text = ""
Text23.Text = ""
Label9.Caption = ""
Text20.Text = ""
Text21.Text = ""
Text24.Text = ""
Label40.Caption = ""
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
Text6.Locked = True
Text7.Locked = True
Text8.Locked = True
Text9.Locked = True
Text10.Locked = True
Text11.Locked = True
Text12.Locked = True
Text13.Locked = True
Text14.Locked = True
Text15.Locked = True
Text16.Locked = True
Text17.Locked = True
Text18.Locked = True
Text19.Locked = True
Text23.Locked = True
Text24.Locked = True

Text22.SetFocus
Text22.Text = ""
Image8.Enabled = False


End Sub

Private Sub Image8_Click()
Image4.Enabled = False
Image5.Enabled = False
Text23.Visible = True
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
Frame3.Visible = True
Text21.Visible = True
Text20.Visible = True
Combo1.Visible = False
Combo2.Visible = False
Image8.Enabled = False
Combo1.Locked = True
Combo2.Locked = True
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
Text6.Locked = True
Text7.Locked = True
Text8.Locked = True
Text9.Locked = True
Text10.Locked = True
Text11.Locked = True
Text12.Locked = True
Text13.Locked = True
Text14.Locked = True
Text15.Locked = True
Text16.Locked = True
Text17.Locked = True
Text18.Locked = True
Text19.Locked = True
Text23.Locked = True

End Sub



Private Sub Label42_Click()

End Sub



Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
Text10.Text = Format(DateClicked, "dd mmm yyyy")

MonthView1.Visible = False
Text11.SetFocus
calculateage

End Sub

Private Sub MonthView2_DateClick(ByVal DateClicked As Date)
Text17.Text = Format(DateClicked, "dd mmm yyyy")
MonthView2.Visible = False
Text18.SetFocus
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
Text23.Text = "Male"
Text14.SetFocus
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
Text23.Text = "Female"
Text14.SetFocus
End If
End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
Text23.Text = "Others"
Text14.SetFocus
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text10.SetFocus
End If
End Sub

Private Sub Text1_LostFocus()
Text1.Text = StrConv(Text1.Text, vbProperCase)
End Sub



Private Sub Text10_GotFocus()
MonthView1.Visible = True
End Sub







Private Sub Text10_LostFocus()
calculateage
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text12.SetFocus

If (Text11.Text = "") Then
Text11.Text = "Null"
Text12.Text = "Null"
Text13.Text = "XXXXXXXXXX"
End If
End If
End Sub

Private Sub Text11_LostFocus()
Text11.Text = StrConv(Text11.Text, vbProperCase)

End Sub



Private Sub Text12_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text13.SetFocus
End If
End Sub

Private Sub Text12_LostFocus()
Text12.Text = StrConv(Text12.Text, vbProperCase)
End Sub





Private Sub Text13_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text17.SetFocus
End If

End Sub

Private Sub Text13_LostFocus()
If Len(Text13.Text) > 10 Or Len(Text13.Text) < 10 Then
MsgBox "Enter only number 10"
Text13.Text = ""
Text13.SetFocus
End If
End Sub



Private Sub Text14_GotFocus()
If Text23.Text = "" Then
MsgBox "please enter your Gender"
End If
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text15.SetFocus
End If
End Sub

Private Sub Text14_LostFocus()
Text14.Text = StrConv(Text14.Text, vbProperCase)
End Sub



Private Sub Text15_GotFocus()
If Text14.Text = "" Then
MsgBox "please enter your State"
End If
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text16.SetFocus
End If
End Sub

Private Sub Text15_LostFocus()
Text15.Text = StrConv(Text15.Text, vbProperCase)
End Sub



Private Sub Text16_GotFocus()
If Text15.Text = "" Then
MsgBox "please enter your State"
Text15.SetFocus
End If
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.SetFocus
End If
End Sub

Private Sub Text16_LostFocus()
Text16.Text = StrConv(Text16.Text, vbProperCase)
End Sub

Private Sub Text17_GotFocus()
MonthView2.Visible = True
End Sub





Private Sub Text18_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text19.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text3.SetFocus

End If

End Sub

Private Sub Text2_LostFocus()
Text2.Text = StrConv(Text2.Text, vbProperCase)
End Sub

Private Sub Text22_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text22.Text = UCase(Text22.Text)
Command1.SetFocus
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'Text3.Text = UCase(Text3.Text)
'Dim A As String
 'A = Text3.Text
 'Text3.Text = Label20.Caption & A
Text4.SetFocus
End If
End Sub

Private Sub Text3_LostFocus()

If Len(Text3.Text) > 10 Or Len(Text3.Text) < 10 Then
MsgBox "Enter only number 10"
Text3.Text = ""
Text3.SetFocus
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text5.SetFocus
End If
End Sub

Private Sub Text4_LostFocus()
Text4.Text = StrConv(Text4.Text, vbProperCase)
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text6.SetFocus
End If
End Sub

Private Sub Text5_LostFocus()
Text5.Text = StrConv(Text5.Text, vbProperCase)
End Sub


Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text7.SetFocus
End If
End Sub

Private Sub Text6_LostFocus()
Text6.Text = StrConv(Text6.Text, vbProperCase)
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text8.SetFocus
End If
End Sub

Private Sub Text7_LostFocus()
If Len(Text7.Text) > 10 Or Len(Text7.Text) < 10 Then
MsgBox "Enter only number 10"
Text7.Text = ""
Text7.SetFocus
End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text9.SetFocus
End If
End Sub

Private Sub Text8_LostFocus()
Text8.Text = StrConv(Text8.Text, vbProperCase)
End Sub



Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text11.SetFocus
End If
End Sub

Private Sub Text9_LostFocus()
Text9.Text = StrConv(Text9.Text, vbProperCase)
End Sub

Private Sub Timer1_Timer()
Label1.Caption = Format(Now, "yyyy-MM-dd ")
Label2.Caption = Format(Now, "hh:mm:ss AM/PM")
End Sub
