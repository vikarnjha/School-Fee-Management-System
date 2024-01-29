VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Feepayment 
   BackColor       =   &H00FFFFFF&
   Caption         =   "FEE PAYMENT"
   ClientHeight    =   9615
   ClientLeft      =   1935
   ClientTop       =   1245
   ClientWidth     =   17805
   FillColor       =   &H00FFFFC0&
   ForeColor       =   &H00FFFFC0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9615
   ScaleWidth      =   17805
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8895
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   16935
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
         Height          =   405
         Left            =   11760
         TabIndex        =   55
         Top             =   4920
         Width           =   2055
      End
      Begin VB.TextBox Text14 
         Height          =   375
         Left            =   6960
         TabIndex        =   53
         Text            =   "Text14"
         Top             =   5040
         Width           =   1455
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2100
         Left            =   120
         TabIndex        =   52
         Top             =   4560
         Width           =   1575
      End
      Begin VB.TextBox Text13 
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
         Left            =   4560
         TabIndex        =   51
         Text            =   "Text13"
         Top             =   3840
         Width           =   1695
      End
      Begin VB.TextBox Text12 
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
         Left            =   4560
         TabIndex        =   49
         Text            =   "Text12"
         Top             =   3240
         Width           =   1695
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
         Height          =   375
         Left            =   11760
         TabIndex        =   47
         Top             =   5400
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "Pay"
         DownPicture     =   "Feepayment.frx":0000
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   8040
         Width           =   2415
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2265
         Left            =   5280
         TabIndex        =   42
         Top             =   6120
         Width           =   2535
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   2520
         Style           =   1  'Checkbox
         TabIndex        =   41
         Top             =   6120
         Width           =   2535
      End
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2370
         Left            =   13920
         TabIndex        =   40
         Top             =   3720
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
         StartOfWeek     =   157810690
         CurrentDate     =   45287
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   9240
         TabIndex        =   33
         Top             =   6000
         Width           =   5295
         Begin VB.TextBox Text11 
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
            Left            =   2400
            TabIndex        =   39
            Text            =   "Text11"
            Top             =   1680
            Width           =   2655
         End
         Begin VB.OptionButton Option6 
            BackColor       =   &H80000016&
            Caption         =   "UPI NO"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   3840
            TabIndex        =   37
            Top             =   840
            Width           =   1335
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H80000016&
            Caption         =   "CHEQUE "
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2040
            TabIndex        =   36
            Top             =   960
            Width           =   1335
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H80000016&
            Caption         =   "CASH"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   35
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
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
            Left            =   120
            TabIndex        =   38
            Top             =   1800
            Width           =   2175
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Methods"
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
            Left            =   1680
            TabIndex        =   34
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.TextBox Text10 
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
         Left            =   11760
         TabIndex        =   29
         Top             =   4440
         Width           =   2055
      End
      Begin VB.TextBox Text9 
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
         Left            =   11760
         TabIndex        =   28
         Top             =   3840
         Width           =   1935
      End
      Begin VB.TextBox Text8 
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
         Left            =   11760
         TabIndex        =   26
         Top             =   3240
         Width           =   1935
      End
      Begin VB.TextBox Text7 
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
         Left            =   11760
         TabIndex        =   24
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox Text6 
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
         Left            =   11760
         TabIndex        =   22
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox Text5 
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
         Left            =   4560
         TabIndex        =   20
         Top             =   5040
         Width           =   1815
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H80000016&
         Caption         =   "Yearly"
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
         Left            =   7800
         TabIndex        =   14
         Top             =   4440
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H80000016&
         Caption         =   "Quaterly "
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
         Left            =   6240
         TabIndex        =   13
         Top             =   4440
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000016&
         Caption         =   "Monthly"
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
         Left            =   4560
         TabIndex        =   12
         Top             =   4440
         Width           =   1575
      End
      Begin VB.TextBox Text4 
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
         Left            =   4560
         TabIndex        =   10
         Top             =   2640
         Width           =   3135
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
         Height          =   405
         Left            =   4560
         TabIndex        =   8
         Top             =   2040
         Width           =   3135
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   975
         Left            =   2400
         TabIndex        =   2
         Top             =   960
         Width           =   14055
         Begin VB.ComboBox Combo3 
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
            Left            =   6840
            TabIndex        =   7
            Top             =   360
            Width           =   1575
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
            Left            =   3720
            TabIndex        =   6
            Text            =   "Combo2"
            Top             =   360
            Width           =   1215
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
            Left            =   1200
            TabIndex        =   5
            Text            =   "Combo1"
            Top             =   360
            Width           =   1095
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
            Height          =   495
            Left            =   11520
            TabIndex        =   3
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label9 
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
            Left            =   8760
            TabIndex        =   18
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label Label8 
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
            Left            =   5160
            TabIndex        =   17
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label7 
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
            Height          =   255
            Left            =   2520
            TabIndex        =   16
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label6 
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
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9360
         TabIndex        =   54
         Top             =   4920
         Width           =   1215
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Fine"
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
         Left            =   9360
         TabIndex        =   50
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label22 
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
         Height          =   375
         Left            =   1920
         TabIndex        =   48
         Top             =   5040
         Width           =   1335
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Date"
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
         Left            =   9360
         TabIndex        =   46
         Top             =   5400
         Width           =   2055
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Fee Amount"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   44
         Top             =   5640
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fee Type"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   43
         Top             =   5640
         Width           =   1935
      End
      Begin VB.Image Image5 
         Height          =   1680
         Left            =   0
         Picture         =   "Feepayment.frx":0A57
         Top             =   0
         Width           =   2130
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Dues"
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
         Left            =   9360
         TabIndex        =   32
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label Label16 
         Caption         =   "Label16"
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
         Left            =   13200
         TabIndex        =   31
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Advance"
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
         Left            =   9360
         TabIndex        =   30
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Payable Amount"
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
         Left            =   9360
         TabIndex        =   27
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label13 
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
         Height          =   375
         Left            =   9360
         TabIndex        =   25
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Dues"
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
         Left            =   1920
         TabIndex        =   23
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Type"
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
         Left            =   1920
         TabIndex        =   21
         Top             =   4440
         Width           =   2175
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Advance"
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
         Left            =   1920
         TabIndex        =   19
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Father Name"
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
         Left            =   1920
         TabIndex        =   11
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Left            =   1920
         TabIndex        =   9
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Transcition Id"
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
         Left            =   11400
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
      Begin VB.Line Line1 
         X1              =   7200
         X2              =   8640
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7320
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Image Image4 
      Height          =   1680
      Left            =   480
      Picture         =   "Feepayment.frx":20FD
      Top             =   480
      Width           =   2130
   End
   Begin VB.Image Image3 
      Height          =   1935
      Left            =   480
      Top             =   480
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   1680
      Left            =   480
      Picture         =   "Feepayment.frx":37A3
      Top             =   480
      Width           =   2130
   End
   Begin VB.Image Image1 
      Height          =   13545
      Left            =   0
      Picture         =   "Feepayment.frx":4E49
      Top             =   0
      Width           =   24510
   End
End
Attribute VB_Name = "Feepayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim isfirstpayment As Boolean
Dim isyearlypaid As Boolean
Dim selectedpayment As String
Dim lastpaymentdate As Date




Private Sub Combo1_Click()
List1.Clear
List2.Clear

sql = "select * from fee_structure where class='" + Combo1.Text + "'"
Set r = c.Execute(sql)
 Do While (r.EOF = False)
List1.AddItem r.Fields("FEETYPE").Value
List2.AddItem r.Fields("Amount").Value
r.MoveNext
Loop
Combo2.Clear
sql = "select distinct section from student where class='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Do While Not r.EOF
Combo2.AddItem r.Fields(0)
r.MoveNext
Loop
End Sub

Private Sub Combo2_Click()
Combo3.Clear
sql = "select distinct roll from student where class='" + Combo1.Text + "' and section='" + Combo2.Text + "'"
Set r = c.Execute(sql)
Do While Not r.EOF
Combo3.AddItem r.Fields(0)
r.MoveNext
Loop
End Sub





Private Sub Combo3_Click()
sql = "select regid from student where class='" + Combo1.Text + "' and section='" + Combo2.Text + "' and roll='" + Combo3.Text + "'"
Set r = c.Execute(sql)
Do While Not r.EOF
Text1.Text = r.Fields(0)
r.MoveNext
Loop
Text1.SetFocus
Dim studentID As String
    Dim transactionID As String
    
    ' Replace "ABC123" with the actual student ID
    studentID = "PRJ123"
    
    ' Generate the transaction ID
    transactionID = GenerateTransactionID(studentID)
    
    ' Set the transaction ID as the Caption of the Label control
    Label16.Caption = transactionID
End Sub


Private Sub Command1_Click()
    ' Check which option button is selected
    If Option1.Value = True Then
        PayMonthly
    ElseIf Option2.Value = True Then
        PayQuarterly
    ElseIf Option3.Value = True Then
        PayYearly
    Else
        MsgBox "Please select a payment frequency.", vbExclamation, "Payment Error"
    End If
End Sub

Private Sub PayMonthly()
    ' Assuming you have a database connection object named "YourDatabaseConnection"
    Abc
sql = "insert into payment values('" + Label16.Caption + "','" + Combo1.Text + "','" + Combo2.Text + "'," + Combo3.Text + ",'" + Text1.Text + "','" + Text3.Text + "','" + Text4.Text + "','" + Text14.Text + "'," + Text5.Text + "," + Text6.Text + "," + Text7.Text + "," + Text8.Text + "," + Text9.Text + "," + Text10.Text + ",'" + Format(Text2.Text, "dd mmm yyyy") + "','" + Text11.Text + "')"
Set r = c.Execute(sql)
UpdateDuesTable
    ' Update lastPaymentDate and set isFirstPayment accordingly
    lastpaymentdate = Date
    isfirstpayment = False

    ' Disable/enable options and update the listbox as needed
    EnablePaymentOptions
    updatefeetypelistbox

    MsgBox "Monthly payment processed.", vbInformation, "Payment Success"
End Sub

Private Sub PayQuarterly()
    ' Assuming you have a database connection object named "YourDatabaseConnection"
  Abc
    sql = "insert into payment values('" + Label16.Caption + "','" + Combo1.Text + "','" + Combo2.Text + "'," + Combo3.Text + ",'" + Text1.Text + "','" + Text3.Text + "','" + Text4.Text + "','" + Text14.Text + "'," + Text5.Text + "," + Text6.Text + "," + Text7.Text + "," + Text8.Text + "," + Text9.Text + "," + Text10.Text + ",'" + Format(Text2.Text, "dd mmm yyyy") + "','" + Text11.Text + "')"
Set r = c.Execute(sql)
UpdateDuesTable
    ' Update lastPaymentDate and set isFirstPayment accordingly
    lastpaymentdate = Date
    isfirstpayment = False

    ' Disable/enable options and update the listbox as needed
    EnablePaymentOptions
    updatefeetypelistbox

    MsgBox "Quarterly payment processed.", vbInformation, "Payment Success"
End Sub

Private Sub PayYearly()
    ' Assuming you have a database connection object named "YourDatabaseConnection"
    sql = "insert into payment values('" + Label16.Caption + "','" + Combo1.Text + "','" + Combo2.Text + "'," + Combo3.Text + ",'" + Text1.Text + "','" + Text3.Text + "','" + Text4.Text + "','" + Text14.Text + "'," + Text5.Text + "," + Text6.Text + "," + Text7.Text + "," + Text8.Text + "," + Text9.Text + "," + Text10.Text + ",'" + Format(Text2.Text, "dd mmm yyyy") + "','" + Text11.Text + "'"
Set r = c.Execute(sql)
        UpdateDuesTable


    ' Update lastPaymentDate and set isFirstPayment accordingly
    lastpaymentdate = Date
    isfirstpayment = False
    isyearlypaid = True

    ' Disable/enable options and update the listbox as needed
    EnablePaymentOptions
    updatefeetypelistbox

    MsgBox "Yearly payment processed.", vbInformation, "Payment Success"
If Text13.Text = "0.00" Then
isyearlypaid = True
MsgBox "YEARLY PAYMENT SUCCESSFULL", vbInformation, "PAYMENT STATUS"
Else
MsgBox "CLEAR ALL DUES AND ADVANCE BEFORE MAKING YEARLY PAYMENT", vbExclamation, "PAYMENT STATUS"
End If
End Sub
Private Sub UpdateDuesTable()
Abc
sql = "SELECT * FROM Dues WHERE RegID = '" & Text1.Text & "'"
Set r = c.Execute(sql)
If Not r.EOF Then
sql = "update dues set advance ='" + Text9.Text + "',dues='" + Text10.Text + "',remarks='" + Text15.Text + "' where regid ='" + Text1.Text + "'"
Set r = c.Execute(sql)
End If
End Sub
Private Function GenerateTransactionID(studentID As String) As String
    Dim currentDate As String
    Dim randomNum As Integer
    Dim transactionID As String
    
    ' Format the current date and time as YYYYMMDDHHMMSS
    currentDate = Format(Now, "YYYY")
    
    ' Generate a random number between 1000 and 9999
    Randomize
    randomNum = Int((9999 - 1000 + 1) * Rnd + 1000)
    
    ' Combine the components to create a unique transaction ID
    transactionID = currentDate & "-" & Combo1.Text & "-" & studentID & "-" & CStr(randomNum)
    
    ' Return the generated transaction ID
    GenerateTransactionID = transactionID
End Function
    
Private Sub Form_Load()
MonthView1.Visible = False
Abc
sql = "select distinct  class from class_detail order by class"
Set r = c.Execute(sql)
Do While Not r.EOF
Combo1.AddItem r.Fields(0)
r.MoveNext
Loop
'Combo2.AddItem "A"
'Combo2.AddItem "B"

End Sub



Private Sub List1_Click()
calculatetotal
If Text14.Text = "" Then
MsgBox "First choose payment month"
End If
End Sub
Private Sub calculatetotal()
Dim Total As Double
Dim i As Integer
Dim frequencymultiplier As Integer
If Option1.Value = True Then
frequencymultiplier = 1
ElseIf Option2.Value = True Then
frequencymultiplier = 3
ElseIf Option3.Value = True Then
frequencymultiplier = 12
Else
frequencymultiplier = 1
End If
Total = 0
For i = 0 To List1.ListCount - 1
If List1.Selected(i) Then
Total = Total + Val(List2.List(i)) * frequencymultiplier
End If
Next i
Text5.Text = Format(Total, "###0.00")

End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
Text2.Text = Format(DateClicked, "dd mmm yyyy")
MonthView1.Visible = False

End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
Text14.Text = "Monthly"
List3.Clear
List3.AddItem "April"
List3.AddItem "May"
List3.AddItem "June"
List3.AddItem "July"
List3.AddItem "August"
List3.AddItem "September"
List3.AddItem "October"
List3.AddItem "November"
List3.AddItem "December"
List3.AddItem "January"
List3.AddItem "February"
List3.AddItem "March"
End If
selectedpayment = "Monthly"
EnablePaymentOptions
updatefeetypelistbox
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
Text14.Text = "Quaterly"
List3.Clear
List3.AddItem "April-June"
List3.AddItem "July-Sept"
List3.AddItem "Oct-Dec"
List3.AddItem "Jan-March"
End If
selectedpayment = "Quaterly"
EnablePaymentOptions
updatefeetypelistbox
End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
Text14.Text = "Yearly"
List3.Clear
List3.AddItem "April-March"

End If
selectedpayment = "Yearly"
EnablePaymentOptions
updatefeetypelistbox
End Sub
Private Sub EnablePaymentOptions()
If isfirstpayment Then
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Else
Option1.Enabled = (selectedpayment <> "Monthly" Or Isnextmonthallowed())
Option2.Enabled = (selectedpayment <> "Quaterly" Or Isnextquaterlyallowed())
Option3.Enabled = (selectedpayment <> "Yearly" Or isyearlypaid And ISFULLPAYMENTREQUIRED())
End If
End Sub
Private Function Isnextmonthallowed() As Boolean
Dim today As Date
today = Date
If today > DateAdd("m", 1, lastpaymentdate) Then
Isnextmonthallowed = True
Else
Isnextmonthallowed = False
End If
End Function
Private Function Isnextquaterlyallowed() As Boolean
Dim today As Date
today = Date
If today > DateAdd("m", 3, lastpaymentdate) Then
Isnextquaterlyallowed = True
Else
Isnextquaterlyallowed = False
End If
End Function

Private Sub updatefeetypelistbox()
If Option3.Value = True Or Option1.Value = True Or Option2.Value = True Then
DisableCertainFeeTypesForYearly
Else
ENABLEALLFEETYPES
End If
End Sub


Private Sub DisableCertainFeeTypesForYearly()
    ' This is a placeholder example; replace it with your actual logic

    ' Assuming you have a list of fee types that should not be allowed for yearly payments
    Dim excludedFeeTypes As New Collection
    excludedFeeTypes.Add "Admission"
    excludedFeeTypes.Add "Registration"
    excludedFeeTypes.Add "Examination"

    ' Iterate through the fee types in the ListBox and disable certain ones
    For i = 0 To List1.ListCount - 1
        Dim feeType As String
        feeType = List1.List(i)

        ' Check if the fee type is in the excluded list
        If IsFeeTypeExcluded(feeType, excludedFeeTypes) Then
            List1.Selected(i) = False
        End If
    Next i
    For i = 0 To List3.ListCount - 1
If FEETYPETREQUIRESYEARLYPAYMENT(List3.List(i)) Then
List3.List(i) = False
End If
Next i
End Sub

Private Function IsFeeTypeExcluded(feeType As String, excludedFeeTypes As Collection) As Boolean
    ' Check if the fee type is in the excluded list
    On Error Resume Next
    IsFeeTypeExcluded = Not IsEmpty(excludedFeeTypes(feeType))
    On Error GoTo 0
End Function
Private Sub ENABLEALLFEETYPES()
For i = 0 To List3.ListCount - 1
List3.List(i) = True
Next i
End Sub
Private Function FEETYPETREQUIRESYEARLYPAYMENT(feeType As String) As Boolean
Abc
 Dim sql As String
    sql = "SELECT amount FROM Fee_Structure WHERE FeeType = '" & feeType & "' AND Class = '" & Combo1.Text & "'"
    Set r = c.Execute(sql)
    Dim yearlypayment As Boolean
   If Not r.EOF And Not r.BOF Then
    yearlypayment = r.Fields("amount").Value
    FEETYPETREQUIRESYEARLYPAYMENT = yearlypayment
    FEETYPEREQUIRESYEARLYPAYMENT = True
    Else
    yearlypayment = False
    End If

End Function
Private Function ISFULLPAYMENTREQUIRED() As Boolean
If Text6.Text = "0.00" Then
ISFULLPAYMENTREQUIRED = Not isyearlypaid
Else
ISFULLPAYMENTREQUIRED = False
End If
End Function

Private Sub HANDLEFINE()
Dim FINEAMOUNT As Double
FINEAMOUNT = CDbl(Text6.Text)
If FINEAMOUNT > 0 Then
Text7.Text = Format(CDbl(Text7.Text), FINEAMOUNT, "0.00")
MsgBox "Fine of $" & FINEAMOUNT & " added to payable amount", vbInformation, "Fine Added"
Else
MsgBox "No fine added", vbInformation
End If
End Sub
Private Sub Option4_Click()
If Option4.Value = True Then
Text11.Text = "Cash"
Label19.Visible = False
End If
End Sub

Private Sub Option5_Click()
If Option5.Value = True Then
Label19.Visible = True
Label19.Caption = "Cheque Number"
Text11.Text = ""

End If
End Sub

Private Sub Option6_Click()
If Option6.Value = True Then
Label19.Visible = True
Label19.Caption = "UPI Number"
Text11.Text = ""

End If

End Sub

Private Sub Text1_GotFocus()
If Combo1.ListIndex >= 0 And Combo2.ListIndex >= 0 And Combo3.ListIndex >= 0 Then
        ' Display a message when the textbox gets focus
        MsgBox "Class: " & Combo1.Text & vbCrLf & "Section: " & Combo2.Text & vbCrLf & "Roll: " & Combo3.Text, vbInformation, "Selected Values"
Dim selectedClass As String
Dim selectedSection As String
Dim selectedRoll As String
selectedClass = Combo1.List(Combo1.ListIndex)
selectedSection = Combo2.List(Combo2.ListIndex)
selectedRoll = Combo3.List(Combo3.ListIndex)

Dim regId As String
regId = Getpredefined(selectedClass, selectedSection, selectedRoll)
Text1.Text = regId
If isnewEntry(regId) Then
Dim studentinfo As Object
Set studentinfo = getstudentinfo(regId)
MsgBox "New Entry" & CStr(regId), vbInformation, "New Student Information"
End If

If haspayments(regId) Then
MsgBox "Payments found for Student " & regId, vbInformation, "Student Information"
Else
MsgBox " No Payments found for Student " & regId, vbInformation, "Payment Status"
End If

Else

        ' Display a message if class, section, or roll is not selected
        MsgBox "Please select Class, Section, and Roll before entering additional information.", vbExclamation, "Selection Required"
    End If



sql = "select * from student where regid='" + Text1.Text + "'"
Set r = c.Execute(sql)
Do While Not r.EOF
Text3.Text = r.Fields("sname")
Text4.Text = r.Fields("fname")
r.MoveNext
Loop
sql = "select * from dues where regid='" + Text1.Text + "'"
Set r = c.Execute(sql)
Do While Not r.EOF
Text12.Text = r.Fields("Advance")
Text12.Text = Format(Text12.Text, "###0.00")
Text13.Text = r.Fields("Dues")
Text13.Text = Format(Text13.Text, "###0.00")
r.MoveNext
Loop
End Sub
Private Function Getpredefined(selectedClass As String, selectedSection As String, selectedRoll As String) As String
Abc
sql = "select regid from student where class   = '" & selectedClass & "' AND Section = '" & selectedSection & "' AND Roll = '" & selectedRoll & "'"
Set r = c.Execute(sql)
If Not r.EOF Then
Getpredefined = r.Fields("regid").Value
Else
getpredifined = ""
End If

End Function
Private Function isnewEntry(regId As String) As Boolean
Abc
sql = "select * from student where regid='" + Text1.Text + "'"
Set r = c.Execute(sql)
isnewEntry = r.EOF
End Function
Private Function getstudentinfo(regId As String) As Object
Dim r As Object

sql = "SELECT sname,fname FROM student where regid='" + Text1.Text + "'"
Set r = c.Execute(sql)
If Not r.EOF Then
        Set getstudentinfo = "Name: " & r.Fields("sname").Value & vbCrLf & "Class: " & r.Fields("Fname").Value ' Adjust fields accordingly
    Else
        Set getstudentinfo = Nothing
    End If
End Function
Private Function haspayments(regId As String) As String
Abc
sql = "select * from payment where regid='" + Text1.Text + "'"
Set r = c.Execute(sql)
haspayments = Not r.EOF
End Function


Private Sub Text2_GotFocus()
MonthView1.Visible = True
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text6.Text = Format(Text6.Text, "###0.00")
Dim value1 As Double
Dim value2 As Double
Dim value3 As Double
Dim value4 As Double
If IsNumeric(Text5.Text) Then
value1 = CDbl(Text5.Text)
Else
MsgBox "Invalid input in text5"
Exit Sub
End If
If IsNumeric(Text6.Text) Then
value2 = CDbl(Text6.Text)
Else
MsgBox "Invalid input in text6"
Exit Sub
End If
If IsNumeric(Text12.Text) Then
value3 = CDbl(Text12.Text)
Else
MsgBox "Invalid input in text12"
Exit Sub
End If
If IsNumeric(Text13.Text) Then
value4 = CDbl(Text13.Text)
Else
MsgBox "Invalid input in text14"
Exit Sub
End If
Dim result As Double
If (Text12.Text = "0.00" And Text13.Text <> "0.00") Then
result = value1 + value2 + value3 + value4
ElseIf (Text12.Text <> "0.00" And Text13.Text = "0.00") Then
result = value1 + value2 + value4 - value3
Else
result = value1 + value2
End If
Text7.Text = Format(result, "###0.00")
keyasii = 0
End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Dim value4 As Double
If IsNumeric(Text8.Text) Then
value4 = CDbl(Text8.Text)
Else
MsgBox "Invalid input in text8"
Exit Sub
End If
Text8.Text = Format(value4, "###0.00")
Dim intitalAmount As Double
Dim userinput As Double
Dim advanceAmount As Double
Dim duesAmount As Double
On Error Resume Next
intitalAmount = CDbl(Text7.Text)
userinput = CDbl(Text8.Text)
On Error GoTo 0
If userinput > intitalAmount Then
advanceAmount = userinput - intitalAmount
duesAmount = 0
Text15.Text = "Advance present"
ElseIf userinput < intitalAmount Then
advanceAmount = 0
duesAmount = intitalAmount - userinput
 Text15.Text = "Dues present"
Else
advanceAmount = 0
duesAmount = 0
Text15.Text = "No dues and No present"
End If
Text9.Text = Format(advanceAmount, "###0.00")
Text10.Text = Format(duesAmount, "###0.00")
End If

End Sub
