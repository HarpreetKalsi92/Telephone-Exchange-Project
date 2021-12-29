VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFFF&
   Caption         =   "New Connection"
   ClientHeight    =   9645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13200
   ForeColor       =   &H00000080&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9645
   ScaleWidth      =   13200
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text14 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3840
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Residence"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   2160
      TabIndex        =   15
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CheckBox Check7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Abbreviated Dialing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   8760
      TabIndex        =   25
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CheckBox Check6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Call Forwording"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   6960
      TabIndex        =   24
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Conferencing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   5400
      TabIndex        =   23
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Hotline"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4200
      TabIndex        =   22
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "CLI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   3600
      TabIndex        =   21
      Top             =   6960
      Width           =   615
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ISD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   435
      Left            =   2640
      TabIndex        =   20
      Top             =   6840
      Width           =   735
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Govt."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   6120
      TabIndex        =   18
      Top             =   6480
      Width           =   975
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Business"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4800
      TabIndex        =   17
      Top             =   6480
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "PSU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   3840
      TabIndex        =   16
      Top             =   6480
      Width           =   735
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2640
      TabIndex        =   14
      Top             =   5880
      Width           =   1695
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   53
      Top             =   5880
      Width           =   1470
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   52
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   12
      Top             =   4320
      Width           =   2055
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   11
      Top             =   4440
      Width           =   2415
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   10
      Top             =   3720
      Width           =   2055
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   8
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   6
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   3
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Addnew"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   51
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H0080FF80&
      Caption         =   "Return To Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   44
      Top             =   7800
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   43
      Top             =   7800
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Control Button"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6960
      TabIndex        =   42
      Top             =   7440
      Width           =   4695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Last"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   41
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000016&
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   40
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   39
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "First"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   38
      Top             =   7800
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Navigation  Button"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   975
      Left            =   240
      TabIndex        =   37
      Top             =   7440
      Width           =   4095
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "STD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   1800
      TabIndex        =   19
      Top             =   6840
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2640
      TabIndex        =   13
      Top             =   5400
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Frame Purpose 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Purpose"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   615
      Left            =   1800
      TabIndex        =   55
      Top             =   6240
      Width           =   5655
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0FFFF&
      Caption         =   "New Allotted Telephone No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   960
      TabIndex        =   54
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Concessional Group No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   5280
      TabIndex        =   50
      Top             =   6000
      Width           =   2175
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   6360
      TabIndex        =   49
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pin Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   8400
      TabIndex        =   48
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Street/Road"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   8160
      TabIndex        =   47
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0FFFF&
      Caption         =   "PAN/GIR No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   8040
      TabIndex        =   46
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Name of The Joint Applicant,If Any"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   600
      TabIndex        =   45
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Facility Required"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   240
      TabIndex        =   36
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Concessional Group Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   5880
      Width           =   2415
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Category Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   240
      TabIndex        =   34
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Name Of the Customer/Company"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   600
      TabIndex        =   33
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "E-Mail Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   7800
      TabIndex        =   32
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "For Billing Crosspondence Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   480
      TabIndex        =   31
      Top             =   4560
      Width           =   3015
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "City/Distt."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   3480
      TabIndex        =   30
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "House No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   3600
      TabIndex        =   29
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Complete  Postal Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   480
      TabIndex        =   28
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Telephone No. Working"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1080
      TabIndex        =   27
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Name of The Father/Group"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   6840
      TabIndex        =   26
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Form For New Telephone Connection"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   555
      Left            =   2640
      TabIndex        =   0
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As ADODB.Connection
Dim rec As ADODB.Recordset

Private Sub Combo1_click()
If Combo1.Text = "01" Then
Text12.Text = "N-OUT-Gen"
ElseIf Combo1.Text = "02" Then
Text12.Text = "N-OYT-Spl"
ElseIf Combo1.Text = "03" Then
Text12.Text = "N-OYT-SS"
ElseIf Combo1.Text = "04" Then
Text12.Text = "N-OYT-SWS"
ElseIf Combo1.Text = "05" Then
Text12.Text = "N-OYT-G-SE-DOT"
ElseIf Combo1.Text = "06" Then
Text12.Text = "OYT-Gen"
ElseIf Combo1.Text = "07" Then
Text12.Text = "OYT-Spl"
ElseIf Combo1.Text = "08" Then
Text12.Text = "TATKAL"
End If

End Sub

Private Sub Combo2_click()
If Combo2.Text = "01" Then
Text13.Text = "Freedom Fighters"
ElseIf Combo2.Text = "02" Then
Text13.Text = "Gallantry Award Winners"
ElseIf Combo2.Text = "03" Then
Text13.Text = "War Widows"
ElseIf Combo2.Text = "04" Then
Text13.Text = "Disabled Soldiers"
ElseIf Combo2.Text = "05" Then
Text13.Text = "Blind"
ElseIf Combo2.Text = "06" Then
Text13.Text = "Senior Citizens"
ElseIf Combo2.Text = "07" Then
Text13.Text = "Retired DOT Employees"
ElseIf Combo2.Text = "08" Then
Text13.Text = "Serving DOT Employees"
ElseIf Combo2.Text = "09" Then
Text13.Text = "Recognized Educational Institutes"
ElseIf Combo2.Text = "10" Then
Text13.Text = "Orphanages"
ElseIf Combo2.Text = "11" Then
Text13.Text = "Homes for aged,infirm,spastics,handicapped,deaf-dump-mute persons/voluntary,organization for tribal welfare,institutions recorgnized by Govt."
End If
End Sub

Private Sub Command1_Click()

rec.MoveFirst
Text1.Text = rec!cname
Text2.Text = rec!JName
Text3.Text = rec!fgname
Text4.Text = rec!PGno
Text5.Text = rec!WTno
Text6.Text = rec!Hno
Text7.Text = rec!ST
Text8.Text = rec!City
Text9.Text = rec!Pin
Text10.Text = rec!caddress
Text11.Text = rec!EMail
Text12.Text = rec!Cat
Text13.Text = rec!Cgname
Text14.Text = rec!Tno
Combo1.Text = rec!Ccode
Combo2.Text = rec!Cgcode

Dim a As String
a = rec!Purpose
If a = "Residence" Then
Option1.Value = True
End If
If a = "PSU" Then
Option2.Value = True
End If
If a = "Business" Then
Option3.Value = True
End If
If a = "Govt." Then
Option4.Value = True
End If
Dim b1, b2, b3, b4, b5, b6, b7 As String
b1 = rec!STD
b2 = rec!ISD
b3 = rec!CLI
b4 = rec!Hotline
b5 = rec!Conf
b6 = rec!CF
b7 = rec!AD
  If b1 = "Yes" Then
  Check1.Value = 1
   Else
   Check1.Value = 0
End If
If b2 = "Yes" Then
  Check2.Value = 1
   Else
   Check2.Value = 0
End If
If b3 = "Yes" Then
  Check3.Value = 1
   Else
   Check3.Value = 0
End If
If b4 = "Yes" Then
  Check4.Value = 1
   Else
   Check4.Value = 0
End If
If b5 = "Yes" Then
  Check5.Value = 1
   Else
   Check5.Value = 0
End If
If b6 = "Yes" Then
  Check6.Value = 1
   Else
   Check6.Value = 0
End If
If b7 = "Yes" Then
  Check7.Value = 1
   Else
   Check7.Value = 0
End If

End Sub

Private Sub Command2_Click()
rec.MoveNext
If rec.EOF = True Then
rec.MoveLast
End If
Text1.Text = rec!cname
Text2.Text = rec!JName
Text3.Text = rec!fgname
Text4.Text = rec!PGno
Text5.Text = rec!WTno
Text6.Text = rec!Hno
Text7.Text = rec!ST
Text8.Text = rec!City
Text9.Text = rec!Pin
Text10.Text = rec!caddress
Text11.Text = rec!EMail
Text12.Text = rec!Cat
Text13.Text = rec!Cgname
Text14.Text = rec!Tno
Combo1.Text = rec!Ccode
Combo2.Text = rec!Cgcode

Dim a As String
a = rec!Purpose
If a = "Residence" Then
Option1.Value = True
End If
If a = "PSU" Then
Option2.Value = True
End If
If a = "Business" Then
Option3.Value = True
End If
If a = "Govt." Then
Option4.Value = True
End If
Dim b1, b2, b3, b4, b5, b6, b7 As String
b1 = rec!STD
b2 = rec!ISD
b3 = rec!CLI
b4 = rec!Hotline
b5 = rec!Conf
b6 = rec!CF
b7 = rec!AD
  If b1 = "Yes" Then
  Check1.Value = 1
   Else
   Check1.Value = 0
End If
If b2 = "Yes" Then
  Check2.Value = 1
   Else
   Check2.Value = 0
End If
If b3 = "Yes" Then
  Check3.Value = 1
   Else
   Check3.Value = 0
End If
If b4 = "Yes" Then
  Check4.Value = 1
   Else
   Check4.Value = 0
End If
If b5 = "Yes" Then
  Check5.Value = 1
   Else
   Check5.Value = 0
End If
If b6 = "Yes" Then
  Check6.Value = 1
   Else
   Check6.Value = 0
End If
If b7 = "Yes" Then
  Check7.Value = 1
   Else
   Check7.Value = 0
End If


End Sub

Private Sub Command3_Click()
rec.MovePrevious

If rec.BOF = True Then
rec.MoveFirst
End If
Text1.Text = rec!cname
Text2.Text = rec!JName
Text3.Text = rec!fgname
Text4.Text = rec!PGno
Text5.Text = rec!WTno
Text6.Text = rec!Hno
Text7.Text = rec!ST
Text8.Text = rec!City
Text9.Text = rec!Pin
Text10.Text = rec!caddress
Text11.Text = rec!EMail
Text12.Text = rec!Cat
Text13.Text = rec!Cgname
Text14.Text = rec!Tno
Combo1.Text = rec!Ccode
Combo2.Text = rec!Cgcode

Dim a As String
a = rec!Purpose
If a = "Residence" Then
Option1.Value = True
End If
If a = "PSU" Then
Option2.Value = True
End If
If a = "Business" Then
Option3.Value = True
End If
If a = "Govt." Then
Option4.Value = True
End If
Dim b1, b2, b3, b4, b5, b6, b7 As String
b1 = rec!STD
b2 = rec!ISD
b3 = rec!CLI
b4 = rec!Hotline
b5 = rec!Conf
b6 = rec!CF
b7 = rec!AD
  If b1 = "Yes" Then
  Check1.Value = 1
   Else
   Check1.Value = 0
End If
If b2 = "Yes" Then
  Check2.Value = 1
   Else
   Check2.Value = 0
End If
If b3 = "Yes" Then
  Check3.Value = 1
   Else
   Check3.Value = 0
End If
If b4 = "Yes" Then
  Check4.Value = 1
   Else
   Check4.Value = 0
End If
If b5 = "Yes" Then
  Check5.Value = 1
   Else
   Check5.Value = 0
End If
If b6 = "Yes" Then
  Check6.Value = 1
   Else
   Check6.Value = 0
End If
If b7 = "Yes" Then
  Check7.Value = 1
   Else
   Check7.Value = 0
End If

End Sub

Private Sub Command4_Click()
rec.MoveLast
Text1.Text = rec!cname
Text2.Text = rec!JName
Text3.Text = rec!fgname
Text4.Text = rec!PGno
Text5.Text = rec!WTno
Text6.Text = rec!Hno
Text7.Text = rec!ST
Text8.Text = rec!City
Text9.Text = rec!Pin
Text10.Text = rec!caddress
Text11.Text = rec!EMail
Text12.Text = rec!Cat
Text13.Text = rec!Cgname
Text14.Text = rec!Tno
Combo1.Text = rec!Ccode
Combo2.Text = rec!Cgcode

Dim a As String
a = rec!Purpose
If a = "Residence" Then
Option1.Value = True
End If
If a = "PSU" Then
Option2.Value = True
End If
If a = "Business" Then
Option3.Value = True
End If
If a = "Govt." Then
Option4.Value = True
End If
Dim b1, b2, b3, b4, b5, b6, b7 As String
b1 = rec!STD
b2 = rec!ISD
b3 = rec!CLI
b4 = rec!Hotline
b5 = rec!Conf
b6 = rec!CF
b7 = rec!AD
  If b1 = "Yes" Then
  Check1.Value = 1
   Else
   Check1.Value = 0
End If
If b2 = "Yes" Then
  Check2.Value = 1
   Else
   Check2.Value = 0
End If
If b3 = "Yes" Then
  Check3.Value = 1
   Else
   Check3.Value = 0
End If
If b4 = "Yes" Then
  Check4.Value = 1
   Else
   Check4.Value = 0
End If
If b5 = "Yes" Then
  Check5.Value = 1
   Else
   Check5.Value = 0
End If
If b6 = "Yes" Then
  Check6.Value = 1
   Else
   Check6.Value = 0
End If
If b7 = "Yes" Then
  Check7.Value = 1
   Else
   Check7.Value = 0
End If

End Sub



Private Sub Command6_Click()
Dim a As Integer
rec.MoveFirst
flag = False
 While Not rec.EOF = True
  If Text14.Text = rec!Tno Then
    flag = True
    MsgBox "This Telephone number has already Allotted to Someone else."
    MsgBox "Please Choose another Number, Thank you."
    Text14.Text = ""
    Text14.SetFocus
 End If
    rec.MoveNext
Wend
If flag = False Then

If Trim(Text1.Text) = "" Or Trim(Text2.Text) = "" Or Trim(Text3.Text) = "" Or Trim(Text4.Text) = "" Or Trim(Text5.Text) = "" Or Trim(Text6.Text) = "" Or Trim(Text7.Text) = "" Or Trim(Text8.Text) = "" Or Trim(Text9.Text) = "" Or Trim(Text10.Text) = "" Or Trim(Text11.Text) = "" Or Trim(Text12.Text) = "" Or Trim(Text13.Text) = "" Or Trim(Text14.Text) = "" Then
MsgBox "No one field should be blank", vbCritical
Text14.SetFocus
Else
rec.AddNew
rec!cname = Text1.Text
rec!JName = Text2.Text
rec!fgname = Text3.Text
rec!PGno = Text4.Text
rec!WTno = Text5.Text
rec!Hno = Text6.Text
rec!ST = Text7.Text
rec!City = Text8.Text
rec!Pin = Text9.Text
rec!caddress = Text10.Text
rec!EMail = Text11.Text
rec!Ccode = Combo1.Text
rec!Cat = Text12.Text
rec!Tno = Text14.Text
rec!Cgcode = Combo2.Text
rec!Cgname = Text13.Text
rec!dis = "no"
If Option1.Value = True Then
rec!Purpose = "Residence"
End If
If Option2.Value = True Then
rec!Purpose = "PSU"
End If
If Option3.Value = True Then
rec!Purpose = "Business"
End If
If Option4.Value = True Then
rec!Purpose = "Govt."
End If
If Check1.Value = 1 Then
rec!STD = "Yes"
Else
rec!STD = "no"
End If

If Check2.Value = 1 Then
rec!ISD = "Yes"
Else
rec!ISD = "no"
End If

If Check3.Value = 1 Then
rec!CLI = "Yes"
Else
rec!CLI = "no"
End If

If Check4.Value = 1 Then
rec!Hotline = "Yes"
Else
rec!Hotline = "no"
End If

If Check5.Value = 1 Then
rec!Conf = "Yes"
Else
rec!Conf = "no"
End If

If Check6.Value = 1 Then
rec!CF = "Yes"
Else
rec!CF = "no"
End If

If Check7.Value = 1 Then
rec!AD = "Yes"
Else
rec!AD = "no"
End If
MsgBox "Record Saved"
rec.Update
End If
End If

End Sub







Private Sub Command8_Click()
c = MsgBox("Are you sure to exit this program?", vbYesNo, "Confirm Box")
If c = vbYes Then
form13.Show
Unload Me
End If
End Sub

Private Sub Command9_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
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
Combo1.Text = "Select Code"
Combo2.Text = "Select Code"

Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False

Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
Check4.Value = 0
Check5.Value = 0
Check6.Value = 0
Check7.Value = 0

Text14.SetFocus
End Sub

Private Sub Form_Load()
Set db = New ADODB.Connection
Set rec = New ADODB.Recordset
db.ConnectionString = "dsn=phone;uid=;pwd=;"
db.Provider = "Microsoft.jet.Oledb.4.0"
db.Open App.Path & "\Telephone.mdb"
rec.Open "newconn1", db, adOpenKeyset, adLockPessimistic

Combo1.AddItem "01"
Combo1.AddItem "02"
Combo1.AddItem "03"
Combo1.AddItem "04"
Combo1.AddItem "05"
Combo1.AddItem "06"
Combo1.AddItem "07"
Combo1.AddItem "08"
Combo2.AddItem "01"
Combo2.AddItem "02"
Combo2.AddItem "03"
Combo2.AddItem "04"
Combo2.AddItem "05"
Combo2.AddItem "06"
Combo2.AddItem "07"
Combo2.AddItem "08"
rec!dis = "No"
End Sub

