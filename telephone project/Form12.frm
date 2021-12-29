VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0080FFFF&
   Caption         =   "New Connection"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.OptionButton Option1 
      BackColor       =   &H0080FFFF&
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
      Height          =   255
      Left            =   2040
      TabIndex        =   54
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CheckBox Check7 
      BackColor       =   &H0080FFFF&
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
      Height          =   255
      Left            =   8760
      TabIndex        =   53
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CheckBox Check6 
      BackColor       =   &H0080FFFF&
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
      Height          =   255
      Left            =   6960
      TabIndex        =   52
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H0080FFFF&
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
      Height          =   255
      Left            =   5400
      TabIndex        =   51
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H0080FFFF&
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
      Height          =   255
      Left            =   4200
      TabIndex        =   50
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H0080FFFF&
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
      Height          =   435
      Left            =   3480
      TabIndex        =   49
      Top             =   6840
      Width           =   615
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H0080FFFF&
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
      Height          =   435
      Left            =   2640
      TabIndex        =   48
      Top             =   6840
      Width           =   735
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H0080FFFF&
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
      Height          =   255
      Left            =   5640
      TabIndex        =   47
      Top             =   6480
      Width           =   975
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H0080FFFF&
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
      Height          =   255
      Left            =   4440
      TabIndex        =   46
      Top             =   6480
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H0080FFFF&
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
      Height          =   255
      Left            =   3600
      TabIndex        =   45
      Top             =   6480
      Width           =   735
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H0080C0FF&
      Height          =   315
      Left            =   2640
      TabIndex        =   44
      Top             =   5880
      Width           =   1695
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H0080C0FF&
      Height          =   375
      Left            =   7560
      TabIndex        =   43
      Top             =   5880
      Width           =   1470
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H0080C0FF&
      Height          =   375
      Left            =   6720
      TabIndex        =   42
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H0080C0FF&
      Height          =   375
      Left            =   3360
      TabIndex        =   41
      Top             =   4800
      Width           =   2535
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H0080C0FF&
      Height          =   375
      Left            =   3360
      TabIndex        =   40
      Top             =   4320
      Width           =   2775
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H0080C0FF&
      Height          =   375
      Left            =   6240
      TabIndex        =   39
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H0080C0FF&
      Height          =   375
      Left            =   2640
      TabIndex        =   38
      Text            =   "fdss"
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H0080C0FF&
      Height          =   375
      Left            =   6120
      TabIndex        =   37
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H0080C0FF&
      Height          =   375
      Left            =   2640
      TabIndex        =   36
      Text            =   "r34"
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H0080C0FF&
      Height          =   375
      Left            =   3000
      TabIndex        =   35
      Text            =   "4345"
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H0080C0FF&
      Height          =   375
      Left            =   7920
      TabIndex        =   34
      Text            =   "4564"
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H0080C0FF&
      Height          =   375
      Left            =   3000
      TabIndex        =   33
      Text            =   "hhhd"
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H0080C0FF&
      Height          =   375
      Left            =   9120
      TabIndex        =   32
      Text            =   "eer"
      Top             =   600
      Width           =   1695
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
      Left            =   7080
      TabIndex        =   31
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
      Left            =   9960
      TabIndex        =   24
      Top             =   7800
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Exit"
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
      Left            =   9000
      TabIndex        =   23
      Top             =   7800
      Width           =   975
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
      Left            =   8040
      TabIndex        =   22
      Top             =   7800
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FFFF&
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
      TabIndex        =   21
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
      TabIndex        =   20
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton Command3 
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
      TabIndex        =   19
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
      TabIndex        =   18
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
      TabIndex        =   17
      Top             =   7800
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
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
      Height          =   975
      Left            =   240
      TabIndex        =   16
      Top             =   7320
      Width           =   4095
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0080FFFF&
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
      Height          =   375
      Left            =   1800
      TabIndex        =   15
      Top             =   6840
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H0080C0FF&
      Height          =   315
      Left            =   2640
      TabIndex        =   12
      Top             =   5400
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0080C0FF&
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Text            =   "eee"
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label19 
      BackColor       =   &H0080FFFF&
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
      Height          =   255
      Left            =   5280
      TabIndex        =   30
      Top             =   6000
      Width           =   2175
   End
   Begin VB.Label Label18 
      BackColor       =   &H0080FFFF&
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
      Height          =   255
      Left            =   5280
      TabIndex        =   29
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label Label17 
      BackColor       =   &H0080FFFF&
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
      Height          =   255
      Left            =   4920
      TabIndex        =   28
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label16 
      BackColor       =   &H0080FFFF&
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
      Height          =   255
      Left            =   4800
      TabIndex        =   27
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label15 
      BackColor       =   &H0080FFFF&
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
      Height          =   255
      Left            =   5880
      TabIndex        =   26
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label14 
      BackColor       =   &H0080FFFF&
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
      Height          =   255
      Left            =   5880
      TabIndex        =   25
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label Label13 
      BackColor       =   &H0080FFFF&
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
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackColor       =   &H0080FFFF&
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
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label Label11 
      BackColor       =   &H0080FFFF&
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
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   5880
      Width           =   2415
   End
   Begin VB.Label Label10 
      BackColor       =   &H0080FFFF&
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
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
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
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label9 
      BackColor       =   &H0080FFFF&
      Caption         =   "E-Mail Address,If Any"
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
      Left            =   120
      TabIndex        =   7
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackColor       =   &H0080FFFF&
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
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   3015
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080FFFF&
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
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080FFFF&
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
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FFFF&
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
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FFFF&
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
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FFFF&
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
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Form For New Telephone Connection"
      BeginProperty Font 
         Name            =   "Galliard BT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2640
      TabIndex        =   0
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As ADODB.Connection
Dim rec As ADODB.Recordset

Private Sub Command1_Click()
rec.MoveFirst
Text1.Text = rec!CName
Text2.Text = rec!JName
Text3.Text = rec!FGName
Text4.Text = rec!PGno
Text5.Text = rec!WTno
Text6.Text = rec!Hno
Text7.Text = rec!ST
Text8.Text = rec!City
Text9.Text = rec!Pin
Text10.Text = rec!CAddress
Text11.Text = rec!EMail
Text12.Text = rec!CAT
Text13.Text = rec!Cgname
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
Text1.Text = rec!CName
Text2.Text = rec!JName
Text3.Text = rec!FGName
Text4.Text = rec!PGno
Text5.Text = rec!WTno
Text6.Text = rec!Hno
Text7.Text = rec!ST
Text8.Text = rec!City
Text9.Text = rec!Pin
Text10.Text = rec!CAddress
Text11.Text = rec!EMail
Text12.Text = rec!CAT
Text13.Text = rec!Cgname
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
Text1.Text = rec!CName
Text2.Text = rec!JName
Text3.Text = rec!FGName
Text4.Text = rec!PGno
Text5.Text = rec!WTno
Text6.Text = rec!Hno
Text7.Text = rec!ST
Text8.Text = rec!City
Text9.Text = rec!Pin
Text10.Text = rec!CAddress
Text11.Text = rec!EMail
Text12.Text = rec!CAT
Text13.Text = rec!Cgname
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
Text1.Text = rec!CName
Text2.Text = rec!JName
Text3.Text = rec!FGName
Text4.Text = rec!PGno
Text5.Text = rec!WTno
Text6.Text = rec!Hno
Text7.Text = rec!ST
Text8.Text = rec!City
Text9.Text = rec!Pin
Text10.Text = rec!CAddress
Text11.Text = rec!EMail
Text12.Text = rec!CAT
Text13.Text = rec!Cgname
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
rec.AddNew
rec!CName = Text1.Text
rec!JName = Text2.Text
rec!FGName = Text3.Text
rec!PGno = Text4.Text
rec!WTno = Text5.Text
rec!Hno = Text6.Text
rec!ST = Text7.Text
rec!City = Text8.Text
rec!Pin = Text9.Text
rec!CAddress = Text10.Text
rec!EMail = Text11.Text
rec!Ccode = Combo1.Text
rec!CAT = Text12.Text
rec!Cgcode = Combo2.Text
rec!Cgname = Text13.Text
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
End Sub


Private Sub Command7_Click()
Unload Me
End Sub


Private Sub Command8_Click()
Form1.Show
Unload Me
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


End Sub

Private Sub Form_Load()
Set db = New ADODB.Connection
Set rec = New ADODB.Recordset
db.ConnectionString = "dsn=phone;uid=;pwd=;"
db.Open
rec.Open "newconn1", db, adOpenDynamic, adLockOptimistic
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
End Sub

