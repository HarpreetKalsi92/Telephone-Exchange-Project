VERSION 5.00
Begin VB.Form menu 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Form13"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form13"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   " New Telephone Connection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1560
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   " Payment During  Connection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   11
      Top             =   1560
      Width           =   3495
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "  ADD On Facility Opening"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   10
      Top             =   2520
      Width           =   3495
   End
   Begin VB.CommandButton Command5 
      Caption         =   " Add On Facility Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   9
      Top             =   3360
      Width           =   3615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Prepare TheTelephone Bill"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   8
      Top             =   4320
      Width           =   3615
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H000000FF&
      Caption         =   " Pay The Telephone Bill "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   7
      Top             =   4320
      Width           =   3495
   End
   Begin VB.CommandButton Command9 
      Caption         =   " Exit From Here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   6
      Top             =   7080
      Width           =   3495
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H000000FF&
      Caption         =   " Bill Form "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   5
      Top             =   3360
      Width           =   3495
   End
   Begin VB.CommandButton Command11 
      Caption         =   " Disconnection Of Phone Connection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   4
      Top             =   5160
      Width           =   3615
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Employee Detail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   3
      Top             =   6120
      Width           =   3735
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Employee Payroll System"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   2
      Top             =   6120
      Width           =   3495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Shifting the Phone Connection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   2520
      Width           =   3615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Reconnect The Connection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   0
      Top             =   5160
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "      Menu Form"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   735
      Left            =   4200
      TabIndex        =   13
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As ADODB.Connection
Dim rec As ADODB.Recordset
Dim rec1 As ADODB.Recordset
Dim rec2 As ADODB.Recordset
Private Sub Command1_Click()
Unload Me
Load Form1
Form1.Show
End Sub

Private Sub Command10_Click()
Unload Me
Load Form9
Form9.Show
End Sub

Private Sub Command11_Click()
Unload Me
Load Form2
Form2.Show
End Sub

Private Sub Command12_Click()
Unload Me
Load Form7
Form7.Show
End Sub

Private Sub Command13_Click()
Unload Me
Load Form8
Form8.Show
End Sub

Private Sub Command2_Click()
Unload Me
Load Form3
Form3.Show
End Sub

Private Sub Command3_Click()
Unload Me
Load Form11
Form11.Show
End Sub

Private Sub Command4_Click()
Unload Me
Load Form5
Form5.Show
End Sub

Private Sub Command5_Click()
Unload Me
Load Form6
Form6.Show
End Sub

Private Sub Command6_Click()
Unload Me
Load Form14
Form14.Show
End Sub

Private Sub Command7_Click()
Unload Me
Load Form10
Form10.Show
End Sub

Private Sub Command8_Click()
Unload Me
Load Form3
Form3.Show
End Sub

Private Sub Command9_Click()
MsgBox ("Thanks To U For Using My Project")
MsgBox ("Have A Nice Day! Good Bye")
Unload Me
End Sub

Private Sub Form_Load()
'StatusBar1.Panels(6).Text = "SELECT ANY FORM FROM THIS MENU"
'Set db = New ADODB.Connection
'Set rec = New ADODB.Recordset
'Set rec1 = New ADODB.Recordset
'db.ConnectionString = "dsn=neha;uid=sa;pwd=;"
'db.Open
'rec.Open "newconn", db, adOpenDynamic, adLockOptimistic
'rec1.Open "paybill", db, adOpenDynamic, adLockOptimistic


End Sub


