VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Form4"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7680
   LinkTopic       =   "Form4"
   ScaleHeight     =   4170
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   1065
      Left            =   2640
      Picture         =   "Form15.frx":0000
      ScaleHeight     =   1005
      ScaleWidth      =   4125
      TabIndex        =   7
      Top             =   240
      Width           =   4185
   End
   Begin VB.CommandButton Command2 
      Caption         =   " E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   3
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   8000
      Left            =   240
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Left            =   240
      Top             =   600
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0FF&
      Height          =   495
      Left            =   6240
      TabIndex        =   2
      Text            =   " "
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0C0FF&
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   6240
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   " "
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   " Enter The User Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   " Enter the Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0E0FF&
      Caption         =   "     Login Form"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   3360
      TabIndex        =   4
      Top             =   1800
      Width           =   3495
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a As String
Dim b As String
a = Text1.Text
b = Text2.Text
If a = "abc" And b = "1234" Then
  Form1.Show
  Unload Me
  
Else
  MsgBox ("Enter Correct Name and password")
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = " "
Text2.Text = " "

End Sub
