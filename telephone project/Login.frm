VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Form4"
   ClientHeight    =   6795
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9480
   LinkTopic       =   "Form4"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   1065
      Left            =   2760
      Picture         =   "Login.frx":0000
      ScaleHeight     =   1005
      ScaleWidth      =   4125
      TabIndex        =   7
      Top             =   480
      Width           =   4185
   End
   Begin VB.CommandButton Command2 
      Caption         =   " E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   3
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Interval        =   2200
      Left            =   4800
      Top             =   120
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Text            =   " "
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   5640
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   " "
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
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
      Height          =   495
      Left            =   2160
      TabIndex        =   8
      Top             =   7680
      Width           =   7935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   " Enter The User Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   " Enter the Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0E0FF&
      Caption         =   "     Login Form"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   3000
      TabIndex        =   4
      Top             =   2640
      Width           =   3495
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Dim b As Integer
Private Sub Command1_Click()
Dim a As String
Dim b As String
a = Text1.Text
b = Text2.Text
If a = "HELLO" And b = "123" Then
  form13.Show
  Unload Me
  
Else
  MsgBox ("Enter Correct Name and password")
End If

End Sub

Private Sub Command2_Click()
c = MsgBox("Are you sure to exit this program?", vbYesNo, "Confirm Box")
If c = vbYes Then

Unload Me
End If
End Sub

Private Sub Form_Load()
Text1.Text = "HELLO"
Text2.Text = ""
End Sub





Private Sub Timer2_Timer()
If b = 0 Then
Label3.Caption = "Wel Come To You Telephone Exchange System"
b = 1
ElseIf b = 1 Then
Label3.Caption = "This Project Title Is Based On BSNL"
b = 2
ElseIf b = 2 Then
Label3.Caption = "Please Enter The Password"

b = 3
ElseIf b = 3 Then
Label3.Caption = " Sir Please Enter The Password"
b = 4
ElseIf b = 4 Then
Label3.Caption = "Sir Please Enter The Correct Password"
b = 0
Else
b = 0
End If
End Sub
