VERSION 5.00
Begin VB.Form Form12 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Form12"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form12"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   7920
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   11160
      Top             =   5520
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Continue"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Submitted By"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   495
      Left            =   3960
      TabIndex        =   6
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "  Name                         Jagmohan,Harjeet And Paramjeet"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Top             =   1320
      Width           =   7575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "  Class                          MCA       6th Sem"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   1920
      Width           =   7335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "  Registration No.         200031623188"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   2520
      Width           =   5895
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Submitted To"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   4440
      Width           =   3015
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   7800
      Width           =   9495
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer, b As Integer

Private Sub Command1_Click()
Unload Me
'Load Form2
'Form2.Show
End Sub

Private Sub Form_Load()
a = 0
b = 0
End Sub

Private Sub Timer1_Timer()
'Dim a As Integer

'If a = 0 Then
'Picture = LoadPicture("c:\nehag\CONNECT.jpg")
'a = 1
If a = 1 Then
Picture = LoadPicture("c:\program Files\microsoft visual studio\vb98\project\REG.jpg")
a = 2
ElseIf a = 2 Then
Picture = LoadPicture("c:\program Files\microsoft visual studio\vb98\project\Eula.jpg")
a = 3
ElseIf a = 3 Then
Picture = LoadPicture("c:\program Files\microsoft visual studio\vb98\project\DIALTONE.jpg")
a = 4
ElseIf a = 4 Then
Picture = LoadPicture("c:\program Files\microsoft visual studio\vb98\project\TAPI.jpg")
a = 1
Else
a = 1
'Picture1.Picture = LoadPicture("c:\pict\CONNECT.jpg")
End If
End Sub
Private Sub Timer2_Timer()
If b = 0 Then
Label7.Caption = "This Project Title Is Based On BSNL"
b = 1
ElseIf b = 1 Then
Label7.Caption = "This Project IS Prepared by Jagmohan,Harjeet And Paramjeet"
b = 2
ElseIf b = 2 Then
Label7.Caption = "I am The Student Of MCA 6th Sem Of PTU Study Centre ,Barnala "
b = 3
ElseIf b = 3 Then
Label7.Caption = "This Project Is Completed In The Guidance Of Mr.Surinder Singh Sir."
b = 4
ElseIf b = 4 Then
Label7.Caption = "My Project Title Is Telephone System Automation "
b = 0
Else
b = 0
End If
End Sub

