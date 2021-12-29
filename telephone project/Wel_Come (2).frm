VERSION 5.00
Begin VB.Form Form12 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Form12"
   ClientHeight    =   9210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13650
   LinkTopic       =   "Form12"
   ScaleHeight     =   9210
   ScaleWidth      =   13650
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   120
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Continue"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0E0FF&
      Caption         =   " Roll No. :-9204480065"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   3480
      Width           =   3855
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Name:-Sandeep Singh"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Top             =   3000
      Width           =   3855
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Roll No. :-10204480079"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   840
      TabIndex        =   8
      Top             =   2400
      Width           =   3975
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Name:-Gurdeep Singh"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   840
      TabIndex        =   7
      Top             =   1920
      Width           =   4095
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Mrs. Rajni Rani"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   7320
      TabIndex        =   6
      Top             =   4440
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Submitted By:-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   840
      TabIndex        =   5
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "  Name  :-Sarbjit Singh                         "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   960
      Width           =   4215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "  Roll No. :- 9204480066                "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   1320
      Width           =   4215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Submitted To:-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   7320
      TabIndex        =   2
      Top             =   3840
      Width           =   3015
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   5880
      Width           =   14055
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
Load Form4
Form4.Show
End Sub

Private Sub Form_Load()
a = 0
b = 0
End Sub




Private Sub Label7_Click()

End Sub

Private Sub Timer2_Timer()
If b = 0 Then
Label7.Caption = "This Project Title Is Based On BSNL"
b = 1
ElseIf b = 1 Then
Label7.Caption = "This Project IS Prepared by Sarbjit Singh,Gurdeep Singh,Sandeep Singh"
b = 2
ElseIf b = 2 Then
Label7.Caption = "We are the Students Of Bsc 6th Sem Of Bagga computer center ,Ahmedgarh"
b = 3
ElseIf b = 3 Then
Label7.Caption = "This Project Is Completed In The Guidance Of Mrs. Rajni Rani"
ElseIf b = 4 Then
Label7.Caption = "My Project Title Is Telephone System Automation "
b = 0
Else
b = 0
End If
End Sub

