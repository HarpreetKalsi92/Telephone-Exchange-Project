VERSION 5.00
Begin VB.Form form13 
   Caption         =   "menu"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12375
   LinkTopic       =   "Form5"
   Picture         =   "menu1.frx":0000
   ScaleHeight     =   9495
   ScaleWidth      =   12375
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   400
      Left            =   8040
      Top             =   600
   End
   Begin VB.Timer Timer1 
      Interval        =   2200
      Left            =   3720
      Top             =   600
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00C0C000&
      Caption         =   " Complaint Form"
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
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6360
      Width           =   3615
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C000&
      Caption         =   " Add On Open Facility"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2160
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
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
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   3615
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C000&
      Caption         =   "Shifting The Phone Connection"
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
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   3615
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C000&
      Caption         =   "Add  On Facility Close"
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
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   3615
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C000&
      Caption         =   "Prepare The Telephone Bill"
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
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3840
      Width           =   3615
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C000&
      Caption         =   "Exit"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6360
      Width           =   3615
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C000&
      Caption         =   "Disconnect Of Phone Connection"
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
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4680
      Width           =   3615
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C000&
      Caption         =   "Employee Payrol System"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5520
      Width           =   3615
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C000&
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
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5520
      Width           =   3615
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C000&
      Caption         =   "Bill Form"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3000
      Width           =   3615
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C000&
      Caption         =   "Pay The Telephone Bill"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3840
      Width           =   3615
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C000&
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4680
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C000&
      Caption         =   " Payment During Connection"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label Label16 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   29
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label15 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5880
      TabIndex        =   28
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label14 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   27
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label Label13 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5880
      TabIndex        =   26
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   5880
      TabIndex        =   25
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   5880
      TabIndex        =   24
      Top             =   5640
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   23
      Top             =   6360
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1440
      TabIndex        =   22
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   21
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1440
      TabIndex        =   20
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   19
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   1440
      TabIndex        =   18
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   1440
      TabIndex        =   17
      Top             =   5640
      Width           =   255
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   1440
      TabIndex        =   16
      Top             =   6480
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   615
      Left            =   1560
      TabIndex        =   15
      Top             =   7080
      Width           =   8895
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
      TabIndex        =   1
      Top             =   480
      Width           =   3735
   End
End
Attribute VB_Name = "form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As ADODB.Connection
Dim rec As ADODB.Recordset
Dim rec1 As ADODB.Recordset
Dim rec2 As ADODB.Recordset
Dim b As Integer
Dim a As Integer
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

Private Sub Command14_Click()
Unload Me
Load Form17
Form17.Show
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
Load Forma1
Forma1.Show
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
Load Form16
Form16.Show
End Sub

Private Sub Command9_Click()
e = MsgBox("Are you sure to exit this program?", vbYesNo, "Confirm Box")
If e = vbYes Then
MsgBox "Have A Nice Day! Good Bye", vbOKOnly, "Thanks"

Unload Me
End If

End Sub


Private Sub Form_Load()
b = 0
a = 0
End Sub





Private Sub Timer1_Timer()
If b = 0 Then
Label2.Caption = "<----------Wel Come To 'Telephone Exchange System'---------->"
b = 1
ElseIf b = 1 Then
Label2.Caption = "  If Do You Want Any Service,Then Click Particulur Button"
b = 2
ElseIf b = 2 Then
Label2.Caption = "<----------Wel Come To 'Telephone Exchange System'---------->"
b = 3
ElseIf b = 3 Then
Label2.Caption = "  If Do You Want Any Service,Then Click Particulur Button"
b = 4
ElseIf b = 4 Then
Label2.Caption = "<----------Wel Come To 'Telephone Exchange System'--------->"
b = 5
ElseIf b = 5 Then
Label2.Caption = "  If Do You Want Any Service,Then Click Particulur Button"
b = 0
Else
b = 0
End If
End Sub

Private Sub Timer2_Timer()
If a = 0 Then
Label3.Caption = "*"
a = 1
ElseIf a = 1 Then
Label5.Caption = "*"
a = 2
ElseIf a = 2 Then
Label6.Caption = "*"
a = 3
ElseIf a = 3 Then
Label7.Caption = "*"
a = 4
ElseIf a = 4 Then
Label8.Caption = "*"
a = 5
ElseIf a = 5 Then
Label9.Caption = "*"
a = 6
ElseIf a = 6 Then
Label10.Caption = "*"
a = 7
ElseIf a = 7 Then
Label4.Caption = "*"
a = 8
ElseIf a = 8 Then
Label11.Caption = "*"
a = 9
ElseIf a = 9 Then
Label12.Caption = "*"
a = 10
ElseIf a = 10 Then
Label13.Caption = "*"
a = 11
ElseIf a = 11 Then
Label14.Caption = "*"
a = 12
ElseIf a = 12 Then
Label15.Caption = "*"
a = 13
ElseIf a = 13 Then
Label16.Caption = "*"
a = 0
Else
b = 0
End If
End Sub
