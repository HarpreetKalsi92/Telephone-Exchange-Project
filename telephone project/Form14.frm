VERSION 5.00
Begin VB.Form Form3 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0FFC0&
   Caption         =   "payment during connention"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   495
      Left            =   4680
      TabIndex        =   50
      Top             =   6360
      Width           =   1095
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H00C0C0FF&
      Height          =   405
      Left            =   9240
      TabIndex        =   49
      Top             =   5400
      Width           =   1695
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00C0C0FF&
      Height          =   405
      Left            =   2520
      TabIndex        =   48
      Top             =   5520
      Width           =   1815
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00C0C0FF&
      Height          =   405
      Left            =   9240
      TabIndex        =   47
      Top             =   4920
      Width           =   1695
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00C0C0FF&
      Height          =   405
      Left            =   2520
      TabIndex        =   46
      Top             =   5040
      Width           =   1815
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00C0C0FF&
      Height          =   405
      Left            =   9240
      TabIndex        =   45
      Top             =   4440
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00C0C0FF&
      Height          =   405
      Left            =   9240
      TabIndex        =   44
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   2400
      TabIndex        =   43
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0C0FF&
      Height          =   405
      Left            =   9240
      TabIndex        =   42
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0C0FF&
      Height          =   405
      Left            =   2400
      TabIndex        =   41
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   4080
      TabIndex        =   40
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Manuplation Button"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5520
      TabIndex        =   35
      Top             =   7440
      Width           =   4935
      Begin VB.CommandButton Command5 
         Caption         =   "&Payment"
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
         Index           =   1
         Left            =   120
         TabIndex        =   39
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command8 
         Caption         =   "E&xit"
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
         Index           =   1
         Left            =   3720
         TabIndex        =   38
         ToolTipText     =   "Exit"
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command7 
         Caption         =   "M&enu"
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
         Index           =   1
         Left            =   2640
         TabIndex        =   37
         ToolTipText     =   "Menu"
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Save/Update"
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
         Index           =   1
         Left            =   1320
         TabIndex        =   36
         ToolTipText     =   "Save/Update"
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Navigation Button"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1080
      TabIndex        =   30
      Top             =   7440
      Width           =   3975
      Begin VB.CommandButton Command1 
         Caption         =   "&<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   34
         ToolTipText     =   "MoveFirst"
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&>>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   2880
         TabIndex        =   33
         ToolTipText     =   "MoveNext"
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   1920
         TabIndex        =   32
         ToolTipText     =   "MoveLast"
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   1080
         TabIndex        =   31
         ToolTipText     =   "MovePrevious"
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   4080
      TabIndex        =   13
      Text            =   " "
      Top             =   840
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0C0FF&
      Height          =   315
      Left            =   9240
      TabIndex        =   12
      Text            =   " "
      Top             =   3480
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00C0C0FF&
      Height          =   315
      Left            =   4080
      TabIndex        =   11
      Text            =   " "
      Top             =   3960
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   9480
      MaxLength       =   6
      TabIndex        =   10
      Top             =   840
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Payment Mode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   7
      Top             =   1920
      Width           =   5415
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Cash"
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
         Left            =   600
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Draft/Cheque"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2400
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Internal Wiring Required"
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
      Left            =   360
      TabIndex        =   4
      Top             =   6120
      Width           =   3735
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Yes"
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
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "No"
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
         Left            =   2160
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Telephone Instrument  Required"
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
      Left            =   6480
      TabIndex        =   1
      Top             =   6120
      Width           =   4935
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Yes"
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
         Left            =   480
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "No"
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
         Left            =   2040
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00C0C0FF&
      Height          =   315
      Left            =   4080
      TabIndex        =   0
      Text            =   " "
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "        Payment During Connection"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   2520
      TabIndex        =   29
      Top             =   0
      Width           =   6135
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   " Name Of Customer/Company"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   " Name Of Joint Applicant(If Any)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "  Net Amount"
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
      Left            =   360
      TabIndex        =   26
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Draft No."
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
      Left            =   600
      TabIndex        =   25
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Name Of Bank"
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
      Left            =   6600
      TabIndex        =   24
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      Caption         =   " Payment Date"
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
      Left            =   6600
      TabIndex        =   23
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      Caption         =   " Exchange Capacity"
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
      Left            =   480
      TabIndex        =   22
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      Caption         =   " Area"
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
      Left            =   6600
      TabIndex        =   21
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Catagory Code"
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
      Left            =   480
      TabIndex        =   20
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFC0&
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
      Left            =   480
      TabIndex        =   19
      Top             =   4560
      Width           =   2415
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFC0&
      Caption         =   " Installation Fee"
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
      Left            =   480
      TabIndex        =   18
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFC0&
      Caption         =   " Rent/Two Month"
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
      Left            =   6600
      TabIndex        =   17
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Category Name"
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
      Left            =   6600
      TabIndex        =   16
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Concessional Group Name"
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
      Left            =   6600
      TabIndex        =   15
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Phone Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   6720
      TabIndex        =   14
      Top             =   960
      Width           =   2055
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As ADODB.Connection
Dim rec As ADODB.Recordset


Private Sub Combo1_click()
 Dim e As Long, ar As String
e = Val(Text6.Text)
ar = Combo1.Text
If e < 100 Then
If ar = "Rural" Or ar = "Urban" Then
Text10.Text = 100
End If
ElseIf e > 100 And e <= 1000 Then
If ar = "Urban" Then
Text10.Text = 240
End If
If ar = "Rural" Then
Text10.Text = 100
End If
ElseIf e > 1000 And e <= 30000 Then
If ar = "Urban" Then
Text10.Text = 240
End If
If ar = "Rural" Then
Text10.Text = 220
End If
ElseIf e > 30000 And e <= 100000 Then
If ar = "Urban" Then
Text10.Text = 360
End If
If ar = "Rural" Then
Text10.Text = 300
End If
ElseIf e > 100000 And e <= 300000 Then
If ar = "Urban" Then
Text10.Text = 500
End If
If ar = "Rural" Then
Text10.Text = 420
End If
ElseIf e > 30000 Then
If ar = "Urban" Then
Text10.Text = 300
End If
If ar = "Rural" Then
Text10.Text = 420
End If
End If
End Sub

Private Sub Combo2_click()
If Combo2.Text = "01" Then
Text7.Text = "N-OUT-Gen"
ElseIf Combo2.Text = "02" Then
Text7.Text = "N-OYT-Spl"
ElseIf Combo2.Text = "03" Then
Text7.Text = "N-OYT-SS"
ElseIf Combo2.Text = "04" Then
Text7.Text = "N-OYT-SWS"
ElseIf Combo2.Text = "05" Then
Text7.Text = "N-OYT-G-SE-DOT"
ElseIf Combo2.Text = "06" Then
Text7.Text = "OYT-Gen"
ElseIf Combo2.Text = "07" Then
Text7.Text = "OYT-Spl"
ElseIf Combo2.Text = "08" Then
Text7.Text = "TATKAL"
End If

End Sub

Private Sub Combo3_Click()
If Combo3.Text = "01" Then
Text8.Text = "Freedom Fighters"
ElseIf Combo3.Text = "02" Then
Text8.Text = "Gallantry Award Winners"
ElseIf Combo3.Text = "03" Then
Text8.Text = "War Widows"
ElseIf Combo3.Text = "04" Then
Text8.Text = "Disabled Soldiers"
ElseIf Combo3.Text = "05" Then
Text8.Text = "Blind"
ElseIf Combo3.Text = "06" Then
Text8.Text = "Senior Citizens"
ElseIf Combo3.Text = "07" Then
Text8.Text = "Retired DOT Employees"
ElseIf Combo3.Text = "08" Then
Text8.Text = "Serving DOT Employees"
ElseIf Combo3.Text = "09" Then
Text8.Text = "Recognized Educational Institutes"
ElseIf Combo3.Text = "10" Then
Text8.Text = "Orphanages"
ElseIf Combo3.Text = "11" Then
Text8.Text = "Homes for aged,infirm,spastics,handicapped,deaf-dump-mute persons/voluntary,organization for tribal welfare,institutions recorgnized by Govt."
End If

End Sub

Private Sub Command1_Click(Index As Integer)
rec.MoveFirst
Text1.Text = rec!Cuname
Text2.Text = rec!Pno
Text3.Text = rec!JName
Text4.Text = rec!Dno
Text5.Text = rec!Bname
Text6.Text = rec!Exc
Text7.Text = rec!CName
Text8.Text = rec!Cgname
Text9.Text = rec!Ifee
Text10.Text = rec!Rmonth
Text11.Text = rec!Namount
Text12.Text = rec!PDate
Combo1.Text = rec!Area
Combo2.Text = rec!Ccode
Combo3.Text = rec!Cgcode
Dim a As String
a = rec!pmode
If a = "Cash" Then
  Option1.Value = True
End If

If a = "Draft" Then
  Option2.Value = True
End If


End Sub

Private Sub Command2_Click(Index As Integer)
rec.MovePrevious
If rec.BOF = True Then
rec.MoveFirst
End If
Text1.Text = rec!Cuname
Text2.Text = rec!Pno
Text3.Text = rec!JName
Text4.Text = rec!Dno
Text5.Text = rec!Bname
Text6.Text = rec!Exc
Text7.Text = rec!CName
Text8.Text = rec!Cgname
Text9.Text = rec!Ifee
Text10.Text = rec!Rmonth
Text11.Text = rec!Namount
Text12.Text = rec!PDate
Combo1.Text = rec!Area
Combo2.Text = rec!Ccode
Combo3.Text = rec!Cgcode

End Sub

Private Sub Command3_Click(Index As Integer)
rec.MoveLast
Text1.Text = rec!Cuname
Text2.Text = rec!Pno
Text3.Text = rec!JName
Text4.Text = rec!Dno
Text5.Text = rec!Bname
Text6.Text = rec!Exc
Text7.Text = rec!CName
Text8.Text = rec!Cgname
Text9.Text = rec!Ifee
Text10.Text = rec!Rmonth
Text11.Text = rec!Namount
Text12.Text = rec!PDate
Combo1.Text = rec!Area
Combo2.Text = rec!Ccode
Combo3.Text = rec!Cgcode

End Sub

Private Sub Command4_Click(Index As Integer)
rec.MoveNext
If rec.EOF = True Then
rec.MoveLast
End If
Text1.Text = rec!Cuname
Text2.Text = rec!Pno
Text3.Text = rec!JName
Text4.Text = rec!Dno
Text5.Text = rec!Bname
Text6.Text = rec!Exc
Text7.Text = rec!CName
Text8.Text = rec!Cgname
Text9.Text = rec!Ifee
Text10.Text = rec!Rmonth
Text11.Text = rec!Namount
Text12.Text = rec!PDate
Combo1.Text = rec!Area
Combo2.Text = rec!Ccode
Combo3.Text = rec!Cgcode

Dim a As String
a = rec!pmode
If a = "cash" Then
Option1.Value = True
End If

If a = "Draft/cheque" Then
Option2.Value = True
End If
End Sub

Private Sub Command5_Click(Index As Integer)

Dim x As String, b As String, c As String
If Option3.Value = True Then
x = "Yes"
End If
If Option4.Value = True Then
x = "No"
End If
If Option5.Value = True Then
b = "Yes"
End If
If Option6.Value = True Then
b = "No"
End If

If Combo2.Text = "01" Or Combo2.Text = "02" Or Combo2.Text = "03" Then
If x = "Yes" And b = "Yes" Then
Text11.Text = Val(Text9.Text) + 2160
ElseIf x = "Yes" And b = "No" Then
Text11.Text = (Val(Text9.Text) + (2160 - 500))
ElseIf x = "No" And b = "Yes" Then
Text11.Text = (Val(Text9.Text) + (2160 - 250))
ElseIf x = "No" And b = "No" Then
Text11.Text = (Val(Text9.Text) + (2160 - 250 - 500))
End If
End If

If Combo2.Text = "04" Or Combo2.Text = "05" Then

If x = "Yes" And b = "Yes" Then
Text11.Text = Val(Text9.Text)
ElseIf x = "Yes" And b = "No" Then
Text11.Text = Val(Text9.Text) - 500
ElseIf x = "No" And b = "Yes" Then
Text11.Text = Val(Text9.Text) - 250
ElseIf x = "No" And b = "No" Then
Text11.Text = Val(Text9.Text) - 250 - 500
End If
End If

If Combo2.Text = "06" Or Combo2.Text = "07" Then
r = Val(Text10.Text)
Text10.Text = r - 80
If x = "Yes" And b = "Yes" Then
Text11.Text = Val(Text9.Text) + 10000
ElseIf x = "Yes" And b = "No" Then
Text11.Text = (Val(Text9.Text) + (10000 - 500))
ElseIf x = "No" And b = "Yes" Then
Text11.Text = (Val(Text9.Text) + (10000 - 250))
ElseIf x = "No" And b = "No" Then
Text11.Text = (Val(Text9.Text) + (10000 - 250 - 500))
End If
End If
If Combo2.Text = "08" Then
If x = "Yes" And b = "Yes" Then
Text11.Text = Val(Text9.Text) + 30000
ElseIf x = "Yes" And b = "No" Then
Text11.Text = (Val(Text9.Text) + (30000 - 500))
ElseIf x = "No" And b = "Yes" Then
Text11.Text = (Val(Text9.Text) + (30000 - 250))
ElseIf x = "No" And b = "No" Then
Text11.Text = (Val(Text9.Text) + (30000 - 250 - 500))
End If
End If
If Combo3.Text = "01" Or Combo3.Text = "02" Or Combo3.Text = "03" Or Combo3.Text = "04" Then
Text9.Text = 0
'Text3.Text = 2160
End If

If Combo3.Text = "01" Or Combo3.Text = "03" Or Combo3.Text = "04" Then
  con = Val(Text10.Text)
  Text10.Text = con / 2
End If
 
 If Combo3.Text = "05" And Combo2.Text = "01" Then
  con = Val(Text10.Text)
  Text10.Text = con / 2
  End If
  
  If Combo2.Text = "04" And Combo3.Text = "01" Then
  Text11.Text = 0
  Text9.Text = 0
  con = Val(Text10.Text)
  Text10.Text = con / 2
  End If
  
 If Combo3.Text = "09" Or Combo3.Text = "10" Or Combo3.Text = "11" Then
 rn = Val(Text10.Text)
 Text10.Text = rn - (rn * 25 / 100)
 End If

 If Combo3.Text = "02" Then
 Text10.Text = 0
 Text11.Text = 0
 Text9.Text = 0

End If
kk = Val(Text2.Text)
If MsgBox(kk, vbOKOnly, "U Have Alloted the Phone No") = vbOK Then
End If
Text12.Text = Date

End Sub

Private Sub Command6_Click(Index As Integer)

rec.AddNew
Text1.Text = rec!Cuname
Text2.Text = rec!Pno
Text3.Text = rec!JName
Text4.Text = rec!Dno
Text5.Text = rec!Bname
Text6.Text = rec!Exc
Text7.Text = rec!CName
Text8.Text = rec!Cgname
Text9.Text = rec!Ifee
Text10.Text = rec!Rmonth
Text11.Text = rec!Namount
Text12.Text = rec!PDate
Combo1.Text = rec!Area
Combo2.Text = rec!Ccode
Combo3.Text = rec!Cgcode
MsgBox "Record saved"
rec.Update


End Sub

Private Sub Command8_Click(Index As Integer)
Unload Me
End Sub

Private Sub Form_Load()
Set db = New ADODB.Connection
Set rec = New ADODB.Recordset
db.ConnectionString = "dsn=phone;uid=;pwd=;"
db.Open
rec.Open "payment", db, adOpenDynamic, adLockOptimistic
Combo1.AddItem "Rural"
Combo1.AddItem "Urban"
Combo2.AddItem "01"
Combo2.AddItem "02"
Combo2.AddItem "03"
Combo2.AddItem "04"
Combo2.AddItem "05"
Combo2.AddItem "06"
Combo2.AddItem "07"
Combo2.AddItem "08"

Combo3.AddItem "01"
Combo3.AddItem "02"
Combo3.AddItem "03"
Combo3.AddItem "04"
Combo3.AddItem "05"
Combo3.AddItem "06"
Combo3.AddItem "07"
Combo3.AddItem "08"
Combo3.AddItem "09"
Combo3.AddItem "10"


End Sub

Private Sub Option1_Click()
Text4.Text = "0"
Text5.Text = "Null"
Text4.Enabled = False
Text5.Enabled = False

End Sub

Private Sub Option2_Click()
Text4.Enabled = True
Text5.Enabled = True

End Sub

Private Sub Text6_lostfocus()
a = Val(Text6.Text)
If a <= 500 Then
Text9.Text = 300 + (300 * 5 / 100)
Else
Text9.Text = 800 + (800 * 5 / 100)
End If
End Sub

