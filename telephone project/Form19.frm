VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form8"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form8"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   4080
      TabIndex        =   34
      Top             =   6000
      Width           =   1695
   End
   Begin VB.TextBox Text11 
      Height          =   405
      Left            =   8760
      TabIndex        =   33
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox Text10 
      Height          =   405
      Left            =   4080
      TabIndex        =   32
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox Text9 
      Height          =   405
      Left            =   8760
      TabIndex        =   31
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   4080
      TabIndex        =   30
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      Height          =   405
      Left            =   8760
      TabIndex        =   29
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      Height          =   405
      Left            =   4080
      TabIndex        =   28
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Height          =   405
      Left            =   8760
      TabIndex        =   27
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   4080
      TabIndex        =   26
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4080
      TabIndex        =   25
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   8760
      TabIndex        =   24
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4080
      TabIndex        =   23
      Top             =   960
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Control Buttons"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2400
      TabIndex        =   1
      Top             =   6720
      Width           =   6135
      Begin VB.CommandButton Command9 
         BackColor       =   &H00C0FFFF&
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
         Left            =   4680
         TabIndex        =   22
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00C0FFFF&
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
         Left            =   3600
         TabIndex        =   21
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0FFFF&
         Caption         =   "U&pdate"
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
         Left            =   1440
         TabIndex        =   20
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cancel"
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
         Left            =   2520
         TabIndex        =   19
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Calculate"
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
         Left            =   360
         TabIndex        =   18
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "House Loan Issue "
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
      Left            =   720
      TabIndex        =   0
      Top             =   4920
      Width           =   5775
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NO"
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
         Left            =   2760
         TabIndex        =   36
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
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
         Height          =   435
         Left            =   600
         TabIndex        =   35
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Employee GPF No"
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
      Left            =   1440
      TabIndex        =   17
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Employee Name"
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
      Left            =   6840
      TabIndex        =   16
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Basic Pay"
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
      Left            =   1440
      TabIndex        =   15
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "House Rent Allowance"
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
      Left            =   1440
      TabIndex        =   14
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Dearness Allowance"
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
      Left            =   6720
      TabIndex        =   13
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "General Travelling Allowance"
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
      Left            =   1440
      TabIndex        =   12
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFFF&
      Caption         =   " General Provident Fund"
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
      Left            =   1320
      TabIndex        =   11
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Income Tax"
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
      Left            =   6840
      TabIndex        =   10
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Net Salary"
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
      Left            =   2280
      TabIndex        =   9
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Allowances"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Total"
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
      Left            =   6840
      TabIndex        =   7
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Deductions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0FFFF&
      Caption         =   "GPF_ Advance"
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
      Left            =   1440
      TabIndex        =   4
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Total Deductions"
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
      Left            =   6840
      TabIndex        =   3
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      Form For Employee Payroll System"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   2
      Top             =   0
      Width           =   7215
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As ADODB.Connection
Dim rec As ADODB.Recordset
Private Sub Command5_Click(Index As Integer)
Dim a, b, c As Double
 a = Val(Text8.Text)
 b = Val(Text9.Text)
 c = Val(Text10.Text)
 Text11.Text = Val(a + b + c)
 Dim a1, a2, a3 As Double
 a1 = Val(Text3.Text)
  a2 = Val(Text7.Text)
 a3 = Val(Text11.Text)
Text12.Text = a1 + a2 - a3


 End Sub

Private Sub Command7_Click(Index As Integer)
rec.MoveFirst
While Not rec.EOF = True
  If rec!egpfno = Text1.Text Then
    rec!ename = Text2.Text
    rec!BPay = Text3.Text
    rec!HRA = Text4.Text
    rec!DA = Text5.Text
    rec!GTA = Text6.Text
    rec!total = Text7.Text
    rec!GPF = Text8.Text
    rec!ITax = Text9.Text
    rec!Gpfad = Text10.Text
    rec!TD = Text11.Text
    rec!NSalary = Text12.Text
    If Option1.Value = True Then
        rec!HLIssue = "Yes"
    End If
    If Option2.Value = True Then
        rec!HLIssue = "No"
    End If
    rec.Update
    MsgBox "record saved"
 End If
  rec.MoveNext
Wend
End Sub

Private Sub Command8_Click(Index As Integer)
Unload Me
End Sub

Private Sub Form_Load()
Set db = New ADODB.Connection
Set rec = New ADODB.Recordset
db.ConnectionString = "dsn=phone;uid=;pwd=;"
db.Open
rec.Open "emp_pay", db, adOpenDynamic, adLockOptimistic

End Sub

Private Sub Text1_LostFocus()
rec.MoveFirst
Dim a As Integer
a = Text1.Text
While Not rec.EOF = True
  If a = rec!egpfno Then
    Text2.Text = rec!ename
  End If
  rec.MoveNext
Wend

End Sub

Private Sub Text3_LostFocus()
Dim a As Double
Dim h As Double, d As Double, t As Double
a = Text3.Text
If a < 5000 Then
  h = (a * 25) / 100
  Text4.Text = h
  d = (a * 55) / 100
  Text5.Text = d
  t = (a * 15) / 100
  Text6.Text = t
  Text7.Text = h + d + t
End If
If a >= 5000 And a < 10000 Then
h = (a * 30) / 100
Text4.Text = h
  d = (a * 60) / 100
  Text5.Text = d
  t = (a * 20) / 100
  Text6.Text = t
  Text7.Text = h + d + t
End If
If a > 10000 Then
h = (a * 40) / 100
Text4.Text = h
  d = (a * 70) / 100
  Text5.Text = d
  t = (a * 30) / 100
  Text6.Text = t
  Text7.Text = h + d + t
End If
Dim pf As Double, tax As Double
a = Text3.Text
If a < 5000 Then
  pf = (a * 10) / 100
  Text8.Text = pf
  tax = (a * 10) / 100
  Text9.Text = tax
    Text11.Text = pf + tax
End If
If a >= 5000 And a < 10000 Then
  pf = (a * 10) / 100
  Text8.Text = pf
  tax = (a * 10) / 100
  Text9.Text = tax
    Text11.Text = pf + tax
End If
If a > 10000 Then
  pf = (a * 10) / 100
  Text8.Text = pf
  tax = (a * 10) / 100
  Text9.Text = tax
    Text11.Text = pf + tax
End If

End Sub
