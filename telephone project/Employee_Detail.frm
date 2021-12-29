VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form7"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11685
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form7"
   ScaleHeight     =   8430
   ScaleWidth      =   11685
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   9120
      TabIndex        =   7
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   9120
      TabIndex        =   13
      Top             =   5160
      Width           =   1935
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00C0C0FF&
      Height          =   405
      Left            =   9120
      TabIndex        =   11
      Top             =   4320
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Text            =   " "
      ToolTipText     =   "Can't Enter In This Box"
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Text            =   " "
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   9120
      TabIndex        =   5
      Text            =   " "
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Text            =   " "
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Text            =   " "
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Navigation Buttons"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   16
      Top             =   6360
      Width           =   4455
      Begin VB.CommandButton Command3 
         Caption         =   "&Previous"
         Height          =   495
         Index           =   1
         Left            =   2160
         TabIndex        =   32
         ToolTipText     =   "MoveNext"
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Last"
         Height          =   495
         Index           =   1
         Left            =   3120
         TabIndex        =   31
         ToolTipText     =   "MoveLast"
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Next"
         Height          =   495
         Index           =   1
         Left            =   1320
         TabIndex        =   30
         ToolTipText     =   "MovePrevious"
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&First"
         Height          =   495
         Index           =   1
         Left            =   480
         TabIndex        =   29
         ToolTipText     =   "MoveFirst"
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Control Button"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5640
      TabIndex        =   14
      Top             =   6360
      Width           =   4575
      Begin VB.CommandButton Command10 
         Caption         =   "MENU SHOW"
         Height          =   495
         Index           =   1
         Left            =   2640
         TabIndex        =   36
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton Command8 
         Caption         =   "E&XIT"
         Height          =   495
         Index           =   1
         Left            =   1560
         TabIndex        =   35
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command7 
         Caption         =   "S&AVE"
         Height          =   495
         Index           =   1
         Left            =   480
         TabIndex        =   15
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         Caption         =   "DE&LETE"
         Height          =   495
         Index           =   1
         Left            =   2160
         TabIndex        =   34
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "A&DD"
         Height          =   495
         Index           =   1
         Left            =   840
         TabIndex        =   33
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0C0FF&
      Height          =   315
      Left            =   3000
      TabIndex        =   10
      Text            =   " "
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Sex"
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
      Height          =   975
      Left            =   5760
      TabIndex        =   1
      Top             =   600
      Width           =   4815
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Male"
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
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Female"
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
         Left            =   2160
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00C0C0FF&
      Height          =   315
      Left            =   3000
      TabIndex        =   12
      Text            =   " "
      Top             =   5280
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   9120
      TabIndex        =   9
      Text            =   " "
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "     Form For Employee Detail"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   3240
      TabIndex        =   28
      Top             =   0
      Width           =   6255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Employee GPF No."
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
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Name OF Employee"
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
      Left            =   120
      TabIndex        =   26
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Father's Name"
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
      Left            =   5760
      TabIndex        =   25
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Employee Address"
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
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Date Of Birth"
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
      Left            =   0
      TabIndex        =   23
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Date Of Joining"
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
      Left            =   5640
      TabIndex        =   22
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Department"
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
      Left            =   120
      TabIndex        =   21
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Group"
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
      Height          =   495
      Left            =   5640
      TabIndex        =   20
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Designation"
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
      Left            =   0
      TabIndex        =   19
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Basic Pay"
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
      Left            =   5640
      TabIndex        =   18
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pincode"
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
      Left            =   5640
      TabIndex        =   17
      Top             =   2640
      Width           =   2055
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As ADODB.Connection
Dim rec As ADODB.Recordset
Dim rec1 As ADODB.Recordset
Dim flag As Boolean
 
Private Sub Command1_Click(Index As Integer)
rec.MoveFirst
Text1.Text = rec!EGPFno
Text2.Text = rec!EName
Text3.Text = rec!FName
Text4.Text = rec!EAddress
Text5.Text = rec!Pin
Text6.Text = rec!DBirth
Text7.Text = rec!DJoining
Text8.Text = rec!Group
Text9.Text = rec!BPay
Combo1.Text = rec!Department
Combo2.Text = rec!Destination
Dim a As String

a = rec!Sex
If a = "Male" Then
Option1.Value = True
End If
If a = "Female" Then
Option2.Value = True
End If
End Sub

Private Sub Command10_Click(Index As Integer)
form13.Show
Unload Me
End Sub

Private Sub Command2_Click(Index As Integer)
rec.MoveNext
If rec.EOF = True Then
rec.MoveLast
End If
Text1.Text = rec!EGPFno
Text2.Text = rec!EName
Text3.Text = rec!FName
Text4.Text = rec!EAddress
Text5.Text = rec!Pin
Text6.Text = rec!DBirth
Text7.Text = rec!DJoining
Text8.Text = rec!Group
Text9.Text = rec!BPay
Combo1.Text = rec!Department
Combo2.Text = rec!Destination
Dim a As String
a = rec!Sex
If a = "Male" Then

Option1.Value = True
End If

If a = "Female" Then
Option2.Value = True
End If

End Sub

Private Sub Command3_Click(Index As Integer)
rec.MovePrevious
If rec.BOF = True Then
rec.MoveFirst
End If
Text1.Text = rec!EGPFno
Text2.Text = rec!EName
Text3.Text = rec!FName
Text4.Text = rec!EAddress
Text5.Text = rec!Pin
Text6.Text = rec!DBirth
Text7.Text = rec!DJoining
 Text8.Text = rec!Group
Text9.Text = rec!BPay
Combo1.Text = rec!Department
Combo2.Text = rec!Destination
Dim a As String
a = rec!Sex
If a = "Male" Then

Option1.Value = True
End If

If a = "Female" Then
Option2.Value = True
End If

End Sub

Private Sub Command4_Click(Index As Integer)
rec.MoveLast
Text1.Text = rec!EGPFno
Text2.Text = rec!EName
Text3.Text = rec!FName
Text4.Text = rec!EAddress
Text5.Text = rec!Pin
Text6.Text = rec!DBirth
Text7.Text = rec!DJoining
 Text8.Text = rec!Group
Text9.Text = rec!BPay
Combo1.Text = rec!Department
Combo2.Text = rec!Destination
Dim a As String
a = rec!Sex
If a = "Male" Then
Option1.Value = True
End If

If a = "Female" Then
Option2.Value = True
End If
End Sub

Private Sub Command5_Click(Index As Integer)
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
 Text9.Text = ""
 Combo1.Text = ""
Combo2.Text = ""
Option1.Value = False
Option2.Value = False

End Sub

Private Sub Command6_Click(Index As Integer)
rec.Delete
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
 Text9.Text = ""
 Combo1.Text = ""
Combo2.Text = ""
Option1.Value = False
Option2.Value = False
MsgBox "Record Deleted"


rec.Close

rec.MoveFirst
End Sub

Private Sub Command7_Click(Index As Integer)
rec.AddNew
rec1.AddNew

rec!EGPFno = Text1.Text
rec1!EGPFno = Text1.Text
rec1!EName = Text2.Text
rec!EName = Text2.Text
rec!FName = Text3.Text
rec!EAddress = Text4.Text
rec!Pin = Text5.Text
rec!DBirth = Text6.Text
rec!DJoining = Text7.Text
rec!Group = Text8.Text
rec!BPay = Text9.Text
rec!Department = Combo1.Text
rec!Destination = Combo2.Text
If Option1.Value = True Then
rec!Sex = "Male"
End If
If Option2.Value = True Then
rec!Sex = "Female"
End If
MsgBox "Record Saved"
rec.Update
rec1.Update
End Sub

Private Sub Command8_Click(Index As Integer)
Unload Me
End Sub

Private Sub Form_Load()
Set db = New ADODB.Connection
Set rec = New ADODB.Recordset
Set rec1 = New ADODB.Recordset
db.ConnectionString = "dsn=phone;uid=;pwd=;"
db.Provider = "Microsoft.jet.Oledb.4.0"
db.Open App.Path & "\Telephone.mdb"

rec.Open "emp", db, adOpenDynamic, adLockOptimistic
rec1.Open "emp_pay", db, adOpenDynamic, adLockOptimistic

Combo1.AddItem "Maintance"
Combo1.AddItem "Account"
Combo1.AddItem "Finance"
Combo2.AddItem "Barnala"
Combo2.AddItem "Moga"
Combo2.AddItem "Patiala"

End Sub

Private Sub Text1_LostFocus()
Dim a As Integer
rec.MoveFirst
flag = False
 While Not rec.EOF = True
  If Text1.Text = rec!EGPFno And Text1.Text = rec!EGPFno Then
    flag = True
    MsgBox "This number has already exist"
    MsgBox "Please Enter Other Number"
    Text1.Text = ""
    Text1.SetFocus
 End If
    rec.MoveNext
Wend
If flag = False Then
 Text2.SetFocus
  End If
 
End Sub
