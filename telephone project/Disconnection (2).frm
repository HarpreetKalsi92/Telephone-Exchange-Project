VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form2"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8310
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   6810
   ScaleWidth      =   8310
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox DTPicker1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      ScaleHeight     =   315
      ScaleWidth      =   1755
      TabIndex        =   20
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2760
      TabIndex        =   19
      Top             =   3960
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2280
      TabIndex        =   18
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2280
      TabIndex        =   17
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton Command9 
      Caption         =   "&Menu"
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
      TabIndex        =   16
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Disconnect"
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
      Left            =   5520
      TabIndex        =   15
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   -18240
      TabIndex        =   14
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Manuplation Button"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5040
      TabIndex        =   9
      Top             =   5160
      Width           =   3975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Navigation Button"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   600
      TabIndex        =   8
      Top             =   5160
      Width           =   3735
      Begin VB.CommandButton Command5 
         Caption         =   "&>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   13
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&>>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   12
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   11
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Purpose of Disconnection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   4095
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "WithDraw"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   5
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Bill Not Pay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Date of Disconnection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Date of Bill Pay"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Name of Subscriber"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Telephone No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Form For Disconnection"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As ADODB.Connection
Dim rec As ADODB.Recordset
Dim rec1 As ADODB.Recordset
Dim rec2 As ADODB.Recordset
Private Sub Command1_Click()
If rec.State = 0 Then
rec.MoveFirst
Text1.Text = rec!Tno
Text2.Text = rec!namesub
DTPicker1.Value = rec!dbillp
Text4.Text = rec!ddiss
Dim a As String
a = rec!Purpose
If a = "Bill Not Pay" Then
Option1.Value = True
End If

If a = "WithDraw" Then
Option2.Value = True
End If

End Sub



Private Sub Command2_Click()
rec.MovePrevious
If rec.BOF = True Then
rec.MoveFirst
End If
Text1.Text = rec!Tno
Text2.Text = rec!namesub
DTPicker1.Value = rec!dbillp
Text4.Text = rec!ddiss
Dim a As String
a = rec!Purpose
If a = "Bill Not Pay" Then
Option1.Value = True
End If

If a = "WithDraw" Then
Option2.Value = True
End If


End Sub

Private Sub Command3_Click()
rec.MoveNext
If rec.EOF = True Then
rec.MoveLast
End If
Text1.Text = rec!Tno
Text2.Text = rec!namesub
DTPicker1.Value = rec!dbillp
Text4.Text = rec!ddiss
Dim a As String
a = rec!Purpose
If a = "Bill Not Pay" Then
Option1.Value = True
End If

If a = "WithDraw" Then
Option2.Value = True
End If


End Sub



Private Sub Command5_Click()
rec.MoveLast
Text1.Text = rec!Tno
Text2.Text = rec!namesub
DTPicker1.Value = rec!dbillp
Text4.Text = rec!ddiss
Dim a As String
a = rec!Purpose
If a = "Bill Not Pay" Then
Option1.Value = True
End If

If a = "WithDraw" Then
Option2.Value = True
End If


End Sub
Private Sub Command7_Click()
Dim a As Integer
a = Text1.Text
rec2.MoveFirst
While Not rec2.EOF = True
If a = rec2!Tno Then
rec2!dis = "yes"
rec2.Update
End If
rec2.MoveNext
Wend
rec.AddNew
rec!Tno = Text1.Text
rec!namesub = Text2.Text
rec!dbillp = DTPicker1.Value
rec!ddiss = Text4.Text
If Option1.Value = True Then
rec!Purpose = "Bill Not Pay"
End If
If Option2.Value = True Then
rec!Purpose = "Withdraw"
End If
MsgBox "This Number Is Disconnect"
rec.Update
End Sub

Private Sub Command9_Click()
c = MsgBox("Are you sure to exit this program?", vbYesNo, "Confirm Box")
If c = vbYes Then
form13.Show
Unload Me
End If
End Sub

Private Sub Form_Load()
Set db = New ADODB.Connection
Set rec = New ADODB.Recordset
Set rec1 = New ADODB.Recordset
Set rec2 = New ADODB.Recordset
db.ConnectionString = "dsn=phone;uid=;pwd=;"
db.Provider = "Microsoft.jet.Oledb.4.0"
db.Open App.Path & "\Telephone.mdb"
rec.Open "disconn", db, adOpenDynamic, adLockOptimistic
rec1.Open "payment", db, adOpenDynamic, adLockOptimistic
rec2.Open "newconn1", db, adOpenDynamic, adLockOptimistic
Text4.Text = Date
End Sub

Private Sub Text1_LostFocus()
Dim a As Integer
rec1.MoveFirst
While Not rec1.EOF = True
  If Text1.Text = rec1!Pno Then
  Text2.Text = rec1!cuname
  End If
  rec1.MoveNext
Wend
End Sub
