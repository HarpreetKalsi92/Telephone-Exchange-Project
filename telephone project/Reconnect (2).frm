VERSION 5.00
Begin VB.Form Form14 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Form14"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10395
   LinkTopic       =   "Form14"
   ScaleHeight     =   6285
   ScaleWidth      =   10395
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFC0&
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
      Left            =   6000
      TabIndex        =   5
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Reconnect"
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
      Left            =   3480
      TabIndex        =   4
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFC0&
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
      Left            =   6000
      TabIndex        =   3
      Text            =   " "
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFC0&
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
      Left            =   6000
      TabIndex        =   2
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Menu"
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
      Left            =   4800
      TabIndex        =   1
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Exit"
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
      Left            =   6120
      TabIndex        =   0
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Reconnection  The Telephone Connection"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   3240
      TabIndex        =   9
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "     Telephone No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   1800
      TabIndex        =   8
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "  Date Of  Reconnection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "   Payable Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   2760
      Width           =   1815
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As ADODB.Connection
Dim rec As ADODB.Recordset
Dim rec1 As ADODB.Recordset
Dim rec2 As ADODB.Recordset

Private Sub Command1_Click()
rec1.AddNew
  rec1!Tno = Text1.Text
  rec1!dreconn = Text2.Text
   rec1!pamount = Text3.Text
 ' rec1.Update
 
rec.MoveFirst
Dim a As Integer
a = Text1.Text
rec2.MoveFirst
While Not rec2.EOF = True
If a = rec2!Tno Then
rec2!dis = "no"
rec2.Update
MsgBox "This Telephone Number Is Reconnected"
End If
rec2.MoveNext
Wend

rec.MoveFirst
a = Text1.Text
While Not rec.EOF = True
  If a = rec!Tno Then
    rec.Delete
    MsgBox "This Record Deleted In Disconnect Table"
  End If
  rec.MoveNext
Wend



End Sub

Private Sub Command2_Click()
c = MsgBox("Are you sure to exit this program?", vbYesNo, "Confirm Box")
If c = vbYes Then
form13.Show
Unload Me
End If
End Sub

Private Sub Command3_Click()
c = MsgBox("Are you sure to exit this program?", vbYesNo, "Confirm Box")
If c = vbYes Then

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
rec1.Open "reconnect", db, adOpenDynamic, adLockOptimistic
rec2.Open "newconn1", db, adOpenDynamic, adLockOptimistic
Text2.Text = Date
End Sub
Private Sub Text1_LostFocus()
Dim a As Integer
Dim b As Integer
b = 0
a = Text1.Text
rec.MoveFirst
While Not rec.EOF = True

If a = rec1!Tno Then

b = 1
Text3.Text = 200
End If
rec.MoveNext
Wend
If b = 1 Then
Text2.SetFocus
Else

 'MsgBox ("Reply")
 Text1.Text = ""
 Text1.SetFocus
 Text3.Text = 200
End If

End Sub

Private Sub Text3_Change()
Text3.Text = 200
End Sub
