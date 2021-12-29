VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form10"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form10"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   495
      Left            =   3000
      TabIndex        =   17
      Top             =   7320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   6480
      TabIndex        =   9
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   6480
      TabIndex        =   8
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   6480
      TabIndex        =   7
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   6480
      TabIndex        =   6
      Top             =   3480
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   6480
      TabIndex        =   5
      Top             =   4200
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   6480
      TabIndex        =   4
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
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
      Left            =   4320
      TabIndex        =   3
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
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
      Left            =   6960
      TabIndex        =   2
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
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
      Left            =   5640
      TabIndex        =   1
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Calculate"
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
      Left            =   3000
      TabIndex        =   0
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "  Form For Prepare The  Telephone Bill"
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
      Left            =   1920
      TabIndex        =   16
      Top             =   0
      Width           =   7695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Telephone No."
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
      Left            =   1680
      TabIndex        =   15
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Previous  Meter reading"
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
      Left            =   1680
      TabIndex        =   14
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Current Meter Reading"
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
      Left            =   1680
      TabIndex        =   13
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Metered Calls"
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
      Left            =   1680
      TabIndex        =   12
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Net Chargeable Calls"
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
      Left            =   1680
      TabIndex        =   11
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Net  Amount"
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
      Left            =   1680
      TabIndex        =   10
      Top             =   5040
      Width           =   2175
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As ADODB.Connection
Dim rec As ADODB.Recordset
Dim rec1 As ADODB.Recordset
Dim rec2 As ADODB.Recordset
Dim i, j, k, s
Dim am As Integer


Private Sub Command1_Click()
rec.AddNew
rec!tno = Text1.Text
rec!PMR = Text2.Text
rec!CMR = Text3.Text
rec!Mcalls = Text4.Text
rec!NCcalls = Text5.Text
rec!Namount = Text6.Text
MsgBox "Record saved"
rec.Update

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command4_Click()

Dim a1 As Integer, a2 As Integer, a3 As Integer
Dim x As Integer
a1 = Val(Text2.Text)
a2 = Val(Text3.Text)
a3 = Val(Text5.Text)
x = a2 - a1
Text4.Text = x
Dim s As String, s1 As Integer
s = rec2!Area
    If s = "urban" Then
    s1 = x - 50
    Text5.Text = s1
    End If
    If s = "rural" Then
    s1 = x - 75
    Text5.Text = s1
    End If
Dim c As Integer
  rec2.MoveFirst
  c = Text5.Text
  While rec2.EOF = False
    If rec2!Pno = Text1.Text Then
    Set k = rec2!Area
    If k = "rural" Then
    If c >= 1 And c <= 150 Then
    Text6.Text = am
    am = c * 0
    ElseIf c > 150 And c <= 400 Then
    am = 150 * 0 + (c - 150) * 0.8
    Text6.Text = am
    ElseIf c > 400 And c <= 500 Then
    am = 150 * 0 + 200 * 0.8 + (c - 400) * 1
    am = Val(Text6.Text)
    ElseIf c > 500 Then
    am = 150 * 0 + 200 * 0.8 + 100 * 1 + (c - 500) * 1.2
    Text6.Text = am
   End If
   End If
   End If
 rec2.MoveNext
 Wend

rec2.MoveFirst
   c = Text5.Text
  While rec2.EOF = False
    If rec2!Pno = Text1.Text Then
    Set k = rec2!Area
    If k = "urban" Then
    If c >= 1 And c <= 100 Then
    Text6.Text = am
    am = c * 0
    ElseIf c > 100 And c <= 400 Then
    am = 100 * 0 + (c - 100) * 0.8
    Text6.Text = am
    ElseIf c > 400 And c <= 500 Then
    am = 100 * 0 + 200 * 0.8 + (c - 400) * 1
    am = Text6.Text
    ElseIf c > 500 Then
    am = 100 * 0 + 200 * 0.8 + 100 * 1 + (c - 500) * 1.2
    Text6.Text = am
   End If
   End If
   End If
 rec2.MoveNext
 Wend
 
 Dim x1 As Integer, x2 As Integer, x3 As Integer
Dim q1, q2, q3, q4, q5, q6, q7 As String
rec1.MoveFirst
x1 = 0
While rec1.EOF = False
  If rec1!tno = Text1.Text Then
    q1 = rec1!STD
    q2 = rec1!ISD
    q3 = rec1!CLI
    q4 = rec1!Hotline
    q5 = rec1!Conf
    q6 = rec1!CF
    q7 = rec1!AD
    If q1 = "Yes" Then
      x1 = x1 + 1
    End If
    If q2 = "Yes" Then
      x1 = x1 + 1
    End If
    If q3 = "Yes" Then
      x1 = x1 + 1
    End If
    If q4 = "Yes" Then
      x1 = x1 + 1
    End If
    If q5 = "Yes" Then
      x1 = x1 + 1
    End If
    If q6 = "Yes" Then
      x1 = x1 + 1
    End If
    If q7 = "Yes" Then
      x1 = x1 + 1
    End If
    
    If x1 >= 3 Then
       x2 = 50
    ElseIf x1 = 1 Then
      x2 = 20
    ElseIf x1 = 2 Then
      x2 = 40
    
    End If
    MsgBox (x2)
    Text6.Text = Val(Text6.Text) + x2
 End If
 rec1.MoveNext
Wend
 
 
End Sub

Private Sub Form_Load()
 
    Set db = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set rec1 = New ADODB.Recordset
    Set rec2 = New ADODB.Recordset
    db.ConnectionString = "dsn=phone;uid=;pwd=;"
    db.Open
    rec.Open "bill_perpare", db, adOpenDynamic, adLockOptimistic
    rec1.Open "newconn1", db, adOpenDynamic, adLockOptimistic
    rec2.Open "payment", db, adOpenDynamic, adLockOptimistic

End Sub
Private Sub Text1_LostFocus()
Dim a As Integer
rec.MoveFirst
While Not rec1.EOF = True
  If Text1.Text = rec1!tno Then
  
   'Else
   'MsgBox "number Does Not Exist"
  End If
  rec1.MoveNext
Wend
End Sub
