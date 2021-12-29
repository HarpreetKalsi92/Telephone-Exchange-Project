VERSION 5.00
Begin VB.Form Form16 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form16"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11520
   LinkTopic       =   "Form16"
   ScaleHeight     =   8640
   ScaleWidth      =   11520
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   6720
      TabIndex        =   10
      Top             =   3240
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Draft"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   8280
      TabIndex        =   9
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
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
      Left            =   6720
      MaxLength       =   6
      TabIndex        =   8
      Text            =   " "
      Top             =   4800
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
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
      Left            =   6720
      TabIndex        =   7
      Text            =   " "
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pay/Submit"
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
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
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
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Re&set"
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
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
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
      Left            =   6720
      TabIndex        =   3
      Text            =   " "
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
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
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
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
      Left            =   6720
      TabIndex        =   1
      Text            =   " "
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
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
      Left            =   6720
      TabIndex        =   0
      Text            =   " "
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "  Form For Pay The Telephone Bill"
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
      TabIndex        =   17
      Top             =   120
      Width           =   6735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Telephone No"
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
      Left            =   1440
      TabIndex        =   16
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Total Payable Amout"
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
      Left            =   1320
      TabIndex        =   15
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Mode Of Payment"
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
      Left            =   1440
      TabIndex        =   14
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Payment Date"
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
      Left            =   1440
      TabIndex        =   13
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Draft No"
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
      Left            =   1440
      TabIndex        =   12
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Drawn On Bank"
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
      Left            =   1320
      TabIndex        =   11
      Top             =   5640
      Width           =   2055
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As ADODB.Connection
Dim rec As ADODB.Recordset
Dim rec1 As ADODB.Recordset
Private Sub display()
Dim a As String
a = rec!Mode
Text1.Text = rec!Tno
Text2.Text = rec!TPAmount
Text3.Text = rec!PDate
Text4.Text = rec!DNo
Text6.Text = rec!DBank
If a = "Cash" Then
Option1.Value = True
End If
If a = "Demand Draft" Then
Option2.Value = True
End If
End Sub
Private Sub Command1_Click()
rec.AddNew
If Option1.Value = True Then
a = "Cash"
Text4.Text = 0
Text6.Text = "Nil"
Option2.Enabled = False
End If
If Option2.Value = True Then
a = "Draft"
Option1.Enabled = False
End If
Text3.Text = Date
'rec!Mode = a
rec!Tno = Text1.Text
rec!TPAmount = Text2.Text
rec!PDate = Text3.Text
rec!DNo = Text4.Text
rec!DBank = Text6.Text
rec.Update
End Sub

Private Sub Command2_Click()
c = MsgBox("Are you sure to exit this program?", vbYesNo, "Confirm Box")
If c = vbYes Then
form13.Show
Unload Me
End If
End Sub

Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text6.Text = ""
Option1.Value = False
Option2.Value = False
Text1.SetFocus
End Sub
Private Sub Command4_Click()
Dim dat As Date
issue = Date
duedt = issue + 21
disdt = issue + 35
If MsgBox(issue, vbOKOnly, " Bill Issue Date") = vbOK Then
If MsgBox(duedt, vbOKOnly, " Bill can be pay upto Date Without surcharge") = vbOK Then
If MsgBox(disdt, vbOKOnly, "Bill can be pay upto this date with surcharges") = vbOK Then

End If
End If
End If
dat = Text3.Text
MsgBox (dat)
am = Val(Text2.Text)
If MsgBox(am, vbOKOnly, "The Actual Amount Is") = vbOK Then
End If

rec1.MoveFirst

While rec1.EOF = False
  If rec1!Tno = Text1.Text Then
    If dat >= issue And dat <= duedt Then
       MsgBox ("You Pay Your Bill  Before Last Date,So Here is no any Surcharge")
       charge = 0
     End If
     If dat > duedt And dat <= disdt Then
       MsgBox ("U Pay Your Bill After The Due Date,So Pay It With Surcharge")
       GoTo start
    End If
    If dat > disdt Then
      MsgBox ("UR Phone Has Disconnected Now")
  End If
start:
  If am >= 1 And am <= 500 Then
 charge = 10
 ElseIf am > 500 And am <= 1000 Then
 charge = 20
ElseIf am > 1000 And am <= 2000 Then
 charge = 40
ElseIf am > 2000 And am <= 3500 Then
charge = 70
ElseIf am > 3500 And am <= 5000 Then
charge = 100
ElseIf am > 5000 And am <= 7500 Then
charge = 150
ElseIf am > 7500 And am <= 10000 Then
charge = 200
ElseIf am > 10000 And am <= 20000 Then
charge = 400
ElseIf am > 20000 And am <= 50000 Then
charge = 1000
ElseIf am > 50000 Then
charge = 2000
End If
Text2.Text = am + charge
End If
rec1.MoveNext
Wend
End Sub



Private Sub Form_Load()
Set db = New ADODB.Connection
Set rec = New ADODB.Recordset
Set rec1 = New ADODB.Recordset
db.ConnectionString = "dsn=phone;uid=;pwd=;"
db.Provider = "Microsoft.jet.Oledb.4.0"
db.Open App.Path & "\Telephone.mdb"
rec.Open "pay_bill", db, adOpenDynamic, adLockOptimistic
rec1.Open "bill_Prepare", db, adOpenDynamic, adLockOptimistic
Option1.Value = True
Option2.Value = True
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
Text4.Text = 0
Text6.Text = "Nil"
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
Text4.Text = ""
Text6.Text = ""
'Text4.SetFocus
End If
End Sub

Private Sub Option2_LostFocus()
Text4.SetFocus
End Sub

Private Sub Text1_LostFocus()
rec1.MoveFirst
While rec1.EOF = False
If rec1!Tno = Text1.Text Then
Text2.Text = rec1!Namount
Text3.Text = dat
End If
rec1.MoveNext
Wend
End Sub


