VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form15 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Form15"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form15"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   495
      Left            =   10680
      TabIndex        =   37
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   495
      Left            =   8280
      TabIndex        =   36
      Top             =   7920
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   9840
      TabIndex        =   34
      Top             =   4200
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   20709377
      CurrentDate     =   39300
   End
   Begin VB.TextBox Text18 
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
      Left            =   9840
      TabIndex        =   33
      Top             =   7320
      Width           =   1935
   End
   Begin VB.TextBox Text17 
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
      Left            =   3840
      TabIndex        =   32
      Top             =   7560
      Width           =   2055
   End
   Begin VB.TextBox Text16 
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
      Left            =   3840
      TabIndex        =   31
      Top             =   6360
      Width           =   2055
   End
   Begin VB.TextBox Text15 
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
      Left            =   9840
      TabIndex        =   30
      Top             =   6720
      Width           =   1935
   End
   Begin VB.TextBox Text13 
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
      Left            =   9840
      TabIndex        =   29
      Top             =   6120
      Width           =   1935
   End
   Begin VB.TextBox Text12 
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
      Left            =   3840
      TabIndex        =   28
      Top             =   5760
      Width           =   2055
   End
   Begin VB.TextBox Text11 
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
      Left            =   9840
      TabIndex        =   27
      Top             =   5280
      Width           =   1935
   End
   Begin VB.TextBox Text10 
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
      Left            =   3840
      TabIndex        =   26
      Top             =   5160
      Width           =   2055
   End
   Begin VB.TextBox Text9 
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
      Left            =   9840
      TabIndex        =   25
      Top             =   4800
      Width           =   1935
   End
   Begin VB.TextBox Text8 
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
      Left            =   3840
      TabIndex        =   24
      Top             =   4560
      Width           =   2055
   End
   Begin VB.TextBox Text7 
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
      Left            =   3840
      TabIndex        =   23
      Top             =   3840
      Width           =   2055
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
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Text            =   " "
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox Text3 
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
      Left            =   3840
      TabIndex        =   4
      Text            =   " "
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox Text5 
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
      Left            =   3840
      TabIndex        =   3
      Text            =   " "
      ToolTipText     =   "In MM-DD-YY Format "
      Top             =   3240
      Width           =   2055
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
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Text            =   " "
      Top             =   2280
      Width           =   2055
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
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "B&ack"
      Height          =   495
      Left            =   9480
      TabIndex        =   0
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFC0&
      Caption         =   " Net Chargable Calls"
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
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   5760
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "     Form For Showing The Bill To Subscriber"
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
      Height          =   540
      Left            =   1200
      TabIndex        =   22
      Top             =   0
      Width           =   7515
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Subscriber's Name"
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
      Left            =   120
      TabIndex        =   21
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   " Subscriber's Address"
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
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Caption         =   " Subsciber's Father Name"
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
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Caption         =   " Phone No:-"
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
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Caption         =   " Bill  Issue Date"
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
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Period Of Bill "
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
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      Caption         =   " Payable Upto"
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
      Height          =   255
      Left            =   6600
      TabIndex        =   15
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Previous Meter Reading"
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
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   4560
      Width           =   2775
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFC0&
      Caption         =   " Current Meter  Reading"
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
      Height          =   255
      Left            =   6480
      TabIndex        =   13
      Top             =   4920
      Width           =   2535
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Metered Calls"
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
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFC0&
      Caption         =   " Free Calls"
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
      Height          =   255
      Left            =   6480
      TabIndex        =   11
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Rent For Two Month"
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
      Height          =   255
      Left            =   6480
      TabIndex        =   10
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Service Tax @ 5% Of Total"
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
      Height          =   255
      Left            =   6480
      TabIndex        =   9
      Top             =   6720
      Width           =   2535
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0FFC0&
      Caption         =   " Total Payable Before Date"
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
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   6480
      Width           =   2535
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0FFC0&
      Caption         =   " Surcharges"
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
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   7680
      Width           =   1695
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Bill Payable After Due Date"
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
      Left            =   6480
      TabIndex        =   6
      Top             =   7320
      Width           =   2415
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As ADODB.Connection
Dim rec As ADODB.Recordset
Dim rec1 As ADODB.Recordset
Dim rec2 As ADODB.Recordset
Dim rec3 As ADODB.Recordset
Dim a3 As Integer

Private Sub Command1_Click()
Form9.Show
End Sub

Private Sub Command2_Click()
rec3.AddNew
rec3!Pno = Text1.Text
rec3!sname = Text2.Text
rec3!SAddress = Text3.Text
rec3!SFname = Text4.Text
rec3!Bill_Issue_Date = Text5.Text
rec3!Period_of_bill = Text7.Text
rec3!previous_meter_reading = Text8.Text
rec3!current_meter_reading = Text9.Text
rec3!metered_Calls = Text10.Text
rec3!free_Calls = Text11.Text
rec3!netchargable_calls = Text12.Text
rec3!service_tax = Text15.Text
rec3!total_payable_before_date = Text16.Text
rec3!Surcharge = Text17.Text
'rec3!Bill_Payable_after_due_date = Text18.Text
rec3!payable_upto = DTPicker1.Value
MsgBox "Record Saved"
rec3.Update
End Sub

Private Sub Command3_Click()
form13.Show
Unload Me
End Sub

Private Sub Form_Load()
Set db = New ADODB.Connection
Set rec = New ADODB.Recordset
Set rec1 = New ADODB.Recordset
Set rec2 = New ADODB.Recordset
Set rec3 = New ADODB.Recordset
db.ConnectionString = "dsn=phone;uid=;pwd=;"
db.Open
rec.Open "newconn1", db, adOpenDynamic, adLockOptimistic
rec1.Open "bill_prepare", db, adOpenDynamic, adLockOptimistic
rec2.Open "payment", db, adOpenDynamic, adLockOptimistic
rec3.Open "show_bill", db, adOpenDynamic, adLockOptimistic

End Sub



Private Sub Text1_LostFocus()
rec.MoveFirst
Dim a As Integer
a = Text1.Text
While Not rec.EOF = True
  If a = rec!Tno Then
   Text2.Text = rec!cname
   Text3.Text = rec!caddress
   Text4.Text = rec!fgname
End If
  rec.MoveNext
Wend
rec1.MoveFirst
Dim a1 As Integer
a1 = Text1.Text
While Not rec1.EOF = True
If a1 = rec1!Tno Then
Text5.Text = Date
Text7.Text = "Two Months"
Text8.Text = rec1!pmr
Text9.Text = rec1!cmr
Text10.Text = rec1!Mcalls
Text12.Text = rec1!nccalls
Text16.Text = rec1!Namount
End If
rec1.MoveNext
Wend
Dim s As String, s1 As Integer
Dim s2 As Integer
  rec2.MoveFirst
  s2 = Text1.Text
  While rec2.EOF = False
    If rec2!Pno = Text1.Text Then
 s = rec2!Area
    If s = "urban" Then
    s1 = 100
    Text11.Text = s1
    Text13.Text = 300
    
   
a3 = Text16.Text
Text15.Text = a3 * (5 / 100)
    Text16.Text = Val(Text16.Text) + Val(Text13.Text) + Val(Text15.Text)
    Text17.Text = a3 * (10 / 100)
    Text18.Text = Val(Text16.Text) + Val(Text17.Text)
    End If
    If s = "rural" Then
    s1 = 150
    Text11.Text = s1
    Text13.Text = 220
    
    
a3 = Text16.Text
Text15.Text = a3 * (5 / 100)
    Text16.Text = Val(Text16.Text) + Val(Text13.Text) + Val(Text15.Text)
    Text17.Text = a3 * (10 / 100)
    Text18.Text = Val(Text16.Text) + Val(Text17.Text)
    End If
    
    End If

rec2.MoveNext
Wend
Text12.Text = Val(Text10.Text) - Val(Text11.Text)
End Sub

