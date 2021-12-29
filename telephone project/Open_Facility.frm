VERSION 5.00
Begin VB.Form Forma1 
   BackColor       =   &H00C0FFC0&
   Caption         =   "open"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10575
   LinkTopic       =   "Form5"
   ScaleHeight     =   7830
   ScaleWidth      =   10575
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command8 
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
      Left            =   5760
      TabIndex        =   35
      Top             =   7680
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Control Button"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5520
      TabIndex        =   27
      Top             =   6960
      Width           =   4335
      Begin VB.CommandButton Command7 
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
         Height          =   495
         Left            =   3120
         TabIndex        =   34
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command6 
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
         Left            =   2160
         TabIndex        =   33
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Save"
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
         Left            =   1200
         TabIndex        =   32
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   1335
      Left            =   600
      TabIndex        =   26
      Top             =   6960
      Width           =   4335
      Begin VB.CommandButton Command4 
         Caption         =   "&Last"
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
         Left            =   3240
         TabIndex        =   31
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Previous"
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
         Left            =   2280
         TabIndex        =   30
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Next"
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
         Left            =   1320
         TabIndex        =   29
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&First"
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
         TabIndex        =   28
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00C0E0FF&
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
      TabIndex        =   25
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00C0E0FF&
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
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0E0FF&
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
      Left            =   3720
      TabIndex        =   12
      Text            =   " "
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0E0FF&
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
      Left            =   2280
      TabIndex        =   11
      Text            =   " "
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0E0FF&
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
      Left            =   6120
      TabIndex        =   10
      Text            =   " "
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0E0FF&
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
      Left            =   2280
      TabIndex        =   9
      Text            =   " "
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0E0FF&
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
      Left            =   6120
      TabIndex        =   8
      Text            =   " "
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "STD"
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
      Left            =   840
      TabIndex        =   7
      Top             =   3960
      Width           =   855
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "ISD"
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
      Left            =   1920
      TabIndex        =   6
      Top             =   3960
      Width           =   735
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "CLI"
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
      Left            =   2640
      TabIndex        =   5
      Top             =   3960
      Width           =   855
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Hotline"
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
      Left            =   3600
      TabIndex        =   4
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Conferencing"
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
      Left            =   4800
      TabIndex        =   3
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CheckBox Check6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Call Forwarding"
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
      Left            =   6360
      TabIndex        =   2
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CheckBox Check7 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Abbreviated Call"
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
      Left            =   8400
      TabIndex        =   1
      Top             =   3960
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
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
      Left            =   3720
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "   Form  For Add On Facility Opening  "
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
      Left            =   2760
      TabIndex        =   23
      Top             =   0
      Width           =   6975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   " Telephone No"
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
      Left            =   120
      TabIndex        =   22
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   " Name Of Customer"
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
      Left            =   0
      TabIndex        =   21
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Caption         =   " Address Where The Phone Is Working"
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
      Left            =   0
      TabIndex        =   20
      Top             =   2040
      Width           =   3855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Caption         =   " House No."
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
      Left            =   360
      TabIndex        =   19
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Caption         =   " Street/Road"
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
      Left            =   4320
      TabIndex        =   18
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      Caption         =   " City/District"
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
      Left            =   240
      TabIndex        =   17
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      Caption         =   " PinCode"
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
      Left            =   4320
      TabIndex        =   16
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Facility To Be Opened"
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
      Left            =   120
      TabIndex        =   15
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Date"
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
      Left            =   120
      TabIndex        =   14
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Amount  During Facility Open"
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
      Left            =   0
      TabIndex        =   13
      Top             =   5160
      Width           =   2895
   End
End
Attribute VB_Name = "Forma1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As ADODB.Connection
Dim rec As ADODB.Recordset
Dim counter As Integer

Private Sub Command1_Click()
rec.MoveFirst
Text1.Text = rec!Tno
Text2.Text = rec!cname
Text3.Text = rec!Hno
Text4.Text = rec!ST
Text5.Text = rec!City
Text6.Text = rec!Pin
Text8.Text = rec!Later_facility
Dim b1, b2, b3, b4, b5, b6, b7 As String
b1 = rec!STD
b2 = rec!ISD
b3 = rec!CLI
b4 = rec!Hotline
b5 = rec!Conf
b6 = rec!CF
b7 = rec!AD
  If b1 = "Yes" Then
  Check1.Value = 1
   Else
   Check1.Value = 0
End If
If b2 = "Yes" Then
  Check2.Value = 1
   Else
   Check2.Value = 0
End If
If b3 = "Yes" Then
  Check3.Value = 1
   Else
   Check3.Value = 0
End If
If b4 = "Yes" Then
  Check4.Value = 1
   Else
   Check4.Value = 0
End If
If b5 = "Yes" Then
  Check5.Value = 1
   Else
   Check5.Value = 0
End If
If b6 = "Yes" Then
  Check6.Value = 1
   Else
   Check6.Value = 0
End If
If b7 = "Yes" Then
  Check7.Value = 1
   Else
   Check7.Value = 0
End If

End Sub

Private Sub Command2_Click()
rec.MoveNext
If rec.EOF = True Then
rec.MoveLast
End If
Text1.Text = rec!Tno
Text2.Text = rec!cname
Text3.Text = rec!Hno
Text4.Text = rec!ST
Text5.Text = rec!City
Text6.Text = rec!Pin
'Text8.Text = rec!Later_facility
Dim b1, b2, b3, b4, b5, b6, b7 As String
b1 = rec!STD
b2 = rec!ISD
b3 = rec!CLI
b4 = rec!Hotline
b5 = rec!Conf
b6 = rec!CF
b7 = rec!AD
  If b1 = "Yes" Then
  Check1.Value = 1
   Else
   Check1.Value = 0
End If
If b2 = "Yes" Then
  Check2.Value = 1
   Else
   Check2.Value = 0
End If
If b3 = "Yes" Then
  Check3.Value = 1
   Else
   Check3.Value = 0
End If
If b4 = "Yes" Then
  Check4.Value = 1
   Else
   Check4.Value = 0
End If
If b5 = "Yes" Then
  Check5.Value = 1
   Else
   Check5.Value = 0
End If
If b6 = "Yes" Then
  Check6.Value = 1
   Else
   Check6.Value = 0
End If
If b7 = "Yes" Then
  Check7.Value = 1
   Else
   Check7.Value = 0
End If


End Sub

Private Sub Command3_Click()
rec.MovePrevious
If rec.BOF = True Then
rec.MoveFirst
End If
Text1.Text = rec!Tno
Text2.Text = rec!cname
Text3.Text = rec!Hno
Text4.Text = rec!ST
Text5.Text = rec!City
Text6.Text = rec!Pin
'Text8.Text = rec!Later_facility
Dim b1, b2, b3, b4, b5, b6, b7 As String
b1 = rec!STD
b2 = rec!ISD
b3 = rec!CLI
b4 = rec!Hotline
b5 = rec!Conf
b6 = rec!CF
b7 = rec!AD
  If b1 = "Yes" Then
  Check1.Value = 1
   Else
   Check1.Value = 0
End If
If b2 = "Yes" Then
  Check2.Value = 1
   Else
   Check2.Value = 0
End If
If b3 = "Yes" Then
  Check3.Value = 1
   Else
   Check3.Value = 0
End If
If b4 = "Yes" Then
  Check4.Value = 1
   Else
   Check4.Value = 0
End If
If b5 = "Yes" Then
  Check5.Value = 1
   Else
   Check5.Value = 0
End If
If b6 = "Yes" Then
  Check6.Value = 1
   Else
   Check6.Value = 0
End If
If b7 = "Yes" Then
  Check7.Value = 1
   Else
   Check7.Value = 0
End If


End Sub

Private Sub Command4_Click()
rec.MoveLast
Text1.Text = rec!Tno
Text2.Text = rec!cname
Text3.Text = rec!Hno
Text4.Text = rec!ST
Text5.Text = rec!City
Text6.Text = rec!Pin
Text8.Text = rec!Later_facility
Dim b1, b2, b3, b4, b5, b6, b7 As String
b1 = rec!STD
b2 = rec!ISD
b3 = rec!CLI
b4 = rec!Hotline
b5 = rec!Conf
b6 = rec!CF
b7 = rec!AD
  If b1 = "Yes" Then
  Check1.Value = 1
   Else
   Check1.Value = 0
End If
If b2 = "Yes" Then
  Check2.Value = 1
   Else
   Check2.Value = 0
End If
If b3 = "Yes" Then
  Check3.Value = 1
   Else
   Check3.Value = 0
End If
If b4 = "Yes" Then
  Check4.Value = 1
   Else
   Check4.Value = 0
End If
If b5 = "Yes" Then
  Check5.Value = 1
   Else
   Check5.Value = 0
End If
If b6 = "Yes" Then
  Check6.Value = 1
   Else
   Check6.Value = 0
End If
If b7 = "Yes" Then
  Check7.Value = 1
   Else
   Check7.Value = 0
End If

End Sub

Private Sub Command5_Click()
Dim a As Integer
a = Text1.Text
rec.MoveFirst
While Not rec.EOF = True
  If a = rec!Tno Then
    rec!Tno = Text1.Text
    rec!cname = Text2.Text
    rec!Hno = Text3.Text
    rec!ST = Text4.Text
    rec!City = Text5.Text
    rec!Pin = Text6.Text
    rec!Later_facility = Text8.Text
    counter = 0
    If Check1.Value = 1 Then
      rec!STD = "Yes"
      counter = counter + 1
    Else
      rec!STD = "No"
    End If
    If Check2.Value = 1 Then
      rec!ISD = "Yes"
      counter = counter + 1
    Else
      rec!ISD = "No"
    End If
    If Check3.Value = 1 Then
      counter = counter + 1
      rec!CLI = "Yes"
    Else
      rec!CLI = "No"
    End If
    If Check4.Value = 1 Then
      rec!Hotline = "Yes"
      counter = counter + 1
    Else
      rec!Hotline = "No"
    End If
    If Check5.Value = 1 Then
      rec!Conf = "Yes"
      counter = counter + 1
    Else
      rec!Conf = "No"
    End If
    If Check6.Value = 1 Then
      rec!CF = "Yes"
      counter = counter + 1
    Else
      rec!CF = "No"
    End If
    If Check7.Value = 1 Then
    counter = counter + 1
      rec!AD = "Yes"
    Else
      rec!AD = "No"
    End If
    If counter = 0 Then
        Text8.Text = 0
    End If
    If counter = 1 Then
    Text8.Text = 30
    End If
    If counter > 1 Then
    Text8.Text = 50
    End If
    MsgBox "record saved"
    rec.Update
  End If
  rec.MoveNext
Wend
Text7.Text = Date

End Sub

Private Sub Command6_Click()
Unload Me

End Sub



Private Sub Command7_Click()
form13.Show
Unload Me
End Sub

Private Sub Command8_Click()
Dim a As Integer
a = Text1.Text
rec.MoveFirst
While Not rec.EOF = True
  If a = rec!Tno Then
    rec!Tno = Text1.Text
    rec!cname = Text2.Text
    rec!Hno = Text3.Text
    rec!ST = Text4.Text
    rec!City = Text5.Text
    rec!Pin = Text6.Text
     counter = 0
    If Check1.Value = 1 Then
      rec!STD = "Yes"
      counter = counter + 1
    Else
      rec!STD = "No"
    End If
    If Check2.Value = 1 Then
      rec!ISD = "Yes"
      counter = counter + 1
    Else
      rec!ISD = "No"
    End If
    If Check3.Value = 1 Then
      counter = counter + 1
      rec!CLI = "Yes"
    Else
      rec!CLI = "No"
    End If
    If Check4.Value = 1 Then
      rec!Hotline = "Yes"
      counter = counter + 1
    Else
      rec!Hotline = "No"
    End If
    If Check5.Value = 1 Then
      rec!Conf = "Yes"
      counter = counter + 1
    Else
      rec!Conf = "No"
    End If
    If Check6.Value = 1 Then
      rec!CF = "Yes"
      counter = counter + 1
    Else
      rec!CF = "No"
    End If
    If Check7.Value = 1 Then
    counter = counter + 1
      rec!AD = "Yes"
    Else
      rec!AD = "No"
    End If
    If counter = 0 Then
        Text8.Text = "0"
    End If
    If counter = 1 Then
    Text8.Text = 30
    End If
    If counter > 1 Then
    Text8.Text = 50
    End If
    End If
  rec.MoveNext
Wend
Text7.Text = Date
End Sub

Private Sub Form_Load()
Set db = New ADODB.Connection
Set rec = New ADODB.Recordset
db.ConnectionString = "dsn=phone;uid=;pwd=;"

db.Provider = "Microsoft.jet.Oledb.4.0"
db.Open App.Path & "\Telephone.mdb"

rec.Open "newconn1", db, adOpenDynamic, adLockOptimistic
Text7.Text = Date
End Sub

Private Sub Text1_LostFocus()
rec.MoveFirst
Dim a As Integer

While Not rec.EOF = True
If Text1.Text = rec!Tno Then
Text2.Text = rec!cname
Text3.Text = rec!Hno
Text4.Text = rec!ST
Text5.Text = rec!City
Text6.Text = rec!Pin
Dim s1, s2, s3, s4, s5, s6, s7 As String
s1 = rec!STD
s2 = rec!ISD
s3 = rec!CLI
s4 = rec!Hotline
s5 = rec!Conf
s6 = rec!CF
s7 = rec!AD
If s1 = "Yes" Then
Check1.Value = 1
Else
Check1.Value = 0
End If

If s2 = "Yes" Then
Check2.Value = 1
Else
Check2.Value = 0
End If
If s3 = "Yes" Then
Check3.Value = 1
Else
Check3.Value = 0
End If
If s4 = "Yes" Then
Check4.Value = 1
Else
Check4.Value = 0
End If
If s5 = "Yes" Then
Check5.Value = 1
Else
Check5.Value = 0
End If
If s6 = "Yes" Then
Check6.Value = 1
Else
Check6.Value = 0
End If
If s7 = "Yes" Then
Check7.Value = 1
Else
Check7.Value = 0
End If
End If
rec.MoveNext
Wend
End Sub


