VERSION 5.00
Begin VB.Form LOGIN 
   BackColor       =   &H00FFC0C0&
   Caption         =   "LOGIN"
   ClientHeight    =   2070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6255
   BeginProperty Font 
      Name            =   "Papyrus"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2274.95
   ScaleMode       =   0  'User
   ScaleWidth      =   7999.183
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "GO>>"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   10
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GO>>"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ADMINISTRATOR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   3000
      TabIndex        =   9
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "NEW USER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   4800
      TabIndex        =   8
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Please login as new user"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label Labelerr 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Invalid username and password!!!!!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "USERNAME"
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
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   2055
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "LOGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Call connect
Dim rs As ADODB.Recordset
 Set rs = New ADODB.Recordset
    
rs.Open "select * from login where username='" & Text1.Text & "' and password ='" & Text2.Text & "'", cn, adOpenStatic, adLockOptimistic


If rs.EOF = False Then
HOME.Show
Unload Me
Else
Labelerr.Visible = True
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus

End If
rs.Close

End Sub


Private Sub Command2_Click()
MsgBox "THANK YOU", vbExclamation
End

End Sub

Private Sub Command3_Click()
Call connect

Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
    
rs.Open "select * from login where username='" & Text1.Text & "' and password ='" & Text2.Text & "'", cn, adOpenStatic, adLockOptimistic


If rs.EOF = False Then
Form1.Show
Unload Me
Else
Labelerr.Visible = True
Text2.Text = ""
Text2.SetFocus

End If
rs.Close

End Sub

Private Sub Form_Load()
Command1.Visible = True
Command3.Visible = False
End Sub

Private Sub Label3_Click()
Label4.Visible = False
Label5.Visible = False

Text1.Text = "administrator"
Text1.Enabled = False
MsgBox "Please Enter password", vbExclamation
Text2.SetFocus
Command1.Visible = False
Command3.Visible = True

End Sub

Private Sub Label5_Click()
Label4.Visible = True
Label3.Visible = False

Text1.Text = "new"
Text2.Text = "lic"
Text1.Enabled = False
Text2.Enabled = False
Command1.SetFocus

End Sub

