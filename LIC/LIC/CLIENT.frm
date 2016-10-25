VERSION 5.00
Begin VB.Form CLIENT 
   BackColor       =   &H00FFC0C0&
   Caption         =   "CLIENT"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   FillColor       =   &H00FFC0C0&
   ForeColor       =   &H00FFC0C0&
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture5 
      Height          =   4575
      Left            =   15960
      Picture         =   "CLIENT.frx":0000
      ScaleHeight     =   4515
      ScaleWidth      =   4155
      TabIndex        =   39
      Top             =   3960
      Width           =   4215
   End
   Begin VB.TextBox Textpno 
      Height          =   285
      Left            =   2400
      TabIndex        =   38
      Text            =   "Text2"
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Textaid 
      Height          =   285
      Left            =   2640
      TabIndex        =   37
      Text            =   "Text1"
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Textcid 
      Height          =   285
      Left            =   2880
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PROCEED>>"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12720
      TabIndex        =   14
      Top             =   9840
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000E&
      Caption         =   "<<BACK TO HOME"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      MaskColor       =   &H00FFFFC0&
      TabIndex        =   15
      Top             =   9840
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "CLIENT INFORMATION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   6615
      Left            =   4440
      TabIndex        =   19
      Top             =   3120
      Width           =   10815
      Begin VB.ComboBox cmbsex 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2160
         TabIndex        =   4
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox cpin 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8040
         TabIndex        =   12
         Top             =   4320
         Width           =   2295
      End
      Begin VB.TextBox cnation 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   13
         Top             =   5160
         Width           =   2415
      End
      Begin VB.TextBox cemail 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   11
         Top             =   4320
         Width           =   4815
      End
      Begin VB.TextBox coff 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8040
         TabIndex        =   10
         Top             =   3720
         Width           =   2295
      End
      Begin VB.TextBox cmob 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   9
         Top             =   3720
         Width           =   2535
      End
      Begin VB.TextBox cres 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2160
         TabIndex        =   8
         Top             =   3000
         Width           =   8175
      End
      Begin VB.TextBox flname 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7680
         TabIndex        =   7
         Top             =   1920
         Width           =   2655
      End
      Begin VB.TextBox fmname 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   6
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox ffname 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2160
         TabIndex        =   5
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox clname 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   7680
         TabIndex        =   3
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox cmname 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   4920
         TabIndex        =   2
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox cfname 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2160
         TabIndex        =   1
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFC0C0&
         Caption         =   "*PINCODE"
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
         Left            =   7080
         TabIndex        =   35
         Top             =   4440
         Width           =   1935
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFC0C0&
         Caption         =   "OFFICE"
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
         Left            =   7200
         TabIndex        =   34
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MOBILE NO"
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
         Left            =   2160
         TabIndex        =   33
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFC0C0&
         Caption         =   "LAST NAME"
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
         Left            =   8760
         TabIndex        =   32
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIDDLE NAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6000
         TabIndex        =   31
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "FIRST NAME"
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
         Left            =   2880
         TabIndex        =   30
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFC0C0&
         Caption         =   "LAST NAME"
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
         Left            =   8640
         TabIndex        =   29
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIDDLE NAME"
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
         TabIndex        =   28
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFC0C0&
         Caption         =   "FIRST NAME"
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
         Left            =   2880
         TabIndex        =   27
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "*NATIONALITY"
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
         Left            =   240
         TabIndex        =   26
         Top             =   5160
         Width           =   2415
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "EMAIL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   25
         Top             =   4440
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "TELEPHONE NO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   24
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "*RESIDENTIAL ADD"
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
         TabIndex        =   23
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "*FATHER'S FULLNAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   22
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "*SEX"
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
         TabIndex        =   21
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "*FULLNAME"
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
         TabIndex        =   20
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture4 
      Height          =   3015
      Left            =   360
      ScaleHeight     =   2955
      ScaleWidth      =   3195
      TabIndex        =   18
      Top             =   6360
      Width           =   3255
      Begin VB.Image Image2 
         Height          =   3600
         Left            =   0
         Picture         =   "CLIENT.frx":1C2E7
         Top             =   0
         Width           =   4800
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   2775
      Left            =   360
      ScaleHeight     =   2715
      ScaleWidth      =   3195
      TabIndex        =   17
      Top             =   3120
      Width           =   3255
      Begin VB.Image Image3 
         Height          =   2940
         Left            =   0
         Picture         =   "CLIENT.frx":1FBEC
         Top             =   0
         Width           =   4815
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   1935
      Left            =   4080
      ScaleHeight     =   1875
      ScaleWidth      =   11595
      TabIndex        =   16
      Top             =   600
      Width           =   11655
      Begin VB.Image Image1 
         Height          =   2400
         Left            =   2160
         Picture         =   "CLIENT.frx":22E18
         Top             =   0
         Width           =   6450
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   480
      ScaleHeight     =   675
      ScaleWidth      =   1395
      TabIndex        =   0
      Top             =   600
      Width           =   1455
      Begin VB.Image Image4 
         Height          =   765
         Left            =   0
         Picture         =   "CLIENT.frx":29432
         Top             =   0
         Width           =   1470
      End
   End
End
Attribute VB_Name = "CLIENT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
HOME.Show
End Sub

Private Sub Command2_Click()
If (cfname(1).Text = "" Or cmname(2).Text = "" Or clname(3).Text = "" Or cmbsex.Text = "" Or ffname.Text = "" Or fmname.Text = "" Or flname.Text = "" Or cres.Text = "" Or cpin.Text = "" Or cnation.Text = "") Then
  MsgBox "All * marked fields are compulsory!!!!", vbExclamation
Else
  Call max
  cn.Execute ("insert into client_info values(" + Textcid.Text + ",'" + cfname(1).Text + "','" + cmname(2).Text + "','" + clname(3).Text + "','" + cmbsex.Text + "','" + ffname.Text + "','" + fmname.Text + "','" + flname.Text + "','" + cres.Text + "'," + cmob.Text + "," + coff.Text + ",'" + cemail.Text + "'," + cpin.Text + ",'" + cnation.Text + "'," + Textaid.Text + "," + Textpno.Text + ")")
        Unload Me
        POLICY1.Show
    End If
End Sub

Private Sub Form_Load()
cmbsex.AddItem "Male"
cmbsex.AddItem "Female"
    Call connect
End Sub

Public Function max()
    Dim clientid As Integer
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "select max(client_id) from client_info", cn, adOpenStatic, adLockOptimistic
    If IsNull(rs.Fields(0)) Then
        clientid = 1
        Textcid.Text = clientid
    Else
        clientid = rs.Fields(0) + 1
        Textcid.Text = clientid
    End If
    Textpno.Text = Textcid + 1000
    Textaid.Text = Textcid + 3000
    rs.Close

End Function



Private Sub cfname_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8) Then
        KeyAscii = 0
        MsgBox "Please Enter character Only!!!!!", vbExclamation
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

End Sub



Private Sub clname_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8) Then
        KeyAscii = 0
        MsgBox "Please Enter character Only!!!!!", vbExclamation
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub cmname_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8) Then
        KeyAscii = 0
        MsgBox "Please Enter character Only!!!!!", vbExclamation
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub



Private Sub cmob_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8) Then
        KeyAscii = 0
        MsgBox "Please Enter Integer Only!!!!!", vbExclamation
    End If
End Sub

Private Sub cnation_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8) Then
        KeyAscii = 0
        MsgBox "Please Enter character Only!!!!!", vbExclamation
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub


Private Sub cpin_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8) Then
        KeyAscii = 0
        MsgBox "Please Enter Integer Only!!!!!", vbExclamation
    End If
End Sub


Private Sub ffname_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8) Then
        KeyAscii = 0
        MsgBox "Please Enter character Only!!!!!", vbExclamation
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub



Private Sub flname_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8) Then
        KeyAscii = 0
        MsgBox "Please Enter character Only!!!!!", vbExclamation
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub fmname_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8) Then
        KeyAscii = 0
        MsgBox "Please Enter character Only!!!!!", vbExclamation
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub



Private Sub coff_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8) Then
        KeyAscii = 0
        MsgBox "Please Enter Integer Only!!!!!", vbExclamation
    End If
End Sub

