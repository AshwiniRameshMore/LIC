VERSION 5.00
Begin VB.Form frmclient 
   Caption         =   "Insurance Form"
   ClientHeight    =   3090
   ClientLeft      =   450
   ClientTop       =   615
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11520
   ScaleWidth      =   19080
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtpno 
      Height          =   195
      Left            =   12360
      TabIndex        =   35
      Top             =   10440
      Width           =   150
   End
   Begin VB.TextBox txtaid 
      Height          =   195
      Left            =   12120
      TabIndex        =   34
      Top             =   10440
      Width           =   150
   End
   Begin VB.TextBox txtcid 
      Height          =   195
      Left            =   11880
      TabIndex        =   33
      Top             =   10440
      Width           =   150
   End
   Begin VB.CommandButton cmdhome 
      Caption         =   "<< &Back To Home"
      Height          =   495
      Left            =   6360
      TabIndex        =   32
      Top             =   10200
      Width           =   2295
   End
   Begin VB.CommandButton cmdproceed 
      Caption         =   "&Proceed >>"
      Height          =   495
      Left            =   15360
      TabIndex        =   25
      Top             =   10200
      Width           =   2295
   End
   Begin VB.Frame frmcinfo 
      Caption         =   "Client Information"
      ForeColor       =   &H000000FF&
      Height          =   6855
      Left            =   6360
      TabIndex        =   0
      Top             =   3240
      Width           =   11295
      Begin VB.TextBox txtfn 
         Height          =   285
         Left            =   2280
         TabIndex        =   31
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtnation 
         Height          =   285
         Left            =   2280
         TabIndex        =   26
         Top             =   6120
         Width           =   1935
      End
      Begin VB.TextBox txtfmn 
         Height          =   285
         Left            =   5280
         TabIndex        =   23
         Top             =   2520
         Width           =   2415
      End
      Begin VB.TextBox txtffn 
         Height          =   285
         Left            =   2280
         TabIndex        =   22
         Top             =   2520
         Width           =   2415
      End
      Begin VB.TextBox txtfsn 
         Height          =   285
         Left            =   8280
         TabIndex        =   21
         Top             =   2520
         Width           =   2415
      End
      Begin VB.ComboBox cmbsex 
         Height          =   315
         ItemData        =   "Form2.frx":0000
         Left            =   2280
         List            =   "Form2.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtpin 
         Height          =   315
         Left            =   8640
         TabIndex        =   17
         Top             =   5280
         Width           =   2055
      End
      Begin VB.TextBox txtemail 
         Height          =   315
         Left            =   2280
         TabIndex        =   15
         Top             =   5280
         Width           =   4455
      End
      Begin VB.TextBox txtoff 
         Height          =   315
         Left            =   8640
         TabIndex        =   13
         Top             =   4440
         Width           =   2055
      End
      Begin VB.TextBox txtmob 
         Height          =   315
         Left            =   3240
         TabIndex        =   12
         Top             =   4440
         Width           =   1935
      End
      Begin VB.TextBox txtra 
         Height          =   285
         Left            =   2280
         TabIndex        =   8
         Top             =   3480
         Width           =   8415
      End
      Begin VB.TextBox txtmn 
         Height          =   285
         Left            =   5280
         TabIndex        =   3
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtsn 
         Height          =   285
         Left            =   8280
         TabIndex        =   2
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "Last Name"
         Height          =   255
         Left            =   8280
         TabIndex        =   30
         Top             =   2880
         Width           =   2415
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "First Name"
         Height          =   255
         Left            =   2160
         TabIndex        =   29
         Top             =   2880
         Width           =   2415
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Middle Name"
         Height          =   255
         Left            =   5160
         TabIndex        =   28
         Top             =   2880
         Width           =   2415
      End
      Begin VB.Label fathertxt 
         Caption         =   "* Father's Full Name :"
         Height          =   495
         Left            =   480
         TabIndex        =   24
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label17 
         Caption         =   "* Nationality :"
         Height          =   375
         Left            =   480
         TabIndex        =   20
         Top             =   6120
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "* Sex :"
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "* Pin Code :"
         Height          =   375
         Left            =   7320
         TabIndex        =   16
         Top             =   5280
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "E-mail :"
         Height          =   255
         Left            =   600
         TabIndex        =   14
         Top             =   5280
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Office :"
         Height          =   375
         Left            =   7440
         TabIndex        =   11
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Mobile :"
         Height          =   375
         Left            =   2280
         TabIndex        =   10
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Telephone No:"
         Height          =   375
         Left            =   600
         TabIndex        =   9
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "* Residencial Address :"
         Height          =   375
         Left            =   480
         TabIndex        =   7
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Middle Name"
         Height          =   255
         Left            =   5160
         TabIndex        =   6
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "First Name"
         Height          =   255
         Left            =   2160
         TabIndex        =   5
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Last Name"
         Height          =   255
         Left            =   8160
         TabIndex        =   4
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "* Full Name :"
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Label Label18 
      Caption         =   "PROPOSAL  FOR  INSURANCE  ON  OWN  LIFE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   8520
      TabIndex        =   27
      Top             =   2760
      Width           =   6855
   End
   Begin VB.Image Image4 
      Height          =   765
      Left            =   360
      Picture         =   "Form2.frx":001C
      Top             =   240
      Width           =   1470
   End
   Begin VB.Image Image3 
      Height          =   2940
      Left            =   360
      Picture         =   "Form2.frx":646F
      Top             =   2280
      Width           =   4815
   End
   Begin VB.Image Image2 
      Height          =   3600
      Left            =   360
      Picture         =   "Form2.frx":969B
      Top             =   7200
      Width           =   4800
   End
   Begin VB.Image Image1 
      Height          =   2400
      Left            =   8760
      Picture         =   "Form2.frx":CFA0
      Top             =   120
      Width           =   6450
   End
End
Attribute VB_Name = "frmclient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbsex_Click()
    txtffn.SetFocus
End Sub

Private Sub cmdhome_Click()
    Unload Me
    frmhome.Show
End Sub

Private Sub cmdproceed_Click()
    If (txtfn.Text = "" Or txtmn.Text = "" Or txtsn.Text = "" Or cmbsex.Text = "" Or txtffn.Text = "" Or txtfmn.Text = "" Or txtfsn.Text = "" Or txtra.Text = "" Or txtpin.Text = "" Or txtnation.Text = "") Then
        MsgBox "All * Marked Fields Are Compulsory!!!!!", vbExclamation
    Else
        Call max
        con.Execute ("insert into client_info values(" + txtcid.Text + ",'" + txtfn.Text + "','" + txtmn.Text + "','" + txtsn.Text + "','" + cmbsex.Text + "','" + txtffn.Text + "','" + txtfmn.Text + "','" + txtfsn.Text + "','" + txtra.Text + "'," + txtmob.Text + "," + txtoff.Text + ",'" + txtemail.Text + "'," + txtpin.Text + ",'" + txtnation.Text + "'," + txtaid.Text + "," + txtpno.Text + ")")
        Unload Me
        frmpolicy.Show
    End If
End Sub

Private Sub Form_Load()
    Call conn
    txtcid.Visible = False
    txtaid.Visible = False
    txtpno.Visible = False
End Sub

Public Function max()
    Dim clientid As Integer
    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    rst.Open "select max(client_id) from client_info", con, adOpenStatic, adLockOptimistic
    If IsNull(rst.Fields(0)) Then
        clientid = 1
        txtcid.Text = clientid
    Else
        clientid = rst.Fields(0) + 1
        txtcid.Text = clientid
    End If
    txtpno.Text = txtcid + 1000
    txtaid.Text = txtcid + 3000
    rst.Close
End Function

Private Sub txtemail_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        txtpin.SetFocus
    End If
End Sub

Private Sub txtffn_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789`-=[]\;',./~!@#$%^&*()_+{}|:<>?"
    i = InStr(nstr, Chr(KeyAscii))
    If (i > 0) Then
        KeyAscii = 0
        MsgBox "Please Enter Character Only!!!!!", vbExclamation
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    If (KeyAscii = 13) Then
        txtfmn.SetFocus
    End If
End Sub

Private Sub txtfmn_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789`-=[]\;',./~!@#$%^&*()_+{}|:<>?"
    i = InStr(nstr, Chr(KeyAscii))
    If (i > 0) Then
        KeyAscii = 0
        MsgBox "Please Enter Character Only!!!!!", vbExclamation
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    If (KeyAscii = 13) Then
        txtfsn.SetFocus
    End If
End Sub

Private Sub txtfn_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789`-=[]\;',./~!@#$%^&*()_+{}|:<>?"
    i = InStr(nstr, Chr(KeyAscii))
    If (i > 0) Then
        KeyAscii = 0
        MsgBox "Please Enter Character Only!!!!!", vbExclamation
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    If (KeyAscii = 13) Then
        txtmn.SetFocus
    End If
End Sub

Private Sub txtfsn_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789`-=[]\;',./~!@#$%^&*()_+{}|:<>?"
    i = InStr(nstr, Chr(KeyAscii))
    If (i > 0) Then
        KeyAscii = 0
        MsgBox "Please Enter Character Only!!!!!", vbExclamation
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    If (KeyAscii = 13) Then
        txtra.SetFocus
    End If
End Sub

Private Sub txtmn_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789`-=[]\;',./~!@#$%^&*()_+{}|:<>?"
    i = InStr(nstr, Chr(KeyAscii))
    If (i > 0) Then
        KeyAscii = 0
        MsgBox "Please Enter Character Only!!!!!", vbExclamation
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    If (KeyAscii = 13) Then
        txtsn.SetFocus
    End If
End Sub

Private Sub txtmob_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8 And KeyAscii <> 13) Then
        KeyAscii = 0
        MsgBox "Please Enter Integer Only!!!!!", vbExclamation
    End If
    If (KeyAscii = 13) Then
        txtoff.SetFocus
    End If
End Sub

Private Sub txtnation_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789`-=[]\;',./~!@#$%^&*()_+{}|:<>?"
    i = InStr(nstr, Chr(KeyAscii))
    If (i > 0) Then
        KeyAscii = 0
        MsgBox "Please Enter Character Only!!!!!", vbExclamation
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    If (KeyAscii = 13) Then
        cmdproceed.SetFocus
    End If
End Sub

Private Sub txtoff_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8 And KeyAscii <> 13) Then
        KeyAscii = 0
        MsgBox "Please Enter Integer Only!!!!!", vbExclamation
    End If
    If (KeyAscii = 13) Then
        txtemail.SetFocus
    End If
End Sub

Private Sub txtpin_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8 And KeyAscii <> 13) Then
        KeyAscii = 0
        MsgBox "Please Enter Integer Only!!!!!", vbExclamation
    End If
    If (KeyAscii = 13) Then
        txtnation.SetFocus
    End If
End Sub

Private Sub txtra_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "`=[]\;~!@#$%^*+{}|:<>?"
    i = InStr(nstr, Chr(KeyAscii))
    If (i > 0) Then
        KeyAscii = 0
        MsgBox "Special Characters Are Not Allowed!!!!!", vbExclamation
    End If
    If (KeyAscii = 13) Then
        txtmob.SetFocus
    End If
End Sub

Private Sub txtsn_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789`-=[]\;',./~!@#$%^&*()_+{}|:<>?"
    i = InStr(nstr, Chr(KeyAscii))
    If (i > 0) Then
        KeyAscii = 0
        MsgBox "Please Enter Character Only!!!!!", vbExclamation
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    If (KeyAscii = 13) Then
        cmbsex.SetFocus
    End If
End Sub
