VERSION 5.00
Begin VB.Form frmagent 
   Caption         =   "Insurance Form (Continue)"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   11520
   ScaleWidth      =   19080
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtcid 
      Height          =   195
      Left            =   11040
      TabIndex        =   51
      Top             =   10560
      Width           =   150
   End
   Begin VB.TextBox txtnid 
      Height          =   195
      Left            =   11520
      TabIndex        =   50
      Top             =   10560
      Width           =   150
   End
   Begin VB.TextBox txtaid 
      Height          =   195
      Left            =   11280
      TabIndex        =   49
      Top             =   10560
      Width           =   150
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "< &Back"
      Height          =   495
      Left            =   6120
      TabIndex        =   48
      Top             =   10320
      Width           =   2295
   End
   Begin VB.CommandButton cmdhome 
      Caption         =   "<< &Back To Home"
      Height          =   495
      Left            =   3600
      TabIndex        =   47
      Top             =   10320
      Width           =   2295
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   13680
      TabIndex        =   26
      Top             =   10320
      Width           =   2295
   End
   Begin VB.Frame frmainfo 
      Caption         =   "Agent Information"
      ForeColor       =   &H000000FF&
      Height          =   3615
      Left            =   3600
      TabIndex        =   14
      Top             =   6600
      Width           =   12375
      Begin VB.TextBox txtamob 
         Height          =   315
         Left            =   4800
         TabIndex        =   41
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox txtaoff 
         Height          =   315
         Left            =   9720
         TabIndex        =   40
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox txtaemail 
         Height          =   315
         Left            =   3120
         TabIndex        =   39
         Top             =   3000
         Width           =   4455
      End
      Begin VB.TextBox txtapin 
         Height          =   315
         Left            =   9720
         TabIndex        =   38
         Top             =   3000
         Width           =   1935
      End
      Begin VB.ComboBox cmbwa 
         Height          =   315
         ItemData        =   "Form5.frx":0000
         Left            =   9720
         List            =   "Form5.frx":002B
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox txtasn 
         Height          =   285
         Left            =   9120
         TabIndex        =   19
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox txtaage 
         Height          =   315
         Left            =   3120
         TabIndex        =   18
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtafn 
         Height          =   285
         Left            =   3120
         TabIndex        =   17
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox txtamn 
         Height          =   285
         Left            =   6120
         TabIndex        =   16
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtaadd 
         Height          =   315
         Left            =   3120
         TabIndex        =   15
         Top             =   1200
         Width           =   8535
      End
      Begin VB.Label Label17 
         Caption         =   "Telephone No:"
         Height          =   375
         Left            =   600
         TabIndex        =   46
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "Mobile :"
         Height          =   375
         Left            =   3120
         TabIndex        =   45
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "Office :"
         Height          =   375
         Left            =   8760
         TabIndex        =   44
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label14 
         Caption         =   "E-mail :"
         Height          =   255
         Left            =   600
         TabIndex        =   43
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "* Pin Code:"
         Height          =   375
         Left            =   8640
         TabIndex        =   42
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "* Working  Area :"
         Height          =   255
         Left            =   8040
         TabIndex        =   27
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "* Agent's Full Name:"
         Height          =   375
         Left            =   480
         TabIndex        =   25
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label6 
         Caption         =   "* Age :"
         Height          =   255
         Left            =   480
         TabIndex        =   24
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "* Address:"
         Height          =   375
         Left            =   480
         TabIndex        =   23
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Last Name"
         Height          =   255
         Left            =   9120
         TabIndex        =   22
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "First Name"
         Height          =   255
         Left            =   3120
         TabIndex        =   21
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Middle Name"
         Height          =   255
         Left            =   6120
         TabIndex        =   20
         Top             =   840
         Width           =   2415
      End
   End
   Begin VB.Frame frmninfo 
      Caption         =   "Nominee Information"
      ForeColor       =   &H000000FF&
      Height          =   3735
      Left            =   3600
      TabIndex        =   0
      Top             =   2760
      Width           =   12375
      Begin VB.TextBox txtnmob 
         Height          =   315
         Left            =   4800
         TabIndex        =   32
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox txtnoff 
         Height          =   315
         Left            =   9720
         TabIndex        =   31
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox txtnemail 
         Height          =   315
         Left            =   3120
         TabIndex        =   30
         Top             =   3120
         Width           =   4455
      End
      Begin VB.TextBox txtnpin 
         Height          =   315
         Left            =   9720
         TabIndex        =   29
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox txtnadd 
         Height          =   315
         Left            =   3120
         TabIndex        =   6
         Top             =   1200
         Width           =   8535
      End
      Begin VB.TextBox txtnmn 
         Height          =   285
         Left            =   6120
         TabIndex        =   5
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtnfn 
         Height          =   285
         Left            =   3120
         TabIndex        =   4
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox txtrel 
         Height          =   315
         Left            =   6960
         TabIndex        =   3
         Top             =   1920
         Width           =   4695
      End
      Begin VB.TextBox txtnage 
         Height          =   345
         Left            =   3120
         TabIndex        =   2
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtnsn 
         Height          =   285
         Left            =   9120
         TabIndex        =   1
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label12 
         Caption         =   "Telephone No:"
         Height          =   375
         Left            =   600
         TabIndex        =   37
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Mobile :"
         Height          =   375
         Left            =   3120
         TabIndex        =   36
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Office :"
         Height          =   375
         Left            =   8760
         TabIndex        =   35
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "E-mail :"
         Height          =   255
         Left            =   600
         TabIndex        =   34
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "* Pin Code:"
         Height          =   375
         Left            =   8640
         TabIndex        =   33
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         Caption         =   "Middle Name"
         Height          =   255
         Left            =   6120
         TabIndex        =   13
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         Caption         =   "First Name"
         Height          =   255
         Left            =   3120
         TabIndex        =   12
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         Caption         =   "Last Name"
         Height          =   255
         Left            =   9120
         TabIndex        =   11
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label33 
         Caption         =   "* Address:"
         Height          =   375
         Left            =   480
         TabIndex        =   10
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label23 
         Caption         =   "* Relation With Client :"
         Height          =   375
         Left            =   4800
         TabIndex        =   9
         Top             =   1920
         Width           =   3015
      End
      Begin VB.Label Label22 
         Caption         =   "* Age :"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label21 
         Caption         =   "* Nominee's Full Name:"
         Height          =   375
         Left            =   480
         TabIndex        =   7
         Top             =   480
         Width           =   2895
      End
   End
   Begin VB.Image Image8 
      Height          =   1410
      Left            =   17040
      Picture         =   "Form5.frx":009D
      Top             =   8400
      Width           =   1395
   End
   Begin VB.Image Image7 
      Height          =   1410
      Left            =   960
      Picture         =   "Form5.frx":1339
      Top             =   8280
      Width           =   1395
   End
   Begin VB.Image Image6 
      Height          =   1410
      Left            =   17040
      Picture         =   "Form5.frx":2961
      Top             =   5640
      Width           =   1425
   End
   Begin VB.Image Image5 
      Height          =   1410
      Left            =   17040
      Picture         =   "Form5.frx":3731
      Top             =   2880
      Width           =   1395
   End
   Begin VB.Image Image3 
      Height          =   1440
      Left            =   960
      Picture         =   "Form5.frx":448F
      Top             =   5520
      Width           =   1410
   End
   Begin VB.Image Image2 
      Height          =   1410
      Left            =   960
      Picture         =   "Form5.frx":543A
      Top             =   2880
      Width           =   1395
   End
   Begin VB.Image Image4 
      Height          =   765
      Left            =   360
      Picture         =   "Form5.frx":623C
      Top             =   240
      Width           =   1470
   End
   Begin VB.Image Image1 
      Height          =   2490
      Left            =   5040
      Picture         =   "Form5.frx":C68F
      Top             =   120
      Width           =   9390
   End
End
Attribute VB_Name = "frmagent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbwa_Click()
    txtamob.SetFocus
End Sub

Private Sub cmdback_Click()
    Call max
    If (txtcid.Text = "0") Then
    Else
        con.Execute ("delete from other_info where client_id = " + txtcid.Text + "")
        con.Execute ("delete from premium_info where client_id = " + txtcid.Text + "")
    End If
    If (txtpid.Text = "0") Then
    Else
        con.Execute ("delete from policy_info where pol_no = " + txtpid.Text + "")
    End If
    Unload Me
    frmclient.Show
End Sub

Private Sub cmdhome_Click()
    If (cmdsave.Enabled = True) Then
        If (txtcid.Text = "0") Then
        Else
            con.Execute ("delete from client_info where client_id = " + txtcid.Text + "")
            con.Execute ("delete from other_info where client_id = " + txtcid.Text + "")
            con.Execute ("delete from premium_info where client_id=" + txtcid.Text + "")
        End If
        If (txtpid.Text = "0") Then
        Else
            con.Execute ("delete from policy_info where pol_no = " + txtpid.Text + "")
        End If
    End If
    Unload Me
    frmhome.Show
End Sub

Private Sub cmdsave_Click()
    If (txtnfn.Text = "" Or txtnmn.Text = "" Or txtnsn.Text = "" Or txtnadd.Text = "" Or txtnage.Text = "" Or txtrel.Text = "" Or txtnpin.Text = "" Or txtafn.Text = "" Or txtamn.Text = "" Or txtasn.Text = "" Or txtaadd.Text = "" Or txtaage.Text = "" Or cmbwa.Text = "" Or txtapin.Text = "") Then
        MsgBox "All * Marked Fields Are Compulsory!!!!!", vbExclamation
    Else
        Call max1
        con.Execute ("insert into nominee_info values(" + txtnid.Text + ",'" + txtnfn.Text + "','" + txtnmn.Text + "','" + txtnsn.Text + "','" + txtnadd.Text + "'," + txtnmob.Text + "," + txtnoff.Text + ",'" + txtnemail.Text + "'," + txtnpin.Text + "," + txtnage.Text + ",'" + txtrel.Text + "'," + txtcid.Text + ")")
        con.Execute ("insert into agent_info values(" + txtaid.Text + ",'" + txtafn.Text + "','" + txtamn.Text + "','" + txtasn.Text + "','" + txtaadd.Text + "'," + txtamob.Text + "," + txtaoff.Text + ",'" + txtaemail.Text + "'," + txtapin.Text + "," + txtaage.Text + ",'" + cmbwa.Text + "')")
        MsgBox "Record Saved Successfully.....", vbInformation
        cmdsave.Enabled = False
        cmdback.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    Call conn
    txtcid.Visible = False
    txtaid.Visible = False
    txtnid.Visible = False
End Sub
Public Function max()
    Dim clientid As Integer
    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    rst.Open "select max(client_id) from client_info", con, adOpenStatic, adLockOptimistic
    If IsNull(rst.Fields(0)) Then
        clientid = 0
        txtcid.Text = clientid
    Else
        clientid = rst.Fields(0)
        txtcid.Text = clientid
    End If
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "select max(pol_no) from policy_info", con, adOpenStatic, adLockOptimistic
    If IsNull(rst.Fields(0)) Then
        clientid = 0
        txtpid.Text = clientid
    Else
        clientid = rst.Fields(0)
        txtpid.Text = clientid
    End If
    rst.Close
End Function
Public Function max1()
    Dim id As Integer
    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    rst.Open "select max(agent_id) from agent_info", con, adOpenStatic, adLockOptimistic
    If IsNull(rst.Fields(0)) Then
        id = 3001
        txtaid.Text = id
    Else
        id = rst.Fields(0) + 1
        txtaid.Text = id
    End If
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "select max(nominee_id) from nominee_info", con, adOpenStatic, adLockOptimistic
    If IsNull(rst.Fields(0)) Then
        id = 2001
        txtnid.Text = id
    Else
        id = rst.Fields(0) + 1
        txtnid.Text = id
    End If
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "select max(client_id) from client_info", con, adOpenStatic, adLockOptimistic
    If IsNull(rst.Fields(0)) Then
        id = 0
        txtcid.Text = id
    Else
        id = rst.Fields(0)
        txtcid.Text = id
    End If
    rst.Close
End Function

Private Sub txtaadd_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "=[]\~!@#$%^*+{}|:<>?"
    i = InStr(nstr, Chr(KeyAscii))
    If (i > 0) Then
        KeyAscii = 0
        MsgBox "Special Characters Are Not Allowed!!!!!", vbExclamation
    End If
    If (KeyAscii = 13) Then
        txtaage.SetFocus
    End If
End Sub

Private Sub txtaage_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8 And KeyAscii <> 13) Then
        KeyAscii = 0
        MsgBox "Please Enter Integer Only!!!!!", vbExclamation
    End If
    If (KeyAscii = 13) Then
        cmbwa.SetFocus
    End If
End Sub

Private Sub txtaemail_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        txtapin.SetFocus
    End If
End Sub

Private Sub txtafn_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789`-=[]\;'/~!@#$%^&*()_+{}|:<>?"
    i = InStr(nstr, Chr(KeyAscii))
    If (i > 0) Then
        KeyAscii = 0
        MsgBox "Please Enter Character Only!!!!!", vbExclamation
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    If (KeyAscii = 13) Then
        txtamn.SetFocus
    End If
End Sub

Private Sub txtamn_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789`-=[]\;'/~!@#$%^&*()_+{}|:<>?"
    i = InStr(nstr, Chr(KeyAscii))
    If (i > 0) Then
        KeyAscii = 0
        MsgBox "Please Enter Character Only!!!!!", vbExclamation
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    If (KeyAscii = 13) Then
        txtasn.SetFocus
    End If
End Sub

Private Sub txtamob_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8 And KeyAscii <> 13) Then
        KeyAscii = 0
        MsgBox "Please Enter Integer Only!!!!!", vbExclamation
    End If
    If (KeyAscii = 13) Then
        txtaoff.SetFocus
    End If
End Sub

Private Sub txtaoff_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8 And KeyAscii <> 13) Then
        KeyAscii = 0
        MsgBox "Please Enter Integer Only!!!!!", vbExclamation
    End If
    If (KeyAscii = 13) Then
        txtaemail.SetFocus
    End If
End Sub

Private Sub txtapin_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8 And KeyAscii <> 13) Then
        KeyAscii = 0
        MsgBox "Please Enter Integer Only!!!!!", vbExclamation
    End If
    If (KeyAscii = 13) Then
        cmdsave.SetFocus
    End If
End Sub

Private Sub txtasn_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789`-=[]\;'/~!@#$%^&*()_+{}|:<>?"
    i = InStr(nstr, Chr(KeyAscii))
    If (i > 0) Then
        KeyAscii = 0
        MsgBox "Please Enter Character Only!!!!!", vbExclamation
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    If (KeyAscii = 13) Then
        txtaadd.SetFocus
    End If
End Sub

Private Sub txtnadd_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "=[]\~!@#$%^*+{}|:<>?"
    i = InStr(nstr, Chr(KeyAscii))
    If (i > 0) Then
        KeyAscii = 0
        MsgBox "Special Characters Are Not Allowed!!!!!", vbExclamation
    End If
    If (KeyAscii = 13) Then
        txtnage.SetFocus
    End If
End Sub

Private Sub txtnage_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8 And KeyAscii <> 13) Then
        KeyAscii = 0
        MsgBox "Please Enter Integer Only!!!!!", vbExclamation
    End If
    If (KeyAscii = 13) Then
        txtrel.SetFocus
    End If
End Sub

Private Sub txtnemail_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        txtnpin.SetFocus
    End If
End Sub

Private Sub txtnfn_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789`-=[]\;'/~!@#$%^&*()_+{}|:<>?"
    i = InStr(nstr, Chr(KeyAscii))
    If (i > 0) Then
        KeyAscii = 0
        MsgBox "Please Enter Character Only!!!!!", vbExclamation
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    If (KeyAscii = 13) Then
        txtnmn.SetFocus
    End If
End Sub

Private Sub txtnmn_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789`-=[]\;'/~!@#$%^&*()_+{}|:<>?"
    i = InStr(nstr, Chr(KeyAscii))
    If (i > 0) Then
        KeyAscii = 0
        MsgBox "Please Enter Character Only!!!!!", vbExclamation
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    If (KeyAscii = 13) Then
        txtnsn.SetFocus
    End If
End Sub

Private Sub txtnmob_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8 And KeyAscii <> 13) Then
        KeyAscii = 0
        MsgBox "Please Enter Integer Only!!!!!", vbExclamation
    End If
    If (KeyAscii = 13) Then
        txtnoff.SetFocus
    End If
End Sub

Private Sub txtnoff_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8 And KeyAscii <> 13) Then
        KeyAscii = 0
        MsgBox "Please Enter Integer Only!!!!!", vbExclamation
    End If
    If (KeyAscii = 13) Then
        txtnemail.SetFocus
    End If
End Sub

Private Sub txtnpin_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8 And KeyAscii <> 13) Then
        KeyAscii = 0
        MsgBox "Please Enter Integer Only!!!!!", vbExclamation
    End If
    If (KeyAscii = 13) Then
        txtafn.SetFocus
    End If
End Sub

Private Sub txtnsn_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789`-=[]\;'/~!@#$%^&*()_+{}|:<>?"
    i = InStr(nstr, Chr(KeyAscii))
    If (i > 0) Then
        KeyAscii = 0
        MsgBox "Please Enter Character Only!!!!!", vbExclamation
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    If (KeyAscii = 13) Then
        txtnadd.SetFocus
    End If
End Sub

Private Sub txtrel_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789`-=[]\;'/~!@#$%^&*()_+{}|:<>?"
    i = InStr(nstr, Chr(KeyAscii))
    If (i > 0) Then
        KeyAscii = 0
        MsgBox "Please Enter Character Only!!!!!", vbExclamation
    End If
    If (KeyAscii = 13) Then
        txtnmob.SetFocus
    End If
End Sub
