VERSION 5.00
Begin VB.Form frmlogin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "LIC Management System- Login"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtpassword 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2025
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   2325
   End
   Begin VB.TextBox txtusername 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   2325
   End
   Begin VB.CommandButton cmdgo 
      Caption         =   "&GO >>"
      Height          =   340
      Left            =   4680
      TabIndex        =   2
      Top             =   240
      Width           =   1140
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      Height          =   340
      Left            =   4680
      TabIndex        =   3
      Top             =   720
      Width           =   1140
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Height          =   340
      Left            =   4680
      TabIndex        =   4
      Top             =   480
      Width           =   1140
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      Height          =   340
      Left            =   4680
      TabIndex        =   5
      Top             =   480
      Width           =   1140
   End
   Begin VB.Label lbladministrator 
      Caption         =   "Please Login As Administrator....."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label lblusername 
      Caption         =   "&User Name :"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblpassword 
      Caption         =   "&Password :"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblfp 
      Caption         =   "Forgot Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblcp 
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblnu 
      Caption         =   "New User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lbllogin 
      Caption         =   "Back To Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4680
      TabIndex        =   9
      Top             =   1320
      Width           =   1095
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Boolean
Dim flag1 As Boolean

Private Sub cmdCancel_Click()
    MsgBox "Thank You!!!!!", vbInformation
    End
End Sub

Private Sub cmdgo_Click()
    Dim rst As ADODB.Recordset
    If (txtusername.Text = "") Then
        MsgBox "Please Enter User Name!!!!!", vbExclamation
    ElseIf (txtpassword.Text = "") Then
        MsgBox "Please Enter Password!!!!!", vbExclamation
    Else
        Set rst = New ADODB.Recordset
        rst.Open "select password from login where username='" + txtusername.Text + "'", con, adOpenStatic, adLockOptimistic
        If (rst.RecordCount = 0) Then
            MsgBox "Invalid User Name!!!!!", vbCritical
        ElseIf (txtpassword.Text = rst.Fields(0)) Then
            If (flag1 = False) Then
                flag = True
                flag1 = True
                lblusername.Caption = "Enter User Name :"
                lblpassword.Caption = "Enter Password :"
                txtpassword.PasswordChar = ""
                txtusername.Text = ""
                txtpassword.Text = ""
                cmdgo.Visible = False
                cmdcancel.Visible = False
                cmdsave.Visible = True
                cmdok.Visible = False
                lbladministrator.Visible = False
                lblfp.Visible = False
                lblcp.Visible = False
                lblnu.Visible = False
                lbllogin.Visible = True
            Else
                MsgBox "Welcome To LIC Management System!!!!!", vbInformation
                Unload Me
                frmhome.Show
            End If
        Else
            MsgBox "Invalid Password!!!!!", vbExclamation
        End If
        rst.Close
    End If
End Sub

Private Sub cmdok_Click()
    Dim rst As ADODB.Recordset
    If (txtusername.Text = "") Then
        MsgBox "Please Enter User Name!!!!!", vbExclamation
    Else
        Set rst = New ADODB.Recordset
        rst.Open "select password from login where username='" + txtusername.Text + "'", con, adOpenStatic, adLockOptimistic
        If (rst.RecordCount = 0) Then
            MsgBox "Invalid User Name!!!!!", vbExclamation
        Else
            txtpassword.Text = rst.Fields(0)
        End If
        rst.Close
    End If
End Sub

Private Sub cmdsave_Click()
    Dim rst As ADODB.Recordset
    If (txtusername.Text = "") Then
        MsgBox "Please Enter User Name!!!!!", vbExclamation
    ElseIf (txtpassword.Text = "") Then
        MsgBox "Please Enter Password!!!!!", vbExclamation
    Else
        Set rst = New ADODB.Recordset
        If (flag = True) Then
            rst.Open "select * from login where username='" + txtusername.Text + "'", con, adOpenStatic, adLockOptimistic
            If (rst.RecordCount = 0) Then
                con.Execute ("insert into login values('" + txtusername.Text + "','" + txtpassword.Text + "')")
                MsgBox "User Created Successfully.....", vbInformation
            Else
                MsgBox "User Name Already Present!!!!!", vbExclamation
            End If
        Else
            rst.Open "select * from login where username='" + txtusername.Text + "'", con, adOpenStatic, adLockOptimistic
            If (rst.RecordCount = 0) Then
                MsgBox "Invalid User Name!!!!!", vbExclamation
            Else
                con.Execute ("update login set password='" + txtpassword.Text + "' where username='" + txtusername.Text + "'")
                MsgBox "Password Changed Successfully.....", vbInformation
            End If
        End If
        rst.Close
    End If
End Sub

Private Sub Form_Load()
    Call conn
    flag = True
    flag1 = True
    cmdsave.Visible = False
    cmdok.Visible = False
    lbladministrator.Visible = False
End Sub

Private Sub lblcp_Click()
    flag = False
    lblusername.Caption = "Enter User Name :"
    lblpassword.Caption = "Enter New Password :"
    txtpassword.PasswordChar = ""
    txtusername.Text = ""
    txtpassword.Text = ""
    cmdgo.Visible = False
    cmdcancel.Visible = False
    cmdsave.Visible = True
    cmdok.Visible = False
    lblfp.Visible = False
    lblcp.Visible = False
    lblnu.Visible = False
    lbllogin.Visible = True
End Sub

Private Sub lblfp_Click()
    lblusername.Caption = "Enter User Name :"
    lblpassword.Caption = "Password :"
    txtpassword.PasswordChar = ""
    txtusername.Text = ""
    txtpassword.Text = ""
    cmdgo.Visible = False
    cmdcancel.Visible = False
    cmdsave.Visible = False
    cmdok.Visible = True
    lblfp.Visible = False
    lblcp.Visible = False
    lblnu.Visible = False
    lbllogin.Visible = True
End Sub

Private Sub lbllogin_Click()
    lblusername.Caption = "&User Name :"
    lblpassword.Caption = "&Password :"
    txtpassword.PasswordChar = "*"
    txtusername.Text = ""
    txtpassword.Text = ""
    cmdgo.Visible = True
    cmdcancel.Visible = True
    cmdsave.Visible = False
    cmdok.Visible = False
    lblfp.Visible = True
    lblcp.Visible = True
    lblnu.Visible = True
    lbllogin.Visible = False
End Sub

Private Sub lblnu_Click()
    flag1 = False
    lblfp.Visible = False
    lblcp.Visible = False
    lblnu.Visible = False
    lbladministrator.Visible = True
    txtusername.Text = "Administrator"
    txtpassword.SetFocus
End Sub

