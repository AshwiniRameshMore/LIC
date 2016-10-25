VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmreceipt 
   Caption         =   "Premium Receipt"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4680
   LinkTopic       =   "Form8"
   MDIChild        =   -1  'True
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt 
      Height          =   195
      Left            =   11160
      TabIndex        =   35
      Top             =   10680
      Width           =   150
   End
   Begin MSComCtl2.DTPicker dpkpdate 
      Height          =   255
      Left            =   10800
      TabIndex        =   34
      Top             =   10680
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Format          =   61538305
      CurrentDate     =   39732
   End
   Begin MSComCtl2.DTPicker dpkdate 
      Height          =   255
      Left            =   10560
      TabIndex        =   33
      Top             =   10680
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Format          =   61538305
      CurrentDate     =   39732
   End
   Begin VB.ComboBox cmbcid 
      Height          =   315
      Left            =   10080
      TabIndex        =   22
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Height          =   495
      Left            =   12840
      TabIndex        =   21
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmdhome 
      Caption         =   "<< &Back To Home"
      Height          =   495
      Left            =   6840
      TabIndex        =   20
      Top             =   10560
      Width           =   2295
   End
   Begin VB.Frame frmpreceipt 
      Caption         =   "Premium  Receipt"
      Height          =   6855
      Left            =   6840
      TabIndex        =   0
      Top             =   3240
      Width           =   11295
      Begin VB.TextBox txtstatus 
         Height          =   285
         Left            =   9120
         TabIndex        =   31
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtpamt 
         Height          =   285
         Left            =   3240
         TabIndex        =   29
         Top             =   4440
         Width           =   1455
      End
      Begin VB.TextBox txtpolamt 
         Height          =   285
         Left            =   3240
         TabIndex        =   27
         Top             =   3720
         Width           =   1455
      End
      Begin VB.TextBox txtpdate 
         Height          =   285
         Left            =   3240
         TabIndex        =   26
         Top             =   5160
         Width           =   1455
      End
      Begin VB.TextBox txtddate 
         Height          =   285
         Left            =   3240
         TabIndex        =   25
         Top             =   6000
         Width           =   1455
      End
      Begin VB.TextBox txtdamt 
         Height          =   285
         Left            =   9120
         TabIndex        =   23
         Top             =   6000
         Width           =   1455
      End
      Begin VB.TextBox txtaid 
         Height          =   285
         Left            =   3240
         TabIndex        =   16
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox txtpid 
         Height          =   285
         Left            =   3240
         TabIndex        =   14
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtpamt1 
         Height          =   285
         Left            =   9120
         TabIndex        =   12
         Top             =   5160
         Width           =   1455
      End
      Begin VB.TextBox txtlname 
         Height          =   285
         Left            =   3240
         TabIndex        =   7
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtfname 
         Height          =   285
         Left            =   6000
         TabIndex        =   6
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtmname 
         Height          =   285
         Left            =   8640
         TabIndex        =   5
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtbranch 
         Height          =   285
         Left            =   3240
         TabIndex        =   1
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "Status :"
         Height          =   255
         Left            =   6600
         TabIndex        =   32
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Paid Amount :"
         Height          =   255
         Left            =   840
         TabIndex        =   30
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Label Label13 
         Caption         =   "Policy Amount :"
         Height          =   255
         Left            =   840
         TabIndex        =   28
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Due Amount :"
         Height          =   255
         Left            =   6600
         TabIndex        =   24
         Top             =   6000
         Width           =   1695
      End
      Begin VB.Label Label11 
         Caption         =   "Due Date :"
         Height          =   255
         Left            =   840
         TabIndex        =   17
         Top             =   6000
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Agent ID :"
         Height          =   255
         Left            =   840
         TabIndex        =   15
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Policy ID :"
         Height          =   255
         Left            =   840
         TabIndex        =   13
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Premium Amount :"
         Height          =   255
         Left            =   6600
         TabIndex        =   11
         Top             =   5160
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Last Name"
         Height          =   255
         Left            =   3000
         TabIndex        =   10
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "First Name"
         Height          =   255
         Left            =   5760
         TabIndex        =   9
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Middle Name"
         Height          =   255
         Left            =   8400
         TabIndex        =   8
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Client Name :"
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Premium Date :"
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   5160
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Branch :"
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Label Label12 
      Caption         =   "Enter The Client ID :"
      Height          =   255
      Left            =   7680
      TabIndex        =   19
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   4515
      Left            =   240
      Picture         =   "Form8.frx":0000
      Top             =   3360
      Width           =   6000
   End
   Begin VB.Label Label18 
      Caption         =   "PREMIUM  RECEIPT"
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
      Left            =   11160
      TabIndex        =   18
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Image Image4 
      Height          =   765
      Left            =   360
      Picture         =   "Form8.frx":4167
      Top             =   240
      Width           =   1470
   End
End
Attribute VB_Name = "frmreceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbcid_Click()
    cmdok.SetFocus
End Sub

Private Sub cmdhome_Click()
    Unload Me
    frmhome.Show
End Sub

Private Sub cmdok_Click()
    If (cmbcid.Text = "") Then
        MsgBox "Please Select Client-ID!!!!!", vbExclamation
    Else
        Dim mode As String
        Dim rst As ADODB.Recordset
        Set rst = New ADODB.Recordset
        rst.Open "select claim_id from claim_info where client_id=" + cmbcid.Text + "", con, adOpenStatic, adLockOptimistic
        If (rst.RecordCount = 0) Then
            rst.Close
            dpkdate.Value = Date
            rst.Open "select pol_no from client_info where client_id =" + cmbcid.Text + "", con, adOpenStatic, adLockOptimistic
            txt.Text = rst.Fields(0)
            rst.Close
            rst.Open "select pol_mode from policy_info where pol_no=" + txt.Text + "", con, adOpenStatic, adLockOptimistic
            mode = rst.Fields(0)
            rst.Close
            rst.Open "select premium_date,status from premium_info where client_id =" + cmbcid.Text + "", con, adOpenStatic, adLockOptimistic
            dpkpdate.Value = rst.Fields(0)
            txt.Text = rst.Fields(1)
            rst.Close
            If (mode = "Yearly") Then
                dpkpdate.Value = DateSerial(dpkpdate.Year, dpkpdate.Month, dpkpdate.Day)
                dpkpdate.Value = DateAdd("yyyy", 1, dpkpdate.Value)
            Else
                dpkpdate.Value = DateSerial(dpkpdate.Year, dpkpdate.Month, dpkpdate.Day)
                dpkpdate.Value = DateAdd("m", 6, dpkpdate.Value)
            End If
            If (dpkdate.Value >= dpkpdate.Value And txt.Text = "Paid") Then
                txt.Text = "Unpaid"
                dpkpdate.Value = DateSerial(dpkpdate.Year, dpkpdate.Month, dpkpdate.Day)
                dpkdate.Value = DateAdd("m", 1, dpkpdate.Value)
                con.Execute ("update premium_info set premium_date='" + CStr(dpkpdate.Value) + "',due_date='" + CStr(dpkdate.Value) + "',status='" + txt.Text + "' where client_id=" + cmbcid.Text + "")
            End If
            rst.Open "select client_fname,client_mname,client_lname from client_info where client_id=" + cmbcid.Text + "", con, adOpenStatic, adLockOptimistic
            txtfname.Text = rst.Fields(0)
            txtmname.Text = rst.Fields(1)
            txtlname.Text = rst.Fields(2)
            rst.Close
            rst.Open "select agent_id,pol_id,total,paid,premium_amt,premium_date,due_date,due_amt,status from premium_info where client_id=" + cmbcid.Text + "", con, adOpenStatic, adLockOptimistic
            txtaid.Text = rst.Fields(0)
            txtpid.Text = rst.Fields(1)
            txtpolamt.Text = rst.Fields(2)
            txtpamt.Text = rst.Fields(3)
            txtpamt1.Text = rst.Fields(4)
            txtpdate.Text = rst.Fields(5)
            txtddate.Text = rst.Fields(6)
            txtdamt.Text = rst.Fields(7)
            txtstatus.Text = rst.Fields(8)
            rst.Close
            rst.Open "select work_area from agent_info where agent_id=" + txtaid.Text + "", con, adOpenStatic, adLockOptimistic
            txtbranch.Text = rst.Fields(0)
            rst.Close
        Else
            MsgBox "Your Policy Has Been Closed!!!!!", vbExclamation
        End If
    End If
End Sub

Private Sub Form_Load()
    Call conn
    Call fill
    frmpreceipt.Enabled = False
    txt.Visible = False
    dpkdate.Visible = False
    dpkpdate.Visible = False
End Sub
Public Function fill()
    Dim rst As ADODB.Recordset
    cmbcid.Clear
    Set rst = New ADODB.Recordset
    rst.Open "select distinct client_id from client_info order by client_id", con, adOpenStatic, adLockOptimistic
    While Not rst.EOF
        cmbcid.AddItem rst.Fields(0)
        rst.MoveNext
    Wend
End Function

