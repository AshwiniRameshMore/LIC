VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form PREMIUMRECEIPT 
   BackColor       =   &H00FFC0C0&
   Caption         =   "RECEIPT"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Height          =   855
      Left            =   600
      Picture         =   "PREMIUMRECEIPT1.frx":0000
      ScaleHeight     =   795
      ScaleWidth      =   1515
      TabIndex        =   34
      Top             =   360
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   4575
      Left            =   120
      Picture         =   "PREMIUMRECEIPT1.frx":6453
      ScaleHeight     =   4515
      ScaleWidth      =   5955
      TabIndex        =   32
      Top             =   2400
      Width           =   6015
   End
   Begin VB.ComboBox cmbcid 
      Height          =   315
      Left            =   11160
      TabIndex        =   1
      Top             =   1320
      Width           =   2775
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
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
      Left            =   15240
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "PREMIUM RECEIPT"
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
      Height          =   7575
      Left            =   7080
      TabIndex        =   6
      Top             =   2280
      Width           =   10095
      Begin VB.TextBox status 
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
         Left            =   1800
         TabIndex        =   18
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtclname 
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
         Left            =   1800
         TabIndex        =   17
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtcfname 
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
         Left            =   4560
         TabIndex        =   16
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtcmname 
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
         Left            =   7320
         TabIndex        =   15
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtpolid 
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
         Left            =   1800
         TabIndex        =   14
         Top             =   2760
         Width           =   2055
      End
      Begin VB.TextBox txtaid 
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
         Left            =   1800
         TabIndex        =   13
         Top             =   3600
         Width           =   2055
      End
      Begin VB.TextBox txtpamt 
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
         Left            =   1800
         TabIndex        =   12
         Top             =   4320
         Width           =   2055
      End
      Begin VB.TextBox txtpaid 
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
         Left            =   1800
         TabIndex        =   11
         Top             =   5280
         Width           =   2055
      End
      Begin VB.TextBox txtpdate 
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
         Left            =   1800
         TabIndex        =   10
         Top             =   6120
         Width           =   2055
      End
      Begin VB.TextBox txtpamt1 
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
         Left            =   7560
         TabIndex        =   9
         Top             =   6120
         Width           =   2055
      End
      Begin VB.TextBox txtddate 
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
         Left            =   1800
         TabIndex        =   8
         Top             =   6840
         Width           =   2055
      End
      Begin VB.TextBox txtdamt 
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
         Left            =   7560
         TabIndex        =   7
         Top             =   6720
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "STATUS"
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
         TabIndex        =   31
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "CLIENT NAME"
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
         TabIndex        =   30
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "       LAST NAME"
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
         Left            =   1920
         TabIndex        =   29
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "           FIRST NAME"
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
         TabIndex        =   28
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFC0C0&
         Caption         =   "       MIDDLE NAME"
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
         Left            =   7440
         TabIndex        =   27
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFC0C0&
         Caption         =   "POLICY ID"
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
         TabIndex        =   26
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFC0C0&
         Caption         =   "AGENT ID"
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
         TabIndex        =   25
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "POLICY AMOUNT"
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
         TabIndex        =   24
         Top             =   4320
         Width           =   1455
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFC0C0&
         Caption         =   "PAID AMOUNT"
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
         TabIndex        =   23
         Top             =   5280
         Width           =   1455
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFC0C0&
         Caption         =   "PREMIUM DATE"
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
         TabIndex        =   22
         Top             =   6120
         Width           =   1575
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFC0C0&
         Caption         =   "PREMIUM AMOUNT"
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
         TabIndex        =   21
         Top             =   6120
         Width           =   1935
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFC0C0&
         Caption         =   "DUE DATE"
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
         TabIndex        =   20
         Top             =   6840
         Width           =   1455
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "DUE AMOUNT"
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
         Left            =   6000
         TabIndex        =   19
         Top             =   6840
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command2 
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
      Left            =   1320
      TabIndex        =   3
      Top             =   8400
      Width           =   3495
   End
   Begin VB.TextBox txt 
      Height          =   285
      Left            =   4080
      TabIndex        =   0
      Text            =   "Text14"
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComCtl2.DTPicker dpkpdate 
      Height          =   135
      Left            =   3720
      TabIndex        =   4
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   238
      _Version        =   393216
      Format          =   73596929
      CurrentDate     =   40456
   End
   Begin MSComCtl2.DTPicker dpkdate 
      Height          =   255
      Left            =   4440
      TabIndex        =   5
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Format          =   73596929
      CurrentDate     =   40456
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "PREMIUM RECEIPT"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   8880
      TabIndex        =   35
      Top             =   240
      Width           =   5535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ENTER THE CLIENT ID"
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
      Left            =   9000
      TabIndex        =   33
      Top             =   1320
      Width           =   3015
   End
End
Attribute VB_Name = "PREMIUMRECEIPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdok_Click()
    If (cmbcid.Text = "") Then
        MsgBox "Please Select Client-ID!!!!!", vbExclamation
    Else
        Dim mode As String
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        rs.Open "select claim_id from claim_info where client_id=" + cmbcid.Text + "", cn, adOpenStatic, adLockOptimistic
        If (rs.RecordCount = 0) Then
            rs.Close
            dpkdate.Value = Date
            rs.Open "select pol_no from client_info where client_id =" + cmbcid.Text + "", cn, adOpenStatic, adLockOptimistic
            txt.Text = rs.Fields(0)
            rs.Close
            rs.Open "select pol_mode from policy_info where pol_no=" + txt.Text + "", cn, adOpenStatic, adLockOptimistic
            mode = rs.Fields(0)
            rs.Close
            rs.Open "select premium_date,status from premium_info where client_id =" + cmbcid.Text + "", cn, adOpenStatic, adLockOptimistic
            dpkpdate.Value = rs.Fields(0)
            txt.Text = rs.Fields(1)
            rs.Close
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
                cn.Execute ("update premium_info set premium_date='" + CStr(dpkpdate.Value) + "',due_date='" + CStr(dpkdate.Value) + "',status='" + txt.Text + "' where client_id=" + cmbcid.Text + "")
            End If
            rs.Open "select client_fname,client_mname,client_lname from client_info where client_id=" + cmbcid.Text + "", cn, adOpenStatic, adLockOptimistic
            txtcfname.Text = rs.Fields(0)
            txtcmname.Text = rs.Fields(1)
            txtclname.Text = rs.Fields(2)
            rs.Close
            rs.Open "select agent_id,pol_id,total,paid,premium_amt,premium_date,due_date,due_amt,status from premium_info where client_id=" + cmbcid.Text + "", cn, adOpenStatic, adLockOptimistic
            txtaid.Text = rs.Fields(0)
            txtpolid.Text = rs.Fields(1)
            txtpamt.Text = rs.Fields(2)
            txtpaid.Text = rs.Fields(3)
            txtpamt1.Text = rs.Fields(4)
            txtpdate.Text = rs.Fields(5)
            txtddate.Text = rs.Fields(6)
            txtdamt.Text = rs.Fields(7)
            status.Text = rs.Fields(8)
            rs.Close
        Else
            MsgBox "Your Policy Has Been Closed!!!!!", vbExclamation
        End If
    End If
End Sub

Private Sub Command2_Click()
HOME.Show

End Sub

Private Sub Form_Load()
    Call connect
    Call fill
    Frame1.Enabled = False
    txt.Visible = False
    dpkdate.Visible = False
    dpkpdate.Visible = False
     
End Sub
Public Function fill()
    Dim rs As ADODB.Recordset
    cmbcid.Clear
    Set rs = New ADODB.Recordset
    rs.Open "select distinct client_id from client_info order by client_id", cn, adOpenStatic, adLockOptimistic
    While Not rs.EOF
        cmbcid.AddItem rs.Fields(0)
        rs.MoveNext
    Wend
End Function




