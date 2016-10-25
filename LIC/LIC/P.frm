VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form PREMIUM1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "PREMIUM"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      Height          =   855
      Left            =   840
      Picture         =   "P.frx":0000
      ScaleHeight     =   795
      ScaleWidth      =   1515
      TabIndex        =   18
      Top             =   360
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   4575
      Left            =   480
      Picture         =   "P.frx":6453
      ScaleHeight     =   4515
      ScaleWidth      =   5955
      TabIndex        =   17
      Top             =   3720
      Width           =   6015
   End
   Begin VB.PictureBox Picture2 
      Height          =   2775
      Left            =   6000
      Picture         =   "P.frx":DA95
      ScaleHeight     =   2715
      ScaleWidth      =   11475
      TabIndex        =   16
      Top             =   360
      Width           =   11535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "PAY PREMIUM"
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
      Height          =   4455
      Left            =   7680
      TabIndex        =   10
      Top             =   3600
      Width           =   8415
      Begin VB.ComboBox cmbcid 
         Height          =   315
         Left            =   2040
         TabIndex        =   1
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtpamt 
         Height          =   375
         Left            =   2040
         TabIndex        =   12
         Top             =   3480
         Width           =   2175
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
         Left            =   5280
         TabIndex        =   2
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton CMDGETRECEIPT 
         Caption         =   "GET PREMIUM RECEIPT"
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
         Left            =   5280
         TabIndex        =   3
         Top             =   1920
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker dpkdate 
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   2040
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   73596929
         CurrentDate     =   40456
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ENTER CLIENT ID"
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
         TabIndex        =   15
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "DATE"
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
         TabIndex        =   14
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label3 
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
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   3480
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "<<GO BACK"
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
      Left            =   10080
      TabIndex        =   4
      Top             =   8520
      Width           =   3135
   End
   Begin VB.TextBox txtpaid 
      Height          =   375
      Left            =   5520
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   4320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtclaimid 
      Height          =   285
      Left            =   5160
      TabIndex        =   8
      Text            =   "Text3"
      Top             =   5040
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox TXTSTATUS 
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   4440
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtpid 
      Height          =   285
      Left            =   5280
      TabIndex        =   6
      Text            =   "Text4"
      Top             =   4560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtpdur 
      Height          =   285
      Left            =   4560
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   4560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtpamt1 
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Text            =   "Text5"
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "PREMIUM1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbcid_Click()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "select claim_id from claim_info where client_id=" + cmbcid.Text + "", cn, adOpenStatic, adLockOptimistic
    If (rs.RecordCount > 0) Then
        MsgBox "Your Policy Has Been Closed!!!!!", vbExclamation
        cmdok.Enabled = False
    Else
        cmdok.Enabled = True
        rs.Close
        rs.Open "select due_date from premium_info where client_id=" + cmbcid.Text + "", cn, adOpenStatic, adLockOptimistic
        If (dpkdate.Value <= rs.Fields(0)) Then
            rs.Close
            rs.Open "select premium_amt from premium_info where client_id=" + cmbcid.Text + "", cn, adOpenStatic, adLockOptimistic
        Else
            rs.Close
            rs.Open "select due_amt from premium_info where client_id=" + cmbcid.Text + "", cn, adOpenStatic, adLockOptimistic
        End If
        txtpamt.Text = rs.Fields(0)
        rs.Close
    End If
End Sub

Private Sub cmdok_Click()
    If (cmbcid.Text = "") Then
        MsgBox "Please Select Client-ID!!!!!", vbExclamation
    Else
        Dim paid As Integer
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        txtstatus.Text = "Paid"
        rs.Open "select paid from premium_info where client_id=" + cmbcid.Text + "", cn, adOpenStatic, adLockOptimistic
        paid = rs.Fields(0)
        txtpaid.Text = paid + txtpamt.Text
         cn.Execute ("update premium_info set paid=" + txtpaid.Text + ",status='" + txtstatus.Text + "' where client_id=" + cmbcid.Text + "")
        MsgBox "Premium Paid Successfully.....", vbInformation
    End If
End Sub

Private Sub CMDGETRECEIPT_Click()
    Unload Me
    PREMIUMRECEIPT.Show

End Sub

Private Sub Command4_Click()
HOME.Show

End Sub

Private Sub Form_Load()
    Call connect
    Call fill
   
        dpkdate.Enabled = False
        dpkdate.Value = Date
        txtpamt.Enabled = False
    End Sub
Public Function fill()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "select distinct client_id from client_info order by client_id", cn, adOpenStatic, adLockOptimistic
    While Not rs.EOF
        cmbcid.AddItem rs.Fields(0)
        rs.MoveNext
    Wend
End Function

Public Function max()
    Dim claimid As Integer
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "select max(claim_id) from claim_info", cn, adOpenStatic, adLockOptimistic
    If IsNull(rs.Fields(0)) Then
        claimid = 4001
        txtclaimid.Text = claimid
    Else
        claimid = rs.Fields(0) + 1
        txtclaimid.Text = claimid
    End If
    rs.Close
End Function



