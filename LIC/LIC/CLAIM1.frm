VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CLAIM1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "CLAIM"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11490
   ScaleWidth      =   19080
   WindowState     =   2  'Maximized
   Begin VB.Frame frmclaim 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Get Claim Amount"
      Height          =   4455
      Left            =   7680
      TabIndex        =   8
      Top             =   3600
      Width           =   9255
      Begin VB.CommandButton cmdok1 
         Caption         =   "&OK"
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
         Left            =   6600
         TabIndex        =   13
         Top             =   1920
         Width           =   2055
      End
      Begin VB.ComboBox cmbcid1 
         Height          =   315
         ItemData        =   "CLAIM1.frx":0000
         Left            =   3480
         List            =   "CLAIM1.frx":0002
         TabIndex        =   12
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtcamt 
         Height          =   285
         Left            =   3480
         TabIndex        =   11
         Top             =   3360
         Width           =   1455
      End
      Begin VB.ComboBox cmbcod 
         Height          =   315
         ItemData        =   "CLAIM1.frx":0004
         Left            =   3480
         List            =   "CLAIM1.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2640
         Width           =   1455
      End
      Begin VB.ComboBox cmbstatus 
         Height          =   315
         Left            =   3480
         TabIndex        =   9
         Top             =   1920
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dpkcdate 
         Height          =   375
         Left            =   3480
         TabIndex        =   14
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   11730945
         CurrentDate     =   39732
      End
      Begin VB.Label Label8 
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
         Height          =   255
         Left            =   720
         TabIndex        =   19
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "CLAIM AMOUNT"
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
         Left            =   720
         TabIndex        =   18
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label Label5 
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
         Height          =   375
         Left            =   720
         TabIndex        =   17
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "CAUSE OF DEATH"
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
         Left            =   720
         TabIndex        =   16
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ENTER CLIENT ID:"
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
         Left            =   720
         TabIndex        =   15
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.TextBox txtpaid 
      Height          =   285
      Left            =   2040
      TabIndex        =   7
      Top             =   4680
      Width           =   150
   End
   Begin VB.TextBox txtpdur 
      Height          =   195
      Left            =   1920
      TabIndex        =   6
      Top             =   4800
      Width           =   150
   End
   Begin VB.TextBox txtpamt1 
      Height          =   195
      Left            =   2280
      TabIndex        =   5
      Top             =   4800
      Width           =   150
   End
   Begin VB.TextBox txtpid 
      Height          =   195
      Left            =   1920
      TabIndex        =   4
      Top             =   5040
      Width           =   150
   End
   Begin VB.TextBox txtstatus 
      Height          =   195
      Left            =   2160
      TabIndex        =   3
      Top             =   5040
      Width           =   150
   End
   Begin VB.TextBox txtclaimid 
      Height          =   195
      Left            =   2280
      TabIndex        =   2
      Top             =   4560
      Width           =   150
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<<GO BACK TO HOME"
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
      Left            =   11160
      TabIndex        =   1
      Top             =   8400
      Width           =   2775
   End
   Begin VB.PictureBox Picture4 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      Picture         =   "CLAIM1.frx":0027
      ScaleHeight     =   795
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   2880
      Left            =   6960
      Picture         =   "CLAIM1.frx":647A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   10635
   End
   Begin VB.Image Image2 
      Height          =   4500
      Left            =   0
      Picture         =   "CLAIM1.frx":12A87
      Top             =   3600
      Width           =   6720
   End
End
Attribute VB_Name = "CLAIM1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbcid1_Click()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "select claim_id from claim_info where client_id=" + cmbcid1.Text + "", cn, adOpenStatic, adLockOptimistic
    If (rs.RecordCount > 0) Then
        MsgBox "Your Policy Has Been Closed!!!!!", vbExclamation
        cmdok1.Enabled = False
    Else
        cmdok1.Enabled = True
    End If
End Sub

Private Sub cmbstatus_Click()
    If (cmbstatus.Text = "Client Death") Then
        cmbcod.Enabled = True
    Else
        cmbcod.Enabled = False
    End If
End Sub
Private Sub cmdok1_Click()
    If (cmbcid1.Text = "" Or cmbcod.Text = "") Then
        MsgBox "All Fields Are Compulsory!!!!!", vbExclamation
    Else
        Dim bonus As Double
        Dim pamt As Double
        Dim amt As Double
        Dim rs As ADODB.Recordset
        Call max
        Set rs = New ADODB.Recordset
        rs.Open "select pol_no from client_info where client_id=" + cmbcid1.Text + "", cn, adOpenStatic, adLockOptimistic
        txtpid.Text = rs.Fields(0)
        rs.Close
        rs.Open "select pol_id,pol_duration,pol_amount from policy_info where pol_no=" + txtpid.Text + "", cn, adOpenStatic, adLockOptimistic
        txtpid.Text = rs.Fields(0)
        txtpdur.Text = rs.Fields(1)
        txtpamt1.Text = rs.Fields(2)
        rs.Close
        rs.Open "select paid from premium_info where client_id=" + cmbcid1.Text + "", cn, adOpenStatic, adLockOptimistic
        txtpaid.Text = rs.Fields(0)
        rs.Close
        If (txtpid.Text = "149") Then
            amt = txtpamt1.Text / 1000
            bonus = 65 * txtpdur.Text * amt
                If (cmbcod.Text = "Natural") Then
                    pamt = bonus + txtpamt1.Text
                Else
                    pamt = bonus + (txtpamt1.Text * 2)
                End If
            
        ElseIf (txtpid.Text = "102") Then
                amt = txtpaid.Text / 1000
                bonus = 45 * txtpdur.Text * amt
                If (cmbcod.Text = "Natural") Then
                    pamt = bonus + txtpamt1.Text
                Else
                    pamt = bonus + (txtpamt1.Text * 2)
                End If
            
        ElseIf (txtpid.Text = "89") Then
                amt = txtpaid.Text / 1000
                bonus = 46 * txtpdur.Text * amt
                If (cmbcod.Text = "Natural") Then
                    pamt = bonus + txtpamt1.Text
                Else
                    pamt = bonus + (txtpamt1.Text * 2)
                End If
            
        End If
        txtcamt.Text = pamt
        cn.Execute ("insert into claim_info values(" + txtclaimid.Text + ",'" + CStr(dpkcdate.Value) + "','" + cmbstatus.Text + "','" + cmbcod.Text + "'," + txtcamt.Text + "," + cmbcid1.Text + ")")
        MsgBox "Your Claim Amount Is Sanctioned Successfully.....", vbInformation
            
       End If
End Sub

Private Sub Command1_Click()
HOME.Show

End Sub

Private Sub Form_Load()
    Call connect
    Call fill
    txtclaimid.Visible = False
    txtpid.Visible = False
    txtpamt1.Visible = False
    txtpdur.Visible = False
    txtstatus.Visible = False
    txtpaid.Visible = False
        cmbstatus.AddItem "Client Death"
        
        Call max
        dpkcdate.Enabled = False
        dpkcdate.Value = Date
        txtcamt.Enabled = False
        cmbcod.Enabled = False
    
End Sub


Public Function fill()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "select distinct client_id from client_info order by client_id", cn, adOpenStatic, adLockOptimistic
    While Not rs.EOF
        cmbcid1.AddItem rs.Fields(0)
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



