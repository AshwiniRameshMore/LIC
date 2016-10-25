VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmpremium 
   Caption         =   "Premium"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4680
   LinkTopic       =   "Form7"
   MDIChild        =   -1  'True
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtpaid 
      Height          =   195
      Left            =   12360
      TabIndex        =   23
      Top             =   10200
      Width           =   150
   End
   Begin VB.TextBox txtpdur 
      Height          =   195
      Left            =   13320
      TabIndex        =   21
      Top             =   10200
      Width           =   150
   End
   Begin VB.TextBox txtpamt1 
      Height          =   195
      Left            =   13560
      TabIndex        =   20
      Top             =   10200
      Width           =   150
   End
   Begin VB.TextBox txtpid 
      Height          =   195
      Left            =   13080
      TabIndex        =   19
      Top             =   10200
      Width           =   150
   End
   Begin VB.TextBox txtstatus 
      Height          =   195
      Left            =   12840
      TabIndex        =   18
      Top             =   10200
      Width           =   150
   End
   Begin VB.TextBox txtclaimid 
      Height          =   195
      Left            =   12600
      TabIndex        =   17
      Top             =   10200
      Width           =   150
   End
   Begin VB.CommandButton cmdhome 
      Caption         =   "<< &Back To Home"
      Height          =   495
      Left            =   8040
      TabIndex        =   12
      Top             =   9960
      Width           =   2295
   End
   Begin VB.Frame frmclaim 
      Caption         =   "Get Claim Amount"
      Height          =   3975
      Left            =   8040
      TabIndex        =   1
      Top             =   5640
      Width           =   9255
      Begin VB.CommandButton cmdok1 
         Caption         =   "&OK"
         Height          =   375
         Left            =   6600
         TabIndex        =   27
         Top             =   1920
         Width           =   2055
      End
      Begin VB.ComboBox cmbcid1 
         Height          =   315
         ItemData        =   "Form7.frx":0000
         Left            =   3480
         List            =   "Form7.frx":0002
         TabIndex        =   15
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtcamt 
         Height          =   285
         Left            =   3480
         TabIndex        =   13
         Top             =   3360
         Width           =   1455
      End
      Begin VB.ComboBox cmbcod 
         Height          =   315
         ItemData        =   "Form7.frx":0004
         Left            =   3480
         List            =   "Form7.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2640
         Width           =   1455
      End
      Begin VB.ComboBox cmbstatus 
         Height          =   315
         ItemData        =   "Form7.frx":0027
         Left            =   3480
         List            =   "Form7.frx":0031
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1920
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dpkcdate 
         Height          =   375
         Left            =   3480
         TabIndex        =   25
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20643841
         CurrentDate     =   39732
      End
      Begin VB.Label Label8 
         Caption         =   "Date :"
         Height          =   255
         Left            =   720
         TabIndex        =   22
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Claim Amount :"
         Height          =   375
         Left            =   720
         TabIndex        =   14
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Status :"
         Height          =   375
         Left            =   720
         TabIndex        =   10
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Cause Of Death :"
         Height          =   375
         Left            =   720
         TabIndex        =   8
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Enter The Client-ID :"
         Height          =   375
         Left            =   720
         TabIndex        =   6
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.Frame frmpaypremium 
      Caption         =   "Pay Premium"
      Height          =   2415
      Left            =   8040
      TabIndex        =   0
      Top             =   3000
      Width           =   9255
      Begin VB.CommandButton cmdok 
         Caption         =   "&OK"
         Height          =   375
         Left            =   6600
         TabIndex        =   26
         Top             =   720
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker dpkdate 
         Height          =   375
         Left            =   3480
         TabIndex        =   24
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20643841
         CurrentDate     =   39732
      End
      Begin VB.ComboBox cmbcid 
         Height          =   315
         ItemData        =   "Form7.frx":0053
         Left            =   3480
         List            =   "Form7.frx":0055
         TabIndex        =   16
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton cmdget 
         Caption         =   "&Get Premium Receipt"
         Height          =   375
         Left            =   6600
         TabIndex        =   5
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtpamt 
         Height          =   285
         Left            =   3480
         TabIndex        =   4
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Date :"
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Premium Amount :"
         Height          =   375
         Left            =   720
         TabIndex        =   3
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Enter The Client ID :"
         Height          =   375
         Left            =   720
         TabIndex        =   2
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Image Image4 
      Height          =   765
      Left            =   360
      Picture         =   "Form7.frx":0057
      Top             =   240
      Width           =   1470
   End
   Begin VB.Image Image2 
      Height          =   6000
      Left            =   600
      Picture         =   "Form7.frx":64AA
      Top             =   3360
      Width           =   5940
   End
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   8520
      Picture         =   "Form7.frx":A558
      Top             =   600
      Width           =   8475
   End
End
Attribute VB_Name = "frmpremium"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbcid_Click()
    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    rst.Open "select claim_id from claim_info where client_id=" + cmbcid.Text + "", con, adOpenStatic, adLockOptimistic
    If (rst.RecordCount > 0) Then
        MsgBox "Your Policy Has Been Closed!!!!!", vbExclamation
        cmdok.Enabled = False
    Else
        cmdok.Enabled = True
        rst.Close
        rst.Open "select due_date from premium_info where client_id=" + cmbcid.Text + "", con, adOpenStatic, adLockOptimistic
        If (dpkdate.Value <= rst.Fields(0)) Then
            rst.Close
            rst.Open "select premium_amt from premium_info where client_id=" + cmbcid.Text + "", con, adOpenStatic, adLockOptimistic
        Else
            rst.Close
            rst.Open "select due_amt from premium_info where client_id=" + cmbcid.Text + "", con, adOpenStatic, adLockOptimistic
        End If
        txtpamt.Text = rst.Fields(0)
        rst.Close
    End If
End Sub

Private Sub cmbcid1_Click()
    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    rst.Open "select claim_id from claim_info where client_id=" + cmbcid1.Text + "", con, adOpenStatic, adLockOptimistic
    If (rst.RecordCount > 0) Then
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

Private Sub cmdget_Click()
    Unload Me
    frmreceipt.Show
End Sub

Private Sub cmdhome_Click()
    Unload Me
    frmhome.Show
End Sub

Private Sub cmdok_Click()
    If (cmbcid.Text = "") Then
        MsgBox "Please Select Client-ID!!!!!", vbExclamation
    Else
        Dim paid As Integer
        Dim rst As ADODB.Recordset
        Set rst = New ADODB.Recordset
        txtstatus.Text = "Paid"
        rst.Open "select paid from premium_info where client_id=" + cmbcid.Text + "", con, adOpenStatic, adLockOptimistic
        paid = rst.Fields(0)
        txtpaid.Text = paid + txtpamt.Text
        con.Execute ("update premium_info set paid=" + txtpaid.Text + ",status='" + txtstatus.Text + "' where client_id=" + cmbcid.Text + "")
        MsgBox "Premium Paid Successfully.....", vbInformation
    End If
End Sub

Private Sub cmdok1_Click()
    If (cmbstatus.Text = "Policy Matured") Then
        cmbcod.Text = "NA"
    End If
    If (cmbcid1.Text = "" Or cmbstatus.Text = "" Or cmbcod.Text = "") Then
        MsgBox "All Fields Are Compulsory!!!!!", vbExclamation
    Else
        Dim bonus As Double
        Dim pamt As Double
        Dim amt As Double
        Dim rst As ADODB.Recordset
        Call max
        Set rst = New ADODB.Recordset
        rst.Open "select pol_no from client_info where client_id=" + cmbcid1.Text + "", con, adOpenStatic, adLockOptimistic
        txtpid.Text = rst.Fields(0)
        rst.Close
        rst.Open "select pol_id,pol_duration,pol_amount from policy_info where pol_no=" + txtpid.Text + "", con, adOpenStatic, adLockOptimistic
        txtpid.Text = rst.Fields(0)
        txtpdur.Text = rst.Fields(1)
        txtpamt1.Text = rst.Fields(2)
        rst.Close
        rst.Open "select paid from premium_info where client_id=" + cmbcid1.Text + "", con, adOpenStatic, adLockOptimistic
        txtpaid.Text = rst.Fields(0)
        rst.Close
        If (txtpid.Text = "149") Then
            amt = txtpamt1.Text / 1000
            bonus = 65 * txtpdur.Text * amt
            If (cmbstatus.Text = "Policy Matured") Then
                pamt = bonus + txtpamt1.Text
            Else
                If (cmbcod.Text = "Natural") Then
                    pamt = bonus + txtpamt1.Text
                Else
                    pamt = bonus + (txtpamt1.Text * 2)
                End If
            End If
        ElseIf (txtpid.Text = "102") Then
            If (cmbstatus.Text = "Policy Matured") Then
                amt = txtpamt1.Text / 1000
                bonus = 45 * txtpdur.Text * amt
                pamt = bonus + txtpamt1.Text
            Else
                amt = txtpaid.Text / 1000
                bonus = 45 * txtpdur.Text * amt
                If (cmbcod.Text = "Natural") Then
                    pamt = bonus + txtpamt1.Text
                Else
                    pamt = bonus + (txtpamt1.Text * 2)
                End If
            End If
        ElseIf (txtpid.Text = "89") Then
            If (cmbstatus.Text = "Policy Matured") Then
                amt = txtpamt1.Text / 1000
                bonus = 46 * txtpdur.Text * amt
                pamt = bonus + txtpamt1.Text
            Else
                amt = txtpaid.Text / 1000
                bonus = 46 * txtpdur.Text * amt
                If (cmbcod.Text = "Natural") Then
                    pamt = bonus + txtpamt1.Text
                Else
                    pamt = bonus + (txtpamt1.Text * 2)
                End If
            End If
        ElseIf (txtpid.Text = "91") Then
            If (cmbstatus.Text = "Policy Matured") Then
                amt = txtpamt1.Text / 1000
                bonus = 45 * txtpdur.Text * amt
                pamt = bonus + txtpamt1.Text
            Else
                amt = txtpaid.Text / 1000
                bonus = 45 * txtpdur.Text * amt
                If (cmbcod.Text = "Natural") Then
                    pamt = bonus + txtpamt1.Text
                Else
                    pamt = bonus + (txtpamt1.Text * 2)
                End If
            End If
        ElseIf (txtpid.Text = "160") Then
            If (cmbstatus.Text = "Policy Matured") Then
                amt = txtpamt1.Text / 1000
                bonus = 50 * txtpdur.Text * amt
                pamt = bonus + txtpamt1.Text
            Else
                amt = txtpaid.Text / 1000
                bonus = 50 * txtpdur.Text * amt
                If (cmbcod.Text = "Natural") Then
                    pamt = bonus + txtpamt1.Text
                Else
                    pamt = bonus + (txtpamt1.Text * 2)
                End If
            End If
        ElseIf (txtpid.Text = "164") Then
            If (cmbstatus.Text = "Policy Matured") Then
                pamt = 0
            Else
                If (cmbcod.Text = "Natural") Then
                    pamt = txtpamt1.Text
                Else
                    pamt = txtpamt1.Text
                End If
            End If
        End If
        txtcamt.Text = pamt
        con.Execute ("insert into claim_info values(" + txtclaimid.Text + ",'" + CStr(dpkcdate.Value) + "','" + cmbstatus.Text + "','" + cmbcod.Text + "'," + txtcamt.Text + "," + cmbcid1.Text + ")")
        MsgBox "Your Claim Amount Is Sanctioned Successfully.....", vbInformation
        MsgBox "Your Policy Has Been Closed.....", vbInformation
    End If
End Sub

Private Sub Form_Load()
    Call conn
    Call fill
    txtclaimid.Visible = False
    txtpid.Visible = False
    txtpamt1.Visible = False
    txtpdur.Visible = False
    txtstatus.Visible = False
    txtpaid.Visible = False
    If (frmpaypremium.Enabled = True) Then
        dpkdate.Enabled = False
        dpkdate.Value = Date
        txtpamt.Enabled = False
    Else
        Call max
        dpkcdate.Enabled = False
        dpkcdate.Value = Date
        txtcamt.Enabled = False
        cmbcod.Enabled = False
    End If
End Sub
Public Function fill()
    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    rst.Open "select distinct client_id from client_info order by client_id", con, adOpenStatic, adLockOptimistic
    While Not rst.EOF
        cmbcid.AddItem rst.Fields(0)
        cmbcid1.AddItem rst.Fields(0)
        rst.MoveNext
    Wend
End Function

Public Function max()
    Dim claimid As Integer
    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    rst.Open "select max(claim_id) from claim_info", con, adOpenStatic, adLockOptimistic
    If IsNull(rst.Fields(0)) Then
        claimid = 4001
        txtclaimid.Text = claimid
    Else
        claimid = rst.Fields(0) + 1
        txtclaimid.Text = claimid
    End If
    rst.Close
End Function

