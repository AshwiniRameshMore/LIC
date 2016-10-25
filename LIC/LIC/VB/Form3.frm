VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmpolicy 
   Caption         =   "Insurance Form (Continue)"
   ClientHeight    =   3120
   ClientLeft      =   660
   ClientTop       =   825
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   11520
   ScaleWidth      =   19080
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker dpkddate 
      Height          =   255
      Left            =   11880
      TabIndex        =   42
      Top             =   10320
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Format          =   61210625
      CurrentDate     =   39732
   End
   Begin VB.TextBox txtdamt 
      Height          =   195
      Left            =   10200
      TabIndex        =   41
      Top             =   10320
      Width           =   150
   End
   Begin VB.TextBox txtpno 
      Height          =   195
      Left            =   11640
      TabIndex        =   40
      Top             =   10320
      Width           =   150
   End
   Begin VB.TextBox txtamt 
      Height          =   195
      Left            =   11400
      TabIndex        =   37
      Top             =   10320
      Width           =   150
   End
   Begin VB.TextBox txtstatus 
      Height          =   195
      Left            =   10680
      TabIndex        =   36
      Top             =   10320
      Width           =   150
   End
   Begin VB.TextBox txtpamt 
      Height          =   195
      Left            =   10440
      TabIndex        =   35
      Top             =   10320
      Width           =   150
   End
   Begin VB.TextBox txtaid 
      Height          =   195
      Left            =   11160
      TabIndex        =   34
      Top             =   10320
      Width           =   150
   End
   Begin VB.TextBox txtcid 
      Height          =   195
      Left            =   10920
      TabIndex        =   31
      Top             =   10320
      Width           =   150
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "< &Back"
      Height          =   495
      Left            =   6120
      TabIndex        =   30
      Top             =   10200
      Width           =   2295
   End
   Begin VB.CommandButton cmdhome 
      Caption         =   "<< &Back To Home"
      Height          =   495
      Left            =   3600
      TabIndex        =   29
      Top             =   10200
      Width           =   2295
   End
   Begin VB.Frame frmpinfo 
      Caption         =   "Policy Information"
      ForeColor       =   &H000000FF&
      Height          =   3015
      Left            =   3600
      TabIndex        =   18
      Top             =   3000
      Width           =   12375
      Begin VB.ComboBox cmbcage 
         Height          =   315
         ItemData        =   "Form3.frx":0000
         Left            =   3120
         List            =   "Form3.frx":0019
         TabIndex        =   38
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox txtpname 
         Height          =   285
         Left            =   9000
         TabIndex        =   33
         Top             =   600
         Width           =   1695
      End
      Begin VB.ComboBox cmbamt 
         Height          =   315
         ItemData        =   "Form3.frx":003D
         Left            =   9000
         List            =   "Form3.frx":0044
         TabIndex        =   25
         Top             =   1200
         Width           =   1695
      End
      Begin VB.ComboBox cmbpid 
         Height          =   315
         ItemData        =   "Form3.frx":004F
         Left            =   3120
         List            =   "Form3.frx":0068
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   480
         Width           =   1455
      End
      Begin VB.ComboBox cmbmode 
         Height          =   315
         ItemData        =   "Form3.frx":008C
         Left            =   9000
         List            =   "Form3.frx":0096
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1800
         Width           =   1695
      End
      Begin VB.ComboBox cmbdur 
         Height          =   315
         ItemData        =   "Form3.frx":00AF
         Left            =   3120
         List            =   "Form3.frx":00B1
         TabIndex        =   19
         Top             =   1080
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpckpd 
         Height          =   375
         Left            =   3120
         TabIndex        =   28
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   61210625
         CurrentDate     =   39708
      End
      Begin VB.Label Label6 
         Caption         =   "* Client Age :"
         Height          =   255
         Left            =   480
         TabIndex        =   39
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "* Name :"
         Height          =   375
         Left            =   6480
         TabIndex        =   32
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "* Proposal Date :"
         Height          =   255
         Left            =   480
         TabIndex        =   27
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "* Amount :"
         Height          =   255
         Left            =   6480
         TabIndex        =   26
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lable24 
         Caption         =   "* Mode :"
         Height          =   255
         Left            =   6480
         TabIndex        =   23
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "* Policy ID :"
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "* Duration (Years) :"
         Height          =   375
         Left            =   480
         TabIndex        =   21
         Top             =   1080
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdproceed 
      Caption         =   "&Proceed >>"
      Height          =   495
      Left            =   13680
      TabIndex        =   17
      Top             =   10200
      Width           =   2295
   End
   Begin VB.Frame frmoinfo 
      Caption         =   "Other Information"
      ForeColor       =   &H000000FF&
      Height          =   3975
      Left            =   3600
      TabIndex        =   0
      Top             =   6120
      Width           =   12375
      Begin VB.TextBox txtloswh 
         Height          =   285
         Left            =   9000
         TabIndex        =   8
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txtnope 
         Height          =   285
         Left            =   3120
         TabIndex        =   7
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txtenod 
         Height          =   285
         Left            =   9000
         TabIndex        =   6
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtpo 
         Height          =   285
         Left            =   3120
         TabIndex        =   5
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtpn 
         Height          =   285
         Left            =   9000
         TabIndex        =   4
         Top             =   3240
         Width           =   2775
      End
      Begin VB.TextBox txtsoi 
         Height          =   285
         Left            =   3120
         TabIndex        =   3
         Top             =   3240
         Width           =   2775
      End
      Begin VB.TextBox txtai 
         Height          =   285
         Left            =   9000
         TabIndex        =   2
         Top             =   2400
         Width           =   2775
      End
      Begin VB.TextBox txteq 
         Height          =   285
         Left            =   3120
         TabIndex        =   1
         Top             =   2400
         Width           =   2775
      End
      Begin VB.Label Label27 
         Caption         =   "Length Of Services With Him (Months):"
         Height          =   375
         Left            =   6600
         TabIndex        =   16
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label26 
         Caption         =   "Name Of Present Employer :"
         Height          =   495
         Left            =   600
         TabIndex        =   15
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label25 
         Caption         =   "* Exact Nature Of Duties :"
         Height          =   375
         Left            =   6480
         TabIndex        =   14
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label24 
         Caption         =   "* Present Occupation :"
         Height          =   615
         Left            =   480
         TabIndex        =   13
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label31 
         Caption         =   "Are You An Income Tax Assessee? If Yes Give Pan No.:"
         Height          =   615
         Left            =   6600
         TabIndex        =   12
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label Label30 
         Caption         =   "* Sources Of Income :"
         Height          =   375
         Left            =   480
         TabIndex        =   11
         Top             =   3240
         Width           =   2535
      End
      Begin VB.Label Label29 
         Caption         =   "* Annual Income :"
         Height          =   375
         Left            =   6480
         TabIndex        =   10
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label28 
         Caption         =   "Educational Qualification :"
         Height          =   495
         Left            =   600
         TabIndex        =   9
         Top             =   2400
         Width           =   2415
      End
   End
   Begin VB.Image Image10 
      Height          =   2835
      Left            =   3960
      Picture         =   "Form3.frx":00B3
      Top             =   120
      Width           =   11685
   End
   Begin VB.Image Image9 
      Height          =   765
      Left            =   360
      Picture         =   "Form3.frx":7D36
      Top             =   240
      Width           =   1470
   End
   Begin VB.Image Image8 
      Height          =   1410
      Left            =   16920
      Picture         =   "Form3.frx":E189
      Top             =   8520
      Width           =   1395
   End
   Begin VB.Image Image7 
      Height          =   1365
      Left            =   1080
      Picture         =   "Form3.frx":112A7
      Top             =   8640
      Width           =   1380
   End
   Begin VB.Image Image6 
      Height          =   1410
      Left            =   16920
      Picture         =   "Form3.frx":1202B
      Top             =   5640
      Width           =   1395
   End
   Begin VB.Image Image4 
      Height          =   1365
      Left            =   16920
      Picture         =   "Form3.frx":12E39
      Top             =   2760
      Width           =   1365
   End
   Begin VB.Image Image3 
      Height          =   1410
      Left            =   1080
      Picture         =   "Form3.frx":13C29
      Top             =   5640
      Width           =   1395
   End
   Begin VB.Image Image1 
      Height          =   1410
      Left            =   1080
      Picture         =   "Form3.frx":14B9E
      Top             =   2760
      Width           =   1395
   End
End
Attribute VB_Name = "frmpolicy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbmode_Click()
    If (cmbpid.Text = "160") Then
        If (cmbmode.Text = "Half-Yearly") Then
            MsgBox "Half-Yearly Mode Is Not Present!!!!!", vbExclamation
        End If
    End If
End Sub

Private Sub cmbpid_Click()
    If (cmbpid.Text = "149") Then
        cmbamt.Clear
        cmbdur.Clear
        cmbcage.Clear
        txtpname.Text = "Jeevan Anand"
        cmbamt.AddItem "100000"
        cmbamt.AddItem "500000"
        cmbamt.AddItem "1000000"
        cmbdur.AddItem "10"
        cmbdur.AddItem "15"
        cmbdur.AddItem "20"
        cmbdur.AddItem "25"
        cmbcage.AddItem "19"
        cmbcage.AddItem "20"
        cmbcage.AddItem "21"
        cmbcage.AddItem "22"
        cmbcage.AddItem "23"
    ElseIf (cmbpid.Text = "102") Then
        cmbamt.Clear
        cmbdur.Clear
        cmbcage.Clear
        txtpname.Text = "Jeevan Kishor"
        cmbamt.AddItem "50000"
        cmbamt.AddItem "100000"
        cmbamt.AddItem "200000"
        cmbdur.AddItem "20"
        cmbdur.AddItem "25"
        cmbdur.AddItem "30"
        cmbcage.AddItem "0"
        cmbcage.AddItem "1"
        cmbcage.AddItem "2"
        cmbcage.AddItem "3"
        cmbcage.AddItem "4"
        cmbcage.AddItem "5"
        cmbcage.AddItem "6"
        cmbcage.AddItem "7"
    ElseIf (cmbpid.Text = "89") Then
        cmbamt.Clear
        cmbdur.Clear
        cmbcage.Clear
        txtpname.Text = "Jeevan Saathi"
        cmbamt.AddItem "50000"
        cmbamt.AddItem "100000"
        cmbamt.AddItem "200000"
        cmbdur.AddItem "20"
        cmbdur.AddItem "25"
        cmbdur.AddItem "30"
        cmbcage.AddItem "24"
        cmbcage.AddItem "25"
        cmbcage.AddItem "26"
        cmbcage.AddItem "27"
    ElseIf (cmbpid.Text = "91") Then
        cmbamt.Clear
        cmbdur.Clear
        cmbcage.Clear
        txtpname.Text = "Jana Raksha"
        cmbamt.AddItem "30000"
        cmbamt.AddItem "50000"
        cmbamt.AddItem "100000"
        cmbdur.AddItem "15"
        cmbdur.AddItem "20"
        cmbdur.AddItem "25"
        cmbcage.AddItem "20"
        cmbcage.AddItem "21"
        cmbcage.AddItem "22"
        cmbcage.AddItem "23"
    ElseIf (cmbpid.Text = "160") Then
        cmbamt.Clear
        cmbdur.Clear
        cmbcage.Clear
        txtpname.Text = "Jeevan Bharati"
        cmbamt.AddItem "50000"
        cmbamt.AddItem "100000"
        cmbamt.AddItem "200000"
        cmbamt.AddItem "500000"
        cmbdur.AddItem "15"
        cmbdur.AddItem "20"
        cmbcage.AddItem "25"
        cmbcage.AddItem "26"
        cmbcage.AddItem "27"
        cmbcage.AddItem "28"
        cmbcage.AddItem "29"
        cmbcage.AddItem "30"
        cmbcage.AddItem "31"
        cmbcage.AddItem "32"
        cmbcage.AddItem "33"
        cmbcage.AddItem "34"
        cmbcage.AddItem "35"
    ElseIf (cmbpid.Text = "164") Then
        cmbamt.Clear
        cmbdur.Clear
        cmbcage.Clear
        txtpname.Text = "Anmol Jeevan"
        cmbamt.AddItem "1000000"
        cmbdur.AddItem "10"
        cmbdur.AddItem "15"
        cmbdur.AddItem "20"
        cmbdur.AddItem "25"
        cmbcage.AddItem "20"
        cmbcage.AddItem "21"
        cmbcage.AddItem "22"
        cmbcage.AddItem "23"
    End If
End Sub

Private Sub cmdback_Click()
    Call max
    If (txtcid.Text = "0") Then
    Else
        con.Execute ("delete from client_info where client_id = " + txtcid.Text + "")
    End If
    Unload Me
    frmclient.Show
End Sub

Private Sub cmdhome_Click()
    Call max
    If (txtcid.Text = "0") Then
    Else
        con.Execute ("delete from client_info where client_id = " + txtcid.Text + "")
    End If
    Unload Me
    frmhome.Show
End Sub

Private Sub cmdproceed_Click()
    If (cmbpid.Text = "" Or txtpname.Text = "" Or cmbdur.Text = "" Or cmbamt.Text = "" Or cmbmode.Text = "" Or dtpckpd.Value = "" Or cmbcage.Text = "" Or txtpo.Text = "" Or txtenod.Text = "" Or txtai.Text = "" Or txtsoi.Text = "") Then
        MsgBox "All * Marked Fields Are Compulsory!!!!!", vbExclamation
    Else
        Dim fine As Integer
        Dim rst As ADODB.Recordset
        Call max1
        con.Execute ("insert into policy_info values(" + txtpno.Text + "," + cmbpid.Text + ",'" + txtpname.Text + "'," + cmbamt.Text + "," + cmbdur.Text + ",'" + cmbmode.Text + "','" + CStr(dtpckpd.Value) + "'," + cmbcage.Text + ")")
        'con.Execute ("insert into other_info values(" + txtcid.Text + ",'" + txtpo.Text + "','" + txtenod.Text + "','" + txtnope.Text + "'," + txtloswh.Text + ",'" + txteq.Text + "'," + txtai.Text + ",'" + txtsoi.Text + "'," + txtpn.Text + ")")
        Set rst = New ADODB.Recordset
        rst.Open "select agent_id from client_info where client_id=" + txtcid.Text + "", con, adOpenStatic, adLockOptimistic
        txtaid.Text = rst.Fields(0)
        rst.Close
        If (cmbpid.Text = "149") Then
            If (cmbamt.Text = "100000") Then
                If (cmbmode.Text = "Yearly") Then
                    rst.Open "select onelac_yr from jeevan_anand where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", con, adOpenStatic, adLockOptimistic
                Else
                    rst.Open "select onelac_hr from jeevan_anand where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", con, adOpenStatic, adLockOptimistic
                End If
            ElseIf (cmbamt.Text = "500000") Then
                If (cmbmode.Text = "Yearly") Then
                    rst.Open "select fivelac_yr from jeevan_anand where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", con, adOpenStatic, adLockOptimistic
                Else
                    rst.Open "select fivelac_hr from jeevan_anand where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", con, adOpenStatic, adLockOptimistic
                End If
            Else
                If (cmbmode.Text = "Yearly") Then
                    rst.Open "select tenlac_yr from jeevan_anand where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", con, adOpenStatic, adLockOptimistic
                Else
                    rst.Open "select tenlac_hr from jeevan_anand where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", con, adOpenStatic, adLockOptimistic
                End If
            End If
        ElseIf (cmbpid.Text = "102") Then
            If (cmbamt.Text = "50000") Then
                If (cmbmode.Text = "Yearly") Then
                    rst.Open "select fifty_yr from jeevan_kishor where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", con, adOpenStatic, adLockOptimistic
                Else
                    rst.Open "select fifty_hr from jeevan_kishor where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", con, adOpenStatic, adLockOptimistic
                End If
            ElseIf (cmbamt.Text = "100000") Then
                If (cmbmode.Text = "Yearly") Then
                    rst.Open "select onelac_yr from jeevan_kishor where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", con, adOpenStatic, adLockOptimistic
                Else
                    rst.Open "select onelac_hr from jeevan_kishor where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", con, adOpenStatic, adLockOptimistic
                End If
            Else
                If (cmbmode.Text = "Yearly") Then
                    rst.Open "select twolac_yr from jeevan_kishor where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", con, adOpenStatic, adLockOptimistic
                Else
                    rst.Open "select twolac_hr from jeevan_kishor where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", con, adOpenStatic, adLockOptimistic
                End If
            End If
        ElseIf (cmbpid.Text = "89") Then
            If (cmbamt.Text = "50000") Then
                If (cmbmode.Text = "Yearly") Then
                    rst.Open "select fifty_yr from jeevan_saathi where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", con, adOpenStatic, adLockOptimistic
                Else
                    rst.Open "select fifty_hr from jeevan_saathi where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", con, adOpenStatic, adLockOptimistic
                End If
            ElseIf (cmbamt.Text = "100000") Then
                If (cmbmode.Text = "Yearly") Then
                    rst.Open "select onelac_yr from jeevan_saathi where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", con, adOpenStatic, adLockOptimistic
                Else
                    rst.Open "select onelac_hr from jeevan_saathi where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", con, adOpenStatic, adLockOptimistic
                End If
            Else
                If (cmbmode.Text = "Yearly") Then
                    rst.Open "select twolac_yr from jeevan_saathi where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", con, adOpenStatic, adLockOptimistic
                Else
                    rst.Open "select twolac_hr from jeevan_saathi where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", con, adOpenStatic, adLockOptimistic
                End If
            End If
        ElseIf (cmbpid.Text = "91") Then
            If (cmbamt.Text = "30000") Then
                If (cmbmode.Text = "Yearly") Then
                    rst.Open "select thirty_yr from jana_raksha where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", con, adOpenStatic, adLockOptimistic
                Else
                    rst.Open "select thirty_hr from jana_raksha where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", con, adOpenStatic, adLockOptimistic
                End If
            ElseIf (cmbamt.Text = "50000") Then
                If (cmbmode.Text = "Yearly") Then
                    rst.Open "select fifty_yr from jana_raksha where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", con, adOpenStatic, adLockOptimistic
                Else
                    rst.Open "select fifty_hr from jana_raksha where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", con, adOpenStatic, adLockOptimistic
                End If
            Else
                If (cmbmode.Text = "Yearly") Then
                    rst.Open "select onelac_yr from jana_raksha where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", con, adOpenStatic, adLockOptimistic
                Else
                    rst.Open "select onelac_hr from jana_raksha where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", con, adOpenStatic, adLockOptimistic
                End If
            End If
        ElseIf (cmbpid.Text = "160") Then
            If (cmbamt.Text = "50000") Then
                rst.Open "select fifty_yr from jeevan_bharati where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", con, adOpenStatic, adLockOptimistic
            ElseIf (cmbamt.Text = "100000") Then
                rst.Open "select onelac_yr from jeevan_bharati where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", con, adOpenStatic, adLockOptimistic
            ElseIf (cmbamt.Text = "200000") Then
                rst.Open "select twolac_yr from jeevan_bharati where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", con, adOpenStatic, adLockOptimistic
            Else
                rst.Open "select fivelac_yr from jeevan_bharati where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", con, adOpenStatic, adLockOptimistic
            End If
        ElseIf (cmbpid.Text = "164") Then
            If (cmbmode.Text = "Yearly") Then
                rst.Open "select tenlac_yr from anmol_jeevan where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", con, adOpenStatic, adLockOptimistic
            Else
                rst.Open "select tenlac_hr from anmol_jeevan where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", con, adOpenStatic, adLockOptimistic
            End If
        End If
        txtpamt.Text = rst.Fields(0)
        fine = txtpamt * 9 / 100
        txtdamt.Text = txtpamt.Text + fine
        rst.Close
        txtamt.Text = "0"
        txtstatus.Text = "Unpaid"
        dtpckpd.Value = DateSerial(dtpckpd.Year, dtpckpd.Month, dtpckpd.Day)
        dpkddate.Value = DateAdd("m", 1, dtpckpd.Value)
        con.Execute "insert into premium_info values(" + txtcid.Text + "," + txtaid.Text + "," + cmbpid.Text + "," + cmbamt.Text + "," + txtamt.Text + "," + txtpamt.Text + ",'" + CStr(dtpckpd.Value) + "','" + CStr(dpkddate.Value) + "'," + txtdamt.Text + ",'" + txtstatus.Text + "')"
        Unload Me
        frmagent.Show
    End If
End Sub

Private Sub Form_Load()
    Call conn
    dtpckpd.Value = Date
    txtpname.Enabled = False
    txtdamt.Visible = False
    txtamt.Visible = False
    txtaid.Visible = False
    txtstatus.Visible = False
    txtpamt.Visible = False
    txtcid.Visible = False
    txtpno.Visible = False
    dpkddate.Visible = False
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
End Function
Public Function max1()
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
        clientid = 1001
        txtpno.Text = clientid
    Else
        clientid = rst.Fields(0) + 1
        txtpno.Text = clientid
    End If
    rst.Close
End Function

Private Sub txtai_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8 And KeyAscii <> 13) Then
        KeyAscii = 0
        MsgBox "Please Enter Integer Only!!!!!", vbExclamation
    End If
    If (KeyAscii = 13) Then
        txtsoi.SetFocus
    End If
End Sub

Private Sub txtenod_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789`-=[]\;'/~!@#$%^&*()_+{}|:<>?"
    i = InStr(nstr, Chr(KeyAscii))
    If (i > 0) Then
        KeyAscii = 0
        MsgBox "Please Enter Character Only!!!!!", vbExclamation
    End If
    If (KeyAscii = 13) Then
        txtnope.SetFocus
    End If
End Sub

Private Sub txteq_KeyPress(KeyAscii As Integer)
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
       txtai.SetFocus
    End If
End Sub

Private Sub txtloswh_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8 And KeyAscii <> 13) Then
        KeyAscii = 0
        MsgBox "Please Enter Integer Only!!!!!", vbExclamation
    End If
    If (KeyAscii = 13) Then
        txteq.SetFocus
    End If
End Sub

Private Sub txtnope_KeyPress(KeyAscii As Integer)
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
        txtloswh.SetFocus
    End If
End Sub

Private Sub txtpn_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8 And KeyAscii <> 13) Then
        KeyAscii = 0
        MsgBox "Please Enter Integer Only!!!!!", vbExclamation
    End If
    If (KeyAscii = 13) Then
        cmdproceed.SetFocus
    End If
End Sub

Private Sub txtpo_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789`-=[]\;'/~!@#$%^&*()_+{}|:<>?"
    i = InStr(nstr, Chr(KeyAscii))
    If (i > 0) Then
        KeyAscii = 0
        MsgBox "Please Enter Character Only!!!!!", vbExclamation
    End If
    If (KeyAscii = 13) Then
        txtenod.SetFocus
    End If
End Sub

Private Sub txtsoi_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789`-=[]\;'/~!@#$%^&*()_+{}|:<>?"
    i = InStr(nstr, Chr(KeyAscii))
    If (i > 0) Then
        KeyAscii = 0
        MsgBox "Please Enter Character Only!!!!!", vbExclamation
    End If
    If (KeyAscii = 13) Then
        txtpn.SetFocus
    End If
End Sub
