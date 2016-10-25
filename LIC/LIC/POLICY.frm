VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form POLICY1 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "POLICY"
   ClientHeight    =   10170
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   16515
   FillColor       =   &H00FFC0C0&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00FFC0C0&
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "POLICY.frx":0000
   ScaleHeight     =   10170
   ScaleWidth      =   16515
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture8 
      Height          =   855
      Left            =   720
      Picture         =   "POLICY.frx":6453
      ScaleHeight     =   795
      ScaleWidth      =   1395
      TabIndex        =   51
      Top             =   480
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker dpkddate 
      Height          =   375
      Left            =   1080
      TabIndex        =   49
      Top             =   10080
      Visible         =   0   'False
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   661
      _Version        =   393216
      Format          =   100728833
      CurrentDate     =   40456
   End
   Begin VB.TextBox txtpno 
      Height          =   285
      Left            =   1800
      TabIndex        =   48
      Text            =   "Text3"
      Top             =   9960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtamt 
      Height          =   285
      Left            =   360
      TabIndex        =   47
      Text            =   "Text2"
      Top             =   9960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtaid 
      Height          =   285
      Left            =   600
      TabIndex        =   46
      Text            =   "Text1"
      Top             =   9960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtcid 
      Height          =   285
      Left            =   2040
      TabIndex        =   45
      Text            =   "Text4"
      Top             =   9960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtstatus 
      Height          =   285
      Left            =   1320
      TabIndex        =   44
      Text            =   "Text3"
      Top             =   9960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtpamt 
      Height          =   285
      Left            =   1560
      TabIndex        =   43
      Text            =   "Text2"
      Top             =   9960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtdamt 
      Height          =   285
      Left            =   840
      TabIndex        =   42
      Top             =   9960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Frame frame3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "POLICY"
      Height          =   32055
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   51255
      Begin VB.CommandButton PROCEED 
         Caption         =   "PROCEED>>"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   12480
         TabIndex        =   14
         Top             =   9240
         Width           =   2295
      End
      Begin VB.CommandButton back 
         Caption         =   "<BACK"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9720
         TabIndex        =   16
         Top             =   9240
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "<<BACK TO HOME"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5520
         TabIndex        =   15
         Top             =   9240
         Width           =   2535
      End
      Begin VB.PictureBox Picture7 
         Height          =   2295
         Left            =   5280
         Picture         =   "POLICY.frx":C8A6
         ScaleHeight     =   2235
         ScaleWidth      =   9315
         TabIndex        =   41
         Top             =   360
         Width           =   9375
      End
      Begin VB.PictureBox Picture6 
         Height          =   1455
         Left            =   16080
         Picture         =   "POLICY.frx":10CDA
         ScaleHeight     =   1395
         ScaleWidth      =   1395
         TabIndex        =   24
         Top             =   7440
         Width           =   1455
      End
      Begin VB.PictureBox Picture5 
         Height          =   1455
         Left            =   16200
         Picture         =   "POLICY.frx":11ACA
         ScaleHeight     =   1395
         ScaleWidth      =   1395
         TabIndex        =   23
         Top             =   5040
         Width           =   1455
      End
      Begin VB.PictureBox Picture4 
         Height          =   1455
         Left            =   16200
         Picture         =   "POLICY.frx":1284E
         ScaleHeight     =   1395
         ScaleWidth      =   1395
         TabIndex        =   22
         Top             =   2400
         Width           =   1455
      End
      Begin VB.PictureBox Picture3 
         Height          =   1455
         Left            =   2520
         Picture         =   "POLICY.frx":1361E
         ScaleHeight     =   1395
         ScaleWidth      =   1395
         TabIndex        =   21
         Top             =   7320
         Width           =   1455
      End
      Begin VB.PictureBox Picture2 
         Height          =   1455
         Left            =   2520
         Picture         =   "POLICY.frx":14582
         ScaleHeight     =   1395
         ScaleWidth      =   1395
         TabIndex        =   20
         Top             =   5040
         Width           =   1455
      End
      Begin VB.PictureBox Picture1 
         Height          =   1455
         Left            =   2520
         Picture         =   "POLICY.frx":15384
         ScaleHeight     =   1395
         ScaleWidth      =   1395
         TabIndex        =   19
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Frame frame3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "OTHER INFORMATION"
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
         Height          =   3375
         Index           =   0
         Left            =   5520
         TabIndex        =   18
         Top             =   5760
         Width           =   9255
         Begin VB.TextBox opan 
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
            Left            =   6840
            TabIndex        =   13
            Top             =   2520
            Width           =   2175
         End
         Begin VB.TextBox oincome 
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
            Left            =   6840
            TabIndex        =   12
            Top             =   1800
            Width           =   2175
         End
         Begin VB.TextBox olen 
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
            Left            =   6840
            TabIndex        =   11
            Top             =   1080
            Width           =   2175
         End
         Begin VB.TextBox oduty 
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
            Left            =   6840
            TabIndex        =   10
            Top             =   360
            Width           =   2175
         End
         Begin VB.TextBox osource 
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
            Left            =   2280
            TabIndex        =   9
            Top             =   2520
            Width           =   2295
         End
         Begin VB.TextBox oqual 
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
            Left            =   2280
            TabIndex        =   8
            Top             =   1800
            Width           =   2295
         End
         Begin VB.TextBox oname 
            BackColor       =   &H00FFFFFF&
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
            Left            =   2280
            TabIndex        =   7
            Top             =   1080
            Width           =   2295
         End
         Begin VB.TextBox oocc 
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
            Left            =   2280
            TabIndex        =   6
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label Label15 
            BackColor       =   &H00FFC0C0&
            Caption         =   "ARE YOU ON INCOME TAX ASSESSEE?IF YES,GIVE PAN NO.:"
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
            Left            =   4680
            TabIndex        =   39
            Top             =   2520
            Width           =   2295
         End
         Begin VB.Label Label14 
            BackColor       =   &H00FFC0C0&
            Caption         =   "ANNUAL INCOME:"
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
            Left            =   4680
            TabIndex        =   38
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label Label13 
            BackColor       =   &H00FFC0C0&
            Caption         =   "LENTH OF SERVICES WITH HIM(MONTHS):"
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
            Left            =   4680
            TabIndex        =   37
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label Label12 
            BackColor       =   &H00FFC0C0&
            Caption         =   "EXACT NATURE OF DUTIES:"
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
            Left            =   4680
            TabIndex        =   36
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label11 
            BackColor       =   &H00FFC0C0&
            Caption         =   "SOURCES OF INCOME:"
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
            TabIndex        =   35
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Label Label10 
            BackColor       =   &H00FFC0C0&
            Caption         =   "EDUCATIONAL QUALIFICATION:"
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
            TabIndex        =   34
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label Label9 
            BackColor       =   &H00FFC0C0&
            Caption         =   "NAME OF PRESENT EMPLOYER:"
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
            TabIndex        =   33
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FFC0C0&
            Caption         =   "PRESENT OCCUPATION:"
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
            TabIndex        =   32
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "POLICY INFORMATION"
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
         Height          =   3255
         Left            =   5520
         TabIndex        =   17
         Top             =   2400
         Width           =   9255
         Begin MSComCtl2.DTPicker dtprodate 
            Height          =   375
            Left            =   1920
            TabIndex        =   50
            Top             =   1920
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   661
            _Version        =   393216
            Format          =   100728833
            CurrentDate     =   40456
         End
         Begin VB.TextBox pname 
            Height          =   375
            Left            =   5760
            TabIndex        =   40
            Top             =   360
            Width           =   2535
         End
         Begin VB.ComboBox cmbmode 
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
            Left            =   5760
            TabIndex        =   4
            Top             =   1800
            Width           =   2535
         End
         Begin VB.ComboBox cmbamt 
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
            Left            =   5760
            TabIndex        =   3
            Top             =   1080
            Width           =   2535
         End
         Begin VB.ComboBox cmbcage 
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
            Left            =   1920
            TabIndex        =   5
            Top             =   2640
            Width           =   2535
         End
         Begin VB.ComboBox cmbpid 
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
            Left            =   1920
            TabIndex        =   1
            Top             =   360
            Width           =   2535
         End
         Begin VB.ComboBox cmbdur 
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
            Left            =   1920
            TabIndex        =   2
            Top             =   1200
            Width           =   2535
         End
         Begin VB.Label Label7 
            BackColor       =   &H00FFC0C0&
            Caption         =   "MODE:"
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
            Left            =   4800
            TabIndex        =   31
            Top             =   1800
            Width           =   1935
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFC0C0&
            Caption         =   "AMOUNT:"
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
            Left            =   4800
            TabIndex        =   30
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFC0C0&
            Caption         =   "NAME:"
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
            Left            =   4800
            TabIndex        =   29
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "CLIENT AGE:"
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
            Left            =   240
            TabIndex        =   28
            Top             =   2640
            Width           =   2055
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "PROPOSAL DATE:"
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
            TabIndex        =   27
            Top             =   1920
            Width           =   1935
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "DURATION(YEARS):"
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
            TabIndex        =   26
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "POLICY ID:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   615
            Left            =   240
            TabIndex        =   25
            Top             =   360
            Width           =   2055
         End
      End
   End
End
Attribute VB_Name = "POLICY1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub back_Click()
 Call max
    If (txtcid.Text = "0") Then
    Else
    cn.Execute ("delete from client_info where client_id = " + txtcid.Text + "")
    End If
    Unload Me
    CLIENT.Show

End Sub

 Private Sub cmbpid_Click()
  If (cmbpid.Text = "149") Then
        cmbamt.Clear
        cmbdur.Clear
        cmbcage.Clear
        pname.Text = "Jeevan Anand"
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
        pname.Text = "Jeevan Kishore"
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
        pname.Text = "Jeevan Saral"
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
    End If

End Sub

Private Sub Command1_Click()
Call max
    If (txtcid.Text = "0") Then
    Else
        cn.Execute ("delete from client_info where client_id = " + txtcid.Text + "")
    End If
    Unload Me
    HOME.Show

End Sub

Public Function max()
 Dim clientid As Integer
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "select max(client_id) from client_info", cn, adOpenStatic, adLockOptimistic
    If IsNull(rs.Fields(0)) Then
        clientid = 0
        txtcid.Text = clientid
    Else
        clientid = rs.Fields(0)
        txtcid.Text = clientid
    End If
    rs.Close
End Function

Private Sub Form_Load()
Call connect
   dtprodate.Value = Date
cmbpid.AddItem "149"
cmbpid.AddItem "102"
cmbpid.AddItem "89"
cmbmode.AddItem "yearly"
cmbmode.AddItem "half-yearly"

End Sub



Private Sub PROCEED_Click()
 If (cmbpid.Text = "" Or pname.Text = "" Or cmbdur.Text = "" Or cmbamt.Text = "" Or cmbmode.Text = "" Or dtprodate.Value = "" Or cmbcage.Text = "" Or oocc.Text = "" Or oduty.Text = "" Or osource.Text = "" Or oincome.Text = "") Then
        MsgBox "All * Marked Fields Are Compulsory!!!!!", vbExclamation
    Else
        Dim fine As Integer
        Dim rs As ADODB.Recordset
        Call MAX1
        cn.Execute ("insert into policy_info values(" + txtpno.Text + "," + cmbpid.Text + ",'" + pname.Text + "'," + cmbamt.Text + "," + cmbdur.Text + ",'" + cmbmode.Text + "','" + CStr(dtprodate.Value) + "'," + cmbcage.Text + ")")
        Set rs = New ADODB.Recordset
        rs.Open "select agent_id from client_info where client_id=" + txtcid.Text + "", cn, adOpenStatic, adLockOptimistic
        txtaid.Text = rs.Fields(0)
        rs.Close
         If (cmbpid.Text = "149") Then
            If (cmbamt.Text = "100000") Then
                If (cmbmode.Text = "Yearly") Then
                    rs.Open "select onelac_yr from jeevan_anand where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", cn, adOpenStatic, adLockOptimistic
                Else
                    rs.Open "select onelac_hr from jeevan_anand where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", cn, adOpenStatic, adLockOptimistic
                End If
            ElseIf (cmbamt.Text = "500000") Then
                If (cmbmode.Text = "Yearly") Then
                    rs.Open "select fivelac_yr from jeevan_anand where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", cn, adOpenStatic, adLockOptimistic
                Else
                    rs.Open "select fivelac_hr from jeevan_anand where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", cn, adOpenStatic, adLockOptimistic
                End If
            Else
                If (cmbmode.Text = "Yearly") Then
                    rs.Open "select tenlac_yr from jeevan_anand where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", cn, adOpenStatic, adLockOptimistic
                Else
                    rs.Open "select tenlac_hr from jeevan_anand where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", cn, adOpenStatic, adLockOptimistic
                End If
            End If
        ElseIf (cmbpid.Text = "102") Then
            If (cmbamt.Text = "50000") Then
                If (cmbmode.Text = "Yearly") Then
                    rs.Open "select fifty_yr from jeevan_kishor where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", cn, adOpenStatic, adLockOptimistic
                Else
                    rs.Open "select fifty_hr from jeevan_kishor where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", cn, adOpenStatic, adLockOptimistic
                End If
            ElseIf (cmbamt.Text = "100000") Then
                If (cmbmode.Text = "Yearly") Then
                    rs.Open "select onelac_yr from jeevan_kishor where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", cn, adOpenStatic, adLockOptimistic
                Else
                    rs.Open "select onelac_hr from jeevan_kishor where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", cn, adOpenStatic, adLockOptimistic
                End If
            Else
                If (cmbmode.Text = "Yearly") Then
                    rs.Open "select twolac_yr from jeevan_kishor where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", cn, adOpenStatic, adLockOptimistic
                Else
                    rs.Open "select twolac_hr from jeevan_kishor where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", cn, adOpenStatic, adLockOptimistic
                End If
            End If
        ElseIf (cmbpid.Text = "89") Then
            If (cmbamt.Text = "50000") Then
                If (cmbmode.Text = "Yearly") Then
                    rs.Open "select fifty_yr from jeevan_saral where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", cn, adOpenStatic, adLockOptimistic
                Else
                    rs.Open "select fifty_hr from jeevan_saral where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", cn, adOpenStatic, adLockOptimistic
                End If
            ElseIf (cmbamt.Text = "100000") Then
                If (cmbmode.Text = "Yearly") Then
                    rs.Open "select onelac_yr from jeevan_saral where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", cn, adOpenStatic, adLockOptimistic
                Else
                    rs.Open "select onelac_hr from jeevan_saral where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", cn, adOpenStatic, adLockOptimistic
                End If
            Else
                If (cmbmode.Text = "Yearly") Then
                    rs.Open "select twolac_yr from jeevan_saral where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", cn, adOpenStatic, adLockOptimistic
                Else
                    rs.Open "select twolac_hr from jeevan_saral where pol_id=" + cmbpid.Text + " and client_age=" + cmbcage.Text + " and pol_duration ='" + cmbdur.Text + "'", cn, adOpenStatic, adLockOptimistic
                End If
            End If
        End If
        txtpamt.Text = rs.Fields(0)
        fine = txtpamt * 9 / 100
        txtdamt.Text = txtpamt.Text + fine
        rs.Close
        txtamt.Text = "0"
        txtstatus.Text = "Unpaid"
       
        dpkddate.Value = DateAdd("m", 1, dtprodate.Value)
        cn.Execute "insert into premium_info values(" + txtcid.Text + "," + txtaid.Text + "," + cmbpid.Text + "," + cmbamt.Text + "," + txtamt.Text + "," + txtpamt.Text + ",'" + CStr(dtprodate.Value) + "','" + CStr(dpkddate.Value) + "'," + txtdamt.Text + ",'" + txtstatus.Text + "')"
        Unload Me
        AGENT.Show
    End If
End Sub


Public Function MAX1()
 Dim clientid As Integer
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "select max(client_id) from client_info", cn, adOpenStatic, adLockOptimistic
    If IsNull(rs.Fields(0)) Then
        clientid = 0
        txtcid.Text = clientid
    Else
        clientid = rs.Fields(0)
        txtcid.Text = clientid
    End If
    rs.Close
    Set rs = New ADODB.Recordset
    rs.Open "select max(pol_no) from policy_info", cn, adOpenStatic, adLockOptimistic
    If IsNull(rs.Fields(0)) Then
        clientid = 1001
        txtpno.Text = clientid
    Else
        clientid = rs.Fields(0) + 1
        txtpno.Text = clientid
    End If
    rs.Close
End Function

Private Sub oduty_KeyPress(KeyAscii As Integer)
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



Private Sub oincome_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8) Then
        KeyAscii = 0
        MsgBox "Please Enter Integer Only!!!!!", vbExclamation
    End If
End Sub

Private Sub olen_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8) Then
        KeyAscii = 0
        MsgBox "Please Enter Integer Only!!!!!", vbExclamation
    End If
End Sub

Private Sub oname_KeyPress(KeyAscii As Integer)
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

Private Sub oocc_KeyPress(KeyAscii As Integer)
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


Private Sub oqual_KeyPress(KeyAscii As Integer)
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

Private Sub osource_KeyPress(KeyAscii As Integer)
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

