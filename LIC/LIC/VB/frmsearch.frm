VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmsearch 
   Caption         =   "Search"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11520
   ScaleWidth      =   19080
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdhome 
      Caption         =   "<< &Back To Home"
      Height          =   495
      Left            =   5520
      TabIndex        =   6
      Top             =   9960
      Width           =   2055
   End
   Begin VB.Frame frmsearch 
      Caption         =   "Search"
      ForeColor       =   &H000000FF&
      Height          =   6375
      Left            =   5520
      TabIndex        =   0
      Top             =   3120
      Width           =   12855
      Begin VB.TextBox txtsearchof 
         Height          =   285
         Left            =   3960
         TabIndex        =   9
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtsearchby 
         Height          =   285
         Left            =   8040
         TabIndex        =   8
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "&OK"
         Height          =   495
         Left            =   8160
         TabIndex        =   7
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtskey 
         Height          =   285
         Left            =   3960
         TabIndex        =   1
         Top             =   1320
         Width           =   1815
      End
      Begin MSDataGridLib.DataGrid dgsearch 
         Height          =   3855
         Left            =   1200
         TabIndex        =   2
         Top             =   2040
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   6800
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Search By :"
         Height          =   375
         Left            =   6240
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Enter The Key For Search :"
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Search Of :"
         Height          =   255
         Left            =   1080
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Image Image2 
      Height          =   4800
      Left            =   120
      Picture         =   "frmsearch.frx":0000
      Top             =   3840
      Width           =   5250
   End
   Begin VB.Image Image1 
      Height          =   2760
      Left            =   6120
      Picture         =   "frmsearch.frx":4347
      Top             =   120
      Width           =   11460
   End
   Begin VB.Image Image9 
      Height          =   765
      Left            =   360
      Picture         =   "frmsearch.frx":92A6
      Top             =   240
      Width           =   1470
   End
End
Attribute VB_Name = "frmsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdhome_Click()
    Unload Me
    frmhome.Show
End Sub

Private Sub cmdok_Click()
    If (txtskey.Text = "") Then
        MsgBox "Please Enter Search Key!!!!!", vbExclamation
    Else
        Set rs = New ADODB.Recordset
        If (txtsearchof.Text = "Client") Then
            If (txtsearchby.Text = "Client ID") Then
                rs.Open "select * from client_info where client_id=" + txtskey.Text + "", con, adOpenStatic, adLockOptimistic
            ElseIf (txtsearchby.Text = "Client Name") Then
                rs.Open "select * from client_info where client_fname='" + txtskey.Text + "' or client_mname='" + txtskey.Text + "' or client_lname='" + txtskey.Text + "'", con, adOpenStatic, adLockOptimistic
            ElseIf (txtsearchby.Text = "Agent ID") Then
                rs.Open "select * from client_info where agent_id=" + txtskey.Text + "", con, adOpenStatic, adLockOptimistic
            Else
                rs.Open "select * from client_info where pol_id=" + txtskey.Text + "", con, adOpenStatic, adLockOptimistic
            End If
        ElseIf (txtsearchof.Text = "Nominee") Then
            If (txtsearchbyy.Text = "Nominee ID") Then
                rs.Open "select * from nominee_info where nominee_id=" + txtskey.Text + "", con, adOpenStatic, adLockOptimistic
            ElseIf (txtsearchby.Text = "Nominee Name") Then
                rs.Open "select * from nominee_info where nom_fname='" + txtskey.Text + "' or nom_mname='" + txtskey.Text + "' or nom_lname='" + txtskey.Text + "'", con, adOpenStatic, adLockOptimistic
            Else
                rs.Open "select * from nominee_info where agent_id=" + txtskey.Text + "", con, adOpenStatic, adLockOptimistic
            End If
        Else
            If (txtsearchby.Text = "Agent ID") Then
                rs.Open "select * from agent_info where agent_id=" + txtskey.Text + "", con, adOpenStatic, adLockOptimistic
            Else
                rs.Open "select * from nominee_info where agent_fname='" + txtskey.Text + "' or agent_mname='" + txtskey.Text + "' or agent_lname='" + txtskey.Text + "'", con, adOpenStatic, adLockOptimistic
            End If
        End If
        If (rs.RecordCount = 0) Then
            MsgBox "Record Not Present!!!!!", vbExclamation
        Else
            MsgBox "Record Found.....", vbInformation
            Set dgsearch.DataSource = rs
        End If
    End If
End Sub

Private Sub Form_Load()
    Call conn
End Sub

Private Sub txtskey_Change()

End Sub

Private Sub txtskey_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        cmdok.SetFocus
    End If
End Sub
