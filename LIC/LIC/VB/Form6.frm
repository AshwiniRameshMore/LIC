VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmdetails 
   Caption         =   "Details"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4680
   LinkTopic       =   "Form6"
   MDIChild        =   -1  'True
   ScaleHeight     =   11520
   ScaleWidth      =   19080
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdhome 
      Caption         =   "<< &Back To Home"
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   9960
      Width           =   2295
   End
   Begin VB.Frame frmdetails 
      Caption         =   "Details:"
      ForeColor       =   &H000000FF&
      Height          =   8055
      Left            =   5400
      TabIndex        =   0
      Top             =   1320
      Width           =   12855
      Begin VB.TextBox txtdetailsof 
         Height          =   285
         Left            =   3120
         TabIndex        =   6
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "&OK"
         Height          =   495
         Left            =   6000
         TabIndex        =   5
         Top             =   600
         Width           =   1695
      End
      Begin MSDataGridLib.DataGrid dgdetails 
         Height          =   6015
         Left            =   1200
         TabIndex        =   3
         Top             =   1560
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   10610
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
      Begin VB.Label Label3 
         Caption         =   "Details Of :"
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Label Label18 
      Caption         =   "DETAILS"
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
      Left            =   11520
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   5760
      Left            =   120
      Picture         =   "Form6.frx":0000
      Top             =   2400
      Width           =   5115
   End
   Begin VB.Image Image4 
      Height          =   765
      Left            =   360
      Picture         =   "Form6.frx":68E2
      Top             =   240
      Width           =   1470
   End
End
Attribute VB_Name = "frmdetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdhome_Click()
    Unload Me
    frmhome.Show
End Sub

Private Sub cmdok_Click()
    Set rs = New ADODB.Recordset
    If (txtdetailsof.Text = "Client") Then
        rs.Open "select * from client_info", con, adOpenStatic, adLockOptimistic
    ElseIf (txtdetailsof.Text = "Nominee") Then
        rs.Open "select * from nominee_info", con, adOpenStatic, adLockOptimistic
    ElseIf (txtdetailsof.Text = "Agent") Then
        rs.Open "select * from agent_info", con, adOpenStatic, adLockOptimistic
    ElseIf (txtdetailsof.Text = "Policy") Then
        rs.Open "select * from policy_info", con, adOpenStatic, adLockOptimistic
    ElseIf (txtdetailsof.Text = "Premium") Then
        rs.Open "select * from premium_info", con, adOpenStatic, adLockOptimistic
    Else
        rs.Open "select * from claim_info", con, adOpenStatic, adLockOptimistic
    End If
    If (rs.RecordCount = 0) Then
        MsgBox "Details Not Present!!!!!", vbExclamation
    Else
        Set dgdetails.DataSource = rs
    End If
End Sub

Private Sub Form_Load()
    Call conn
End Sub
