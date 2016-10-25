VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form DETAILS1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "DETAILS"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   FillColor       =   &H00FFC0C0&
   ForeColor       =   &H00FFC0C0&
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Height          =   5295
      Left            =   360
      Picture         =   "DETAILS.frx":0000
      ScaleHeight     =   5235
      ScaleWidth      =   3795
      TabIndex        =   8
      Top             =   2880
      Width           =   3855
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   360
      Picture         =   "DETAILS.frx":7BF8
      ScaleHeight     =   795
      ScaleWidth      =   1515
      TabIndex        =   7
      Top             =   480
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "DETAILS"
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
      Height          =   6735
      Left            =   4440
      TabIndex        =   0
      Top             =   2280
      Width           =   14415
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4575
         Left            =   360
         TabIndex        =   5
         Top             =   1800
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   8070
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
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
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
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
         Height          =   495
         Left            =   7080
         TabIndex        =   1
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   2760
         TabIndex        =   3
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "DETAILS OF"
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
         Left            =   1560
         TabIndex        =   4
         Top             =   960
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
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
      Height          =   495
      Left            =   8160
      TabIndex        =   2
      Top             =   9480
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "DETAIL INFORMATION"
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
      Left            =   7920
      TabIndex        =   6
      Top             =   480
      Width           =   5535
   End
End
Attribute VB_Name = "DETAILS1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
HOME.Show
End Sub

Private Sub Command3_Click()
Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    If (Text1.Text = "CLIENT") Then
        rs.Open "select * from client_info order by client_id", cn, adOpenStatic, adLockOptimistic
    ElseIf (Text1.Text = "NOMINEE") Then
        rs.Open "select * from nominee_info order by nominee_id", cn, adOpenStatic, adLockOptimistic
    ElseIf (Text1.Text = "AGENT") Then
        rs.Open "select * from agent_info order by agent_id", cn, adOpenStatic, adLockOptimistic
    ElseIf (Text1.Text = "POLICY") Then
        rs.Open "select * from policy_info order by pol_no", cn, adOpenStatic, adLockOptimistic
    ElseIf (Text1.Text = "PREMIUM") Then
        rs.Open "select * from premium_info order by client_id", cn, adOpenStatic, adLockOptimistic

    End If
    
    If (rs.RecordCount = 0) Then
        MsgBox "Details Not Present!!!!!", vbExclamation
    Else
            Set DataGrid1.DataSource = rs
    End If

End Sub

Private Sub Form_Load()
Call connect
End Sub
