VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form SEARCH1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "SEARCH"
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
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   720
      MaskColor       =   &H00FFC0C0&
      TabIndex        =   3
      Top             =   8880
      UseMaskColor    =   -1  'True
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "SEARCH"
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
      Height          =   6495
      Left            =   6480
      TabIndex        =   6
      Top             =   3840
      Width           =   12375
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4335
         Left            =   360
         TabIndex        =   10
         Top             =   1800
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   7646
         _Version        =   393216
         BackColor       =   16777215
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
      Begin VB.TextBox Txtserkey 
         Height          =   495
         Left            =   2400
         TabIndex        =   1
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtser 
         Height          =   495
         Left            =   2400
         TabIndex        =   9
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H000000FF&
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
         Left            =   6840
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ENTER ID TO SEARCH:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         TabIndex        =   8
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "SEARCH:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   3735
      Left            =   360
      Picture         =   "SEARCH.frx":0000
      ScaleHeight     =   3675
      ScaleWidth      =   5715
      TabIndex        =   5
      Top             =   4320
      Width           =   5775
   End
   Begin VB.PictureBox Picture2 
      Height          =   3495
      Left            =   7200
      Picture         =   "SEARCH.frx":D4FA
      ScaleHeight     =   3435
      ScaleWidth      =   10995
      TabIndex        =   4
      Top             =   120
      Width           =   11055
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   480
      Picture         =   "SEARCH.frx":17266
      ScaleHeight     =   795
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "SEARCH1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    If (Txtserkey.Text = "") Then
        MsgBox "Please Enter Search Key!!!!!", vbExclamation
    Else
        Set rs = New ADODB.Recordset
        If (txtser.Text = "CLIENT") Then
                rs.Open "select * from client_info where client_id=" + Txtserkey.Text + "", cn, adOpenStatic, adLockOptimistic

        ElseIf (txtser.Text = "NOMINEE") Then
                 rs.Open "select * from nominee_info where nominee_id=" + Txtserkey.Text + "", cn, adOpenStatic, adLockOptimistic
        
        Else
                rs.Open "select * from agent_info where agent_id=" + Txtserkey.Text + "", cn, adOpenStatic, adLockOptimistic
        End If
   End If
        
        If (rs.RecordCount = 0) Then
            MsgBox "Record Not Present!!!!!", vbExclamation
        Else
            Set DataGrid1.DataSource = rs
        End If

End Sub

Private Sub Command2_Click()
HOME.Show
End Sub

Private Sub Form_Load()
Call connect

End Sub


