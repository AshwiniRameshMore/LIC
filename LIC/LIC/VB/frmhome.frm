VERSION 5.00
Begin VB.Form frmhome 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Home"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4680
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   15
      Left            =   5400
      Top             =   360
   End
   Begin VB.Timer timer 
      Interval        =   100
      Left            =   4440
      Top             =   360
   End
   Begin VB.Image imglic 
      Height          =   1035
      Left            =   7320
      Picture         =   "frmhome.frx":0000
      Top             =   120
      Width           =   4800
   End
   Begin VB.Label lblwelcome 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Welcome To Life Insurance Corporation Of India!!!!!"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   4200
      TabIndex        =   4
      Top             =   1440
      Width           =   11055
   End
   Begin VB.Label lbldy 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Dhome Yogesh K."
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   14520
      TabIndex        =   3
      Top             =   9720
      Width           =   2895
   End
   Begin VB.Label lblgn 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Gugale Nitin N."
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   14520
      TabIndex        =   2
      Top             =   10200
      Width           =   2895
   End
   Begin VB.Label lblbt 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Bhapkar Tushar S."
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   14520
      TabIndex        =   1
      Top             =   9240
      Width           =   2895
   End
   Begin VB.Label lblpb 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Project By....."
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   12360
      TabIndex        =   0
      Top             =   8640
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   6120
      Left            =   1680
      Picture         =   "frmhome.frx":2787
      Top             =   2400
      Width           =   15570
   End
End
Attribute VB_Name = "frmhome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Boolean
Dim flag1 As Boolean
Private Sub Form_Load()
    flag = True
    flag1 = True
    lblwelcome.Left = 19200
    lblpb.Left = 19200
    lblbt.Left = 19200
    lbldy.Left = 19200
    lblgn.Left = 19200
End Sub

Private Sub timer_Timer()
    If (flag1 = False) Then
        If (lblwelcome.Left = -10800) Then
            lblwelcome.Left = 19080
        Else
            If (flag = True) Then
                lblwelcome.ForeColor = &HFF&
                flag = False
            Else
                lblwelcome.ForeColor = &HFF0000
                flag = True
            End If
            lblwelcome.Left = lblwelcome.Left - 120
        End If
    End If
End Sub

Private Sub Timer1_Timer()
    If (lblpb.Left = 12360) Then
        If (lblbt.Left = 14520) Then
            flag1 = False
        Else
            lblbt.Left = lblbt.Left - 120
            lbldy.Left = lbldy.Left - 120
            lblgn.Left = lblgn.Left - 120
        End If
    Else
        lblpb.Left = lblpb.Left - 120
    End If
End Sub
