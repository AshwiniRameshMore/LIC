VERSION 5.00
Begin VB.Form HOME2 
   BackColor       =   &H00FF8080&
   Caption         =   "HOME"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF0000&
      Caption         =   "PROCEED>>"
      DisabledPicture =   "HOME.frx":0000
      DownPicture     =   "HOME.frx":C60D
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
      Left            =   12240
      MaskColor       =   &H0000FFFF&
      Picture         =   "HOME.frx":18C1A
      TabIndex        =   1
      Top             =   9240
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Height          =   6135
      Left            =   1800
      Picture         =   "HOME.frx":25227
      ScaleHeight     =   6075
      ScaleWidth      =   15555
      TabIndex        =   0
      Top             =   1560
      Width           =   15615
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      Caption         =   "PALLAVI LAMBATE"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   4560
      TabIndex        =   5
      Top             =   9600
      Width           =   4095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Caption         =   "ASHWINI R. MORE"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   4560
      TabIndex        =   4
      Top             =   8880
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "Project By.................."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   8160
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "WELCOME TO LIFE INSURANCE CORPORATION OF INDIA"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   4800
      TabIndex        =   2
      Top             =   360
      Width           =   14655
   End
End
Attribute VB_Name = "HOME2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MDIForm1.Show

End Sub

