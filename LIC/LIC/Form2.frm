VERSION 5.00
Begin VB.Form HOME 
   BackColor       =   &H00FFFF00&
   Caption         =   "HOME"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   2.45972e5
   ScaleMode       =   0  'User
   ScaleWidth      =   3.78347e5
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "PROCEED>>"
      Height          =   615
      Left            =   13080
      TabIndex        =   4
      Top             =   10200
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF00&
      Caption         =   "PALLAVI LAMBATE"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10320
      TabIndex        =   3
      Top             =   9360
      Width           =   5535
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      Caption         =   "ASHWINI R.MORE"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10320
      TabIndex        =   2
      Top             =   8640
      Width           =   5295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Project by-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   1
      Top             =   7800
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      Caption         =   "WELCOME TO LIFE INSURANCE CORPORATION OF INDIA!!!!!"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   1920
      TabIndex        =   0
      Top             =   6360
      Width           =   11175
   End
End
Attribute VB_Name = "HOME"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
Load Image("c:\images\47.jpg")
End Sub

Private Sub Command1_Click()
MDIForm1.Show

End Sub
