VERSION 5.00
Begin VB.Form THANKS 
   BackColor       =   &H00004040&
   Caption         =   "THANKS"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   13800
      TabIndex        =   1
      Top             =   10080
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Height          =   8535
      Left            =   3480
      Picture         =   "THANKS.frx":0000
      ScaleHeight     =   8475
      ScaleWidth      =   12795
      TabIndex        =   0
      Top             =   960
      Width           =   12855
   End
End
Attribute VB_Name = "THANKS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub


