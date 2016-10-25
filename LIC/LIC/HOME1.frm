VERSION 5.00
Begin VB.Form HOME 
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   19335
      Left            =   0
      Picture         =   "HOME1.frx":0000
      ScaleHeight     =   19275
      ScaleWidth      =   4620
      TabIndex        =   2
      Top             =   0
      Width           =   4680
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF0000&
         Caption         =   "EXIT"
         DisabledPicture =   "HOME1.frx":313B8
         DownPicture     =   "HOME1.frx":3D9C5
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
         Left            =   14160
         MaskColor       =   &H0000FFFF&
         Picture         =   "HOME1.frx":49FD2
         TabIndex        =   1
         Top             =   10080
         UseMaskColor    =   -1  'True
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF0000&
         Caption         =   "PROCEED>>"
         DisabledPicture =   "HOME1.frx":565DF
         DownPicture     =   "HOME1.frx":62BEC
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
         Left            =   11280
         MaskColor       =   &H0000FFFF&
         Picture         =   "HOME1.frx":6F1F9
         TabIndex        =   0
         Top             =   10080
         UseMaskColor    =   -1  'True
         Width           =   2415
      End
      Begin VB.PictureBox Picture2 
         Height          =   5415
         Left            =   12000
         Picture         =   "HOME1.frx":7B806
         ScaleHeight     =   5355
         ScaleWidth      =   6795
         TabIndex        =   5
         Top             =   3360
         Width           =   6855
      End
      Begin VB.PictureBox Picture3 
         Height          =   5055
         Left            =   5520
         Picture         =   "HOME1.frx":7D9D9
         ScaleHeight     =   4995
         ScaleWidth      =   6195
         TabIndex        =   4
         Top             =   1800
         Width           =   6255
      End
      Begin VB.PictureBox Picture4 
         Height          =   4575
         Left            =   120
         Picture         =   "HOME1.frx":863BB
         ScaleHeight     =   4515
         ScaleWidth      =   5235
         TabIndex        =   3
         Top             =   480
         Width           =   5295
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ASHWINI  R. MORE"
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
         Height          =   375
         Left            =   2160
         TabIndex        =   12
         Top             =   8640
         Width           =   3255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   5880
         TabIndex        =   11
         Top             =   360
         Width           =   14655
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
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
         Left            =   360
         TabIndex        =   10
         Top             =   7800
         Width           =   3135
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
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
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Top             =   9240
         Width           =   3255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "JEEVAN SARAL"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   8
         Top             =   5040
         Width           =   4215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "JEEVAN KISHORE"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6360
         TabIndex        =   7
         Top             =   6840
         Width           =   4695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "JEEVAN ANAND"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   13560
         TabIndex        =   6
         Top             =   8760
         Width           =   4215
      End
   End
End
Attribute VB_Name = "HOME"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MDIForm1.Show
End Sub

Private Sub Command2_Click()
End
End Sub
