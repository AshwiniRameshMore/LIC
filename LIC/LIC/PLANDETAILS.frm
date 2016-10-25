VERSION 5.00
Begin VB.Form PLANDETAILS 
   BackColor       =   &H00FFC0C0&
   Caption         =   "PLANDETAILS"
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
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   480
      Picture         =   "PLANDETAILS.frx":0000
      ScaleHeight     =   795
      ScaleWidth      =   1515
      TabIndex        =   14
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
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
      Height          =   735
      Left            =   1080
      TabIndex        =   4
      Top             =   9840
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
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
      Left            =   2400
      TabIndex        =   3
      Top             =   5040
      Width           =   1935
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2280
      TabIndex        =   2
      Top             =   4200
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2280
      TabIndex        =   1
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "PLAN DETAILS"
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
      Height          =   735
      Left            =   8400
      TabIndex        =   13
      Top             =   720
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   4815
      Left            =   4680
      Top             =   3120
      Width           =   5415
   End
   Begin VB.OLE OLEBLANK 
      BackColor       =   &H00FFC0C0&
      Class           =   "Word.Document.12"
      Height          =   7935
      Left            =   11520
      OleObjectBlob   =   "PLANDETAILS.frx":6453
      SourceDoc       =   "H:\Doc1.docx"
      TabIndex        =   12
      Top             =   2640
      Width           =   6495
   End
   Begin VB.OLE OLESB 
      Class           =   "Word.Document.12"
      Height          =   7935
      Left            =   11520
      OleObjectBlob   =   "PLANDETAILS.frx":9C6B
      SourceDoc       =   "H:\lic.doc\pol\Jeevan saral benefits.docx"
      TabIndex        =   11
      Top             =   2640
      Width           =   6495
   End
   Begin VB.OLE OLESF 
      Class           =   "Word.Document.12"
      Height          =   7935
      Left            =   11520
      OleObjectBlob   =   "PLANDETAILS.frx":10C83
      SourceDoc       =   "H:\lic.doc\pol\Jeevan saral features.docx"
      TabIndex        =   10
      Top             =   2640
      Width           =   6495
   End
   Begin VB.OLE OLEKB 
      Class           =   "Word.Document.12"
      Height          =   7935
      Left            =   11520
      OleObjectBlob   =   "PLANDETAILS.frx":1949B
      SourceDoc       =   "H:\lic.doc\pol\Jeevan kishore benefits.docx"
      TabIndex        =   9
      Top             =   2640
      Width           =   6495
   End
   Begin VB.OLE OLEKF 
      Class           =   "Word.Document.12"
      Height          =   7935
      Left            =   11520
      OleObjectBlob   =   "PLANDETAILS.frx":206B3
      SourceDoc       =   "H:\lic.doc\pol\Jeevan kishore features.docx"
      TabIndex        =   8
      Top             =   2640
      Width           =   6495
   End
   Begin VB.OLE OLEAB 
      Class           =   "Word.Document.12"
      Height          =   7935
      Left            =   11520
      OleObjectBlob   =   "PLANDETAILS.frx":276CB
      SourceDoc       =   "H:\lic.doc\pol\Jeevan anand benefits.docx"
      TabIndex        =   7
      Top             =   2640
      Width           =   6495
   End
   Begin VB.OLE OLEAF 
      Class           =   "Word.Document.12"
      Height          =   7935
      Left            =   11520
      OleObjectBlob   =   "PLANDETAILS.frx":2EAE3
      SourceDoc       =   "H:\lic.doc\pol\Jeevan anand features.docx"
      TabIndex        =   6
      Top             =   2640
      Width           =   6495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "DETAILS ABOUT"
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
      Left            =   600
      TabIndex        =   5
      Top             =   4200
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "POLICY TYPE"
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
      Left            =   600
      TabIndex        =   0
      Top             =   3120
      Width           =   3855
   End
End
Attribute VB_Name = "PLANDETAILS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If (Combo1.Text = "") Then
        MsgBox "Please Select Policy Type!!!!!", vbExclamation
    ElseIf (Combo2.Text = "") Then
        MsgBox "Please Select Details About!!!!!", vbExclamation
    Else
If Combo1.Text = "JEEVAN-ANAND" Then
Image1.Picture = LoadPicture("c:\Images\8.jpg")
   If Combo2.Text = "FEATURES" Then
    OLEAF.Visible = True
    OLEAB.Visible = False
    OLEKF.Visible = False
    OLEKB.Visible = False
    OLESF.Visible = False
    OLESB.Visible = False
    OLEBLANK.Visible = False
   Else
    OLEAF.Visible = False
    OLEAB.Visible = True
    OLEKF.Visible = False
    OLEKB.Visible = False
    OLESF.Visible = False
    OLESB.Visible = False
    OLEBLANK.Visible = False
   End If
ElseIf Combo1.Text = "JEEVAN-KISHORE" Then
Image1.Picture = LoadPicture("c:\Images\5.jpg")
     If Combo2.Text = "FEATURES" Then
    OLEAF.Visible = False
    OLEAB.Visible = False
    OLEKF.Visible = True
    OLEKB.Visible = False
    OLESF.Visible = False
    OLESB.Visible = False
    OLEBLANK.Visible = False
   Else
    OLEAF.Visible = False
    OLEAB.Visible = False
    OLEKF.Visible = False
    OLEKB.Visible = True
    OLESF.Visible = False
    OLESB.Visible = False
    OLEBLANK.Visible = False
   End If
ElseIf Combo1.Text = "JEEVAN-SARAL" Then
Image1.Picture = LoadPicture("c:\Images\10.jpg")
     If Combo2.Text = "FEATURES" Then
    OLEAF.Visible = False
    OLEAB.Visible = False
    OLEKF.Visible = False
    OLEKB.Visible = False
    OLESF.Visible = True
    OLESB.Visible = False
    OLEBLANK.Visible = False
   Else
    OLEAF.Visible = False
    OLEAB.Visible = False
    OLEKF.Visible = False
    OLEKB.Visible = False
    OLESF.Visible = False
    OLESB.Visible = True
    OLEBLANK.Visible = False
   End If

Else
OLEBLANK.Visible = True
End If
End If
End Sub

Private Sub Command2_Click()
HOME.Show

End Sub

Private Sub Form_Load()
Combo1.AddItem "JEEVAN-ANAND"
Combo1.AddItem "JEEVAN-KISHORE"
Combo1.AddItem "JEEVAN-SARAL"
Combo2.AddItem "FEATURES"
Combo2.AddItem "BENEFITS"
End Sub

