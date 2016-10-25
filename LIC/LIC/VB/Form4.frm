VERSION 5.00
Begin VB.Form frmplandetails 
   Caption         =   "Plan Details"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   11520
   ScaleWidth      =   19080
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdhome 
      Caption         =   "<< &Back To Home"
      Height          =   495
      Left            =   840
      TabIndex        =   7
      Top             =   10800
      Width           =   2295
   End
   Begin VB.Frame frmpdetails 
      Caption         =   "Plan Details"
      Height          =   7815
      Left            =   840
      TabIndex        =   0
      Top             =   2760
      Width           =   17775
      Begin VB.CommandButton btnok 
         Caption         =   "&OK"
         Height          =   375
         Left            =   2280
         TabIndex        =   6
         Top             =   1920
         Width           =   1575
      End
      Begin VB.ComboBox cmbda 
         Height          =   315
         ItemData        =   "Form4.frx":0000
         Left            =   3960
         List            =   "Form4.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1200
         Width           =   1695
      End
      Begin VB.ComboBox cmbpt 
         Height          =   315
         ItemData        =   "Form4.frx":0022
         Left            =   3960
         List            =   "Form4.frx":0035
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label iacmb 
         Caption         =   "Details  About :"
         Height          =   255
         Left            =   960
         TabIndex        =   4
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Plan  Type :"
         Height          =   255
         Left            =   960
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
      Begin VB.Image img 
         Height          =   1815
         Left            =   120
         Top             =   2760
         Width           =   2775
      End
      Begin VB.OLE oleblank 
         Appearance      =   0  'Flat
         Class           =   "Word.Document.12"
         Enabled         =   0   'False
         Height          =   7455
         Left            =   8160
         SourceDoc       =   "D:\Projects\VB Projects\LIC\Net Pages\Jeevan Kishor\Features.docx"
         TabIndex        =   17
         Top             =   240
         Width           =   9375
      End
      Begin VB.OLE oleajb 
         Appearance      =   0  'Flat
         Class           =   "Word.Document.12"
         Enabled         =   0   'False
         Height          =   7455
         Left            =   8160
         OleObjectBlob   =   "Form4.frx":0083
         SourceDoc       =   "D:\Projects\VB Projects\LIC\Net Pages\Anmol Jeevan\Benefits.docx"
         TabIndex        =   16
         Top             =   240
         Width           =   9375
      End
      Begin VB.OLE olejsb 
         Appearance      =   0  'Flat
         Class           =   "Word.Document.12"
         Enabled         =   0   'False
         Height          =   7335
         Left            =   8160
         OleObjectBlob   =   "Form4.frx":849B
         SourceDoc       =   "D:\Projects\VB Projects\LIC\Net Pages\Jeevan Saathi\Benefits.docx"
         TabIndex        =   15
         Top             =   240
         Width           =   9375
      End
      Begin VB.OLE olejsf 
         Appearance      =   0  'Flat
         Class           =   "Word.Document.12"
         Enabled         =   0   'False
         Height          =   7455
         Left            =   8160
         OleObjectBlob   =   "Form4.frx":110B3
         SourceDoc       =   "D:\Projects\VB Projects\LIC\Net Pages\Jeevan Saathi\Features.docx"
         TabIndex        =   14
         Top             =   240
         Width           =   9375
      End
      Begin VB.OLE olejbb 
         Appearance      =   0  'Flat
         Class           =   "Word.Document.12"
         Enabled         =   0   'False
         Height          =   7455
         Left            =   8160
         OleObjectBlob   =   "Form4.frx":188CB
         SourceDoc       =   "D:\Projects\VB Projects\LIC\Net Pages\Jeevan Bharati\Benefits.docx"
         TabIndex        =   13
         Top             =   240
         Width           =   9375
      End
      Begin VB.OLE olejbf 
         Appearance      =   0  'Flat
         Class           =   "Word.Document.12"
         Enabled         =   0   'False
         Height          =   7455
         Left            =   8160
         OleObjectBlob   =   "Form4.frx":200E3
         SourceDoc       =   "D:\Projects\VB Projects\LIC\Net Pages\Jeevan Bharati\Features.docx"
         TabIndex        =   12
         Top             =   240
         Width           =   9375
      End
      Begin VB.OLE olejkb 
         Appearance      =   0  'Flat
         Class           =   "Word.Document.12"
         Enabled         =   0   'False
         Height          =   7455
         Left            =   8160
         OleObjectBlob   =   "Form4.frx":294FB
         SourceDoc       =   "D:\Projects\VB Projects\LIC\Net Pages\Jeevan Kishor\Benefits.docx"
         TabIndex        =   11
         Top             =   240
         Width           =   9375
      End
      Begin VB.OLE olejkf 
         Appearance      =   0  'Flat
         Class           =   "Word.Document.12"
         Enabled         =   0   'False
         Height          =   7455
         Left            =   8160
         OleObjectBlob   =   "Form4.frx":30D13
         SourceDoc       =   "D:\Projects\VB Projects\LIC\Net Pages\Jeevan Kishor\Features.docx"
         TabIndex        =   10
         Top             =   240
         Width           =   9375
      End
      Begin VB.OLE olejab 
         Appearance      =   0  'Flat
         Class           =   "Word.Document.12"
         Enabled         =   0   'False
         Height          =   7455
         Left            =   8160
         OleObjectBlob   =   "Form4.frx":3872B
         SourceDoc       =   "D:\Projects\VB Projects\LIC\Net Pages\Jeevan Anand\Benefits.docx"
         TabIndex        =   9
         Top             =   240
         Width           =   9375
      End
      Begin VB.OLE olejaf 
         Appearance      =   0  'Flat
         Class           =   "Word.Document.12"
         Enabled         =   0   'False
         Height          =   7455
         Left            =   8160
         OleObjectBlob   =   "Form4.frx":3FF43
         SourceDoc       =   "D:\Projects\VB Projects\LIC\Net Pages\Jeevan Anand\Features.docx"
         TabIndex        =   8
         Top             =   240
         Width           =   9375
      End
   End
   Begin VB.Label Label18 
      Caption         =   "PLAN  DETAILS"
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
      Left            =   9000
      TabIndex        =   5
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Image Image4 
      Height          =   765
      Left            =   360
      Picture         =   "Form4.frx":4775B
      Top             =   240
      Width           =   1470
   End
   Begin VB.Image Image2 
      Height          =   1920
      Left            =   5760
      Picture         =   "Form4.frx":4DBAE
      Top             =   120
      Width           =   8475
   End
End
Attribute VB_Name = "frmplandetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnok_Click()
    If (cmbpt.Text = "") Then
        MsgBox "Please Select Policy Type!!!!!", vbExclamation
    ElseIf (cmbda.Text = "") Then
        MsgBox "Please Select Details About!!!!!", vbExclamation
    Else
        If (cmbpt = "Jeevan Anand") Then
            img.Picture = LoadPicture("d:\projects\VB Projects\LIC\Images\29.jpg")
            If (cmbda = "Features") Then
                oleblank.Visible = False
                olejaf.Visible = True
                olejab.Visible = False
                olejkf.Visible = False
                olejkb.Visible = False
                olejbf.Visible = False
                olejbb.Visible = False
                olejsf.Visible = False
                olejsb.Visible = False
                oleajb.Visible = False
            Else
                oleblank.Visible = False
                olejaf.Visible = False
                olejab.Visible = True
                olejkf.Visible = False
                olejkb.Visible = False
                olejbf.Visible = False
                olejbb.Visible = False
                olejsf.Visible = False
                olejsb.Visible = False
                oleajb.Visible = False
            End If
        ElseIf (cmbpt = "Jeevan Kishor") Then
            img.Picture = LoadPicture("d:\projects\VB Projects\LIC\Images\5.jpg")
            If (cmbda = "Features") Then
                oleblank.Visible = False
                olejaf.Visible = False
                olejab.Visible = False
                olejkf.Visible = True
                olejkb.Visible = False
                olejbf.Visible = False
                olejbb.Visible = False
                olejsf.Visible = False
                olejsb.Visible = False
                oleajb.Visible = False
            Else
                oleblank.Visible = False
                olejaf.Visible = False
                olejab.Visible = False
                olejkf.Visible = False
                olejkb.Visible = True
                olejbf.Visible = False
                olejbb.Visible = False
                olejsf.Visible = False
                olejsb.Visible = False
                oleajb.Visible = False
                If (cmbda = "Benefits") Then
                Else
                End If
            End If
        ElseIf (cmbpt = "Jeevan Bharati") Then
            img.Picture = LoadPicture("d:\projects\VB Projects\LIC\Images\30.jpg")
            If (cmbda = "Features") Then
                oleblank.Visible = False
                olejaf.Visible = False
                olejab.Visible = False
                olejkf.Visible = False
                olejkb.Visible = False
                olejbf.Visible = True
                olejbb.Visible = False
                olejsf.Visible = False
                olejsb.Visible = False
                oleajb.Visible = False
            Else
                oleblank.Visible = False
                olejaf.Visible = False
                olejab.Visible = False
                olejkf.Visible = False
                olejkb.Visible = False
                olejbf.Visible = False
                olejbb.Visible = True
                olejsf.Visible = False
                olejsb.Visible = False
                oleajb.Visible = False
                If (cmbda = "Benefits") Then
                Else
                End If
            End If
        ElseIf (cmbpt = "Jeevan Saathi") Then
            img.Picture = LoadPicture("d:\projects\VB Projects\LIC\Images\10.jpg")
            If (cmbda = "Features") Then
                oleblank.Visible = False
                olejaf.Visible = False
                olejab.Visible = False
                olejkf.Visible = False
                olejkb.Visible = False
                olejbf.Visible = False
                olejbb.Visible = False
                olejsf.Visible = True
                olejsb.Visible = False
                oleajb.Visible = False
            Else
                oleblank.Visible = False
                olejaf.Visible = False
                olejab.Visible = False
                olejkf.Visible = False
                olejkb.Visible = False
                olejbf.Visible = False
                olejbb.Visible = False
                olejsf.Visible = False
                olejsb.Visible = True
                oleajb.Visible = False
                If (cmbda = "Benefits") Then
                Else
                End If
            End If
        ElseIf (cmbpt = "Anmol Jeevan") Then
            img.Picture = LoadPicture("d:\projects\VB Projects\LIC\Images\7.jpg")
            If (cmbda = "Benefits") Then
                oleblank.Visible = False
                olejaf.Visible = False
                olejab.Visible = False
                olejkf.Visible = False
                olejkb.Visible = False
                olejbf.Visible = False
                olejbb.Visible = False
                olejsf.Visible = False
                olejsb.Visible = False
                oleajb.Visible = True
            End If
        End If
    End If
End Sub

Private Sub cmbpt_Click()
    If (cmbpt = "Anmol Jeevan") Then
        cmbda.Clear
        cmbda.AddItem "Benefits"
    End If
End Sub

Private Sub cmdhome_Click()
    Unload Me
    frmhome.Show
End Sub


Private Sub Form_Load()
    oleblank.Visible = True
    olejaf.Visible = False
    olejab.Visible = False
    olejkf.Visible = False
    olejkb.Visible = False
    olejbf.Visible = False
    olejbb.Visible = False
    olejsf.Visible = False
    olejsb.Visible = False
    oleajb.Visible = False
End Sub
