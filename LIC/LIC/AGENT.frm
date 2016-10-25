VERSION 5.00
Begin VB.Form AGENT 
   BackColor       =   &H00FFC0C0&
   Caption         =   "AGENT"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   DrawStyle       =   5  'Transparent
   FillColor       =   &H00FFC0C0&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00FFC0C0&
   LinkTopic       =   "Form3"
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture6 
      Height          =   2415
      Left            =   1080
      Picture         =   "AGENT.frx":0000
      ScaleHeight     =   2355
      ScaleWidth      =   2355
      TabIndex        =   56
      Top             =   2640
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Height          =   2415
      Left            =   1080
      Picture         =   "AGENT.frx":1B47
      ScaleHeight     =   2355
      ScaleWidth      =   2355
      TabIndex        =   55
      Top             =   6720
      Width           =   2415
   End
   Begin VB.PictureBox Picture4 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      Picture         =   "AGENT.frx":475B
      ScaleHeight     =   795
      ScaleWidth      =   1515
      TabIndex        =   54
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton CMDBACK 
      Caption         =   "<BACK"
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
      Left            =   9720
      TabIndex        =   23
      Top             =   9960
      Width           =   1575
   End
   Begin VB.TextBox txtnid 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   53
      Text            =   "Text1"
      Top             =   7320
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtaid 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2520
      TabIndex        =   52
      Text            =   "Text1"
      Top             =   7320
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtcid 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2880
      TabIndex        =   51
      Text            =   "Text1"
      Top             =   7320
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "AGENT INFORMATION"
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
      Height          =   3375
      Left            =   4800
      TabIndex        =   26
      Top             =   6360
      Width           =   10815
      Begin VB.TextBox apin 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8040
         TabIndex        =   20
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox aemail 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   19
         Top             =   2760
         Width           =   4095
      End
      Begin VB.TextBox aoff 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8040
         TabIndex        =   18
         Top             =   2040
         Width           =   2175
      End
      Begin VB.TextBox amob 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   17
         Top             =   2040
         Width           =   2415
      End
      Begin VB.ComboBox awa 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8400
         TabIndex        =   16
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox aage 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   15
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox aadd 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   14
         Top             =   960
         Width           =   8295
      End
      Begin VB.TextBox alname 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7920
         TabIndex        =   13
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox amname 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   12
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox afname 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   11
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFC0C0&
         Caption         =   "*PINCODE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7080
         TabIndex        =   47
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFC0C0&
         Caption         =   "E-MAIL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   46
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFC0C0&
         Caption         =   "OFFICE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   45
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MOBILE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   44
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFC0C0&
         Caption         =   "TELEPHONE NO"
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
         Left            =   360
         TabIndex        =   43
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFC0C0&
         Caption         =   "*WORKING AREA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         TabIndex        =   42
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFC0C0&
         Caption         =   "*AGE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   41
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFC0C0&
         Caption         =   "*ADDRESS"
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
         Left            =   360
         TabIndex        =   40
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFC0C0&
         Caption         =   "LAST NAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8520
         TabIndex        =   39
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIDDLE NAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   38
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "FIRST NAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   37
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFC0C0&
         Caption         =   "*AGENT'S FULL NAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "NOMINEE INFORMATION"
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
      Height          =   3855
      Left            =   4800
      TabIndex        =   25
      Top             =   2400
      Width           =   10815
      Begin VB.TextBox npin 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7680
         TabIndex        =   10
         Top             =   2880
         Width           =   2535
      End
      Begin VB.TextBox nemail 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   9
         Top             =   3000
         Width           =   3975
      End
      Begin VB.TextBox noff 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7680
         TabIndex        =   8
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox nmob 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   7
         Top             =   2280
         Width           =   2655
      End
      Begin VB.TextBox nrel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7680
         TabIndex        =   6
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox nage 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox nadd 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Top             =   1080
         Width           =   7815
      End
      Begin VB.TextBox nlname 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7680
         TabIndex        =   3
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox nmname 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   2
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox nfname 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFC0C0&
         Caption         =   "LAST NAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8160
         TabIndex        =   50
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIDDLE NAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   49
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFC0C0&
         Caption         =   "FIRST NAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   48
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFC0C0&
         Caption         =   "*PINCODE"
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
         Left            =   6720
         TabIndex        =   35
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFC0C0&
         Caption         =   "OFFICE"
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
         Left            =   6720
         TabIndex        =   34
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MOBILE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   33
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "*RELATION WITH CLIENT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   32
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "E-MAIL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   31
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "TELEPHONE NO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "*AGE"
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
         Left            =   240
         TabIndex        =   29
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "*ADDRESS"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   28
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "*NOMINEE'S FULL NAME"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   3615
      End
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H8000000D&
      Caption         =   "SAVE"
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
      Left            =   12960
      TabIndex        =   21
      Top             =   9960
      Width           =   2655
   End
   Begin VB.CommandButton CMDHOME 
      BackColor       =   &H00FF0000&
      Caption         =   "    <<BACK TO HOME"
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
      Left            =   4800
      MaskColor       =   &H00FFC0C0&
      TabIndex        =   22
      Top             =   9960
      Width           =   2535
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   16320
      Picture         =   "AGENT.frx":ABAE
      ScaleHeight     =   4755
      ScaleWidth      =   3795
      TabIndex        =   24
      Top             =   3600
      Width           =   3855
   End
   Begin VB.PictureBox Picture2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   5880
      Picture         =   "AGENT.frx":EEF5
      ScaleHeight     =   1875
      ScaleWidth      =   8475
      TabIndex        =   0
      Top             =   120
      Width           =   8535
   End
End
Attribute VB_Name = "AGENT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDBACK_Click()
    Call max
    If (txtcid.Text = "0") Then
    Else
        cn.Execute ("delete from other_info where client_id = " + txtcid.Text + "")
        cn.Execute ("delete from premium_info where client_id = " + txtcid.Text + "")
    End If
    If (txtpid.Text = "0") Then
    Else
        cn.Execute ("delete from policy_info where pol_no = " + txtpid.Text + "")
    End If
    Unload Me
    CLIENT.Show

End Sub

Private Sub CMDHOME_Click()
    If (cmdsave.Enabled = True) Then
        If (txtcid.Text = "0") Then
        Else
            cn.Execute ("delete from client_info where client_id = " + txtcid.Text + "")
            cn.Execute ("delete from other_info where client_id = " + txtcid.Text + "")
            cn.Execute ("delete from premium_info where client_id=" + txtcid.Text + "")
        End If
        If (txtpid.Text = "0") Then
        Else
            cn.Execute ("delete from policy_info where pol_no = " + txtpid.Text + "")
        End If
    End If
    Unload Me
    MDIForm1.Show

End Sub


Private Sub cmdsave_Click()
If (nfname.Text = "" Or nmname.Text = "" Or nlname.Text = "" Or nadd.Text = "" Or nage.Text = "" Or nrel.Text = "" Or npin.Text = "" Or afname.Text = "" Or amname.Text = "" Or alname.Text = "" Or aadd.Text = "" Or aage.Text = "" Or awa.Text = "" Or apin.Text = "") Then
        MsgBox "All * Marked Fields Are Compulsory!!!!!", vbExclamation
    Else
        Call MAX1
        cn.Execute ("insert into nominee_info values(" + txtnid.Text + ",'" + nfname.Text + "','" + nmname.Text + "','" + nlname.Text + "','" + nadd.Text + "'," + nmob.Text + "," + noff.Text + ",'" + nemail.Text + "'," + npin.Text + "," + nage.Text + ",'" + nrel.Text + "'," + txtcid.Text + ")")
        cn.Execute ("insert into agent_info values(" + txtaid.Text + ",'" + afname.Text + "','" + amname.Text + "','" + alname.Text + "','" + aadd.Text + "'," + amob.Text + "," + aoff.Text + ",'" + aemail.Text + "'," + apin.Text + "," + aage.Text + ",'" + awa.Text + "')")
        MsgBox "Record Saved Successfully.....", vbInformation
        cmdsave.Enabled = False
        CMDBACK.Enabled = False
    End If

End Sub


Public Function max()
Dim clientid As Integer
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "select max(client_id) from client_info", cn, adOpenStatic, adLockOptimistic
    If IsNull(rs.Fields(0)) Then
        clientid = 0
        txtcid.Text = clientid
    Else
        clientid = rs.Fields(0)
        txtcid.Text = clientid
    End If
    rs.Close
    Set rs = New ADODB.Recordset
    rs.Open "select max(pol_no) from policy_info", cn, adOpenStatic, adLockOptimistic
    If IsNull(rs.Fields(0)) Then
        clientid = 0
        txtpid.Text = clientid
    Else
        clientid = rs.Fields(0)
        txtpid.Text = clientid
    End If
    rs.Close

End Function

Public Function MAX1()
Dim id As Integer
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "select max(agent_id) from agent_info", cn, adOpenStatic, adLockOptimistic
    If IsNull(rs.Fields(0)) Then
        id = 3001
        txtaid.Text = id
    Else
        id = rs.Fields(0) + 1
        txtaid.Text = id
    End If
    rs.Close
    Set rs = New ADODB.Recordset
    rs.Open "select max(nominee_id) from nominee_info", cn, adOpenStatic, adLockOptimistic
    If IsNull(rs.Fields(0)) Then
        id = 2001
        txtnid.Text = id
    Else
        id = rs.Fields(0) + 1
        txtnid.Text = id
    End If
    rs.Close
    Set rs = New ADODB.Recordset
    rs.Open "select max(client_id) from client_info", cn, adOpenStatic, adLockOptimistic
    If IsNull(rs.Fields(0)) Then
        id = 0
        txtcid.Text = id
    Else
        id = rs.Fields(0)
        txtcid.Text = id
    End If
    rs.Close
End Function

Private Sub Form_Load()
Call connect
awa.AddItem "pune"
awa.AddItem "mumbai"
awa.AddItem "nasik"
awa.AddItem "ahmednagar"
awa.AddItem "satara"
awa.AddItem "kolhapur"
awa.AddItem "solapur"
awa.AddItem "alibag"
awa.AddItem "nagpur"
awa.AddItem "latur"
awa.AddItem "sangli"
awa.AddItem "lonavla"
awa.AddItem "jalgaon"
End Sub


Private Sub aage_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8) Then
        KeyAscii = 0
        MsgBox "Please Enter Integer Only!!!!!", vbExclamation
    End If
End Sub

Private Sub afname_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8) Then
        KeyAscii = 0
        MsgBox "Please Enter character Only!!!!!", vbExclamation
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub



Private Sub alname_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8) Then
        KeyAscii = 0
        MsgBox "Please Enter character Only!!!!!", vbExclamation
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub amname_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8) Then
        KeyAscii = 0
        MsgBox "Please Enter character Only!!!!!", vbExclamation
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub


Private Sub amob_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8) Then
        KeyAscii = 0
        MsgBox "Please Enter Integer Only!!!!!", vbExclamation
    End If
End Sub

Private Sub aoff_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8) Then
        KeyAscii = 0
        MsgBox "Please Enter Integer Only!!!!!", vbExclamation
    End If
End Sub


Private Sub apin_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8) Then
        KeyAscii = 0
        MsgBox "Please Enter Integer Only!!!!!", vbExclamation
    End If
End Sub

Private Sub nage_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8) Then
        KeyAscii = 0
        MsgBox "Please Enter Integer Only!!!!!", vbExclamation
    End If
End Sub

Private Sub nfname_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8) Then
        KeyAscii = 0
        MsgBox "Please Enter character Only!!!!!", vbExclamation
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub



Private Sub nlname_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8) Then
        KeyAscii = 0
        MsgBox "Please Enter character Only!!!!!", vbExclamation
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub nmname_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8) Then
        KeyAscii = 0
        MsgBox "Please Enter character Only!!!!!", vbExclamation
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub



Private Sub nmob_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8) Then
        KeyAscii = 0
        MsgBox "Please Enter Integer Only!!!!!", vbExclamation
    End If
End Sub


Private Sub noff_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8) Then
        KeyAscii = 0
        MsgBox "Please Enter Integer Only!!!!!", vbExclamation
    End If
End Sub


Private Sub npin_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "0123456789"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8) Then
        KeyAscii = 0
        MsgBox "Please Enter Integer Only!!!!!", vbExclamation
    End If
End Sub

Private Sub nrel_KeyPress(KeyAscii As Integer)
    Dim nstr As String
    Dim i As Integer
    nstr = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    i = InStr(nstr, Chr(KeyAscii))
    If (i = 0 And KeyAscii <> 8) Then
        KeyAscii = 0
        MsgBox "Please Enter character Only!!!!!", vbExclamation
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub
