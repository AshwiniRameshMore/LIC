VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "MDIForm1"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   1155
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu policy 
      Caption         =   "Policy"
      Begin VB.Menu new 
         Caption         =   "New Client"
      End
   End
   Begin VB.Menu premium 
      Caption         =   "Premium"
      Begin VB.Menu pay 
         Caption         =   "Pay Premium"
      End
      Begin VB.Menu get 
         Caption         =   "Get Premium Receipt"
      End
   End
   Begin VB.Menu claim 
      Caption         =   "Claim"
      Begin VB.Menu getcLAIM 
         Caption         =   "Get claim Amount"
      End
   End
   Begin VB.Menu search 
      Caption         =   "Search"
      Begin VB.Menu sclient 
         Caption         =   "Client"
      End
      Begin VB.Menu snom 
         Caption         =   "Nominee"
      End
      Begin VB.Menu sagent 
         Caption         =   "Agent"
      End
   End
   Begin VB.Menu details 
      Caption         =   "Details"
      Begin VB.Menu dclient 
         Caption         =   "Client"
      End
      Begin VB.Menu dnom 
         Caption         =   "Nominee"
      End
      Begin VB.Menu dagent 
         Caption         =   "Agent"
      End
      Begin VB.Menu dpol 
         Caption         =   "Policy"
      End
      Begin VB.Menu dpre 
         Caption         =   "Premium"
      End
   End
   Begin VB.Menu plan 
      Caption         =   "Plan Details"
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub amt_Click()
Unload Me
CLAIM1.Show

End Sub

Private Sub Command1_Click()
Form1.Show
End Sub

Private Sub dagent_Click()
Unload Me
DETAILS1.Show
DETAILS1.Text1.Enabled = False
DETAILS1.Text1.Text = "AGENT"
End Sub

Private Sub dclient_Click()
Unload Me
DETAILS1.Show
DETAILS1.Text1.Enabled = False
DETAILS1.Text1.Text = "CLIENT"
End Sub

Private Sub dnom_Click()
Unload Me
DETAILS1.Show
DETAILS1.Text1.Enabled = False
DETAILS1.Text1.Text = "NOMINEE"

End Sub

Private Sub dpol_Click()
Unload Me
DETAILS1.Show
DETAILS1.Text1.Enabled = False
DETAILS1.Text1.Text = "POLICY"

End Sub

Private Sub dpre_Click()
Unload Me
DETAILS1.Show
DETAILS1.Text1.Enabled = False
DETAILS1.Text1.Text = "PREMIUM"

End Sub

Private Sub Exit_Click()
THANKS.Show
End Sub


Private Sub get_Click()
Unload Me
PREMIUMRECEIPT.Show

End Sub

Private Sub getcLAIM_Click()

CLAIM1.Show

End Sub

Private Sub MDIForm_Load()
Call connect
End Sub

Private Sub new_Click()
Unload Me
CLIENT.Show
End Sub

Private Sub pay_Click()
Unload Me
PREMIUM1.Show
End Sub

Private Sub plan_Click()

PLANDETAILS.Show
End Sub

Private Sub sagent_Click()
Unload Me
SEARCH1.Show
SEARCH1.txtser.Enabled = False
SEARCH1.txtser.Text = "AGENT"

End Sub

Private Sub sclient_Click()
Unload Me
SEARCH1.Show
SEARCH1.txtser.Enabled = False
SEARCH1.txtser.Text = "CLIENT"

End Sub

Private Sub snom_Click()
Unload Me
SEARCH1.Show
SEARCH1.txtser.Enabled = False
SEARCH1.txtser.Text = "NOMINEE"

End Sub

