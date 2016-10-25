VERSION 5.00
Begin VB.MDIForm MDImaster 
   BackColor       =   &H8000000C&
   Caption         =   "LIC Management System"
   ClientHeight    =   3120
   ClientLeft      =   855
   ClientTop       =   1125
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Menu mnupolicy 
      Caption         =   "&Policy"
      Begin VB.Menu mnunew 
         Caption         =   "&New Client"
      End
      Begin VB.Menu mnu 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnupremium 
      Caption         =   "P&remium"
      Begin VB.Menu mnureceipt 
         Caption         =   "&Get Premium Receipt"
      End
      Begin VB.Menu mnupay 
         Caption         =   "&Pay Premium"
      End
   End
   Begin VB.Menu mnuclaim 
      Caption         =   "&Claim"
      Begin VB.Menu mnuclaimamt 
         Caption         =   "&Get Claim Amount"
      End
   End
   Begin VB.Menu mnusearch 
      Caption         =   "&Search"
      Begin VB.Menu mnusclient 
         Caption         =   "&Client"
         Begin VB.Menu mnucid 
            Caption         =   "By Client ID"
         End
         Begin VB.Menu mnucname 
            Caption         =   "By Client Name"
         End
         Begin VB.Menu mnucagentid 
            Caption         =   "By Agent ID"
         End
         Begin VB.Menu mnucpolid 
            Caption         =   "By Policy ID"
         End
      End
      Begin VB.Menu mnusnominee 
         Caption         =   "&Nominee"
         Begin VB.Menu mnunid 
            Caption         =   "By Nominee ID"
         End
         Begin VB.Menu mnunname 
            Caption         =   "By Nominee Name"
         End
         Begin VB.Menu mnuncid 
            Caption         =   "By Client ID"
         End
      End
      Begin VB.Menu mnusagent 
         Caption         =   "&Agent"
         Begin VB.Menu mnuaid 
            Caption         =   "By Agent ID"
         End
         Begin VB.Menu mnuaname 
            Caption         =   "By Agent name"
         End
      End
   End
   Begin VB.Menu mnudetails 
      Caption         =   "&Details"
      Begin VB.Menu mnudclient 
         Caption         =   "&Client"
      End
      Begin VB.Menu mnudnominee 
         Caption         =   "&Nominee"
      End
      Begin VB.Menu mnudagent 
         Caption         =   "&Agent"
      End
      Begin VB.Menu mnudpolicy 
         Caption         =   "&Policy"
      End
      Begin VB.Menu mnudpremium 
         Caption         =   "P&remium"
      End
      Begin VB.Menu mnudclaim 
         Caption         =   "C&laim"
      End
   End
   Begin VB.Menu mnureports 
      Caption         =   "&Reports"
      Begin VB.Menu mnurclient 
         Caption         =   "&Client"
      End
      Begin VB.Menu mnurnominee 
         Caption         =   "&Nominee"
      End
      Begin VB.Menu mnuagent 
         Caption         =   "&Agent"
      End
      Begin VB.Menu mnurpolicy 
         Caption         =   "&Policy"
      End
      Begin VB.Menu mnurpremium 
         Caption         =   "P&remium"
      End
      Begin VB.Menu mnurclaim 
         Caption         =   "C&laim"
      End
   End
   Begin VB.Menu mnuplandetails 
      Caption         =   "&Plan Details"
   End
End
Attribute VB_Name = "MDImaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    Call conn
End Sub

Private Sub mnuagent_Click()
    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    rst.Open "select agent_id,agent_fname||' '||agent_mname||' '||agent_lname as name,res_address||', '||pincode as address,tel_mob,tel_off,work_area from agent_info order by agent_id", con, adOpenStatic, adLockOptimistic
    If (rst.RecordCount > 0) Then
        Set dragent.DataSource = rst
        dragent.Sections("details").Controls("txtaid").DataField = rst(0).Name
        dragent.Sections("details").Controls("txtname").DataField = rst(1).Name
        dragent.Sections("details").Controls("txtadd").DataField = rst(2).Name
        dragent.Sections("details").Controls("txtmno").DataField = rst(3).Name
        dragent.Sections("details").Controls("txtono").DataField = rst(4).Name
        dragent.Sections("details").Controls("txtwarea").DataField = rst(5).Name
        dragent.Show
    Else
        MsgBox "Agent Details Not Present!!!!!", vbExclamation
    End If
End Sub

Private Sub mnuaid_Click()
    Unload Me
    frmsearch.Show
    frmsearch.txtsearchof.Enabled = False
    frmsearch.txtsearchby.Enabled = False
    frmsearch.txtsearchof.Text = "Agent"
    frmsearch.txtsearchby.Text = "Agent ID"
End Sub

Private Sub mnuaname_Click()
    Unload Me
    frmsearch.Show
    frmsearch.txtsearchof.Enabled = False
    frmsearch.txtsearchby.Enabled = False
    frmsearch.txtsearchof.Text = "Agent"
    frmsearch.txtsearchby.Text = "Agent Name"
End Sub

Private Sub mnucagentid_Click()
    Unload Me
    frmsearch.Show
    frmsearch.txtsearchof.Enabled = False
    frmsearch.txtsearchby.Enabled = False
    frmsearch.txtsearchof.Text = "Client"
    frmsearch.txtsearchby.Text = "Agent ID"
End Sub

Private Sub mnucdob_Click()
    Unload Me
    frmsearch.Show
    frmsearch.txtsearchof.Enabled = False
    frmsearch.txtsearchby.Enabled = False
    frmsearch.txtsearchof.Text = "Client"
    frmsearch.txtsearchby.Text = "Date Of Birth"
End Sub

Private Sub mnucid_Click()
    Unload Me
    frmsearch.Show
    frmsearch.txtsearchof.Enabled = False
    frmsearch.txtsearchby.Enabled = False
    frmsearch.txtsearchof.Text = "Client"
    frmsearch.txtsearchby.Text = "Client ID"
End Sub

Private Sub mnuclaimamt_Click()
    Unload Me
    frmpremium.Show
    frmpremium.frmpaypremium.Enabled = False
    frmpremium.frmclaim.Enabled = True
End Sub

Private Sub mnucname_Click()
    Unload Me
    frmsearch.Show
    frmsearch.txtsearchof.Enabled = False
    frmsearch.txtsearchby.Enabled = False
    frmsearch.txtsearchof.Text = "Client"
    frmsearch.txtsearchby.Text = "Client Name"
End Sub

Private Sub mnucpolid_Click()
    Unload Me
    frmsearch.Show
    frmsearch.txtsearchof.Enabled = False
    frmsearch.txtsearchby.Enabled = False
    frmsearch.txtsearchof.Text = "Client"
    frmsearch.txtsearchby.Text = "Policy ID"
End Sub

Private Sub mnudagent_Click()
    Unload Me
    frmdetails.Show
    frmdetails.txtdetailsof.Enabled = False
    frmdetails.txtdetailsof.Text = "Agent"
End Sub

Private Sub mnudclaim_Click()
    Unload Me
    frmdetails.Show
    frmdetails.txtdetailsof.Enabled = False
    frmdetails.txtdetailsof.Text = "Claim"
End Sub

Private Sub mnudclient_Click()
    Unload Me
    frmdetails.Show
    frmdetails.txtdetailsof.Enabled = False
    frmdetails.txtdetailsof.Text = "Client"
End Sub

Private Sub mnudnominee_Click()
    Unload Me
    frmdetails.Show
    frmdetails.txtdetailsof.Enabled = False
    frmdetails.txtdetailsof.Text = "Nominee"
End Sub

Private Sub mnudpolicy_Click()
    Unload Me
    frmdetails.Show
    frmdetails.txtdetailsof.Enabled = False
    frmdetails.txtdetailsof.Text = "Policy"
End Sub

Private Sub mnudpremium_Click()
    Unload Me
    frmdetails.Show
    frmdetails.txtdetailsof.Enabled = False
    frmdetails.txtdetailsof.Text = "Premium"
End Sub

Private Sub mnuexit_Click()
    MsgBox "Thank You!!!!!", vbInformation
    End
End Sub

Private Sub mnuncid_Click()
    Unload Me
    frmsearch.Show
    frmsearch.txtsearchof.Enabled = False
    frmsearch.txtsearchby.Enabled = False
    frmsearch.txtsearchof.Text = "Nominee"
    frmsearch.txtsearchby.Text = "Client ID"
End Sub

Private Sub mnunew_Click()
    Unload Me
    frmclient.Show
End Sub

Private Sub mnunid_Click()
    Unload Me
    frmsearch.Show
    frmsearch.txtsearchof.Enabled = False
    frmsearch.txtsearchby.Enabled = False
    frmsearch.txtsearchof.Text = "Nominee"
    frmsearch.txtsearchby.Text = "Nominee ID"
End Sub

Private Sub mnunname_Click()
    Unload Me
    frmsearch.Show
    frmsearch.txtsearchof.Enabled = False
    frmsearch.txtsearchby.Enabled = False
    frmsearch.txtsearchof.Text = "Nominee"
    frmsearch.txtsearchby.Text = "Nominee Name"
End Sub

Private Sub mnupay_Click()
    Unload Me
    frmpremium.Show
    frmpremium.frmpaypremium.Enabled = True
    frmpremium.frmclaim.Enabled = False
End Sub

Private Sub mnuplandetails_Click()
    Unload Me
    frmplandetails.Show
End Sub

Private Sub mnurclaim_Click()
    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    rst.Open "select claim_id,client_id,claim_date,status,causeofdeath,amount from claim_info order by claim_id", con, adOpenStatic, adLockOptimistic
    If (rst.RecordCount > 0) Then
        Set drclaim.DataSource = rst
        drclaim.Sections("details").Controls("txtclaimid").DataField = rst(0).Name
        drclaim.Sections("details").Controls("txtcid").DataField = rst(1).Name
        drclaim.Sections("details").Controls("txtcdate").DataField = rst(2).Name
        drclaim.Sections("details").Controls("txtstatus").DataField = rst(3).Name
        drclaim.Sections("details").Controls("txtcod").DataField = rst(4).Name
        drclaim.Sections("details").Controls("txtcamt").DataField = rst(5).Name
        drclaim.Show
    Else
        MsgBox "Claim Details Not Present!!!!!", vbExclamation
    End If
End Sub

Private Sub mnurclient_Click()
    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    rst.Open "select client_id,client_fname||' '||client_mname||' '||client_lname as name,res_address||', '||pincode as address,tel_mob,tel_off,agent_id,pol_no from client_info order by client_id", con, adOpenStatic, adLockOptimistic
    If (rst.RecordCount > 0) Then
        Set drclient.DataSource = rst
        drclient.Sections("details").Controls("txtcid").DataField = rst(0).Name
        drclient.Sections("details").Controls("txtname").DataField = rst(1).Name
        drclient.Sections("details").Controls("txtadd").DataField = rst(2).Name
        drclient.Sections("details").Controls("txtmno").DataField = rst(3).Name
        drclient.Sections("details").Controls("txtono").DataField = rst(4).Name
        drclient.Sections("details").Controls("txtaid").DataField = rst(5).Name
        drclient.Sections("details").Controls("txtpno").DataField = rst(6).Name
        drclient.Show
    Else
        MsgBox "Client Details Not Present!!!!!", vbExclamation
    End If
End Sub

Private Sub mnureceipt_Click()
    Unload Me
    frmreceipt.Show
End Sub

Private Sub mnurnominee_Click()
    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    rst.Open "select nominee_id,nom_fname||' '||nom_mname||' '||nom_lname as name,res_address||', '||pincode as address,tel_mob,tel_off,relation,client_id from nominee_info order by nominee_id", con, adOpenStatic, adLockOptimistic
    If (rst.RecordCount > 0) Then
        Set drnominee.DataSource = rst
        drnominee.Sections("details").Controls("txtnid").DataField = rst(0).Name
        drnominee.Sections("details").Controls("txtname").DataField = rst(1).Name
        drnominee.Sections("details").Controls("txtadd").DataField = rst(2).Name
        drnominee.Sections("details").Controls("txtmno").DataField = rst(3).Name
        drnominee.Sections("details").Controls("txtono").DataField = rst(4).Name
        drnominee.Sections("details").Controls("txtrel").DataField = rst(5).Name
        drnominee.Sections("details").Controls("txtcid").DataField = rst(6).Name
        drnominee.Show
    Else
        MsgBox "Nominee Details Not Present!!!!!", vbExclamation
    End If
End Sub

Private Sub mnurpolicy_Click()
    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    rst.Open "select pol_no,pol_id,pol_name,pol_amount,pol_duration,pol_mode,proposal_date,client_age from policy_info order by pol_no", con, adOpenStatic, adLockOptimistic
    If (rst.RecordCount > 0) Then
        Set drpolicy.DataSource = rst
        drpolicy.Sections("details").Controls("txtpno").DataField = rst(0).Name
        drpolicy.Sections("details").Controls("txtpid").DataField = rst(1).Name
        drpolicy.Sections("details").Controls("txtname").DataField = rst(2).Name
        drpolicy.Sections("details").Controls("txtamt").DataField = rst(3).Name
        drpolicy.Sections("details").Controls("txtdur").DataField = rst(4).Name
        drpolicy.Sections("details").Controls("txtmode").DataField = rst(5).Name
        drpolicy.Sections("details").Controls("txtpdate").DataField = rst(6).Name
        drpolicy.Sections("details").Controls("txtcage").DataField = rst(7).Name
        drpolicy.Show
    Else
        MsgBox "Policy Details Not Present!!!!!", vbExclamation
    End If
End Sub

Private Sub mnurpremium_Click()
    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    rst.Open "select client_id,pol_id,total,paid,premium_amt,premium_date,due_date,due_amt,status from premium_info order by client_id", con, adOpenStatic, adLockOptimistic
    If (rst.RecordCount > 0) Then
        Set drpremium.DataSource = rst
        drpremium.Sections("details").Controls("txtcid").DataField = rst(0).Name
        drpremium.Sections("details").Controls("txtpid").DataField = rst(1).Name
        drpremium.Sections("details").Controls("txtpamt").DataField = rst(2).Name
        drpremium.Sections("details").Controls("txtpaidamt").DataField = rst(3).Name
        drpremium.Sections("details").Controls("txtpramt").DataField = rst(4).Name
        drpremium.Sections("details").Controls("txtpdate").DataField = rst(5).Name
        drpremium.Sections("details").Controls("txtddate").DataField = rst(6).Name
        drpremium.Sections("details").Controls("txtdamt").DataField = rst(7).Name
        drpremium.Sections("details").Controls("txtstatus").DataField = rst(8).Name
        drpremium.Show
    Else
        MsgBox "Premium Details Not Present!!!!!", vbExclamation
    End If
End Sub
