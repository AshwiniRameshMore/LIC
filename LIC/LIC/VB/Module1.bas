Attribute VB_Name = "Module1"
Public con As ADODB.Connection
Public rs As ADODB.Recordset

Public Function conn() As Boolean
    Set con = New ADODB.Connection
    con.Open "dsn=lic", "nitin", "jayganesh"
    conn = True
End Function
