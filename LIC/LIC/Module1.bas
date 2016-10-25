Attribute VB_Name = "Module1"
Public cn As ADODB.Connection



Public Function connect() As Boolean

Set cn = New ADODB.Connection
cn.Open "dsn=pallavi", "ashu", "ashu"

End Function
