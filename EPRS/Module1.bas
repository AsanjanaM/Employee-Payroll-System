Attribute VB_Name = "Module1"
Global con As adodb.Connection
Global rs As adodb.Recordset
Global rs1 As adodb.Recordset
Public Function connectdb()
Set con = New adodb.Connection
con.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\EMPRS.mdb;Persist Security Info=False")
End Function

