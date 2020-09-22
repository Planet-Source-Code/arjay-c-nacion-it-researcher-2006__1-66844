Attribute VB_Name = "modMain"
Public con As New Connection
Public rs As New Recordset
Public rsFav As New Recordset
Public rsReg As New Recordset
Public fav As String
Public item As String


Public Sub OpenCon()
    con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\data.mdb" & ";Persist Security Info=False;Jet OLEDB:Database Password=it2006"
End Sub

Public Sub LoadAll(lst As ListBox)
If rs.State <> adStateClosed Then rs.Close
rs.CursorLocation = adUseClient
rs.Open "Select * from list", con, adOpenStatic, adLockReadOnly
With rs
    Do While Not .EOF
        lst.AddItem rs(0)
        .MoveNext
    Loop
End With
End Sub

Sub Main()
frmSplash.Show
End Sub
