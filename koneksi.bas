Attribute VB_Name = "Module1"
'--------------------------------------
Global dbkoneksi As ADODB.Connection
Global pembelian As ADODB.Recordset
Global stokbenang As ADODB.Recordset
Global mkrajut As ADODB.Recordset
Global maklooncelup As ADODB.Recordset
Global stokkaincelupan As ADODB.Recordset
Global so As ADODB.Recordset
Global totaldata As ADODB.Recordset
'--------------------------------------
Global selek As String
Global edit As String
Global delete As String
'--------------------------------------
Global iList As ListItem
'--------------------------------------

Sub koneksi()
Set dbkoneksi = New ADODB.Connection
    dbkoneksi.Open "Provider=Microsoft.jet.Oledb.4.0; Data Source=" & App.Path & "\database_mbtex.mdb"

End Sub
