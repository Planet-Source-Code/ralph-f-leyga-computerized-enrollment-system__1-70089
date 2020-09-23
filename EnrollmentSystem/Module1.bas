Attribute VB_Name = "Module1"
Public db As New ADODB.Connection
Public rsadd As New ADODB.Recordset
Public rsupdate As New ADODB.Recordset
Public rsRemove As New ADODB.Recordset
Public rs As New ADODB.Recordset
Public rsSY As New ADODB.Recordset
Public rsSYView As New ADODB.Recordset
Public rsPay As New ADODB.Recordset
Public rsView As New ADODB.Recordset
Public rsView1 As New ADODB.Recordset
Public rsCount As New ADODB.Recordset
Public rsinfo As New ADODB.Recordset
Public bol As Boolean
'Global rpt_header As report_header
'Public rpt_header As report_header
Public Sub dbase()
Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= dbase.mdb ;Persist Security Info=False;Jet OLEDB:Database Password=cheese"
End Sub
Public Sub EnableFld(FormName As Form, bVal As Boolean)
    Dim ObjCtrl As Control
    
    For Each ObjCtrl In FormName.Controls
        If TypeOf ObjCtrl Is TextBox Then
            ObjCtrl.Enabled = bVal
        ElseIf TypeOf ObjCtrl Is ComboBox Then
            ObjCtrl.Enabled = bVal
       ' ElseIf TypeOf ObjCtrl Is DTPicker Then
            ObjCtrl.Enabled = bVal
       ' ElseIf TypeOf ObjCtrl Is DataList Then
           ' ObjCtrl.Enabled = bVal
       '' ElseIf TypeOf ObjCtrl Is DataCombo Then
           ' ObjCtrl.Enabled = bVal
        End If
    Next ObjCtrl
    
    Set ObjCtrl = Nothing
End Sub


