Attribute VB_Name = "mod_db"
Public db As ADODB.Connection
Public public_rs As New ADODB.Recordset

Public Sub connect_db()
   On Error GoTo Message
   Set db = New ADODB.Connection
   With db
      .Mode = adModeReadWrite
      .CursorLocation = adUseClient
      .Open "DRIVER={MySQL ODBC 3.51 Driver};" _
            & "SERVER=localhost;" _
            & "PORT=3306;" _
            & "DATABASE=db_inventory;" _
            & "USER=root;" _
            & "PASSWORD=mysqladmin;" _
            & "OPTION=3;"
   End With
   Exit Sub
Message:
   MsgBox Err.Description, vbCritical + vbOKOnly
   If public_rs.State = adStateOpen Then public_rs.Close
   If db.State = adStateOpen Then db.Close
End Sub

Public Sub disconnect_db()
   On Error Resume Next
   If db.State = adStateOpen Then db.Close
   Set db = Nothing
End Sub

Public Function mysql_select(rs As ADODB.Recordset, sql As String) As ADODB.Recordset
    If rs.State = adStateOpen Then rs.Close
    rs.Open sql, db, adOpenStatic, adLockOptimistic
End Function

Public Sub set_datagrid(dg As DataGrid, rs As ADODB.Recordset, sql As String)
    Call mysql_select(rs, sql)
    Set dg.DataSource = rs
End Sub
