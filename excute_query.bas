Attribute VB_Name = "excute_query"
Function send_query() As String
    'Dim cmd As New ADODB.Command
    'Dim sParams As ADODB.Parameter
    Call connectDB
    
    Dim strSql As String
    Dim Syntax As String
    Dim recordsAffected As Long
    strSql = Sheets("query").Range("b2").Value
    
    Syntax = confirm_Query(strSql)
    
    If Syntax = "open" Then
        Sheets("SQLresult").Rows(1 & ":" & Sheets("SQLresult").Rows.Count).Delete
        rs.Open strSql, cn, adLockReadOnly
        
        For intColIndex = 0 To rs.Fields.Count - 1
            Sheets("SQLresult").Range("A1").Offset(0, intColIndex).Value = rs.Fields(intColIndex).Name
        Next
        
        Sheets("SQLresult").Range("a2").CopyFromRecordset rs
        Sheets("SQLresult").Select
        'Cells.Select
        'Cells.EntireColumn.AutoFit
        Range("a1").Select
        
    ElseIf Syntax = "exe" Then
        cn.Execute (strSql)
        If Sheets("query").Range("a2").Value = "보임" Then
            MsgBox "실행됨 : " & strSql
        End If
        'cn.Execute CommandText:=strSql, recordsAffected:=affectedCount
        'Set sParams = cmd.CreateParameter("@MyVal", adVarChar, adParamOutput, 255, "")
        'cmd.Parameters.Append sParams
        '
        'cmd.ActiveConnection = cn
        'cmd.CommandText = strSql
        'cmd.CommandType = adCmdStoredProc
        'cmd.NamedParameters = True
        'cmd.Execute
        '
        ' MsgBox cmd.Parameters("@MyVal").Value
        
    Else
        MsgBox "<" & Syntax & ">는 SQL 쿼리가 아니거나, 허용되지 않았습니다."
    End If
    
    
    If rs.Fields.Count <> 0 Then
        rs.Close '
    End If
    cn.Close
    
    send_query = Syntax
End Function

Sub clean()
    Dim lastRow As Long
    Sheets("query").Unprotect
    lastRow = Sheets("query").Cells(Rows.Count, 3).End(xlUp).Row
    Sheets("query").Range("B2").Cut Destination:=Sheets("query").Range("C" & lastRow + 1)
    Sheets("query").Range("B3:B9999").Cut Destination:=Sheets("query").Range("B2")
    Sheets("query").Protect
End Sub

Sub infinite_query()
    If Range("b2").Value = "" Then
        MsgBox "입력된 쿼리가 없습니다. B2 셀에 쿼리를 입력하세요."
        Exit Sub
    End If
    Dim Syntax As String
    Do
        Syntax = send_query()
        Call clean
        If Range("B2").Value = "" Or Syntax = "open" Then
            Exit Do
        End If
    Loop
End Sub

