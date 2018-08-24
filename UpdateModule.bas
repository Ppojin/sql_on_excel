Attribute VB_Name = "UpdateModule"
Public cn_UM As ADODB.Connection
Public rs_UM As ADODB.Recordset
Public strSql_UM As String
Public currentDir, UserName, dropboxDir As String

Public Function FolderDir()
    currentDir = ThisWorkbook.Path + "\"
    UserName = Split(Split(currentDir, "EZwork")(1), "\")(0)
    dropboxDir = "E:\Dropbox\EZ Data\Synergy\update\"
End Function

Public Function connect_update_DB()
    Dim strServer_Name As String
    Dim strDB_Name As String
    Dim strUser_ID As String
    Dim strPassword As String
    
    strServer_Name = "35.189.159.65"
    strDB_Name = "ezUpdate"
    strUser_ID = "ezadmin"
    strPassword = "esct!##486"
    
    Set cn_UM = New ADODB.Connection
    Set rs_UM = New ADODB.Recordset
    
    cn_UM.Open _
        "DRIVER={MySQL ODBC 5.3 Unicode Driver}" & _
        ";port= 3306" & _
        ";SERVER=" & strServer_Name & _
        ";DATABASE=" & strDB_Name & _
        ";UID=" & strUser_ID & _
        ";PWD=" & strPassword & ""
    
End Function

Sub Update_To_Local()
    Sheets("compareVer").Unprotect
    Call FolderDir
    Call connect_update_DB
    strSql_UM = _
        "select c.filename, x, y, z, myLastVersion, newVersion " + _
        "from " + _
        "    (select " + _
        "        b.filename, b.x_majorUpgrade as x, b.y_minorUpgrade as y, b.z_bugFix as z, a.version as myLastVersion, b.version as newVersion, " + _
        "        (a.x_majorUpgrade*1000000 + a.y_minorUpgrade*1000 + a.z_bugFix) as int_myLastVersion, " + _
        "        (b.x_majorUpgrade*1000000 + b.y_minorUpgrade*1000 + b.z_bugFix) as int_newVersion " + _
        "    from " + _
        "        (select filename, x_majorUpgrade, y_minorUpgrade, z_bugFix, updateDate as updateDate, version " + _
        "        from " + _
        "            (SELECT filename AS f, MAX(id) AS u FROM userUpdateLog WHERE username = '" + userName + "' GROUP BY filename ) y " + _
        "        INNER Join " + _
        "            (SELECT * FROM userUpdateLog WHERE username = '" + userName + "') x " + _
        "        ON x.filename = y.f AND x.id = y.u) a " + _
        "    Right Join " + _
        "        (select filename, x_majorUpgrade, y_minorUpgrade, z_bugFix, updateDate as updateDate, version " + _
        "        from " + _
        "            (SELECT filename AS f, MAX(id) AS u FROM versionLog GROUP BY filename ) y " + _
        "        INNER Join " + _
        "            versionLog x " + _
        "        ON x.filename = y.f AND x.id = y.u) b " + _
        "    on a.filename = b.filename) c " + _
        "where int_myLastVersion < int_newVersion or int_myLastVersion is null;"
    
    Debug.Print strSql_UM
    rs_UM.Open strSql_UM, cn_UM, adLockReadOnly
    
    Dim FN, usedVer, newVer As String
    Dim confirmUpdateDone, luncherUpdateDone As Boolean
    confirmUpdateDone = False
    luncherUpdateDone = False
    Do While rs_UM.EOF = False
        FN = rs_UM.Fields("filename").Value
        If FN = ThisWorkbook.Name Then
            luncherUpdateDone = True
            ThisWorkbook.Save
            ThisWorkbook.SaveAs currentDir + "update_backup\foo"
        End If
        confirmUpdateDone = True
        newVer = rs_UM.Fields("newVersion").Value
        Set isHere = Sheets("CompareVer").Range("d1:d110").Find(FN, LookIn:=xlValues)
        If Dir(currentDir + FN) <> "" Then
            usedVer = rs_UM.Fields("myLastVersion").Value
            FileCopy (currentDir + FN), (currentDir + "update_backup\" + FN + "_" + Format(Now(), "yymmddhhnnss") + "_" + usedVer + "_.backup")
        End If
        FileCopy (dropboxDir + FN), (currentDir + FN)
        strSql_UM = _
            "insert into userUpdateLog(filename, x_majorUpgrade, y_minorUpgrade, z_bugFix, version, username) " + _
            "values( " + _
            "   '" + FN + "'," + _
            "   " + CStr(rs_UM.Fields("x").Value) + ", " + CStr(rs_UM.Fields("y").Value) + ", " + CStr(rs_UM.Fields("z").Value) + ", " + _
            "   '" + rs_UM.Fields("newVersion").Value + "', " + _
            "   '" + userName + "'" + _
            ");"
        Debug.Print (strSql_UM)
        cn_UM.Execute (strSql_UM)
        rs_UM.MoveNext
    Loop
    
    cn_UM.Close
    If luncherUpdateDone Then
        MsgBox "현재 파일이 업데이트되었습니다."
        Workbooks.Open currentDir + "###File_update.xlsm"
        Application.Run "'" + currentDir + "###File_update.xlsm'!Auto_Open"
        ThisWorkbook.Close
    End If
    
    If confirmUpdateDone = True Then
        Call Auto_Open
        MsgBox ("업데이트가 완료되었습니다.")
    Else
        MsgBox ("업데이트할 파일이 없습니다.")
    End If
    Sheets("compareVer").Protect
End Sub
