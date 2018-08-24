Attribute VB_Name = "functionModule"

Public cn As ADODB.Connection
Public rs As ADODB.Recordset
Public strSql As String
Public Function connectDB()
    Dim strServer_Name As String
    Dim strDB_Name As String
    Dim strUser_ID As String
    Dim strPassword As String
    
    strServer_Name = "here"     '''input here your server name
    strDB_Name = "here"         '''input here your database name
    strUser_ID = "here"         '''input here your id
    strPassword = "here"        '''input here your password
    
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    cn.Open _
        "DRIVER={MySQL ODBC 5.3 Unicode Driver}" & _
        ";port= 3306" & _
        ";SERVER=" & strServer_Name & _
        ";DATABASE=" & strDB_Name & _
        ";UID=" & strUser_ID & _
        ";PWD=" & strPassword & ""
    
End Function

Public Function confirm_Query(strSql As String) As String
    Dim Syntax_exe() As String
    Dim Syntax_open() As String
    Dim First_word_of_strsql As String
        
    Syntax_exe() = Split(UCase("create insert alter drop delete"))
    Syntax_open() = Split(UCase("select desc show"))
    First_word_of_strsql = UCase(Split(strSql)(0))
    
    For i = LBound(Syntax_exe) To UBound(Syntax_exe)
        If Syntax_exe(i) = First_word_of_strsql Then
            confirm_Query = "exe"
            Exit Function
        End If
    Next i
    
    For i = LBound(Syntax_open) To UBound(Syntax_open)
        If Syntax_open(i) = First_word_of_strsql Then
            confirm_Query = "open"
            Exit Function
        End If
    Next i
    
    confirm_Query = First_word_of_strsql
    
End Function
