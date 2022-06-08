Imports System
Imports System.Data
Imports System.Data.SqlClient

Module ShareFunction

  Dim Sql As String = ""
  Dim myConnection As SqlConnection
  Dim myCommand As SqlCommand
  Dim myReader As SqlDataReader

  Function GetMemberId(ByVal ConnString As String, ByVal tableName As String) As DataTable

    Dim mytable As New DataTable

    myConnection = New SqlConnection(ConnString)
    Try
      Sql = "SELECT DISTINCT memberId FROM " & tableName
      myConnection.Open()
      myCommand = New SqlCommand(Sql, myConnection)
      myReader = myCommand.ExecuteReader()
      mytable.Load(myReader)
      If Not myReader.IsClosed() Then myReader.Close()
      myCommand.Dispose()
    Catch ex As Exception
      SaveMemberGradeLog(ConnString, "exception", ex.ToString)
      Return mytable
    Finally
      If myConnection.State = ConnectionState.Open Then myConnection.Close()
    End Try
    Return mytable

  End Function

  Function GetMemberId(ByVal ConnString As String, ByVal tableName As String, ByVal whereColumnName As String, _
                       ByVal startDate As String, ByVal endDate As String) As DataTable

    Dim mytable As New DataTable

    myConnection = New SqlConnection(ConnString)
    Try
      Sql = "SELECT DISTINCT memberId FROM " & tableName & " WHERE " & whereColumnName & " BETWEEN @startdate AND @enddate "
      myConnection.Open()
      myCommand = New SqlCommand(Sql, myConnection)
      myCommand.Parameters.AddWithValue("@startdate", startDate)
      myCommand.Parameters.AddWithValue("@enddate", endDate & " 23:59:59")
      myReader = myCommand.ExecuteReader()
      mytable.Load(myReader)
      If Not myReader.IsClosed() Then myReader.Close()
      myCommand.Dispose()
    Catch ex As Exception
      SaveMemberGradeLog(ConnString, "exception", ex.ToString)
      Return mytable
    Finally
      If myConnection.State = ConnectionState.Open Then myConnection.Close()
    End Try
    Return mytable

  End Function

  Function GetDataTableBySql(ByVal ConnString As String, ByVal sql As String) As DataTable

    Dim mytable As New DataTable

    myConnection = New SqlConnection(ConnString)
    Try      
      myConnection.Open()
      myCommand = New SqlCommand(sql, myConnection)
      myReader = myCommand.ExecuteReader()
      mytable.Load(myReader)
      If Not myReader.IsClosed() Then myReader.Close()
      myCommand.Dispose()
    Catch ex As Exception
      SaveMemberGradeLog(ConnString, "exception", ex.ToString)
      Return mytable
    Finally
      If myConnection.State = ConnectionState.Open Then myConnection.Close()
    End Try
    Return mytable

  End Function

  Function GetTotalBySql(ByVal ConnString As String, ByVal sql As String, ByVal memberId As String) As Integer

    Dim total As Integer = 0

    myConnection = New SqlConnection(ConnString)
    Try
      myConnection.Open()
      myCommand = New SqlCommand(Sql, myConnection)
      myCommand.Parameters.AddWithValue("@memberId", memberId)
      total = myCommand.ExecuteScalar
      myCommand.Dispose()
    Catch ex As Exception
      SaveMemberGradeLog(ConnString, "exception", ex.ToString)
      Return 0
    Finally
      If myConnection.State = ConnectionState.Open Then myConnection.Close()
    End Try
    Return total

  End Function

  Function GetTotalBySql(ByVal ConnString As String, ByVal sql As String, ByVal memberId As String, ByVal theDate As String) As Integer

    Dim total As Integer = 0

    myConnection = New SqlConnection(ConnString)
    Try
      myConnection.Open()
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@memberId", memberId)
      myCommand.Parameters.AddWithValue("@theDate", theDate)
      total = myCommand.ExecuteScalar
      myCommand.Dispose()
    Catch ex As Exception
      SaveMemberGradeLog(ConnString, "exception", ex.ToString)
      Return 0
    Finally
      If myConnection.State = ConnectionState.Open Then myConnection.Close()
    End Try
    Return total

  End Function

  Function CheckTableWithContent(ByVal ConnString As String, ByVal tableName As String, ByVal columnName As String, ByVal id As String) As Boolean

    Dim globalFlag As Boolean = False
    myConnection = New SqlConnection(ConnString)
    Try
      myConnection.Open()
      Sql = "SELECT " & columnName & " FROM " & tableName & " WHERE " & columnName & " = @id "
      myCommand = New SqlCommand(Sql, myConnection)
      myCommand.Parameters.AddWithValue("@id", id)
      myReader = myCommand.ExecuteReader
      If myReader.HasRows Then
        globalFlag = True
      Else
        globalFlag = False
      End If
      If Not myReader.IsClosed Then myReader.Close()
      myCommand.Dispose()
    Catch ex As Exception
      SaveMemberGradeLog(ConnString, "exception", ex.ToString)
      Return False
    Finally
      If myConnection.State = ConnectionState.Open Then myConnection.Close()
    End Try
    Return globalFlag

  End Function

  Function CheckMemberGradeLoginWithloginDate(ByVal ConnString As String, ByVal memberId As String) As Boolean

    Dim globalFlag As Boolean = False
    myConnection = New SqlConnection(ConnString)
    Try
      myConnection.Open()
      Sql = "SELECT memberId FROM MemberGradeLogin WHERE memberId = @id1 AND CONVERT(varchar, loginDate, 111) = CONVERT(varchar, GETDATE(), 111) "
      myCommand = New SqlCommand(Sql, myConnection)
      myCommand.Parameters.AddWithValue("@id1", memberId)      
      myReader = myCommand.ExecuteReader
      If myReader.HasRows Then
        globalFlag = True
      Else
        globalFlag = False
      End If
      If Not myReader.IsClosed Then myReader.Close()
      myCommand.Dispose()
    Catch ex As Exception
      SaveMemberGradeLog(ConnString, "exception", ex.ToString)
      Return False
    Finally
      If myConnection.State = ConnectionState.Open Then myConnection.Close()
    End Try
    Return globalFlag

  End Function

  Function CheckMemberGradeLoginWithloginDate(ByVal ConnString As String, ByVal memberId As String, ByVal theDate As String) As Boolean

    Dim globalFlag As Boolean = False
    myConnection = New SqlConnection(ConnString)
    Try
      myConnection.Open()
      Sql = "SELECT memberId FROM MemberGradeLogin WHERE memberId = @id1 AND CONVERT(varchar, loginDate, 111) = CONVERT(varchar, '" & theDate & "', 111) "
      myCommand = New SqlCommand(Sql, myConnection)
      myCommand.Parameters.AddWithValue("@id1", memberId)
      myReader = myCommand.ExecuteReader
      If myReader.HasRows Then
        globalFlag = True
      Else
        globalFlag = False
      End If
      If Not myReader.IsClosed Then myReader.Close()
      myCommand.Dispose()
    Catch ex As Exception
      SaveMemberGradeLog(ConnString, "exception", ex.ToString)
      Return False
    Finally
      If myConnection.State = ConnectionState.Open Then myConnection.Close()
    End Try
    Return globalFlag

  End Function

  Function SelectTableWithContent(ByVal ConnString As String, ByVal tableName As String, ByVal selectColumnName As String, _
                                  ByVal whereColumnName As String, ByVal whereValue As String) As String

    Dim returnStr As String = ""
    myConnection = New SqlConnection(ConnString)
    Try
      myConnection.Open()
      Sql = "SELECT " & selectColumnName & " FROM " & tableName & " WHERE " & whereColumnName & " = @id "      
      myCommand = New SqlCommand(Sql, myConnection)
      myCommand.Parameters.AddWithValue("@id", whereValue)
      myReader = myCommand.ExecuteReader
      If myReader.Read Then
        If myReader(selectColumnName) IsNot DBNull.Value Then
          returnStr = myReader(selectColumnName).ToString.Trim
        End If
      End If
      If Not myReader.IsClosed Then myReader.Close()
      myCommand.Dispose()
    Catch ex As Exception
      SaveMemberGradeLog(ConnString, "exception", ex.ToString)
      Return returnStr
    Finally
      If myConnection.State = ConnectionState.Open Then myConnection.Close()
    End Try
    Return returnStr

  End Function

  Function UpdateTableWithSql(ByVal ConnString As String, ByVal sql As String) As Boolean

    myConnection = New SqlConnection(ConnString)
    Try
      myConnection.Open()
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.ExecuteNonQuery()
      myCommand.Dispose()
    Catch ex As Exception
      SaveMemberGradeLog(ConnString, "exception", ex.ToString)
      Return False
    Finally
      If myConnection.State = ConnectionState.Open Then myConnection.Close()
    End Try
    Return True

  End Function

  Sub SaveMemberGradeLog(ByVal connString As String, ByVal logtype As String, ByVal logdesc As String)

    myConnection = New SqlConnection(ConnString)
    Try
      myConnection.Open()
      Sql = "INSERT INTO MemberGradeLog VALUES(GETDATE(), @logtype, @logdesc) "
      myCommand = New SqlCommand(Sql, myConnection)
      myCommand.Parameters.AddWithValue("@logtype", logtype)
      myCommand.Parameters.AddWithValue("@logdesc", logdesc)
      myCommand.ExecuteNonQuery()
      myCommand.Dispose()
    Catch ex As Exception
    Finally
      If myConnection.State = ConnectionState.Open Then myConnection.Close()
    End Try

  End Sub

  Function GetRaiseMemberId(ByVal ConnString As String, ByVal tableName As String, ByVal whereColumnName As String, _
                      ByVal startDate As String, ByVal endDate As String) As DataTable

    Dim mytable As New DataTable

    myConnection = New SqlConnection(ConnString)
    Try
      Sql = "SELECT DISTINCT USERID, CONVERT(varchar, TODAY, 111) AS TODAY, LOGINCOUNT FROM LBG_ACTIVITY WHERE (TODAY BETWEEN @startdate AND @enddate)"
      myConnection.Open()
      myCommand = New SqlCommand(Sql, myConnection)
      myCommand.Parameters.AddWithValue("@startdate", startDate)
      myCommand.Parameters.AddWithValue("@enddate", endDate)
      myReader = myCommand.ExecuteReader()
      mytable.Load(myReader)
      If Not myReader.IsClosed() Then myReader.Close()
      myCommand.Dispose()
    Catch ex As Exception
      Console.WriteLine(ex.ToString)
      Return mytable
    Finally
      If myConnection.State = ConnectionState.Open Then myConnection.Close()
    End Try
    Return mytable

  End Function

  Function GetRaiseLoginPoint(ByVal connstring As String, ByVal count As Integer) As Integer

    Dim code As String = "st_313"
    Dim str As String = ""
    Dim score As Integer = 0

    str = SelectTableWithContent(connstring, "kpi_set_score", "Rank0_1", "Rank0_2", code)
    If str = "" Then
      score = 0
    Else
      score = Integer.Parse(str)
    End If
    If count <= 3 Then
      score = score * count
    Else
      score = score * 3
    End If

    Return score

  End Function

  Sub UpdateRaiseLoginOnPoint(ByVal connstring As String, ByVal memberId As String, ByVal point As Integer, ByVal count As Integer, ByVal temp As String)

    '---檢查MemberGradeLogin Table中,在當日有沒有此memberId---
    If CheckMemberGradeLoginWithloginDate(connstring, memberId, temp) Then
      UpdateTableWithSql(connstring, "UPDATE MemberGradeLogin SET loginRaiseCount = " & count & ", loginRaiseGrade = " & point & ", loginRaiseDate = '" & temp & "' " & _
                                     " WHERE memberId = '" & memberId & "' AND CONVERT(varchar, loginDate, 111) = CONVERT(varchar, '" & temp & "', 111)")
    Else
      UpdateTableWithSql(connstring, "INSERT INTO MemberGradeLogin(memberId, loginDate, loginRaiseCount, loginRaiseGrade, loginRaiseDate) " & _
                                     "VALUES('" & memberId & "', '" & temp & "', " & count & ", " & point & ", '" & temp & "')")
    End If

  End Sub

End Module
