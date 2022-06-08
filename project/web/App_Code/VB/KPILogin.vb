Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Public Class KPILogin

  Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
  Dim sql As String = ""
  Dim myConnection As SqlConnection
  Dim myCommand As SqlCommand
  Dim myReader As SqlDataReader

  Dim _memberId As String
  Dim _ctNode As String  
  Dim _pageType As String
  Dim _loginGrade As Integer
  Dim _loginCount As Integer
  Dim checkFlag As Boolean

  Dim _oldGrade As Integer
  Dim _loginCode As String
  Dim _loginCode2 As String
  Dim _limitCode As String
  Dim _isOverMax As Boolean

  Public Sub New(ByVal memberId As String, ByVal pageType As String, ByVal ctNode As String)

    If memberId = "" Then Exit Sub

    _memberId = memberId
    _pageType = pageType
    _ctNode = ctNode
    _isOverMax = False
    _oldGrade = 0

    _loginGrade = 0
    _loginCount = 3     '---一天登入3次算分---
    _loginCode = "st_3"
    _limitCode = ""

    '會員登入     st_311
    'IQ大挑戰登入 st_312
    '虛擬養殖登入 st_313

    If _pageType = "loginInterCount" Then _loginCode2 = "st_311"
    If _pageType = "loginIQCount" Then _loginCode2 = "st_312"
    If _pageType = "loginRaiseCount" Then _loginCode2 = "st_313"

    If _loginCode2 = "" Then Exit Sub    

  End Sub

  Public Sub HandleLogin()

    '---檢查KPI是否開啟記錄---
    If Not CheckKpiOpen() Then
      Exit Sub
    End If

    '---檢查瀏覽行為是否公開---
    If Not CheckLoginIsPublic() Then
      Exit Sub
    End If

    '---檢查傳入的memberId是否有登入當時的資料---
    If Not CheckMemberLoginToday() Then
      '---此member沒有今天的記錄---插入一筆新的---
      If Not InsertMemberDefault() Then
        Exit Sub
      End If
    End If

    '---利用pageType來抓值---
    If Not GetLoginGradeByPageType() Then
      Exit Sub
    End If

    '---檢查今日分數是否超過上限---
    If Not CheckGradeLimit() Then
      _isOverMax = True
    End If

    '---利用得到的給點,存入table中---
    If Not SaveGradeIntoDB() Then
      Exit Sub
    End If

  End Sub

  Function CheckKpiOpen() As Boolean

    Dim returnStr As String = ""
    returnStr = Trim(GetID.Get_OneField_ReturnIdStr("CodeMain", "codeMetaID", "kpiOpen", "mCode"))

    If returnStr = "Y" Then
      Return True
    Else
      Return False
    End If

  End Function

  ''' <summary>檢查瀏覽行為是否公開</summary>
  ''' <returns>公開 return true, 關閉 return false</returns>
  ''' <remarks></remarks>
  Function CheckLoginIsPublic() As Boolean

    Dim returnStr As String = ""
    returnStr = Trim(GetID.Get_OneField_ReturnIdStr("kpi_set_ind", "Rank0", _loginCode, "Rank2"))

    If returnStr = "Y" Then
      Return True
    Else
      Return False
    End If

  End Function

  Function CheckMemberLoginToday() As Boolean

    checkFlag = False
    myConnection = New SqlConnection(ConnString)
    Try
      sql = "SELECT memberId, loginDate  FROM MemberGradeLogin WHERE memberId = @memberId "
      sql &= "AND CONVERT(varchar, loginDate, 111) = CONVERT(varchar, GETDATE(), 111) "
      myConnection.Open()
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@memberId", _memberId)
      myReader = myCommand.ExecuteReader
      If myReader.Read Then
        checkFlag = True
      Else
        checkFlag = False
      End If
      If Not myReader.IsClosed Then
        myReader.Close()
      End If
      myCommand.Dispose()

    Catch ex As Exception
      Return False
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

    Return checkFlag

  End Function

  Function InsertMemberDefault() As Boolean

    checkFlag = False
    myConnection = New SqlConnection(ConnString)
    Try
      sql = "INSERT INTO MemberGradeLogin (memberId, loginDate) VALUES(@memberId, GETDATE())"
      myConnection.Open()
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@memberId", _memberId)
      myCommand.ExecuteNonQuery()
      myCommand.Dispose()
      checkFlag = True
    Catch ex As Exception
      Return False
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

    Return checkFlag

  End Function

  Function GetLoginGradeByPageType() As Boolean

    Try
      _loginGrade = GetID.Get_OneField_ReturnIdStr("kpi_set_score", "Rank0_2", _loginCode2, "Rank0_1")
    Catch
      Return False
    End Try

    Return True

  End Function

  ''' <summary>檢查瀏覽次數是否超過當日上限</summary>
  ''' <returns>超過上限return false, 沒超過上限return true</returns>
  ''' <remarks></remarks>
  Function CheckGradeLimit() As Boolean

    Dim totalCount As Integer = 0

    checkFlag = False
    myConnection = New SqlConnection(ConnString)
    Try

      sql = "SELECT memberId, loginDate, " & _pageType & " "
      sql &= " FROM MemberGradeLogin WHERE memberId = @memberId AND CONVERT(varchar, loginDate, 111) = CONVERT(varchar, GETDATE(), 111) "

      myConnection.Open()
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@memberId", _memberId)
      myReader = myCommand.ExecuteReader
      If myReader.Read Then
        If myReader(_pageType) IsNot DBNull.Value Then
          totalCount = Integer.Parse(myReader(_pageType))
        Else
          checkFlag = False
        End If
      Else
        checkFlag = False
      End If
      If Not myReader.IsClosed Then
        myReader.Close()
      End If
      myCommand.Dispose()

      If totalCount >= _loginCount Then
        checkFlag = False
      Else
        checkFlag = True
      End If

    Catch ex As Exception
      Return False
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

    Return checkFlag

  End Function

  Function SaveGradeIntoDB() As Boolean

    'Dim flag As Boolean = CheckGradeLimit()
    checkFlag = False
    myConnection = New SqlConnection(ConnString)
    Try

      sql = " UPDATE MemberGradeLogin SET " & _pageType & " = " & _pageType & " + 1 "

      If Not _isOverMax Then
        If _pageType = "loginInterCount" Then
          sql &= ", loginInterGrade = loginInterGrade + " & _loginGrade & ", loginInterDate = GETDATE() "
        End If
        If _pageType = "loginIQCount" Then
          sql &= ", loginIQGrade = loginIQGrade + " & _loginGrade & ", loginIQDate = GETDATE() "
        End If
        If _pageType = "loginRaiseCount" Then
          sql &= ", loginRaiseGrade = loginRaiseGrade + " & _loginGrade & ", loginRaiseDate = GETDATE() "
        End If
      End If
      sql &= " WHERE memberId = @memberId AND CONVERT(varchar, loginDate, 111) = CONVERT(varchar, GETDATE(), 111) "

      myConnection.Open()
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@memberId", _memberId)
      myCommand.ExecuteNonQuery()
      myCommand.Dispose()

      If Not _isOverMax Then
        sql = " INSERT INTO MemberGradeLoginDetail VALUES(@memberId, GETDATE(), '" & _pageType & "', " & _loginGrade & ", 'Y')"
      Else
        sql = " INSERT INTO MemberGradeLoginDetail VALUES(@memberId, GETDATE(), '" & _pageType & "', " & _loginGrade & ", 'O')"
      End If

      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@memberId", _memberId)
      myCommand.ExecuteNonQuery()
      myCommand.Dispose()

      checkFlag = True
    Catch ex As Exception
      Return False
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

    Return checkFlag

  End Function

End Class
