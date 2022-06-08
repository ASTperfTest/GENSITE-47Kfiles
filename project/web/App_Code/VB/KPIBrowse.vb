Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Public Class KPIBrowse

  Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
  Dim sql As String = ""
  Dim myConnection As SqlConnection
  Dim myCommand As SqlCommand
  Dim myReader As SqlDataReader
  
  Dim _memberId As String
  Dim _pageType As String
  Dim _ctNode As String
  Dim checkFlag As Boolean

  Dim _browseGrade As Integer

  Dim _browseCode As String
  Dim _browseCode2 As String
  Dim _limitCode As String
  Dim _ctnodeCode As String
  Dim _oldGrade As Integer
  Dim _limitScore As Integer
  Dim _isOverMax As Boolean
  Dim _isOverMaxFlag As Boolean

  Public Sub New(ByVal memberId As String, ByVal pageType As String, ByVal ctNode As String)

    If memberId = "" Then Exit Sub

    _memberId = memberId
    _pageType = pageType
    _ctNode = ctNode
    _limitScore = 0
    _isOverMax = False
    _isOverMaxFlag = False
    _oldGrade = 0
    _browseGrade = 0
    _browseCode = "st_2"
    _limitCode = "st_211"
    _ctnodeCode = "st_22"

    If _pageType = "browseInterCP" Then _browseCode2 = "st_212"

    If _pageType = "browseTopicCP" Then _browseCode2 = "st_213"

    If _pageType = "browseCatTreeCP" Then _browseCode2 = "st_214"

    If _pageType = "browseQuestionCP" Then _browseCode2 = "st_215"
    If _pageType = "browseDiscussLP" Then _browseCode2 = "st_215"
    If _pageType = "browseDiscussCP" Then _browseCode2 = "st_215"

    If _pageType = "browsePediaWordCP" Then _browseCode2 = "st_216"
    If _pageType = "browsePediaExplainLP" Then _browseCode2 = "st_216"
    If _pageType = "browsePediaExplainCP" Then _browseCode2 = "st_216"
	If _pageType = "browseInterCP" Then _browseCode2 = "st_217"
	
    If _browseCode2 = "" Then Exit Sub    

  End Sub

  Public Sub HandleBrowse()

    '---檢查KPI是否開啟記錄---
    If Not CheckKpiOpen() Then
      Exit Sub
    End If

    '---檢查瀏覽行為是否公開---
    If Not CheckBrowseIsPublic() Then
      Exit Sub
    End If

    '---檢查傳入的memberId是否有登入當時的資料---
    If Not CheckMemberLoginToday() Then
      '---此member沒有今天的記錄---插入一筆新的---
      If Not InsertMemberDefault() Then
        Exit Sub
      End If
    End If

    '---不論有沒有傳ctnode都先抓值---
    If Not GetBrowseGradeByPageType() Then
      Exit Sub
    End If

    If _ctNode <> "" Then
      If Not GetBrowseGradeByCtNode() Then                
      End If
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
  Function CheckBrowseIsPublic() As Boolean

    Dim returnStr As String = ""
    returnStr = Trim(GetID.Get_OneField_ReturnIdStr("kpi_set_ind", "Rank0", _browseCode, "Rank2"))

    If returnStr = "Y" Then
      Return True
    Else
      Return False
    End If

  End Function

  ''' <summary>檢查此memberid今天是否有登入</summary>
  ''' <returns>if login return true, else return false </returns>
  ''' <remarks></remarks>
  Function CheckMemberLoginToday() As Boolean

    checkFlag = False
    myConnection = New SqlConnection(ConnString)
    Try
      sql = "SELECT memberId, browseDate  FROM MemberGradeBrowse WHERE memberId = @memberId "
      sql &= "AND CONVERT(varchar, browseDate, 111) = CONVERT(varchar, GETDATE(), 111) "
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

      sql = "INSERT INTO MemberGradeBrowse (memberId, browseDate) VALUES(@memberId, GETDATE())"
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

  Function GetBrowseGradeByPageType() As Boolean
    Try
      _browseGrade = GetID.Get_OneField_ReturnIdStr("kpi_set_score", "Rank0_2", _browseCode2, "Rank0_1")
    Catch
      Return False
    End Try

    Return True

  End Function

  ''' <summary>利用ctnode的值來找出特殊的分數</summary>  
  ''' <returns>若有找到ctnode相對應的分數 return true, 若沒有 return false</returns>
  ''' <remarks></remarks>
  Function GetBrowseGradeByCtNode() As Boolean

    checkFlag = False
    myConnection = New SqlConnection(ConnString)
    Try
      sql = " SELECT Rank0_1 FROM kpi_set_score WHERE (Rank0 = @rank0) AND (Rank0_4 = @rank0_4) AND (Rank0_0 = @rank0_0) AND (xStatus = 'Y')"

      myConnection.Open()
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@rank0", _browseCode)
      myCommand.Parameters.AddWithValue("@rank0_4", _ctnodeCode)
      myCommand.Parameters.AddWithValue("@rank0_0", _ctNode)
      myReader = myCommand.ExecuteReader
      If myReader.Read Then
        If myReader("Rank0_1") IsNot DBNull.Value Then
          _browseGrade = Integer.Parse(myReader("Rank0_1"))
          checkFlag = True
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

    Catch ex As Exception
      Return False
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

    Return checkFlag

  End Function

  ''' <summary>檢查瀏覽分數總合是否超過當日上限</summary>
  ''' <returns>超過上限return false, 沒超過上限return true</returns>
  ''' <remarks></remarks>
  Function CheckGradeLimit() As Boolean

    _limitScore = GetID.Get_TwoField("kpi_set_score", "Rank0", _browseCode, "Rank0_2", _limitCode, "Rank0_1")
    Dim totalScore As Integer = 0

    checkFlag = False
    myConnection = New SqlConnection(ConnString)
    Try
      sql = " SELECT browseInterCP + browseTopicCP + browseCatTreeCP + browseQuestionCP + browseDiscussLP + browseDiscussCP "
      sql &= " + browsePediaWordCP + browsePediaExplainLP + browsePediaExplainCP AS total "
      sql &= " FROM MemberGradeBrowse WHERE memberId = @memberId AND CONVERT(varchar, browseDate, 111) = CONVERT(varchar, GETDATE(), 111) "

      myConnection.Open()
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@memberId", _memberId)
      myReader = myCommand.ExecuteReader
      If myReader.Read Then
        If myReader("total") IsNot DBNull.Value Then
          totalScore = Integer.Parse(myReader("total"))
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

      _oldGrade = totalScore

      If _pageType = "browseCatTreeCP" Then        
        If totalScore + _browseGrade + Integer.Parse(_ctNode) >= _limitScore Then
          If totalScore < _limitScore Then
            _isOverMaxFlag = True
          End If
          checkFlag = False
        Else
          checkFlag = True
        End If
      Else
        If totalScore + _browseGrade >= _limitScore Then
          If totalScore < _limitScore Then
            _isOverMaxFlag = True
          End If
          checkFlag = False
        Else
          checkFlag = True
        End If
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

    Dim tempScore As Integer = _limitScore - _oldGrade

    checkFlag = False
    myConnection = New SqlConnection(ConnString)
    Try
      If Not _isOverMax Then
        If _pageType = "browseCatTreeCP" Then
          '---把附件的數量加進去---用_ctnode代值---
          sql = " UPDATE MemberGradeBrowse SET " & _pageType & " = " & _pageType & " + " & _browseGrade & " + " & Integer.Parse(_ctNode) & " "
        Else
          sql = " UPDATE MemberGradeBrowse SET " & _pageType & " = " & _pageType & " + " & _browseGrade & " "
        End If
      Else
        sql = " UPDATE MemberGradeBrowse SET " & _pageType & " = " & _pageType & " + " & tempScore & " "
      End If
      sql &= " WHERE memberId = @memberId AND CONVERT(varchar, browseDate, 111) = CONVERT(varchar, GETDATE(), 111) "

      myConnection.Open()
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@memberId", _memberId)
      myCommand.ExecuteNonQuery()
      myCommand.Dispose()

      If _isOverMax Then
        If _isOverMaxFlag Then         
          sql = "INSERT INTO MemberGradeBrowseDetail VALUES(@memberId, GETDATE(), '" & _pageType & "', " & tempScore & ", " & _oldGrade & ", 'O', " & _ctNode & ") "
        Else
          If _pageType = "browseCatTreeCP" Then
            sql = "INSERT INTO MemberGradeBrowseDetail VALUES(@memberId, GETDATE(), '" & _pageType & "', " & _browseGrade & " + " & Integer.Parse(_ctNode) & ", " & _oldGrade & ", 'O', " & _ctNode & ") "
          Else
            sql = "INSERT INTO MemberGradeBrowseDetail VALUES(@memberId, GETDATE(), '" & _pageType & "', " & _browseGrade & ", " & _oldGrade & ", 'O', " & _ctNode & ") "
          End If
        End If
      Else
        If _pageType = "browseCatTreeCP" Then
          sql = "INSERT INTO MemberGradeBrowseDetail VALUES(@memberId, GETDATE(), '" & _pageType & "', " & _browseGrade & " + " & Integer.Parse(_ctNode) & ", " & _oldGrade & ", 'Y', " & _ctNode & ") "
        Else
          sql = "INSERT INTO MemberGradeBrowseDetail VALUES(@memberId, GETDATE(), '" & _pageType & "', " & _browseGrade & ", " & _oldGrade & ", 'Y', " & _ctNode & ") "
        End If
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
