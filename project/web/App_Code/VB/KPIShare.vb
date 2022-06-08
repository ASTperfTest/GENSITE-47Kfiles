Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Public Class KPIShare

  Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
  Dim sql As String = ""
  Dim myConnection As SqlConnection
  Dim myCommand As SqlCommand
  Dim myReader As SqlDataReader  

  Dim _memberId As String
  Dim _ctNode As String
  Dim _pageType As String
  Dim _shareGrade As Integer
  Dim checkFlag As Boolean

  Dim _shareCode As String
  Dim _shareCode2 As String
  Dim _limitCode As String
  Dim _limitScore As Integer
  Dim _oldGrade As Integer
  Dim _isOverMax As Boolean
  Dim _isOverMaxFlag As Boolean

  Public Sub New(ByVal memberId As String, ByVal pageType As String, ByVal ctNode As String)

    If memberId = "" Then Exit Sub

    _memberId = memberId
    _ctNode = ctNode
    _pageType = pageType

    _shareGrade = 0
    _shareCode = "st_4"
    _limitCode = "st_411"

    _oldGrade = 0
    _limitScore = 0
    _isOverMax = False
    _isOverMaxFlag = False

    '知識家－發問     st_412
    '知識家－討論     st_413
    '知識家－(給)評價 st_414
    '知識家－發表意見 st_415
    '小百科－建議詞彙 st_416
    '小百科－補充解釋 st_417
        '推薦－態度投票   st_418
        '主題館－(給)評價 st_421

    If _pageType = "shareAsk" Then _shareCode2 = "st_412"
    If _pageType = "shareDiscuss" Then _shareCode2 = "st_413"
    If _pageType = "shareCommend" Then _shareCode2 = "st_414"
    If _pageType = "shareOpinion" Then _shareCode2 = "st_415"
    If _pageType = "shareSuggest" Then _shareCode2 = "st_416"
    If _pageType = "shareExplain" Then _shareCode2 = "st_417"
    If _pageType = "shareVote" Then _shareCode2 = "st_418"
        If _pageType = "shareJigsaw" Then _shareCode2 = "st_419"
        If _pageType = "shareSubjectCommend" Then _shareCode2 = "st_421"
	
    If _shareCode2 = "" Then Exit Sub

  End Sub

  Public Sub HandleShare()

    '---檢查KPI是否開啟記錄---
    If Not CheckKpiOpen() Then
      Exit Sub
    End If

    '---檢查分享行為是否公開---
    If Not CheckShareIsPublic() Then
      Exit Sub
    End If

    '---檢查傳入的memberId是否有登入當時的資料---
    If Not CheckMemberShareToday() Then
      '---此member沒有今天的記錄---插入一筆新的---
      If Not InsertMemberDefault() Then
        Exit Sub
      End If
    End If

    '---利用pageType來抓值---
    If Not GetShareGradeByPageType() Then
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

  ''' <summary>檢查分享行為是否公開</summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function CheckShareIsPublic() As Boolean

    Dim returnStr As String = ""
    returnStr = Trim(GetID.Get_OneField_ReturnIdStr("kpi_set_ind", "Rank0", _shareCode, "Rank2"))

    If returnStr = "Y" Then
      Return True
    Else
      Return False
    End If

  End Function

  Function CheckMemberShareToday() As Boolean

    checkFlag = False
    myConnection = New SqlConnection(ConnString)
    Try
      sql = "SELECT memberId, shareDate  FROM MemberGradeShare WHERE memberId = @memberId "
      sql &= "AND CONVERT(varchar, shareDate, 111) = CONVERT(varchar, GETDATE(), 111) "
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
            sql = " INSERT INTO MemberGradeShare (memberId, shareDate) VALUES(@memberId, GETDATE()) "
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

  Function GetShareGradeByPageType() As Boolean

    Try
      _shareGrade = GetID.Get_OneField_ReturnIdStr("kpi_set_score", "Rank0_2", _shareCode2, "Rank0_1")
    Catch
      Return False
    End Try

    Return True

  End Function

  Function CheckGradeLimit() As Boolean

    _limitScore = GetID.Get_TwoField("kpi_set_score", "Rank0", _shareCode, "Rank0_2", _limitCode, "Rank0_1")
    Dim totalScore As Integer = 0

    checkFlag = False
    myConnection = New SqlConnection(ConnString)
    Try
      sql = " SELECT shareAsk + shareDiscuss + shareCommend + shareOpinion + shareSuggest + shareExplain + shareVote AS total "
      sql &= " FROM MemberGradeShare WHERE memberId = @memberId AND CONVERT(varchar, shareDate, 111) = CONVERT(varchar, GETDATE(), 111) "

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
      If totalScore + _shareGrade >= _limitScore Then
        If totalScore < _limitScore Then
          _isOverMaxFlag = True
        End If
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

    checkFlag = False
    Dim tempScore As Integer = _limitScore - _oldGrade
    myConnection = New SqlConnection(ConnString)
    Try

      If Not _isOverMax Then
        sql = " UPDATE MemberGradeShare SET " & _pageType & " = " & _pageType & " + " & _shareGrade & " "
      Else
        sql = " UPDATE MemberGradeShare SET " & _pageType & " = " & _pageType & " + " & tempScore & " "
      End If
      sql &= " WHERE memberId = @memberId AND CONVERT(varchar, shareDate, 111) = CONVERT(varchar, GETDATE(), 111) "

      myConnection.Open()
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@memberId", _memberId)
      myCommand.ExecuteNonQuery()
      myCommand.Dispose()

      If _isOverMax Then
        If _isOverMaxFlag Then
          sql = "INSERT INTO MemberGradeShareDetail VALUES(@memberId, GETDATE(), '" & _pageType & "', " & tempScore & ", " & _oldGrade & ", 'O', " & _ctNode & ",'0') "
        Else
          sql = "INSERT INTO MemberGradeShareDetail VALUES(@memberId, GETDATE(), '" & _pageType & "', " & _shareGrade & ", " & _oldGrade & ", 'O'," & _ctNode & ",'0') "
        End If
      Else
        sql = "INSERT INTO MemberGradeShareDetail VALUES(@memberId, GETDATE(), '" & _pageType & "', " & _shareGrade & ", " & _oldGrade & ", 'Y'," & _ctNode & ",'0') "
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
