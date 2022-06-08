Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System.Web.Configuration

Public Class ActivityFilter

  Dim activityid As String
  Dim memberid As String
  Dim icuitem As String

  Public Sub New(ByVal actid As String, ByVal id As String)

    activityid = actid
    memberid = id

  End Sub

  Public Sub New(ByVal actid As String, ByVal id As String, ByVal item As String)

    activityid = actid
    memberid = id
    icuitem = item

  End Sub

  '---利用這function來判斷登入是否要加分---

  Public Function CheckLoginGrade() As Boolean

    Return True

  End Function

  '---利用這function來判斷發問是否要加分---

  Public Function CheckQuestionGrade() As Boolean

    Return True

  End Function

  '---利用這function來判斷評價是否要加分---

  Public Function CheckStarGrade() As Boolean

    Return True

  End Function

  '---利用這個function來判斷目前活動是否進行中---

  Public Function CheckActivity() As Boolean

    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim sqlString As String = ""
    Dim myConnection As SqlConnection = New SqlConnection(ConnString)
    Dim myCommand As SqlCommand
    Dim myReader As SqlDataReader

    Try
      myConnection.Open()
      sqlString = "SELECT * FROM Activity WHERE (GETDATE() BETWEEN ActivityStartTime AND ActivityEndTime) AND ActivityId = @activityid"
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@activityid", activityid)
      myReader = myCommand.ExecuteReader
      If myReader.Read Then
        Return True
      Else
        Return False
      End If
      myReader.Close()
      myCommand.Dispose()
    Catch ex As Exception
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

  End Function
  
  Public Function CheckStartActivity() As Boolean

    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim sqlString As String = ""
    Dim myConnection As SqlConnection = New SqlConnection(ConnString)
    Dim myCommand As SqlCommand
    Dim myReader As SqlDataReader

    Try
      myConnection.Open()
      sqlString = "SELECT * FROM Activity WHERE (GETDATE() > ActivityStartTime ) AND ActivityId = @activityid"
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@activityid", activityid)
      myReader = myCommand.ExecuteReader
      If myReader.Read Then
        Return True
      Else
        Return False
      End If
      myReader.Close()
      myCommand.Dispose()
    Catch ex As Exception
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

  End Function
  
  Public Function IsActivityEnd() As Boolean

    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim sqlString As String = ""
    Dim myConnection As SqlConnection = New SqlConnection(ConnString)
    Dim myCommand As SqlCommand
    Dim myReader As SqlDataReader

    Try
      myConnection.Open()
      sqlString = "SELECT * FROM Activity WHERE (GETDATE() > ActivityEndTime ) AND ActivityId = @activityid"
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@activityid", activityid)
      myReader = myCommand.ExecuteReader
      If myReader.Read Then
        Return True
      Else
        Return False
      End If
      myReader.Close()
      myCommand.Dispose()
    Catch ex As Exception
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

  End Function
  

  '---全年活動 function, 登入的時候加分---
  Public Sub LoginGradeAdd()

    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim sqlString As String = "", newsql As String = ""
    Dim myConnection As SqlConnection = New SqlConnection(ConnString)
    Dim myCommand As SqlCommand
    Dim myReader As SqlDataReader

    Try
      myConnection.Open()
      sqlString = "SELECT LastLoginTime FROM ActivityMemberNew WHERE MemberId = @memberid"
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@memberid", memberid)
      myReader = myCommand.ExecuteReader
      If myReader.Read Then
        If myReader("LastLoginTime") Is DBNull.Value Then
          newsql = "UPDATE ActivityMemberNew Set LoginGrade = LoginGrade + 2, LastLoginTime = GETDATE() WHERE MemberId = @memberid"
        Else
          If DateDiff("d", myReader("LastLoginTime"), Now) = 0 Then
            newsql = "UPDATE ActivityMemberNew Set LastLoginTime = GETDATE() WHERE MemberId = @memberid"
          Else
            newsql = "UPDATE ActivityMemberNew Set LoginGrade = LoginGrade + 2, LastLoginTime = GETDATE() WHERE MemberId = @memberid"
          End If
        End If
      Else
        newsql = "INSERT INTO ActivityMemberNew (MemberId, LoginGrade, LastLoginTime) VALUES (@memberid, 2, GETDATE()) "
      End If
      myReader.Close()
      myCommand.Dispose()

      myCommand = New SqlCommand(newsql, myConnection)
      myCommand.Parameters.AddWithValue("@memberid", memberid)
      myCommand.ExecuteNonQuery()
      myCommand.Dispose()

    Catch ex As Exception

    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

  End Sub

  '---全年活動 function, 發問的時候加分---
  Public Sub QuestionGradeAdd()

    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim sqlString As String = ""
    Dim myConnection As SqlConnection = New SqlConnection(ConnString)
    Dim myCommand As SqlCommand

    Try
      myConnection.Open()
      sqlString = "UPDATE ActivityMemberNew SET QuestionGrade = QuestionGrade + 2 WHERE MemberId = @memberid"
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@memberid", memberid)
      myCommand.ExecuteNonQuery()
      myCommand.Dispose()
    Catch ex As Exception
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

  End Sub

  '---全年活動 function, 刪除發問的時候減分---
  Public Sub QuestionGradeDelete()

    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim sqlString As String = ""
    Dim myConnection As SqlConnection = New SqlConnection(ConnString)
    Dim myCommand As SqlCommand
    Dim myReader As SqlDataReader
    Dim QuestionGrade As Integer = 0

    Try
      myConnection.Open()
      sqlString = "SELECT QuestionGrade FROM ActivityMemberNew WHERE MemberId = @memberid"
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@memberid", memberid)
      myReader = myCommand.ExecuteReader
      If myReader.Read Then
        QuestionGrade = Integer.Parse(myReader("QuestionGrade").ToString)
      Else
        QuestionGrade = 0
      End If
      If QuestionGrade <= 2 Then
        QuestionGrade = 0
      Else
        QuestionGrade -= 2
      End If
      myReader.Close()
      myCommand.Dispose()


      sqlString = "UPDATE ActivityMemberNew SET QuestionGrade = @questiongrade WHERE MemberId = @memberid"
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@questiongrade", QuestionGrade)
      myCommand.Parameters.AddWithValue("@memberid", memberid)
      myCommand.ExecuteNonQuery()

      myCommand.Dispose()

    Catch ex As Exception
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

  End Sub

  '---全年活動 function, 討論的時候加分---
  Public Sub DiscussGradeAdd()

    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim sqlString As String = ""
    Dim myConnection As SqlConnection = New SqlConnection(ConnString)
    Dim myCommand As SqlCommand

    Try
      myConnection.Open()
      sqlString = "UPDATE ActivityMemberNew SET DiscussGrade = DiscussGrade + 2 WHERE MemberId = @memberid"
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@memberid", memberid)
      myCommand.ExecuteNonQuery()
      myCommand.Dispose()
    Catch ex As Exception
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

  End Sub
  
  

  '---全年活動 function, 評分的時候加分---
  Public Sub StarGradeAdd()

    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim sqlString As String = ""
    Dim myConnection As SqlConnection = New SqlConnection(ConnString)
    Dim myCommand As SqlCommand

    Try
      myConnection.Open()
      sqlString = "UPDATE ActivityMemberNew SET StarGrade = StarGrade + 1 WHERE MemberId = @memberid"
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@memberid", memberid)
      myCommand.ExecuteNonQuery()
      myCommand.Dispose()
    Catch ex As Exception
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

  End Sub

  '---個別活動 function, 取得活動的分數---

  Public Function GetActivityGrade() As String

    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim sqlString As String = ""
    Dim myConnection As New SqlConnection(ConnString)
    Dim myCommand As SqlCommand
    Dim myReader As SqlDataReader

    If activityid = "activity1" Then
      Try
        myConnection.Open()
        sqlString = "SELECT (DiscussCheckGrade + QuestionCheckGrade) AS Grade FROM ActivityMemberNew WHERE MemberID = @memberid"
        myCommand = New SqlCommand(sqlString, myConnection)
        myCommand.Parameters.AddWithValue("@memberid", memberid)
        myReader = myCommand.ExecuteReader
        If myReader.Read Then
          Return myReader("Grade").ToString
        Else
          Return "0"
        End If
        myReader.Close()
        myCommand.Dispose()
      Catch ex As Exception
        Return "0"
      Finally
        If myConnection.State = ConnectionState.Open Then
          myConnection.Close()
        End If
      End Try
    Else
      Return "0"
    End If    

  End Function

  '---每個活動所要處理的動作不同, 利用這function來分辨---

  Public Sub ProcessActivity()

    If activityid = "activity1" Then

      Call Activity1()

    Else
      '
    End If

  End Sub

  Private Sub Activity1()

    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim sqlString As String = ""
    Dim myConnection As New SqlConnection(ConnString)
    Dim myCommand As SqlCommand
    Dim myReader As SqlDataReader

    Dim UserExist As Boolean = False
    Dim CurrentGrade As Double = 0
    Dim CurrentGradeFlag As Boolean = False
    Dim ActFlag As Boolean = False

    Try
      myConnection.Open()

      '---檢查table中是否已有這個人的資料---

      sqlString = "SELECT * FROM ActivityMemberNew WHERE MemberId = @memberid"

      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@memberid", memberid)
      myReader = myCommand.ExecuteReader
      If myReader.Read Then
        UserExist = True
      End If
      myReader.Close()
      myCommand.Dispose()

      If Not UserExist Then
        sqlString = "INSERT INTO ActivityMemberNew (MemberId) VALUES (@memberid)"
        myCommand = New SqlCommand(sqlString, myConnection)
        myCommand.Parameters.AddWithValue("@memberid", memberid)
        myCommand.ExecuteNonQuery()
        myCommand.Dispose()
      End If

      sqlString = "SELECT vGroup FROM CuDTGeneric WHERE icuitem = @icuitem"
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@icuitem", icuitem)
      myReader = myCommand.ExecuteReader
      If myReader.Read Then
        If myReader("vGroup") = "A" Then
          ActFlag = True
        End If
      End If
      myReader.Close()
      myCommand.Dispose()

      If ActFlag Then

        '---討論加分---
        sqlString = "UPDATE ActivityMemberNew SET DiscussCheckGrade = DiscussCheckGrade + 2, Activity1 = 1 WHERE MemberId = @memberid"
        myCommand = New SqlCommand(sqlString, myConnection)
        myCommand.Parameters.AddWithValue("@memberid", memberid)
        myCommand.ExecuteNonQuery()
        myCommand.Dispose()
      End If

    Catch ex As Exception

    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

  End Sub
  
  '檢查是否為活動題目 Add by Ivy 2010.4.26
  Public Function CheckAvtivityKnowledge() As Boolean

    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim sqlString As String = ""
    Dim myConnection As SqlConnection = New SqlConnection(ConnString)
    Dim myCommand As SqlCommand
    Dim myReader As SqlDataReader

    Try
      myConnection.Open()
      sqlString = "SELECT * FROM dbo.KnowledgeActivity WHERE CUItemId = @Cuitemid "
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@Cuitemid", icuitem)
      myReader = myCommand.ExecuteReader
      If myReader.Read Then
        Return True
      Else
        Return False
      End If
      myReader.Close()
      myCommand.Dispose()
    Catch ex As Exception
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

  End Function

End Class
