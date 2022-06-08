Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration

Public Class QuestionAvgDiscuss

  Dim _connStr As String  
  Dim _questionUnit As String
  Dim _discussUnit As String

  Dim sql As String
  Dim myConnection As SqlConnection
  Dim myCommand As SqlCommand
  Dim myReader As SqlDataReader

  Public Sub New(ByVal connstr As String, ByVal startdate As String, ByVal enddate As String, ByVal questionUnit As String, ByVal discussUnit As String)

    _connStr = connstr
    _questionUnit = questionUnit
    _discussUnit = discussUnit

  End Sub

  Public Function HandleQuestionAvgDiscuss() As Boolean

    '---先將有發表問題成功的人取出---
    Dim mytable As New DataTable

    myConnection = New SqlConnection(_connStr)
    Try
      sql = "SELECT DISTINCT CuDTGeneric.iEditor FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem "
      sql &= "WHERE CuDTGeneric.iCTUnit = @questionunit AND KnowledgeForum.Status = 'N' "

      myConnection.Open()
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@questionunit", _questionUnit)      
      myReader = myCommand.ExecuteReader()
      mytable.Load(myReader)
      If Not myReader.IsClosed() Then myReader.Close()
      myCommand.Dispose()
    Catch ex As Exception
      Return False
    Finally
      If myConnection.State = ConnectionState.Open Then myConnection.Close()      
    End Try

    Dim dr As DataRow
    Dim memberId As String = ""
    Dim discussCount As Integer = 0
    Dim questionCount As Integer = 0
    Try
      For i As Integer = 0 To mytable.Rows.Count - 1

        dr = mytable.Rows.Item(i)

        memberId = dr.Item("iEditor")
        discussCount = 0
        questionCount = 0

        myConnection = New SqlConnection(_connStr)
        Try
          sql = "SELECT SUM(DiscussCount) AS DiscussCount, COUNT(iCUItem) as questionCount "
          sql &= "FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem "
          sql &= "WHERE iEditor = @memberid AND CuDTGeneric.iCTUnit = @questionunit AND KnowledgeForum.Status = 'N' "

          myConnection.Open()
          myCommand = New SqlCommand(sql, myConnection)
          myCommand.Parameters.AddWithValue("@questionunit", _questionUnit)
          myCommand.Parameters.AddWithValue("@memberid", memberId)
          myReader = myCommand.ExecuteReader
          If myReader.Read Then
            discussCount = Integer.Parse(myReader("DiscussCount"))
            questionCount = Integer.Parse(myReader("questionCount"))
          End If
        Catch ex As Exception
          Return False
        Finally
          If myConnection.State = ConnectionState.Open Then
            myConnection.Close()
          End If
        End Try

        UpdateMemberGradeContent(memberId, discussCount, questionCount)

      Next
    Catch ex As Exception
      SaveMemberGradeLog(_connStr, "exception", ex.ToString)
      Return False
    End Try
    Return True
  End Function

  Sub UpdateMemberGradeContent(ByVal memberId As String, ByVal discussCount As Integer, ByVal questionCount As Double)

    Dim discussAvg As Double = 0.0
    Dim discussPoint As Integer = 0

    If questionCount <> 0 Then
      discussAvg = FormatNumber(discussCount / questionCount, 2)
    End If

    discussPoint = GetQuestionPointByGrade(discussAvg)

    '---檢查MemberGradeContent Table中,有沒有此memberId---
    If CheckTableWithContent(_connStr, "MemberGradeContent", "memberId", memberId) Then
      '---update---
      UpdateTableWithSql(_connStr, "UPDATE MemberGradeContent SET contentDiscussPoint = " & discussPoint & ", contentDiscussGrade = " & discussAvg & _
                                   ", sumDiscussCount = " & discussCount & ", sumDiscussQuestionCount = " & questionCount & " WHERE memberId = '" & memberId & "'")
    Else
      '---insert---
      UpdateTableWithSql(_connStr, "INSERT INTO MemberGradeContent(memberId, contentDiscussPoint, contentDiscussGrade, sumDiscussCount, sumDiscussQuestionCount) " & _
                                   "VALUES('" & memberId & "'," & discussPoint & "," & discussAvg & "," & discussCount & "," & questionCount & ")")
    End If

  End Sub

  Function GetQuestionPointByGrade(ByVal questionAvg As Double) As Integer

    Dim code As String = ""
    Dim str As String = ""


    If questionAvg < 2 Then code = "st_541"

    If questionAvg < 4.0 And questionAvg >= 2.0 Then code = "st_542"

    If questionAvg >= 4 Then code = "st_543"

    str = SelectTableWithContent(_connStr, "kpi_set_score", "Rank0_1", "Rank0_2", code)
    If str = "" Then
      Return 0
    Else
      Return Integer.Parse(str)
    End If

  End Function

End Class
