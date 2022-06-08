Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration

Public Class QuestionAvgBrowse

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

  Public Function HandleQuestionAvgBrowse() As Boolean

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
    Dim browseCount As Integer = 0
    Dim questionCount As Integer = 0

    Try
      For i As Integer = 0 To mytable.Rows.Count - 1

        dr = mytable.Rows.Item(i)

        memberId = dr.Item("iEditor")
        browseCount = 0
        questionCount = 0

        myConnection = New SqlConnection(_connStr)
        Try
          sql = "SELECT SUM(BrowseCount) AS browseCount, COUNT(iCUItem) AS questionCount "
          sql &= "FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem "
          sql &= "WHERE iEditor = @memberid AND CuDTGeneric.iCTUnit = @questionunit AND KnowledgeForum.Status = 'N' "

          myConnection.Open()
          myCommand = New SqlCommand(sql, myConnection)
          myCommand.Parameters.AddWithValue("@questionunit", _questionUnit)
          myCommand.Parameters.AddWithValue("@memberid", memberId)
          myReader = myCommand.ExecuteReader
          If myReader.Read Then
            browseCount = Integer.Parse(myReader("browseCount"))
            questionCount = Integer.Parse(myReader("questionCount"))
          End If
        Catch ex As Exception
          Return False
        Finally
          If myConnection.State = ConnectionState.Open Then
            myConnection.Close()
          End If
        End Try

        UpdateMemberGradeContent(memberId, browseCount, questionCount)

      Next

    Catch ex As Exception
      SaveMemberGradeLog(_connStr, "exception", ex.ToString)
      Return False
    End Try
    Return True
  End Function

  Sub UpdateMemberGradeContent(ByVal memberId As String, ByVal browseCount As Integer, ByVal questionCount As Double)

    Dim browseAvg As Double = 0.0
    Dim browsePoint As Integer = 0

    If questionCount <> 0 Then
      browseAvg = FormatNumber(browseCount / questionCount, 2)
    End If

    browsePoint = GetBrowsePointByGrade(browseAvg)

    '---檢查MemberGradeContent Table中,有沒有此memberId---
    If CheckTableWithContent(_connStr, "MemberGradeContent", "memberId", memberId) Then
      '---update---
      UpdateTableWithSql(_connStr, "UPDATE MemberGradeContent SET contentBrowsePoint = " & browsePoint & ", contentBrowseGrade = " & browseAvg & _
                                   ", sumBrowseCount = " & browseCount & ", sumBrowseQuestionCount = " & questionCount & " WHERE memberId = '" & memberId & "'")
    Else
      '---insert---
      UpdateTableWithSql(_connStr, "INSERT INTO MemberGradeContent(memberId, contentBrowsePoint, contentBrowseGrade, sumBrowseCount, sumBrowseQuestionCount) " & _
                                   "VALUES('" & memberId & "'," & browsePoint & "," & browseAvg & "," & browseCount & "," & questionCount & ")")
    End If

  End Sub

  Function GetBrowsePointByGrade(ByVal browseAvg As Double) As Integer

    Dim code As String = ""
    Dim str As String = ""


    If browseAvg > 0 And browseAvg < 150 Then code = "st_551"

    If browseAvg < 400 And browseAvg >= 150 Then code = "st_552"

    If browseAvg >= 400 Then code = "st_553"

    str = SelectTableWithContent(_connStr, "kpi_set_score", "Rank0_1", "Rank0_2", code)
    If str = "" Then
      Return 0
    Else
      Return Integer.Parse(str)
    End If

  End Function

End Class
