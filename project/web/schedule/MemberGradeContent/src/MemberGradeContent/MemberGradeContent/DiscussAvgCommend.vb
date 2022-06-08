Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration

Public Class DiscussAvgCommend

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

  Public Function HandleDiscussAvgGrade() As Boolean

    '---先將有發表討論成功的人取出---
    Dim mytable As New DataTable
    myConnection = New SqlConnection(_connStr)
    Try
      sql = "SELECT DISTINCT CuDTGeneric.iEditor FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem "
      sql &= "WHERE CuDTGeneric.iCTUnit = @discussunit AND KnowledgeForum.Status = 'N' "

      myConnection.Open()
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@discussunit", _discussUnit)      
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
    Dim GradePersonCount As Integer = 0
    Dim GradeCount As Double = 0.0

    Try
      For i As Integer = 0 To mytable.Rows.Count - 1

        dr = mytable.Rows.Item(i)
        memberId = dr.Item("iEditor")
        GradePersonCount = 0
        GradeCount = 0.0

        myConnection = New SqlConnection(_connStr)
        Try
          sql = "SELECT SUM(CAST(KnowledgeForum.GradeCount AS float) / CAST(KnowledgeForum.GradePersonCount AS float)) AS GradeCount, "
          sql &= "COUNT(CuDTGeneric.iCUItem) AS GradePersonCount "
          sql &= "FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem "
          sql &= "WHERE (CuDTGeneric.iCTUnit = @discussunit) AND (KnowledgeForum.Status = 'N') AND (CuDTGeneric.iEditor = @memberid) "
          sql &= "AND (KnowledgeForum.GradePersonCount <> 0)"

          myConnection.Open()
          myCommand = New SqlCommand(sql, myConnection)
          myCommand.Parameters.AddWithValue("discussunit", _discussUnit)
          myCommand.Parameters.AddWithValue("memberid", memberId)
          myReader = myCommand.ExecuteReader
          If myReader.Read Then
            GradePersonCount = Integer.Parse(myReader("GradePersonCount"))
            GradeCount = Double.Parse(myReader("GradeCount"))
          End If
          If Not myReader.IsClosed Then myReader.Close()
          myCommand.Dispose()
        Catch ex As Exception
        Finally
          If myConnection.State = ConnectionState.Open Then myConnection.Close()
        End Try

        UpdateMemberGradeContent(memberId, GradePersonCount, GradeCount)

      Next

    Catch ex As Exception
      SaveMemberGradeLog(_connStr, "exception", ex.ToString)
      Return False
    End Try
    Return True

  End Function


  Sub UpdateMemberGradeContent(ByVal memberId As String, ByVal GradePersonCount As Integer, ByVal GradeCount As Double)

    Dim commendAvg As Double = 0.0
    Dim commendPoint As Integer = 0

    If GradePersonCount <> 0 Then
      commendAvg = FormatNumber(GradeCount / GradePersonCount, 2)
    End If

    commendPoint = GetCommendPointByGrade(commendAvg)

    '---檢查MemberGradeContent Table中,有沒有此memberId---
    If CheckTableWithContent(_connStr, "MemberGradeContent", "memberId", memberId) Then
      '---update---
      UpdateTableWithSql(_connStr, "UPDATE MemberGradeContent SET contentCommendPoint = " & commendPoint & ", contentCommendGrade = " & commendAvg & _
                                   ", SumGradePersonCount = " & GradePersonCount & ", SumGradeCount = " & GradeCount & " WHERE memberId = '" & memberId & "'")
    Else
      '---insert---
      UpdateTableWithSql(_connStr, "INSERT INTO MemberGradeContent(memberId, contentCommendPoint, contentCommendGrade, SumGradePersonCount, SumGradeCount) " & _
                                   "VALUES('" & memberId & "'," & commendPoint & "," & commendAvg & "," & GradePersonCount & "," & GradeCount & ")")
    End If

  End Sub

  Function GetCommendPointByGrade(ByVal commendAvg As Double) As Integer

    Dim code As String = ""
    Dim str As String = ""

    If commendAvg < 2 Then code = "st_531"
    If commendAvg < 4.0 And commendAvg >= 2.0 Then code = "st_532"
    If commendAvg >= 4 Then code = "st_533"

    str = SelectTableWithContent(_connStr, "kpi_set_score", "Rank0_1", "Rank0_2", code)
    If str = "" Then
      Return 0
    Else
      Return Integer.Parse(str)
    End If

  End Function

End Class
