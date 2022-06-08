Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration

Public Class IQGameGrade

  Dim _connStr As String  
  Dim _iqUnit As String
  Dim _startDate As String
  Dim _endDate As String

  Dim sql As String
  Dim myConnection As SqlConnection
  Dim myCommand As SqlCommand
  Dim myReader As SqlDataReader

  Public Sub New(ByVal connstr As String, ByVal startdate As String, ByVal enddate As String, ByVal iqunit As String)

    _connStr = connstr
    _iqUnit = iqunit
    _startDate = startdate
    _endDate = enddate

  End Sub

  Public Function HandleIQGameGrade() As Boolean

    Dim mytable As New DataTable

    '---在活動期間至少有登入一次IQ大挑戰---
    myConnection = New SqlConnection(_connStr)
    Try
      sql = "SELECT flashGame.name as name FROM flashGame INNER JOIN CuDTGeneric "
      sql &= "ON flashGame.gicuitem = CuDTGeneric.iCUItem AND CuDTGeneric.iCTUnit = @iqgameunit "
      sql &= "WHERE CuDTGeneric.xPostDate >= @startdate AND CuDTGeneric.xPostDate <= @enddate "
      sql &= "GROUP BY flashGame.name "
      myConnection.Open()
      myCommand = New SqlCommand(sql, myConnection)      
      myCommand.Parameters.AddWithValue("@iqgameunit", _iqUnit)
      myCommand.Parameters.AddWithValue("@startdate", _startDate)
      myCommand.Parameters.AddWithValue("@enddate", _endDate)
      myReader = myCommand.ExecuteReader()
      mytable.Load(myReader)
      If Not myReader.IsClosed() Then myReader.Close()
      myCommand.Dispose()
    Catch ex As Exception
      SaveMemberGradeLog(_connStr, "exception", ex.ToString)
      Return False
    Finally
      If myConnection.State = ConnectionState.Open Then myConnection.Close()
    End Try

    Dim dr As DataRow
    Dim memberId As String = ""
    Dim memberNewGrade As Integer = 0
    Dim memberOldGrade As Integer = 0
    Dim memberGrade As Integer = 0
    Dim memberPoint As Integer = 0
    Dim tempStr As String = ""

    Try
      For i As Integer = 0 To mytable.Rows.Count - 1

        dr = mytable.Rows.Item(i)

        memberId = dr.Item("name")
        memberNewGrade = 0
        memberOldGrade = GetMaxGrade(memberId)  '---取出原有的分數內的最高分---
        memberGrade = 0
        memberPoint = 0

        '---取出原有的分數---
        tempStr = SelectTableWithContent(_connStr, "MemberGradeContent", "contentIQGrade", "memberId", memberId)
        If tempStr = "" Then
          memberNewGrade = 0
        Else
          memberNewGrade = Integer.Parse(tempStr)
        End If

        '---比較分數高低---
        If memberNewGrade > memberOldGrade Then
          memberPoint = GetIQPointByGrade(memberNewGrade)
          memberGrade = memberNewGrade
        Else
          memberPoint = GetIQPointByGrade(memberOldGrade)
          memberGrade = memberOldGrade
        End If

        '---檢查MemberGradeContent Table中,有沒有此memberId---
        If CheckTableWithContent(_connStr, "MemberGradeContent", "memberId", memberId) Then
          '---update---
          UpdateTableWithSql(_connStr, "UPDATE MemberGradeContent SET contentIQPoint = " & memberPoint & ", contentIQGrade = " & memberGrade & " WHERE memberId = '" & memberId & "'")
        Else
          '---insert---
          UpdateTableWithSql(_connStr, "INSERT INTO MemberGradeContent(memberId, contentIQPoint, contentIQGrade) VALUES('" & memberId & "'," & memberPoint & "," & memberGrade & ")")
        End If

      Next
    Catch ex As Exception
      SaveMemberGradeLog(_connStr, "exception", ex.ToString)
      Return False
    End Try
    Return True

  End Function

  Function GetIQPointByGrade(ByVal Grade As Integer) As Integer

    Dim Sql As String = "SELECT Rank0_1 FROM kpi_set_score WHERE Rank0_2= '{0}'"

    If Grade >= 0 And Grade < 2000 Then Sql = Sql.Replace("{0}", "st_511")
    If Grade >= 2000 And Grade < 3000 Then Sql = Sql.Replace("{0}", "st_512")
    If Grade >= 3000 And Grade < 5000 Then Sql = Sql.Replace("{0}", "st_513")    
    If Grade >= 5000 And Grade < 10000 Then Sql = Sql.Replace("{0}", "st_514")  
    If Grade >= 10000 And Grade < 20000 Then Sql = Sql.Replace("{0}", "st_515")    
    If Grade >= 20000 Then Sql = Sql.Replace("{0}", "st_516")
    
    Dim da1 As SqlDataAdapter = New SqlDataAdapter(Sql, myConnection)
    Dim dt As DataTable = New DataTable()
    da1.Fill(dt)
    Return Integer.Parse(dt.Rows(0)("Rank0_1"))

  End Function

  Function GetMaxGrade(ByVal memberId) As Integer

    Dim result As Integer = 0
    myConnection = New SqlConnection(_connStr)
    Try
      myConnection.Open()
      sql = "SELECT MAX(money) AS Grade FROM flashGame WHERE name = @name"
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@name", memberId)
      myReader = myCommand.ExecuteReader
      If myReader.Read Then
        If myReader("Grade") IsNot DBNull.Value Then
          result = Integer.Parse(myReader("Grade"))
        End If
      End If
      If Not myReader.IsClosed Then myReader.Close()
      myCommand.Dispose()
    Catch ex As Exception
      SaveMemberGradeLog(_connStr, "exception", ex.ToString)
    Finally
      If myConnection.State = ConnectionState.Open Then myConnection.Close()
    End Try
    Return result

  End Function

End Class
