Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration

Public Class RaiseGameGrade

  Dim _connStr As String
  Dim _connStrBean As String
  Dim _startDate As String
  Dim _endDate As String

  Dim sql As String
  Dim myConnection As SqlConnection
  Dim myCommand As SqlCommand
  Dim myReader As SqlDataReader

  Public Sub New(ByVal connstr As String, ByVal connstrbean As String, ByVal startdate As String, ByVal enddate As String)

    _connStr = connstr
    _connStrBean = connstrbean
    _startDate = startdate
    _endDate = enddate

  End Sub

  Public Function HandleRaiseGameGrade() As Boolean

    Dim mytable As New DataTable

    '---在活動期間至少有登入一次IQ大挑戰---
    myConnection = New SqlConnection(_connStrBean)
    Try
      sql = "SELECT g.USERID, g.LASTACTION, r.SCORE FROM LBG_GUEST AS g INNER JOIN "
      sql &= "LBG_SCORE_RANKING AS r ON g.UID = r.UID AND g.LASTACTION <= @enddate "
      myConnection.Open()
      myCommand = New SqlCommand(sql, myConnection)
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

        memberId = dr.Item("USERID")
        memberNewGrade = 0
        memberOldGrade = dr.Item("SCORE")
        memberGrade = 0
        memberPoint = 0

        '---取出原有的分數---
        tempStr = SelectTableWithContent(_connStr, "MemberGradeContent", "contentRaiseGrade", "memberId", memberId)
        If tempStr = "" Then
          memberNewGrade = 0
        Else
          memberNewGrade = Integer.Parse(tempStr)
        End If

        '---比較分數高低---
        If memberNewGrade > memberOldGrade Then
          memberPoint = GetRaisePointByGrade(memberNewGrade)
          memberGrade = memberNewGrade
        Else
          memberPoint = GetRaisePointByGrade(memberOldGrade)
          memberGrade = memberOldGrade
        End If

        '---檢查MemberGradeContent Table中,有沒有此memberId---
        If CheckTableWithContent(_connStr, "MemberGradeContent", "memberId", memberId) Then
          '---update---
          UpdateTableWithSql(_connStr, "UPDATE MemberGradeContent SET contentRaisePoint = " & memberPoint & ", contentRaiseGrade = " & memberGrade & " WHERE memberId = '" & memberId & "'")
        Else
          '---insert---
          UpdateTableWithSql(_connStr, "INSERT INTO MemberGradeContent(memberId, contentRaisePoint, contentRaiseGrade) VALUES('" & memberId & "'," & memberPoint & "," & memberGrade & ")")
        End If

      Next
    Catch ex As Exception
      SaveMemberGradeLog(_connStr, "exception", ex.ToString)
      Return False
    End Try
    Return True

  End Function

  Function GetRaisePointByGrade(ByVal Grade As Integer) As Integer

    Dim code As String = ""
    Dim str As String = ""

    If Grade >= 1 And Grade < 3500 Then code = "st_521"
    If Grade >= 3500 And Grade < 7500 Then code = "st_522"
    If Grade >= 7500 Then code = "st_523"

    str = SelectTableWithContent(_connStr, "kpi_set_score", "Rank0_1", "Rank0_2", code)
    If str = "" Then
      Return 0
    Else
      Return Integer.Parse(str)
    End If

  End Function

End Class
