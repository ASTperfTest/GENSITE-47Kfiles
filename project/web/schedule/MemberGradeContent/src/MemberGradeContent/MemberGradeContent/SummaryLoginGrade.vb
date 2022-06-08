Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration

Public Class SummaryLoginGrade

  Dim _connStr As String  
  Dim _sql As String

  Public Sub New(ByVal connstr As String)

    _connStr = connstr    
    _sql = "SELECT SUM(loginInterGrade) + SUM(loginIQGrade) + SUM(loginRaiseGrade) + SUM(loginAdditional) FROM MemberGradeLogin WHERE memberId = @memberId "

  End Sub

  Public Function HandleSummaryLoginGrade(ByVal mytable As DataTable) As Boolean

    Dim dr As DataRow
    Dim memberId As String = ""
    Dim total As Integer = 0

    Try
    For i As Integer = 0 To mytable.Rows.Count - 1

      dr = mytable.Rows.Item(i)
      memberId = dr.Item("memberId")
      total = GetTotalBySql(_connStr, _sql, memberId)

      '---檢查MemberGradeContent Table中,有沒有此memberId---
      If CheckTableWithContent(_connStr, "MemberGradeSummary", "memberId", memberId) Then
        UpdateTableWithSql(_connStr, "UPDATE MemberGradeSummary SET loginTotal = " & total & " WHERE memberId = '" & memberId & "'")
      Else
        UpdateTableWithSql(_connStr, "INSERT INTO MemberGradeSummary(memberId, loginTotal) VALUES('" & memberId & "', " & total & ")")
      End If      

    Next

    Catch ex As Exception
      SaveMemberGradeLog(_connStr, "exception", ex.ToString)
      Return False
    End Try
    Return True

  End Function

End Class
