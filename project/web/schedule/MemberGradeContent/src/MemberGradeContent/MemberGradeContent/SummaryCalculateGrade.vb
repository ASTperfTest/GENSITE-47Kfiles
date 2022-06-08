Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration

Public Class SummaryCalculateGrade

  Dim _connStr As String

  Dim sql As String
  Dim myConnection As SqlConnection
  Dim myCommand As SqlCommand
  Dim myReader As SqlDataReader

  Public Sub New(ByVal connstr As String)

    _connStr = connstr

  End Sub

  Public Function HandleSummaryCalculateGrade(ByVal mytable As DataTable) As Boolean

    Dim dr As DataRow
    Dim memberId As String = ""
    Dim browseTotal As Integer = 0
    Dim loginTotal As Integer = 0
    Dim shareTotal As Integer = 0
    Dim contentTotal As Integer = 0
    Dim additionalTotal As Integer = 0
    Dim total As Integer = 0
    Try
      For i As Integer = 0 To mytable.Rows.Count - 1

        dr = mytable.Rows.Item(i)
        memberId = dr.Item("memberId")
        browseTotal = 0
        loginTotal = 0
        shareTotal = 0
        contentTotal = 0
        additionalTotal = 0
        total = 0

        myConnection = New SqlConnection(_connStr)
        Try
          sql = "SELECT browseTotal, loginTotal, shareTotal, contentTotal, additionalTotal "
          sql &= "FROM MemberGradeSummary WHERE memberId = @memberId "

          myConnection.Open()
          myCommand = New SqlCommand(sql, myConnection)
          myCommand.Parameters.AddWithValue("@memberId", memberId)
          myReader = myCommand.ExecuteReader
          If myReader.Read Then
            browseTotal = Integer.Parse(myReader("browseTotal"))
            loginTotal = Integer.Parse(myReader("loginTotal"))
            shareTotal = Integer.Parse(myReader("shareTotal"))
            contentTotal = Integer.Parse(myReader("contentTotal"))
            additionalTotal = Integer.Parse(myReader("additionalTotal"))
          End If
        Catch ex As Exception
        Finally
          If myConnection.State = ConnectionState.Open Then myConnection.Close()
        End Try

        total = CalulateTotalGrade(browseTotal, loginTotal, shareTotal, contentTotal, additionalTotal)
        UpdateTableWithSql(_connStr, "UPDATE MemberGradeSummary SET calculateTotal = " & total & " WHERE memberId = '" & memberId & "'")

      Next
    Catch ex As Exception
      SaveMemberGradeLog(_connStr, "exception", ex.ToString)
      Return False
    End Try
    Return True

  End Function

  Function CalulateTotalGrade(ByVal browseTotal As Integer, ByVal loginTotal As Integer, ByVal shareTotal As Integer, _
                              ByVal contentTotal As Integer, ByVal additionalTotal As Integer) As Integer

    Dim browseResult As Integer = GetPercentByCode("st_111") * browseTotal
    Dim loginResult As Integer = GetPercentByCode("st_112") * loginTotal
    Dim shareResult As Integer = GetPercentByCode("st_113") * shareTotal
    Dim contentResult As Integer = GetPercentByCode("st_114") * contentTotal
    Dim additionalResult As Integer = GetPercentByCode("st_115") * additionalTotal

    Return (browseResult + loginResult + shareResult + contentResult + additionalResult) / 100

  End Function

  Function GetPercentByCode(ByVal code As String) As Integer

    Dim percent As Integer = 0

    myConnection = New SqlConnection(_connStr)
    Try
      myConnection.Open()

      sql = "SELECT Rank0_1 FROM kpi_set_score WHERE Rank0_2 = @code"
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@code", code)
      myReader = myCommand.ExecuteReader
      If myReader.Read Then
        percent = Integer.Parse(myReader("Rank0_1"))
      End If
      myCommand.Dispose()
    Catch ex As Exception
      SaveMemberGradeLog(_connStr, "exception", ex.ToString)
      Return 0
    Finally
      If myConnection.State = ConnectionState.Open Then myConnection.Close()
    End Try
    Return percent

  End Function

End Class
