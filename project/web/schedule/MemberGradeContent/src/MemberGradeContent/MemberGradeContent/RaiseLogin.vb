Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration

Public Class RaiseLogin

  Dim connstr As String
  Dim beanconnstr As String

  Dim _theStartDate As Date
  Dim _theEndDate As Date

  Public Sub New(ByVal _connstr As String, ByVal _beanconnstr As String)

    _theStartDate = DateAdd("d", -1, Today.ToShortDateString)
    _theEndDate = DateAdd("d", -1, Today.ToShortDateString)

    connstr = _connstr
    beanconnstr = _beanconnstr

  End Sub

  Public Function HandleRaiseLogin() As Boolean

    Dim mytable As DataTable = GetRaiseMemberId(beanconnstr, "LBG_ACTIVITY", "TODAY", _theStartDate.ToShortDateString, _theEndDate.ToShortDateString)

    Dim dr As DataRow
    Dim userId As String = ""
    Dim loginday As String = ""
    Dim logincount As Integer = 0
    Dim point As Integer = 0

    Try
      For i As Integer = 0 To mytable.Rows.Count - 1

        dr = mytable.Rows.Item(i)
        userId = dr.Item("USERID")
        loginday = dr.Item("TODAY")
        logincount = CInt(dr.Item("LOGINCOUNT"))
        point = 0

        '---取出相對應應給的點數---
        point = GetRaiseLoginPoint(connstr, logincount)

        '---更新會員點數---
        UpdateRaiseLoginOnPoint(connstr, userId, point, logincount, loginday)

      Next
    Catch ex As Exception
      Return False
    End Try
    Return True

  End Function

End Class
