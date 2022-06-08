Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration

Public Class LoginAdditional

  Dim _connStr As String
  Dim _theStartDate As Date
  Dim _theEndDate As Date

  Dim tempdate As String

  Public Sub New(ByVal connstr As String)

    _connStr = connstr

    '---2008/09/29---2008/10/06---2008/10/13---2008/10/20---2008/10/27---
    'tempdate = "2008/09/29"
    'tempdate = "2008/10/06"
    'tempdate = "2008/10/13"
    'tempdate = "2008/10/20"
    'tempdate = "2008/10/27"

    '_theStartDate = DateAdd("d", -7, tempdate)
    '_theEndDate = DateAdd("d", -1, tempdate)

    tempdate = Today.ToShortDateString

    _theStartDate = DateAdd("d", -7, Today.ToShortDateString)
    _theEndDate = DateAdd("d", -1, Today.ToShortDateString)

  End Sub

  Public Function HandleLoginAdditional() As Boolean

    Dim mytable As DataTable = GetMemberId(_connStr, "MemberGradeLogin", "loginDate", _theStartDate.ToShortDateString, _theEndDate.ToShortDateString)

    Dim dr As DataRow
    Dim memberId As String = ""
    Dim dayCount As Integer = 0
    Dim aList As ArrayList
    Dim point As Integer = 0

    Try
      For i As Integer = 0 To mytable.Rows.Count - 1

        dr = mytable.Rows.Item(i)
        memberId = dr.Item("memberId")
        dayCount = 0
        aList = Nothing
        point = 0

        '---檢查此會員在這七天內有哪幾天登入---
        aList = CheckMemberLoginDate(memberId, _theStartDate.ToShortDateString, _theEndDate.ToShortDateString)

        '---檢查此會員在登入的天數中是否有瀏覽行為或分享行為---
        dayCount = CheckMemberHaveBrowseOrShare(memberId, aList)

        '---取出相對應應給的點數---
        point = GetLoginAdditionalPoint(dayCount)

        '---更新會員點數---
        UpdateLoginAdditionalOnPoint(memberId, point)

      Next
    Catch ex As Exception
      SaveMemberGradeLog(_connStr, "exception", ex.ToString)
      Return False
    End Try
    Return True

  End Function

  Function CheckMemberLoginDate(ByVal memberId As String, ByVal startDate As String, ByVal endDate As String) As ArrayList

    Dim alist As New ArrayList
    Dim mytable As DataTable
    Dim dr As DataRow
    Dim sql As String = ""
    sql = "SELECT CONVERT(varchar, loginDate, 111) AS loginDate FROM MemberGradeLogin "
    sql &= "WHERE (loginDate BETWEEN '" & startDate & "' AND '" & endDate & " 23:59:59') AND (memberId = '" & memberId & "') ORDER BY loginDate"

    mytable = GetDataTableBySql(_connStr, sql)

    For i As Integer = 0 To mytable.Rows.Count - 1
      dr = mytable.Rows.Item(i)
      alist.Add(dr.Item("loginDate"))
    Next
    Return alist

  End Function

  Function CheckMemberHaveBrowseOrShare(ByVal memberId As String, ByVal aList As ArrayList) As Integer

    Dim dayCount As Integer = 0

    If aList Is Nothing Then Return dayCount

    If aList.Count = 0 Then Return dayCount

    For Each item As String In aList
      If CheckMemberHaveBrowse(memberId, item) Then
        dayCount += 1
      ElseIf CheckMemberHaveShare(memberId, item) Then
        dayCount += 1
      End If
    Next
    Return dayCount

  End Function

  Function CheckMemberHaveBrowse(ByVal memberId As String, ByVal theDate As String) As Boolean

    Dim sql As String = ""

    sql = "SELECT ISNULL(SUM(browseInterCP) + SUM(browseTopicCP) + SUM(browseCatTreeCP) + SUM(browseQuestionCP) + SUM(browseDiscussLP) + SUM(browseDiscussCP), 0) "
    sql &= "FROM MemberGradeBrowse WHERE memberId = @memberId AND CONVERT(varchar, browseDate, 111) = @theDate "

    Dim total As Integer = GetTotalBySql(_connStr, sql, memberId, theDate)
    Dim limitScore As Integer = 0
    limitScore = GetLoginLimit("browse")

    If total >= limitScore Then
      Return True
    Else
      Return False
    End If

  End Function

  Function CheckMemberHaveShare(ByVal memberId As String, ByVal theDate As String) As Boolean

    Dim sql As String = ""

    sql = "SELECT ISNULL(SUM(shareAsk) + SUM(shareDiscuss) + SUM(shareCommend) + SUM(shareOpinion) + SUM(shareSuggest) + SUM(shareExplain) + SUM(shareVote), 0) "
    sql &= "FROM MemberGradeShare WHERE memberId = @memberId AND CONVERT(varchar, shareDate, 111) = @theDate "

    Dim total As Integer = GetTotalBySql(_connStr, sql, memberId, theDate)
    Dim limitScore As Integer = 0
    limitScore = GetLoginLimit("share")
    If total >= limitScore Then
      Return True
    Else
      Return False
    End If

  End Function

  Function GetLoginAdditionalPoint(ByVal dayCount As Integer) As Integer

    Dim code As String = ""
    Dim str As String = ""

    If dayCount < 2 And dayCount > 0 Then code = "st_321"

    If dayCount < 4.0 And dayCount >= 2.0 Then code = "st_322"

    If dayCount >= 4 Then code = "st_323"

    str = SelectTableWithContent(_connStr, "kpi_set_score", "Rank0_1", "Rank0_2", code)
    If str = "" Then
      Return 0
    Else
      Return Integer.Parse(str)
    End If

  End Function

  Function GetLoginLimit(ByVal type As String) As Integer

    Dim code As String = ""
    Dim str As String = ""

    If type = "browse" Then code = "st_331"
    If type = "share" Then code = "st_332"
    
    str = SelectTableWithContent(_connStr, "kpi_set_score", "Rank0_1", "Rank0_2", code)
    If str = "" Then
      Return 0
    Else
      Return Integer.Parse(str)
    End If

  End Function

  Sub UpdateLoginAdditionalOnPoint(ByVal memberId As String, ByVal point As Integer)

    '---檢查MemberGradeLogin Table中,在當日有沒有此memberId---
    If CheckMemberGradeLoginWithloginDate(_connStr, memberId, tempdate) Then
      UpdateTableWithSql(_connStr, "UPDATE MemberGradeLogin SET loginAdditional = " & point & " WHERE memberId = '" & memberId & "' AND CONVERT(varchar, loginDate, 111) = CONVERT(varchar, '" & tempdate & "', 111)")
    Else
      UpdateTableWithSql(_connStr, "INSERT INTO MemberGradeLogin(memberId, loginDate, loginAdditional) VALUES('" & memberId & "', '" & tempdate & "', " & point & ")")
    End If

  End Sub

End Class
