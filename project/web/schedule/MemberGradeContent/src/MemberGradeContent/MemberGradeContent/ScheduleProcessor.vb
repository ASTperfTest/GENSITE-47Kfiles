Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration

Module ScheduleProcessor

  '------------------------------------------------------------------------------------------
  Dim ConnString As String = ConfigurationManager.ConnectionStrings("ConnString").ToString
  Dim BeanConnString As String = ConfigurationManager.ConnectionStrings("BeanConnString").ToString
  Dim Sql As String = ""
  Dim myConnection As SqlConnection
  Dim myCommand As SqlCommand
  Dim myReader As SqlDataReader
  '------------------------------------------------------------------------------------------

  Sub Main()

    '------------------------------------------------------------------------------------------
    Dim QuestionUnit As String = ConfigurationManager.AppSettings("QustionUnit")
    Dim discussUnit As String = ConfigurationManager.AppSettings("DiscussUnit")
    Dim IQGameUnit As String = ConfigurationManager.AppSettings("IQGameUnit")
    '------------------------------------------------------------------------------------------
    Dim startDate As String = ""                '---�Ƶ{�_�l��---
    Dim endDate As String = ""                  '---�Ƶ{������---
    Dim contentFlag As Boolean = False          '---�O�_���椺�e�[�Ȫ��Ƶ{---
    Dim recordFlag As Boolean = False           '---�O�_����ӤH�O�����Ƶ{---
    Dim loginAdditionalFlag As Boolean = False  '---�O�_����n�J�B�~���I---
    '------------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------------
    SaveMemberGradeLog(ConnString, "exe", "Start to execute MemberGradeContent")

    '---�ˬdKPI�O�_�}��---
    If CheckKpiOpen() <> "Y" Then
      SaveMemberGradeLog(ConnString, "exception", "kpi is not open")
      Exit Sub
    Else
      SaveMemberGradeLog(ConnString, "exe", "kpi open")
    End If
    '------------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------------
    '---Ū�Xstartdate and enddate---
    If Not GetKpiStartDateAndEndDate(startDate, endDate) Then
      SaveMemberGradeLog(ConnString, "exception", "can not get startdate and enddate")
      Exit Sub
    Else
      SaveMemberGradeLog(ConnString, "exe", "get startdate and enddate")
    End If
    '------------------------------------------------------------------------------------------

    If CheckContentOpen() <> "Y" Then
      SaveMemberGradeLog(ConnString, "exception", "content flag is false")
      contentFlag = False
    Else
      SaveMemberGradeLog(ConnString, "exe", "content flag is true")
      contentFlag = True
    End If
    If CheckLoginAdditionalOpen() Then
      SaveMemberGradeLog(ConnString, "exe", "CheckLoginAdditionalOpen is true")
      loginAdditionalFlag = True
    Else
      SaveMemberGradeLog(ConnString, "exception", "CheckLoginAdditionalOpen is false")
      loginAdditionalFlag = False
    End If

    recordFlag = True  '---KPI�}�ҴN�Ӱ�---

    '------------------------------------------------------------------------------------------
    '---���e�[��---
    If contentFlag Then

      '---IQ�j�D�Ԥ��Ƨ�s---    
      Dim IQGrade As New IQGameGrade(ConnString, startDate, endDate, IQGameUnit)
      If Not IQGrade.HandleIQGameGrade() Then
        SaveMemberGradeLog(ConnString, "exception", "HandleIQGameGrade is false")
      Else
        SaveMemberGradeLog(ConnString, "exe", "HandleIQGameGrade is true")
      End If

      '---�����i��---
      Dim RaiseGrade As New RaiseGameGrade(ConnString, BeanConnString, startDate, endDate)
      If Not RaiseGrade.HandleRaiseGameGrade() Then
        SaveMemberGradeLog(ConnString, "exception", "HandleRaiseGameGrade is false")
      Else
        SaveMemberGradeLog(ConnString, "exe", "HandleRaiseGameGrade is true")
      End If


      '---�Q�ץ����Q����---
      Dim discussGrade As New DiscussAvgCommend(ConnString, startDate, endDate, QuestionUnit, discussUnit)
      If Not discussGrade.HandleDiscussAvgGrade() Then
        SaveMemberGradeLog(ConnString, "exception", "DiscussAvgCommend is false")
      Else
        SaveMemberGradeLog(ConnString, "exe", "DiscussAvgCommend is true")
      End If


      '---�ҵo�ݪ����D,�����Q�Q�׼�---
      Dim questionGrade As New QuestionAvgDiscuss(ConnString, startDate, endDate, QuestionUnit, discussUnit)
      If Not questionGrade.HandleQuestionAvgDiscuss() Then
        SaveMemberGradeLog(ConnString, "exception", "QuestionAvgDiscuss is false")
      Else
        SaveMemberGradeLog(ConnString, "exe", "QuestionAvgDiscuss is true")
      End If


      '---�ҵo�ݪ����D,�����Q�s����---
      Dim browseGrade As New QuestionAvgBrowse(ConnString, startDate, endDate, QuestionUnit, discussUnit)
      If Not browseGrade.HandleQuestionAvgBrowse() Then
        SaveMemberGradeLog(ConnString, "exception", "QuestionAvgBrowse is false")
      Else
        SaveMemberGradeLog(ConnString, "exe", "QuestionAvgBrowse is true")
      End If


    End If

    '------------------------------------------------------------------------------------------
    '---�n�J�欰�B�~���I---

    If loginAdditionalFlag Then
      Dim loginAdditionalGrade As New LoginAdditional(ConnString)
      If Not loginAdditionalGrade.HandleLoginAdditional() Then
        SaveMemberGradeLog(ConnString, "exception", "loginAdditionalGrade is false")
      Else
        SaveMemberGradeLog(ConnString, "exe", "loginAdditionalGrade is true")
      End If

    End If

    Dim raiseLogin As New RaiseLogin(ConnString, BeanConnString)
    If Not raiseLogin.HandleRaiseLogin() Then
      SaveMemberGradeLog(ConnString, "exception", "raiseLogin is false")
    Else
      SaveMemberGradeLog(ConnString, "exe", "raiseLogin is true")
    End If

    '------------------------------------------------------------------------------------------
    '---�ӤH�O��---

    If recordFlag Then

      '---�O���|�� �s���欰 �`����  MemberGradeSummary---
      Dim sumBrowseGrade As New SummaryBrowseGrade(ConnString)
      If Not sumBrowseGrade.HandleSummaryBrowseGrade(GetMemberId(ConnString, "MemberGradeBrowse")) Then
        SaveMemberGradeLog(ConnString, "exception", "HandleSummaryBrowseGrade is false")
      Else
        SaveMemberGradeLog(ConnString, "exe", "HandleSummaryBrowseGrade is true")
      End If

      '---�O���|�� �n�J�欰 �`����  MemberGradeSummary---
      Dim sumLoginGrade As New SummaryLoginGrade(ConnString)
      If Not sumLoginGrade.HandleSummaryLoginGrade(GetMemberId(ConnString, "MemberGradeLogin")) Then
        SaveMemberGradeLog(ConnString, "exception", "HandleSummaryLoginGrade is false")
      Else
        SaveMemberGradeLog(ConnString, "exe", "HandleSummaryLoginGrade is true")
      End If

      '---�O���|�� ���ɦ欰 �`����  MemberGradeSummary---
      Dim sumShareGrade As New SummaryShareGrade(ConnString)
      If Not sumShareGrade.HandleSummaryShareGrade(GetMemberId(ConnString, "MemberGradeShare")) Then
        SaveMemberGradeLog(ConnString, "exception", "HandleSummaryShareGrade is false")
      Else
        SaveMemberGradeLog(ConnString, "exe", "HandleSummaryShareGrade is true")
      End If

      '---�O���|�� ���e�[�� �`���� MemberGradeSummary---
      Dim sumContentGrade As New SummaryContentGrade(ConnString)
      If Not sumContentGrade.HandleSummaryContentGrade(GetMemberId(ConnString, "MemberGradeContent")) Then
        SaveMemberGradeLog(ConnString, "exception", "HandleSummaryContentGrade is false")
      Else
        SaveMemberGradeLog(ConnString, "exe", "HandleSummaryContentGrade is true")
      End If

      '---�O���|�� ���ʰѻP �`���� MemberGradeSummary---
      Dim sumAdditionalGrade As New SummaryAdditionalGrade(ConnString)
      If Not sumAdditionalGrade.HandleSummaryAdditionalGrade(GetMemberId(ConnString, "MemberGradeAdditional")) Then
        SaveMemberGradeLog(ConnString, "exception", "HandleSummaryAdditionalGrade is false")
      Else
        SaveMemberGradeLog(ConnString, "exe", "HandleSummaryAdditionalGrade is true")
      End If

      '---MemberGradeSummary�U�����ƭ��W��ҥ[�`---
      Dim sumCalculateGrade As New SummaryCalculateGrade(ConnString)
      If Not sumCalculateGrade.HandleSummaryCalculateGrade(GetMemberId(ConnString, "MemberGradeSummary")) Then
        SaveMemberGradeLog(ConnString, "exception", "HandleSummaryCalculateGrade is false")
      Else
        SaveMemberGradeLog(ConnString, "exe", "HandleSummaryCalculateGrade is true")
      End If

    End If
    '------------------------------------------------------------------------------------------

  End Sub

  Function CheckKpiOpen() As String

    Dim str As String = SelectTableWithContent(ConnString, "CodeMain", "mCode", "codeMetaID", "kpiOpen")
    Return str

  End Function

  Function GetKpiStartDateAndEndDate(ByRef startDate As String, ByRef endDate As String) As Boolean

    Dim str As String = ""
    myConnection = New SqlConnection(ConnString)
    Try
      myConnection.Open()
      Sql = "SELECT mCode, mValue FROM CodeMain WHERE codeMetaID = 'kpiTime' "
      myCommand = New SqlCommand(Sql, myConnection)
      myReader = myCommand.ExecuteReader
      While myReader.Read
        If myReader("mCode") IsNot DBNull.Value And myReader("mValue") IsNot DBNull.Value Then
          If myReader("mCode") = "startDate" Then startDate = myReader("mValue").ToString
          If myReader("mCode") = "endDate" Then endDate = myReader("mValue").ToString
        End If
      End While
      If Not myReader.IsClosed Then myReader.Close()
      myCommand.Dispose()
    Catch ex As Exception
      Return False
    Finally
      If myConnection.State = ConnectionState.Open Then myConnection.Close()
    End Try
    Return True

  End Function

  Function CheckContentOpen() As String

    Dim str As String = SelectTableWithContent(ConnString, "kpi_set_ind", "Rank2", "Rank0", "st_5")
    Return str

  End Function

  Function CheckLoginAdditionalOpen() As Boolean

    Dim dayDiff As Integer = 0
    Dim startDateStr As String = ""
    Dim endDateStr As String = ""
    Dim startDate As Date
    Dim theDate As Date = Today.ToShortDateString
    Dim dayFlag As Boolean = False

    '---�ˬd�O�_�t�Z7��---
    GetKpiStartDateAndEndDate(startDateStr, endDateStr)
    startDate = Date.Parse(startDateStr)
    dayDiff = DateDiff("d", startDate, theDate)

    If dayDiff Mod 7 = 0 Then dayFlag = True

    If dayFlag Then
      Return True
    Else
      Return False
    End If

    'If dayFlag Then
    '  Dim str As String = SelectTableWithContent(ConnString, "CodeMain", "mCode", "codeMetaID", "kpiLoginAdditional")
    '  If str = "Y" Then
    '    UpdateTableWithSql(ConnString, "UPDATE CodeMain SET mCode = 'Z' WHERE codeMetaID = 'kpiLoginAdditional' ")
    '    Return True
    '  ElseIf str = "Z" Then
    '    Return False
    '  ElseIf str = "N" Then
    '    UpdateTableWithSql(ConnString, "UPDATE CodeMain SET mCode = 'Y' WHERE codeMetaID = 'kpiLoginAdditional' ")
    '    Return False
    '  End If
    'Else
    '  UpdateTableWithSql(ConnString, "UPDATE CodeMain SET mCode = 'N' WHERE codeMetaID = 'kpiLoginAdditional' ")
    '  Return False
    'End If

  End Function

End Module
