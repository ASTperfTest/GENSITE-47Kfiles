Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration
Imports GSS.Vitals.COA.Data

Partial Class knowledge_grade_add
    Inherits System.Web.UI.Page

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim ScriptNoId As String = "<script>alert('連線逾時或尚未登入，請登入會員');window.location.href='" & Request.UrlReferrer.ToString() &"';</script>"
    Dim ScriptLinkError As String = "<script>alert('連結錯誤!!');history.back();</script>"
    Dim ScriptGradeError As String = "<script>alert('評價錯誤!!');history.back();</script>"
    Dim ScriptHaveGrade As String = "<script>alert('您已對此篇文章給予評價!!');history.back();</script>"
    Dim ScriptGradeSuccess As String = "<script>alert('給予評價成功!!');window.location.href='{0}';</script>"
    Dim ScriptNoGradeToSelf As String = "<script>alert('無法給予自己評價!!');history.back();</script>"

    Dim MemberId As String

    '-----------------------------------------------------------------------------
    If Session("memID") = "" Or Session("memID") Is Nothing Then
      MemberId = ""
      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "NoId", ScriptNoId)
      Exit Sub
    Else
      MemberId = Session("memID").ToString()
    End If
    '-----------------------------------------------------------------------------

    Dim DArticleId As String
    Dim ArticleId As String
    Dim ArticleType As String
    Dim CategoryId As String
    Dim Grade As String

    If Request.QueryString("ArticleId") = "" Then
      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "LinkError", ScriptLinkError)
      Exit Sub
    Else
      ArticleId = Request.QueryString("ArticleId")
    End If

    If Request.QueryString("DArticleId") = "" Then
      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "LinkError", ScriptLinkError)
      Exit Sub
    Else
      DArticleId = Request.QueryString("DArticleId")
    End If

    If Request.QueryString("Grade") <> "1" And Request.QueryString("Grade") <> "2" And Request.QueryString("Grade") <> "3" _
        And Request.QueryString("Grade") <> "4" And Request.QueryString("Grade") <> "5" Then
      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "GradeError", ScriptGradeError)
      Exit Sub
    Else
      Grade = Request.QueryString("Grade")
    End If

    If Request.QueryString("ArticleType") <> "A" And Request.QueryString("ArticleType") <> "B" And Request.QueryString("ArticleType") <> "C" _
               And Request.QueryString("ArticleType") <> "D" And Request.QueryString("ArticleType") <> "E" Then
      ArticleType = "A"
    Else
      ArticleType = Request.QueryString("ArticleType")
    End If

    If Request.QueryString("CategoryId") <> "A" And Request.QueryString("CategoryId") <> "B" And Request.QueryString("CategoryId") <> "C" _
        And Request.QueryString("CategoryId") <> "D" And Request.QueryString("CategoryId") <> "E" And Request.QueryString("CategoryId") <> "F" Then
      CategoryId = ""
    Else
      CategoryId = Request.QueryString("CategoryId")
    End If

    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim sqlString As String = ""
    Dim myConnection As SqlConnection
    Dim myCommand As SqlCommand
    Dim myReader As SqlDataReader
    Dim Flag As Boolean = False
    Dim StarFlag As Boolean = False

    myConnection = New SqlConnection(ConnString)

    '---check if the member give it comment grade.
    Dim CheckEditorFlag As Boolean = False
    Dim CheckEditor As String = ""
    Try
      sqlString = "SELECT CuDTGeneric.iEditor FROM CuDTGeneric WHERE (CuDTGeneric.iCTUnit = @ictunit) AND (CuDTGeneric.iCUItem = @icuitem) "
      myConnection.Open()
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@ictunit", WebConfigurationManager.AppSettings("KnowledgeDiscussCtUnitId"))
      myCommand.Parameters.AddWithValue("@icuitem", DArticleId)
      myReader = myCommand.ExecuteReader
      If myReader.Read Then
        CheckEditorFlag = True
        CheckEditor = myReader("iEditor")
      Else
        CheckEditorFlag = False
      End If
      If Not myReader.IsClosed Then
        myReader.Close()
      End If
      myCommand.Dispose()

    Catch ex As Exception
      If Request.QueryString("debug") = "true" Then
        Response.Write(ex.ToString)
        Response.End()
      End If
      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "GradeError", ScriptGradeError)
      Exit Sub
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

    If CheckEditorFlag Then
      If CheckEditor = MemberId Then
        Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "ScriptNoGradeToSelf", ScriptNoGradeToSelf)
        Exit Sub
      End If
    Else      
      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "ScriptLinkError", ScriptLinkError)
      Exit Sub
    End If

    '---check if the member has gave the commemnt grade.
    Try
      sqlString = "SELECT CuDTGeneric.iCUItem FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem "
      sqlString &= "WHERE (CuDTGeneric.iCTUnit = @ictunit) AND (CuDTGeneric.iEditor = @ieditor) AND (KnowledgeForum.ParentIcuitem = @icuitem) AND (KnowledgeForum.Status = 'N') "
      myConnection.Open()
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@ictunit", WebConfigurationManager.AppSettings("KnowledgeGradeCtUnitId"))
      myCommand.Parameters.AddWithValue("@ieditor", MemberId)
      myCommand.Parameters.AddWithValue("@icuitem", DArticleId)
      myReader = myCommand.ExecuteReader
      If myReader.HasRows Then
        Flag = False
      Else
        Flag = True
      End If
      If Not myReader.IsClosed Then
        myReader.Close()
      End If
      myCommand.Dispose()

    Catch ex As Exception
      If Request.QueryString("debug") = "true" Then
        Response.Write(ex.ToString)
        Response.End()
      End If
      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "GradeError", ScriptGradeError)
      Exit Sub
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

    If Not Flag Then
      '---have grade---
      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "HaveGrade", ScriptHaveGrade)
      Exit Sub
    Else
      '---give grade---
      Dim MaxIcuitem As String
      Try
        '---insert grade into cudtgeneric and knowledgeforum---
        sqlString = "INSERT INTO CuDTGeneric(iBaseDSD, iCTUnit, fCTUPublic, sTitle, iEditor, iDept, siteId) "
        sqlString &= "VALUES (@ibasedsd, @ictunit, 'Y', @stitle, @ieditor, '0', @siteid)"
        myConnection.Open()
        myCommand = New SqlCommand(sqlString, myConnection)
        myCommand.Parameters.AddWithValue("@ibasedsd", WebConfigurationManager.AppSettings("KnowledgeBaseDSD"))
        myCommand.Parameters.AddWithValue("@ictunit", WebConfigurationManager.AppSettings("KnowledgeGradeCtUnitId"))
        myCommand.Parameters.AddWithValue("@stitle", "討論評價-" & DArticleId)
        myCommand.Parameters.AddWithValue("@ieditor", MemberId)
        myCommand.Parameters.AddWithValue("@siteid", WebConfigurationManager.AppSettings("SiteId"))
        myCommand.ExecuteNonQuery()
        myCommand.Dispose()

        sqlString = "SELECT @@identity AS MaxIcuitem"
        Dim MaxCommand As SqlDataAdapter = New SqlDataAdapter(sqlString, myConnection)
        Dim MaxDataSet As DataSet = New DataSet()
        MaxCommand.Fill(MaxDataSet, "IndexRecommend")
        Dim myTable As DataTable = MaxDataSet.Tables("IndexRecommend")
        MaxIcuitem = myTable.Rows(0).Item(0)

        sqlString = "INSERT INTO KnowledgeForum(gicuitem, GradeCount, GradePersonCount, ParentIcuitem) VALUES (@gicuitem, @gradecount, 1, @parenticuitem)"
        myCommand = New SqlCommand(sqlString, myConnection)
        myCommand.Parameters.AddWithValue("@gicuitem", MaxIcuitem)
        myCommand.Parameters.AddWithValue("@gradecount", Convert.ToInt32(Grade))
        myCommand.Parameters.AddWithValue("@parenticuitem", DArticleId)
        myCommand.ExecuteNonQuery()
        myCommand.Dispose()

        '---update the grade and gradeperson of article and its parent---
        sqlString = "UPDATE KnowledgeForum SET GradeCount = GradeCount + @gradecount, GradePersonCount = GradePersonCount + 1 WHERE gicuitem = @gicuitem or gicuitem = @agicuitem"
        myCommand = New SqlCommand(sqlString, myConnection)
        myCommand.Parameters.AddWithValue("@gradecount", Integer.Parse(Grade))
        myCommand.Parameters.AddWithValue("@agicuitem", ArticleId)
        myCommand.Parameters.AddWithValue("@gicuitem", DArticleId)
        myCommand.ExecuteNonQuery()
        myCommand.Dispose()

        StarFlag = True

        '---20080927---vincent---kpi share---
        Dim share As New KPIShare(MemberId, "shareCommend", "")
        share.HandleShare()
                AddPuzzleScore(MemberId)
      Catch ex As Exception
        If Request.QueryString("debug") = "true" Then
          Response.Write(ex.ToString)
          Response.End()
        End If
        Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "GradeError", ScriptGradeError)
        Exit Sub
      Finally
        If myConnection.State = ConnectionState.Open Then
          myConnection.Close()
        End If
      End Try

      '---活動專用--------------------------------------------------------------------------------------------------
      Dim Activity As New ActivityFilter("", MemberId)
      Dim ActivityFlag As Boolean = Activity.CheckStarGrade

      If ActivityFlag Then
        Activity.StarGradeAdd()
      End If
      '-------------------------------------------------------------------------------------------------------------

      If StarFlag Then
        Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "HaveGrade", ScriptGradeSuccess.Replace("{0}", "/knowledge/knowledge_discuss_cp.aspx?ArticleId=" & ArticleId & "&DArticleId=" & DArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId))
        Exit Sub
      End If

    End If

  End Sub

    Sub Activity(ByVal articleid As String, ByVal memberid As String, ByVal grade As String)

        Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
        Dim sqlString As String = ""
        Dim myConnection As New SqlConnection(ConnString)
        Dim myCommand As SqlCommand
        Dim myReader As SqlDataReader

        Dim editor As String = ""
        Dim ArticleCount As Integer = 0
        Dim GradeCount As Integer = 0
        Dim Average As Double = 0
        Dim UserExist As Boolean = False
        Dim CurrentGrade As Double = 0
        Dim CurrentGradeFlag As Boolean = False

        Try
            myConnection.Open()

            '---取出這篇討論誰寫的---
            sqlString = "SELECT iEditor FROM CuDtGeneric WHERE icuitem = @icuitem"
            myCommand = New SqlCommand(sqlString, myConnection)
            myCommand.Parameters.AddWithValue("@icuitem", articleid)
            myReader = myCommand.ExecuteReader
            If myReader.Read Then
                editor = myReader("iEditor")
            End If
            myReader.Close()
            myCommand.Dispose()

            '---檢查table中是否已有這個人的資料---
            sqlString = "SELECT * FROM ActivityMember WHERE MemberId = @memberid"
            myCommand = New SqlCommand(sqlString, myConnection)
            myCommand.Parameters.AddWithValue("@memberid", editor)
            myReader = myCommand.ExecuteReader
            If myReader.Read Then
                UserExist = True
            End If
            myReader.Close()
            myCommand.Dispose()

            If Not UserExist Then
                sqlString = "INSERT INTO ActivityMember VALUES(@memberid, 0, 0, 0, 0, 0, 0)"
                myCommand = New SqlCommand(sqlString, myConnection)
                myCommand.Parameters.AddWithValue("@memberid", editor)
                myCommand.ExecuteNonQuery()
                myCommand.Dispose()
            End If

            '---將資料寫入活動---
            sqlString = "INSERT INTO ActivityInfo VALUES(@articleid, 2, @memberid, GETDATE(), @grade, @memberid1)"
            myCommand = New SqlCommand(sqlString, myConnection)
            myCommand.Parameters.AddWithValue("@articleid", articleid)
            myCommand.Parameters.AddWithValue("@memberid", editor)
            myCommand.Parameters.AddWithValue("@grade", Integer.Parse(grade))
            myCommand.Parameters.AddWithValue("@memberid1", memberid)
            myCommand.ExecuteNonQuery()
            myCommand.Dispose()
        
            '---取出活動中對於這個人的評價的篇數---
            sqlString = "SELECT COUNT(*) FROM ActivityInfo WHERE MemberId = @memberid AND Type = 2"
            myCommand = New SqlCommand(sqlString, myConnection)
            myCommand.Parameters.AddWithValue("@articleid", articleid)
            myCommand.Parameters.AddWithValue("@memberid", editor)
            ArticleCount = myCommand.ExecuteScalar
            myCommand.Dispose()

            '---取出活動中對於這個人的評價的總數---
            sqlString = "SELECT SUM(Grade) FROM ActivityInfo WHERE MemberId = @memberid AND Type = 2"
            myCommand = New SqlCommand(sqlString, myConnection)
            myCommand.Parameters.AddWithValue("@memberid", editor)
            GradeCount = myCommand.ExecuteScalar
            myCommand.Dispose()

            '---計算評價平均值---
            ArticleCount = IIf(ArticleCount <> 0, ArticleCount, 1)
            Average = Math.Round(GradeCount / ArticleCount, 2)

            '---將評價平均值更新至活動中---
            sqlString = "UPDATE ActivityMember SET StarGrade = @grade WHERE MemberId = @memberid"
            myCommand = New SqlCommand(sqlString, myConnection)
            myCommand.Parameters.AddWithValue("@grade", Average)
            myCommand.Parameters.AddWithValue("@memberid", editor)
            GradeCount = myCommand.ExecuteNonQuery
            myCommand.Dispose()

            '---檢查此人目前分數---
            sqlString = "SELECT (ISNULL(AskGrade, 0) + ISNULL(DiscussGrade, 0) + ISNULL(StarGrade, 0)) AS Grade FROM ActivityMember WHERE MemberId = @memberid"
            myCommand = New SqlCommand(sqlString, myConnection)
            myCommand.Parameters.AddWithValue("@memberid", editor)
            myReader = myCommand.ExecuteReader
            If myReader.Read Then
                CurrentGrade = myReader("Grade")
                If CurrentGrade >= 20 Then
                    CurrentGradeFlag = True
                End If
            End If
            myReader.Close()
            myCommand.Dispose()

            '---更新第二關關卡-已大於20分---
            If CurrentGradeFlag Then
                sqlString = "UPDATE ActivityMember SET Game2Flag = 1 WHERE MemberId = @memberid"
                myCommand = New SqlCommand(sqlString, myConnection)
                myCommand.Parameters.AddWithValue("@memberid", editor)
                myCommand.ExecuteNonQuery()
                myCommand.Dispose()
            End If

        Catch ex As Exception
            If Request.QueryString("debug") = "true" Then
                Response.Write(ex.ToString)
            End If
        Finally
            If myConnection.State = ConnectionState.Open Then
                myConnection.Close()
            End If
        End Try

    End Sub

    Private Sub AddPuzzleScore(ByVal meid As String)
        Dim sql As String = " BEGIN TRAN "
        sql += " if not exists( select * from ACCOUNT where login_id =  @login_id) "
        sql += " begin "
        sql += "insert into ACCOUNT (LOGIN_ID,REALNAME,NICKNAME,EMAIL,Energy,CREATE_DATE,MODIFY_DATE,GetEnergy) "
        sql += "    select account,REALNAME,isnull(NICKNAME,''),EMAIL,10,getdate(),getdate(),10 from mGIPcoanew..member where account =  @login_id "
        sql += " End "
        sql += "     else "
        sql += " begin "
        sql += "    update ACCOUNT set REALNAME = m.REALNAME,NICKNAME=isnull(m.NICKNAME,''),email = m.email,Energy=Energy+10,MODIFY_DATE=getdate(),GetEnergy=GetEnergy + 10 "
        sql += "     from (select * from mGIPcoanew..member  where account =  @login_id ) m "
        sql += "    where ACCOUNT.login_id =  @login_id "
        sql += " End "
        sql += " commit "
        SqlHelper.ExecuteNonQuery("PuzzleConnString", sql, DbProviderFactories.CreateParameter("HistoryPictureConnString", "@login_id", "@login_id", meid))
    End Sub
End Class
