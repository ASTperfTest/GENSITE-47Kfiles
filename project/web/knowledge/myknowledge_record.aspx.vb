Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration
Imports GSS.Vitals.COA.Data

Partial Class knowledge_myknowledge_record
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim sb As New StringBuilder
        Dim MemberId As String
        Dim Script As String = "<script>alert('連線逾時或尚未登入，請登入會員');window.location.href='/knowledge/knowledge.aspx';</script>"

        '-----------------------------------------------------------------------------
        If Session("memID") = "" Or Session("memID") Is Nothing Then
            MemberId = ""
            Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "NoId", Script)
            Exit Sub
        Else
            MemberId = Session("memID").ToString()
        End If
        '-----------------------------------------------------------------------------
        
        TabText.Text = WebUtility.GetMyAreaLinks(MyAreaLink.myknowledge_record)



        '新增活動頁籤 
        REM Dim Activity As New ActivityFilter(System.Web.Configuration.WebConfigurationManager.AppSettings("ActivityId"), "")
        REM Dim ActivityFlag As Boolean = Activity.CheckStartActivity
        REM If ActivityFlag Then 
        REM TabText.Text &= "<li ><a href=""/knowledge/myknowledgeActivity_question_lp.aspx"">我的活動發問</a></li>" & _
        REM "<li ><a href=""/knowledge/myknowledgeActivity_discuss_lp.aspx"">我的活動討論</a></li>"
        REM End If


        '-----------------------------------------------------------------------------

        memberIdText.Text = MemberId

        memberPicPath.AlternateText = "個人成長圖示"
        memberPicPath.ImageUrl = "~/xslGip/style3/images/" & GetMemberPicPath(MemberId)



        memberChangePicBtn.AlternateText = "更改圖示"
        memberChangePicBtn.ImageUrl = "/xslGip/style3/images/myset_bt.jpg"
        memberChangePicBtn.PostBackUrl = "/knowledge/myknowledge_changepic.aspx"

        REM If ActivityFlag Then
        REM activityGrade.Text = GetKnowledgeActivityInfo(MemberId)
        REM End If

        sb.Append("<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" summary=""問題列表資料表格"" class=""myset_data"">")
        sb.Append(GetTableContent(MemberId))
        sb.Append("</table>")

        tableContentText.Text = sb.ToString

    End Sub

    Function GetMemberPicPath(ByVal memberId As String) As String

        Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
        Dim sqlStr As String = ""
        Dim myConnection As SqlConnection
        Dim myCommand As SqlCommand
        Dim myReader As SqlDataReader
        Dim picType As String = "A"

        myConnection = New SqlConnection(ConnString)
        Try
            sqlStr = "SELECT pictype FROM Member WHERE account = @memberid"
            myConnection.Open()
            myCommand = New SqlCommand(sqlStr, myConnection)
            myCommand.Parameters.AddWithValue("@memberid", memberId)
            myReader = myCommand.ExecuteReader
            If myReader.Read Then
                If myReader("pictype") IsNot DBNull.Value Then
                    picType = myReader("pictype").ToString()
                End If
            End If
            If Not myReader.IsClosed Then
                myReader.Close()
            End If
            myCommand.Dispose()

            Dim calculateTotal As Integer = 0
            sqlStr = "SELECT calculateTotal FROM MemberGradeSummary WHERE memberId = @memberid"
            myCommand = New SqlCommand(sqlStr, myConnection)
            myCommand.Parameters.AddWithValue("@memberid", memberId)
            myReader = myCommand.ExecuteReader
            If myReader.Read Then
                calculateTotal = myReader("calculateTotal")
            End If
            If Not myReader.IsClosed Then
                myReader.Close()
            End If
            myCommand.Dispose()

            Dim level As Integer = 1
            sqlStr = "SELECT codeMetaID, mCode, mValue, mSortValue FROM CodeMain WHERE (codeMetaID = 'gradelevel') AND (@grade >= mValue) ORDER BY mSortValue DESC"
            myCommand = New SqlCommand(sqlStr, myConnection)
            myCommand.Parameters.AddWithValue("@grade", calculateTotal)
            myReader = myCommand.ExecuteReader
            If myReader.Read Then
                level = Integer.Parse(myReader("mCode"))
            End If
            If Not myReader.IsClosed Then
                myReader.Close()
            End If
            myCommand.Dispose()

            If level = 1 Then
                If picType = "A" Then
                    Return "seticon1-1.jpg"
                ElseIf picType = "B" Then
                    Return "seticon2-1.jpg"
                ElseIf picType = "C" Then
                    Return "seticon3-1.jpg"
                ElseIf picType = "D" Then
                    Return "seticon4-1.jpg"
                End If
            ElseIf level = 2 Then
                If picType = "A" Then
                    Return "seticon1-2.jpg"
                ElseIf picType = "B" Then
                    Return "seticon2-2.jpg"
                ElseIf picType = "C" Then
                    Return "seticon3-2.jpg"
                ElseIf picType = "D" Then
                    Return "seticon4-2.jpg"
                End If
            ElseIf level = 3 Then
                If picType = "A" Then
                    Return "seticon1-3.jpg"
                ElseIf picType = "B" Then
                    Return "seticon2-3.jpg"
                ElseIf picType = "C" Then
                    Return "seticon3-3.jpg"
                ElseIf picType = "D" Then
                    Return "seticon4-3.jpg"
                End If
            ElseIf level = 4 Then
                If picType = "A" Then
                    Return "seticon1-4.jpg"
                ElseIf picType = "B" Then
                    Return "seticon2-4.jpg"
                ElseIf picType = "C" Then
                    Return "seticon3-4.jpg"
                ElseIf picType = "D" Then
                    Return "seticon4-4.jpg"
                End If
            End If
            Return "seticon1-1.jpg"

        Catch ex As Exception
            Return "seticon1-1.jpg"
        Finally
            If myConnection.State = ConnectionState.Open Then
                myConnection.Close()
            End If
        End Try

    End Function

    Function GetTableContent(ByVal memberId As String) As String

        Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
        Dim sqlStr As String = ""
        Dim myConnection As SqlConnection
        Dim myCommand As SqlCommand
        Dim myReader As SqlDataReader
        Dim sb As New StringBuilder

        Dim browseTotal As Integer = 0, loginTotal As Integer = 0, shareTotal As Integer = 0, contentTotal As Integer = 0, contentTotalHistory As Integer = 0
        Dim additionalTotal As Integer = 0, additionalTotalHistory As Integer = 0, calculateTotal As Integer = 0

        myConnection = New SqlConnection(ConnString)
        Try
            myConnection.Open()

            sqlStr = "SELECT * FROM MemberGradeSummary WHERE memberId = @memberid"
            myCommand = New SqlCommand(sqlStr, myConnection)
            myCommand.Parameters.AddWithValue("@memberid", memberId)
            myReader = myCommand.ExecuteReader
            If myReader.Read Then
                browseTotal = myReader("browseTotal")
                loginTotal = myReader("loginTotal")
                shareTotal = myReader("shareTotal")
                contentTotal = myReader("contentTotal")
                additionalTotal = myReader("additionalTotal")
                contentTotalHistory = myReader("contentTotalHistory")
                additionalTotalHistory = myReader("additionalTotalHistory")
                calculateTotal = myReader("calculateTotal")
            End If
            If Not myReader.IsClosed Then
                myReader.Close()
            End If
            myCommand.Dispose()

            sqlStr = "SELECT codeMetaID, mCode, mValue, mSortValue FROM CodeMain WHERE (codeMetaID = 'gradelevel') AND (@grade >= mValue) ORDER BY mSortValue DESC"

            myCommand = New SqlCommand(sqlStr, myConnection)
            myCommand.Parameters.AddWithValue("@grade", calculateTotal)
            myReader = myCommand.ExecuteReader
            If myReader.Read Then
                If myReader("mCode") = "1" Then
                    sb.Append("<tr><th width=""30%"">等級：[入門級會員]</th><th>歷年積分</th><th>本年度積分</th></tr>")
                ElseIf myReader("mCode") = "2" Then
                    sb.Append("<tr><th width=""30%"">等級：[進階級會員]</th><th>歷年積分</th><th>本年度積分</th></tr>")
                ElseIf myReader("mCode") = "3" Then
                    sb.Append("<tr><th width=""30%"">等級：[高手級會員]</th><th>歷年積分</th><th>本年度積分</th></tr>")
                ElseIf myReader("mCode") = "4" Then
                    sb.Append("<tr><th width=""30%"">等級：[達人級會員]</th><th>歷年積分</th><th>本年度積分</th></tr>")
                Else
                    sb.Append("<tr><th width=""30%"">等級</th><th>歷年積分</th><th>本年度積分</th></tr>")
                End If
            End If
            If Not myReader.IsClosed Then
                myReader.Close()
            End If
            myCommand.Dispose()

            Dim browsePercent As Integer = 0, sharePercent As Integer = 0, contentPercent As Integer = 0
            Dim loginPercent As Integer = 0, additionalPercent As Integer = 0

            sqlStr = "SELECT Rank0_1, Rank0_2 FROM kpi_set_score WHERE Rank0 = 'st_1'"
            myCommand = New SqlCommand(sqlStr, myConnection)
            myReader = myCommand.ExecuteReader
            While myReader.Read

                If myReader("Rank0_2") = "st_111" Then
                    browsePercent = Integer.Parse(myReader("Rank0_1"))
                ElseIf myReader("Rank0_2") = "st_112" Then
                    loginPercent = Integer.Parse(myReader("Rank0_1"))
                ElseIf myReader("Rank0_2") = "st_113" Then
                    sharePercent = Integer.Parse(myReader("Rank0_1"))
                ElseIf myReader("Rank0_2") = "st_114" Then
                    contentPercent = Integer.Parse(myReader("Rank0_1"))
                ElseIf myReader("Rank0_2") = "st_115" Then
                    additionalPercent = Integer.Parse(myReader("Rank0_1"))
                End If

            End While
            If Not myReader.IsClosed Then
                myReader.Close()
            End If
            myCommand.Dispose()


            'Bob, 會員年度統計   2009/08/11

            sqlStr = "SELECT isnull(GradeBrowse_thisyear.GradeBrowse, 0) as GradeBrowse, isnull(GradeLogin_thisyear.GradeLogin, 0) as GradeLogin, "
            sqlStr = sqlStr & " isnull(GradeShare_thisyear.GradeShare, 0) as GradeShare, Member.account as memberId , (ISNULL(VCB.ThisYearBrowseGrade,0)+ISNULL(VCC.ThisYearCommendGrade,0)+ISNULL(VCD.ThisYearDiscussGrade,0)+ISNULL(MGCBY.contentIQPoint,0)+ISNULL(MGCBY.contentRaisePoint,0)+ISNULL(MGCBY.contentLogoPoint,0)+ ISNULL(MGCBY.contentPlantLogPoint,0)+ISNULL(MGCBY.contentFishBowlPoint,0)+ ISNULL(MGCBY.contentDrPoint,0)+ ISNULL(MGCBY.contentKnowledgeQAPoint,0)+ ISNULL(MGCBY.contentDr2010Point,0)+ISNULL(MGCBY.contentTreasureHuntPoint,0)) as thisYearContentTotal, ISNULL(VIEWADDTHIS.ThisYearADDpoint, 0) as thisYearADDTotal, "
            sqlStr = sqlStr & " (isnull(GradeBrowse_thisyear.GradeBrowse, 0) * 0.15 + isnull(GradeLogin_thisyear.GradeLogin, 0) * 0.2) + isnull(GradeShare_thisyear.GradeShare, 0) * 0.3 + (ISNULL(VCB.ThisYearBrowseGrade,0)+ISNULL(VCC.ThisYearCommendGrade,0)+ISNULL(VCD.ThisYearDiscussGrade,0)+ISNULL(MGCBY.contentIQPoint,0)+ISNULL(MGCBY.contentRaisePoint,0)+ISNULL(MGCBY.contentLogoPoint,0)+ ISNULL(MGCBY.contentPlantLogPoint,0)+ISNULL(MGCBY.contentFishBowlPoint,0)+ ISNULL(MGCBY.contentDrPoint,0)+ ISNULL(MGCBY.contentKnowledgeQAPoint,0)+ ISNULL(MGCBY.contentDr2010Point,0)+ISNULL(MGCBY.contentTreasureHuntPoint,0)) * 0.20 + ISNULL(VIEWADDTHIS.ThisYearADDpoint, 0) * 0.15 AS total "
            sqlStr = sqlStr & " FROM Member "
            sqlStr = sqlStr & " left JOIN GradeLogin_thisyear  ON Member.account = GradeLogin_thisyear.memberId "
            sqlStr = sqlStr & " left JOIN GradeShare_thisyear  ON Member.account = GradeShare_thisyear.memberId "
            sqlStr = sqlStr & " left JOIN GradeBrowse_thisyear ON Member.account = GradeBrowse_thisyear.memberId "
            sqlStr = sqlStr & " left JOIN MemberGradeSummary B ON Member.account = B.memberId "
            sqlStr = sqlStr & " LEFT JOIN GradeContentBrowse_thisyear VCB ON Member.account = VCB.ieditor  "
            sqlStr = sqlStr & " LEFT JOIN GradeContentCommend_thisyear VCC ON Member.account = VCC.ieditor  "
            sqlStr = sqlStr & " LEFT JOIN GradeContentDiscuss_thisyear VCD ON Member.account = VCD.ieditor "
            sqlStr = sqlStr & " LEFT JOIN View_MemberGradeAdditional_ThisYear VIEWADDTHIS ON Member.account = VIEWADDTHIS.memberId "
            sqlStr = sqlStr & " LEFT JOIN View_MemberGradeContent_ThisYear VIEWConTHIS ON Member.account = VIEWConTHIS.memberId "
            sqlStr = sqlStr & " LEFT JOIN dbo.MemberGradeContentByYear MGCBY ON Member.account = MGCBY.memberId AND (CONVERT(INT, CONVERT(varchar(4), getdate(), 120 ) ) - ISNULL(years,0) = 0) "
            sqlStr = sqlStr & " where Member.account ='" & memberId & "'"

            Dim thisY_Total As String = String.Empty
            Dim thisY_GradeBrowse As String = String.Empty
            Dim thisY_GradeLogin As String = String.Empty
            Dim thisY_Share As String = String.Empty
            Dim thisY_contentTotal As String = String.Empty
            Dim thisY_contentADDTotal As String = String.Empty

            Using thisY_Cmd As SqlCommand = New SqlCommand(sqlStr, myConnection)
                thisY_Cmd.Parameters.AddWithValue("@memberid", memberId)
                Dim thisY_Reader As SqlDataReader = thisY_Cmd.ExecuteReader
                If thisY_Reader.Read Then
                    thisY_Total = thisY_Reader("total").ToString()
                    thisY_GradeBrowse = thisY_Reader("GradeBrowse").ToString()
                    thisY_GradeLogin = thisY_Reader("GradeLogin").ToString()
                    thisY_Share = thisY_Reader("GradeShare").ToString()
                    thisY_contentTotal = thisY_Reader("thisYearContentTotal").ToString()
                    thisY_contentADDTotal = thisY_Reader("thisYearADDTotal").ToString()
                End If

            End Using

            '================================================

            'sb.Append("<tr><th>網站瀏覽得點 (佔" & browsePercent & "%)</th><td width='35%'><span class=""red01"">" & browseTotal & "</span> (瀏覽行為原始得點)</td><td width='35%'>" & thisY_GradeBrowse & "</td></tr>")
            'sb.Append("<tr><th>使用頻率得點 (佔" & loginPercent & "%)</th><td><span class=""red01"">" & loginTotal & "</span> (登入行為原始得點)</td><td>" & thisY_GradeLogin & "</td></tr>")
            'sb.Append("<tr><th>參與互動得點 (佔" & sharePercent & "%)</th><td><span class=""red01"">" & shareTotal & "</span> (互動行為原始得點)</td><td>" & thisY_Share & "</td></tr>")
            'sb.Append("<tr><th>內容價值得點 (佔" & contentPercent & "%)</th><td><span class=""red01"">" & contentTotal & "</span> (內容價值原始得點)</td><td>" & thisY_contentTotal & "</td></tr>")
            'sb.Append("<tr><th>活動加值得點 (佔" & additionalPercent & "%)</th><td><span class=""red01"">" & additionalTotal & "</span> (活動參與原始得點)</td><td>" & (CType(additionalTotal, Integer)-CType(additionalTotalHistory, Integer)).ToString() & "</td></tr>")
            'sb.Append("<tr><th>總得點</th><td><span class=""red01"">" & calculateTotal & "</span> </td><td>" & thisY_Total & "</td></tr>")
            'sb.Append("<tr><td colspan='3'>．總得點 = 各計分項目原始得點*權重%之加總<br/>．請各位會員注意!! 歷年積分為每日排程更新（非即時更新）</td></tr>")

            '隱藏kpi加權比重以及總分
            sb.Append("<tr><th>網站瀏覽得點 </th><td width='35%'><span class=""red01"">" & browseTotal & "</span> (瀏覽行為原始得點)</td><td width='35%'>" & thisY_GradeBrowse & "</td></tr>")
            sb.Append("<tr><th>使用頻率得點 </th><td><span class=""red01"">" & loginTotal & "</span> (登入行為原始得點)</td><td>" & thisY_GradeLogin & "</td></tr>")
            sb.Append("<tr><th>參與互動得點 </th><td><span class=""red01"">" & shareTotal & "</span> (互動行為原始得點)</td><td>" & thisY_Share & "</td></tr>")
            sb.Append("<tr><th>內容價值得點 </th><td><span class=""red01"">" & contentTotal & "</span> (內容價值原始得點)</td><td>" & thisY_contentTotal & "</td></tr>")
            sb.Append("<tr><th>活動加值得點 </th><td><span class=""red01"">" & additionalTotal & "</span> (活動參與原始得點)</td><td>" & thisY_contentADDTotal & "</td></tr>")
            'sb.Append("<tr><th>總得點</th><td><span class=""red01"">" & calculateTotal & "</span> </td><td>" & thisY_Total & "</td></tr>")
            Dim imagesstring As String = ""
            If Integer.Parse(Fix(thisY_Total)) >= 1000 Then
                imagesstring += "<img src=""/images/1.gif "" />&nbsp;&nbsp;"
            Else
                imagesstring += "<img src=""/images/g01.gif "" />&nbsp;&nbsp;"
            End If
            If Integer.Parse(Fix(thisY_Total)) >= 500 Then
                imagesstring += "<img src=""/images/2.gif "" />&nbsp;&nbsp;"
            Else
                imagesstring += "<img src=""/images/g02.gif "" />&nbsp;&nbsp;"
            End If
            If Integer.Parse(thisY_Share) >= 175 Then
                imagesstring += "<img src=""/images/3.gif "" />&nbsp;&nbsp;"
            Else
                imagesstring += "<img src=""/images/g03.gif "" />&nbsp;&nbsp;"
            End If
            If Integer.Parse(thisY_GradeBrowse) >= 1000 Then
                imagesstring += "<img src=""/images/4.gif "" />&nbsp;&nbsp;"
            Else
                imagesstring += "<img src=""/images/g04.gif "" />&nbsp;&nbsp;"
            End If
            If Integer.Parse(thisY_GradeLogin) >= 1250 Then
                imagesstring += "<img src=""/images/5.gif "" />&nbsp;&nbsp;"
            Else
                imagesstring += "<img src=""/images/g05.gif "" />&nbsp;&nbsp;"
            End If

            Dim EEbool As Boolean = False

            If CheckIsFImage(memberId) <= 300 And Integer.Parse(thisY_Share) >= 50 Then
                imagesstring += "<img src=""/images/6.gif "" />&nbsp;&nbsp;"
            Else
                imagesstring += "<img src=""/images/g06.gif "" />&nbsp;&nbsp;"
            End If

            If Integer.Parse(Fix(thisY_Total)) >= 300 And Integer.Parse(thisY_GradeBrowse) >= 800 Then
                imagesstring += "<img src=""/images/7.gif "" />&nbsp;&nbsp;"
            Else
                imagesstring += "<img src=""/images/g07.gif "" />&nbsp;&nbsp;"
            End If
            sb.Append("<tr><th>入圍獎項</th><td colspan='2'>" + imagesstring + "</td></tr>")
            sb.Append("<tr><td colspan='3'>．請各位會員注意!! 歷年積分及本年度積分為每日排程更新（非即時更新）</td></tr>")
        Catch ex As Exception
            Throw ex
        Finally
            If myConnection.State = ConnectionState.Open Then
                myConnection.Close()
            End If
        End Try

        Return sb.ToString

    End Function

    Function CheckIsFImage(ByVal meid As String) As Integer
        Dim returnValue As Integer = 400
        Dim sqlString As String = ""
        sqlString &= "select G2.* from ( "
        sqlString &= "select G1.*, ROW_NUMBER() OVER(ORDER BY calculateTotal DESC) AS rowid from ( "
        sqlString &= "SELECT "
        sqlString &= "Member.account "
        sqlString &= ", Member.realname "
        sqlString &= ", Member.nickname "
        sqlString &= ", isnull(GradeBrowse_thisyear.GradeBrowse, 0) as browseTotal "
        sqlString &= ", isnull(GradeLogin_thisyear.GradeLogin, 0) as loginTotal "
        sqlString &= ", isnull(GradeShare_thisyear.GradeShare, 0) as shareTotal , isnull(B.calculateTotal,0) as allyearstotal "
        sqlString &= ", (isnull(GradeBrowse_thisyear.GradeBrowse, 0) * 0.15 + isnull(GradeLogin_thisyear.GradeLogin, 0) * 0.2) + isnull(GradeShare_thisyear.GradeShare, 0) * 0.3 + (ISNULL(VCB.ThisYearBrowseGrade,0)+ISNULL(VCC.ThisYearCommendGrade,0)+ISNULL(VCD.ThisYearDiscussGrade,0)+ISNULL(MGCBY.contentIQPoint,0)+ISNULL(MGCBY.contentRaisePoint,0)+ISNULL(MGCBY.contentLogoPoint,0)+ ISNULL(MGCBY.contentPlantLogPoint,0)+ISNULL(MGCBY.contentFishBowlPoint,0)+ ISNULL(MGCBY.contentDrPoint,0)+ ISNULL(MGCBY.contentKnowledgeQAPoint,0)+ ISNULL(MGCBY.contentDr2010Point,0)+ISNULL(MGCBY.contentTreasureHuntPoint,0)) * 0.2 + isnull(VIEWADDTHIS.ThisYearADDpoint, 0) * 0.15 AS calculateTotal "
        sqlString &= "FROM Member left JOIN GradeLogin_thisyear ON Member.account = GradeLogin_thisyear.memberId "
        sqlString &= "left JOIN GradeShare_thisyear ON Member.account = GradeShare_thisyear.memberId "
        sqlString &= "left JOIN GradeBrowse_thisyear ON Member.account = GradeBrowse_thisyear.memberId "
        sqlString &= "left JOIN MemberGradeSummary B ON Member.account = B.memberId "
        sqlString &= "LEFT JOIN GradeContentBrowse_thisyear VCB ON Member.account = VCB.ieditor "
        sqlString &= "LEFT JOIN GradeContentCommend_thisyear VCC ON Member.account = VCC.ieditor "
        sqlString &= "LEFT JOIN GradeContentDiscuss_thisyear VCD ON Member.account = VCD.ieditor "
        sqlString &= "LEFT JOIN View_MemberGradeAdditional_ThisYear VIEWADDTHIS ON Member.account = VIEWADDTHIS.memberId "
        sqlString &= "LEFT JOIN dbo.MemberGradeContentByYear MGCBY ON Member.account = MGCBY.memberId AND (CONVERT(INT, CONVERT(varchar(4), getdate(), 120 ) ) - ISNULL(years,0) = 0) "
        sqlString &= ") as G1"
        sqlString &= ") as G2 "
        sqlString &= " where account =@account "
        Dim dt As DataTable
        dt = SqlHelper.GetDataTable("GSSConnString", sqlString, DbProviderFactories.CreateParameter("GSSConnString", "@account", "@account", meid))
        If dt.Rows.Count > 0 Then
            Dim dr As DataRow = dt.Rows(0)
            returnValue = Integer.Parse(dr("rowid"))
        End If
        Return returnValue
    End Function

    Function GetKnowledgeActivityInfo(ByVal memberId As String) As String

        Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
        Dim sqlStr As String = ""
        Dim myConnection As SqlConnection
        Dim myCommand As SqlCommand
        Dim myReader As SqlDataReader
        Dim sb As New StringBuilder
        Dim activityGrade As String = ""

        myConnection = New SqlConnection(ConnString)
        Try
            myConnection.Open()

            sqlStr = "SELECT ISNULL(SUM(Total.Grade),0) As Grade FROM "
            sqlStr &= "(	SELECT SUM(KA.Grade) AS Grade FROM dbo.KnowledgeActivity KA "
            sqlStr &= "	WHERE KA.MemberId = @memberid AND type BETWEEN 1 AND 2 "
            sqlStr &= "	GROUP BY KA.MemberId "
            sqlStr &= "	UNION ALL "
            sqlStr &= "	SELECT SUM(temp.Grade) AS Grade "
            sqlStr &= "	FROM( SELECT CASE  WHEN SUM(KA.Grade)>4 THEN 4 ELSE SUM(KA.Grade) END AS Grade  "
            sqlStr &= "		FROM dbo.KnowledgeActivity KA "
            sqlStr &= "		INNER JOIN dbo.KnowledgeForum KF ON ka.CUItemId = KF.gicuitem "
            sqlStr &= "		WHERE KA.MemberId = @memberid AND KA.State =1 AND KA.type BETWEEN 3 AND 4 AND KF.Status = 'N' "
            sqlStr &= "		GROUP BY KF.ParentIcuitem) AS Temp "
            sqlStr &= ") Total "

            myCommand = New SqlCommand(sqlStr, myConnection)
            myCommand.Parameters.AddWithValue("@memberid", memberId)

            activityGrade = myCommand.ExecuteScalar

            myCommand.Dispose()
            'sb.Append("<div class=""content"">	")
            'sb.Append("<table >")
            'sb.Append("<tr><th scope=""col"">知識家活動得分</th>")
            'sb.Append("<th scope=""col"" >可抽獎次數</th></tr>")
            'sb.Append("</tr>")
            'sb.Append("<tr><td>" & activityGrade & "</td><td>" & GetAwardNum(Integer.Parse(activityGrade)) & "</td></tr>")
            'sb.Append("</table>")
            'sb.Append("</div>")



        Catch ex As Exception
            Throw ex
        Finally
            If myConnection.State = ConnectionState.Open Then
                myConnection.Close()
            End If
        End Try

        Return sb.ToString()
    End Function

    Function GetAwardNum(ByVal grade As Integer) As String

        ' 積分達10分以上即可參加抽獎
        Dim num As String
        If grade < 10 Then
            num = "0"
        ElseIf grade >= 10 And grade < 20 Then
            num = "1"
        ElseIf grade >= 20 And grade < 30 Then
            num = "2"
        ElseIf grade >= 30 And grade < 40 Then
            num = "3"
        ElseIf grade >= 40 And grade < 50 Then
            num = "4"
        ElseIf grade >= 50 Then
            num = "5"
        End If

        Return num

    End Function

End Class
