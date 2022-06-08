
Partial Class knowledge_TabText
    Inherits System.Web.UI.UserControl

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Grace"我要推薦"button
        Dim currentPath As String
        currentPath = Request.PhysicalPath
        If InStr(currentPath, "myknowledge_") Or InStr(currentPath, "recommand_Mylist.aspx") Then
            recomd.Text = "<a href='/recommand/recommand_Add.aspx'><img STYLE='RIGHT:350px; cursor:pointer;' border='0' src='../xslGip/style3/images/推薦bt.gif' alt='推薦好文' /></a>  "
        End If

        GroupText.Text = "<ul class=""group"">"

        If Request.RawUrl.IndexOf("/aintroduction") > 0 Or Request.RawUrl.IndexOf("/arules") > 0 _
           Or Request.RawUrl.IndexOf("/astatement") > 0 Or Request.RawUrl.IndexOf("/knowledge_alp") > 0 Then

            GroupText.Text &= "<li><a href=""/knowledge/knowledge.aspx"">首頁</a></li>"
            GroupText.Text &= "<li><a href=""/knowledge/knowledge_lp.aspx"">知識分類</a></li>"
            GroupText.Text &= "<li><a href=""/knowledge/myknowledge_question_lp.aspx"">我的知識</a></li>"

        ElseIf Request.RawUrl.IndexOf("/knowledge.aspx") > 0 Then

            GroupText.Text &= "<li class=""activity""><a href=""/knowledge/knowledge.aspx"">首頁</a></li>"
            GroupText.Text &= "<li><a href=""/knowledge/knowledge_lp.aspx"">知識分類</a></li>"
            GroupText.Text &= "<li><a href=""/knowledge/myknowledge_question_lp.aspx"">我的知識</a></li>"

        ElseIf Request.RawUrl.IndexOf("/knowledge_lp.aspx") > 0 Or Request.RawUrl.IndexOf("/knowledge_cp.aspx") > 0 _
            Or Request.RawUrl.IndexOf("/knowledge_discuss_cp.aspx") > 0 Or Request.RawUrl.IndexOf("/knowledge_discuss_lp.aspx") > 0 _
            Or Request.RawUrl.IndexOf("/knowledge_message.aspx") > 0 Or Request.RawUrl.IndexOf("/knowledge_professional.aspx") > 0 _
            Or Request.RawUrl.IndexOf("/knowledge_discuss.aspx") > 0 Then

            GroupText.Text &= "<li><a href=""/knowledge/knowledge.aspx"">首頁</a></li>"
            GroupText.Text &= "<li class=""activity""><a href=""/knowledge/knowledge_lp.aspx"">知識分類</a></li>"
            GroupText.Text &= "<li><a href=""/knowledge/myknowledge_question_lp.aspx"">我的知識</a></li>"

        ElseIf Request.RawUrl.IndexOf("/myknowledge_lp.aspx") > 0 Or Request.RawUrl.IndexOf("/myknowledge_discuss.aspx") > 0 _
            Or Request.RawUrl.IndexOf("/myknowledge_question.aspx") > 0 Or Request.RawUrl.IndexOf("/myknowledge_question_draft.aspx") > 0 _
            Or Request.RawUrl.IndexOf("/myknowledge_trace.aspx") > 0 Or Request.RawUrl.IndexOf("/knowledge_question.aspx") > 0 _
            Or Request.RawUrl.IndexOf("/myknowledge_record.aspx") > 0 Or Request.RawUrl.IndexOf("/myknowledge_pedia.aspx") > 0 Then

            GroupText.Text &= "<li><a href=""/knowledge/knowledge.aspx"">首頁</a></li>"
            GroupText.Text &= "<li><a href=""/knowledge/knowledge_lp.aspx"">知識分類</a></li>"
            GroupText.Text &= "<li class=""activity""><a href=""/knowledge/myknowledge_question_lp.aspx"">我的知識</a></li>"

        Else

            GroupText.Text &= "<li><a href=""/knowledge/knowledge.aspx"">首頁</a></li>"
            GroupText.Text &= "<li><a href=""/knowledge/knowledge_lp.aspx"">知識分類</a></li>"
            GroupText.Text &= "<li><a href=""/knowledge/myknowledge_question_lp.aspx"">我的知識</a></li>"

        End If

        '新增活動頁籤 
        Dim Activity As New ActivityFilter(System.Web.Configuration.WebConfigurationManager.AppSettings("ActivityId"), "")
        Dim ActivityFlag As Boolean = Activity.CheckActivity
        If ActivityFlag Then
            GroupText.Text &= "<li class=""knowledgeActivity"" ><a href=""/knowledge/activity_introduce.aspx"" >知識問答你˙我˙他 活動</a></li>"
        End If

        'GroupText.Text &= "<li class=""activity""><a href=""http://kmbeta.coa.gov.tw/knowledge/knowledge_activityrank.aspx"">排行榜測試</a></li>"

        GroupText.Text &= "</ul>"

        AskQuestionText.Text = "<a href=""/knowledge/knowledge_question.aspx?BackUrl=" & HttpUtility.UrlEncode(Request.RawUrl().Replace("&", ";")) & """><img src=""../xslGip/style3/images/btn_1.gif"" alt=""我要發問"" border=""0"" /></a>"


    End Sub
End Class
