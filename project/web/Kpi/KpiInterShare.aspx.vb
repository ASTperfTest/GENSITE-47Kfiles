Imports GSS.Vitals.COA.Data
Imports System.Data
Partial Class Kpi_KpiInterShare
    Inherits System.Web.UI.Page

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        GSS.Vitals.COA.Data.DbConnectionHelper.LoadSetting()
        Dim sharetype As String = Request.QueryString("type")
        If sharetype = "1" Then
            AddShareVote()
        ElseIf sharetype = "2" Then
            AddSsubjectShareVote()
        End If
        'Dim memberId As String = Request.QueryString("memberId")
        'Dim ibaseDSD As Integer = 7
        'Dim iCtUnit As Integer = 2180
        'Dim xItem As String = Request.QueryString("xItem")
        'Dim icuitem As String = Request.QueryString("icuitem")
        'If memberId IsNot Nothing Or memberId <> "" Then
        '    Dim sql As String = " if exists( SELECT iCUItem FROM CuDTGeneric WHERE icuitem = @icuitem and iCtUnit = @iCtUnit and ieditor =@memberId   AND xabstract is null ) "
        '    sql += " begin "
        '    sql += " update CuDTGeneric set xabstract = 1  WHERE icuitem = @icuitem "
        '    sql += " select 1 as idd "
        '    sql += " End "
        '    sql += " Else "
        '    sql += " begin "
        '    sql += " select 0  as idd"
        '    sql += " end "
        '    Dim dt As DataTable = SqlHelper.GetDataTable("ODBCDSN", sql, _
        '           DbProviderFactories.CreateParameter("HistoryPictureConnString", "@icuitem", "@icuitem", icuitem), _
        '           DbProviderFactories.CreateParameter("HistoryPictureConnString", "@iCtUnit", "@iCtUnit", iCtUnit), _
        '           DbProviderFactories.CreateParameter("HistoryPictureConnString", "@memberId", "@memberId", memberId))
        '    If dt.Rows(0)("idd") = 1 Then
        '        Dim share As New KPIShare(memberId, "shareVote", Request.QueryString("xItem"))
        '        share.HandleShare()
        '    End If
        'End If

        'Response.Redirect("/ct.asp?xItem=" & xItem & "&ctNode=" & Request.QueryString("ctNode") & "&mp=" & Request.QueryString("mp") & "&kpi=" & Request.QueryString("kpi"))


    End Sub

    Private Sub AddShareVote()
        Dim memberId As String = Request.QueryString("memberId")
        Dim ibaseDSD As Integer = 7
        Dim iCtUnit As Integer = 2180
        Dim xItem As String = Request.QueryString("xItem")
        Dim icuitem As String = Request.QueryString("icuitem")
        If memberId IsNot Nothing Or memberId <> "" Then
            Dim sql As String = " if exists( SELECT iCUItem FROM CuDTGeneric WHERE icuitem = @icuitem and iCtUnit = @iCtUnit and ieditor =@memberId   AND xabstract is null ) "
            sql += " begin "
            sql += " update CuDTGeneric set xabstract = 1  WHERE icuitem = @icuitem "
            sql += " select 1 as idd "
            sql += " End "
            sql += " Else "
            sql += " begin "
            sql += " select 0  as idd"
            sql += " end "
            Dim dt As DataTable = SqlHelper.GetDataTable("ODBCDSN", sql, _
                   DbProviderFactories.CreateParameter("HistoryPictureConnString", "@icuitem", "@icuitem", icuitem), _
                   DbProviderFactories.CreateParameter("HistoryPictureConnString", "@iCtUnit", "@iCtUnit", iCtUnit), _
                   DbProviderFactories.CreateParameter("HistoryPictureConnString", "@memberId", "@memberId", memberId))
            If dt.Rows(0)("idd") = 1 Then
                Dim share As New KPIShare(memberId, "shareVote", Request.QueryString("xItem"))
                share.HandleShare()
                AddPuzzleScore(memberId)
            End If
        End If

        Response.Redirect("/ct.asp?xItem=" & xItem & "&ctNode=" & Request.QueryString("ctNode") & "&mp=" & Request.QueryString("mp") & "&kpi=" & Request.QueryString("kpi"))
    End Sub

    Private Sub AddSsubjectShareVote()
        Dim memberId As String = Request.QueryString("memberId")
        Dim iCtUnit As Integer = 2751
        Dim icuitem As String = Request.QueryString("icuitem")
        Dim refutl As String = Request.QueryString("refurl")
        If memberId IsNot Nothing Or memberId <> "" Then
            Dim sql As String = " if exists( SELECT iCUItem FROM CuDTGeneric WHERE icuitem = @icuitem and iCtUnit = @iCtUnit and ieditor =@memberId   AND xabstract is null ) "
            sql += " begin "
            sql += " update CuDTGeneric set xabstract = 1  WHERE icuitem = @icuitem "
            sql += " select 1 as idd "
            sql += " End "
            sql += " Else "
            sql += " begin "
            sql += " select 0  as idd"
            sql += " end "

            Dim dt As DataTable = SqlHelper.GetDataTable("ODBCDSN", sql, _
                   DbProviderFactories.CreateParameter("HistoryPictureConnString", "@icuitem", "@icuitem", icuitem), _
                   DbProviderFactories.CreateParameter("HistoryPictureConnString", "@iCtUnit", "@iCtUnit", iCtUnit), _
                   DbProviderFactories.CreateParameter("HistoryPictureConnString", "@memberId", "@memberId", memberId))
            If dt.Rows(0)("idd") = 1 Then
                Dim share As New KPIShare(memberId, "shareSubjectCommend", Request.QueryString("xItem"))
                share.HandleShare()
                AddPuzzleScore(memberId)
            End If
        End If
        'Response.Write("<script language=""javascript"">alert(""評價成功，感謝您對主題館的支持!!"");window.location.href =""" & refutl & """;</script> "")")
        Response.Redirect(refutl)
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
