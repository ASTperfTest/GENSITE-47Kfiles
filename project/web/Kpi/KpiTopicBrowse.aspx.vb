
Partial Class Kpi_KpiTopicBrowse
    Inherits System.Web.UI.Page

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    Dim memberId As String = Session("memID")
    If memberId IsNot Nothing Or memberId <> "" Then
      Dim browse As New KPIBrowse(memberId, "browseTopicCP", Request.QueryString("ctNode"))
      browse.HandleBrowse()
    End If

    Response.Redirect("/subject/ct.asp?xItem=" & Request.QueryString("xItem") & "&ctNode=" & Request.QueryString("ctNode") & "&mp=" & Request.QueryString("mp") & "&kpi=0")

  End Sub

End Class
