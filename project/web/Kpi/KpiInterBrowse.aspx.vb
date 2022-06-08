
Partial Class Kpi_KpiInterBrowse
    Inherits System.Web.UI.Page

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    Dim memberId As String = Request.QueryString("memberId")
    If memberId IsNot Nothing Or memberId <> "" Then
      Dim browse As New KPIBrowse(memberId, "browseInterCP", Request.QueryString("ctNode"))
      browse.HandleBrowse()
    End If

    Response.Redirect("/ct.asp?xItem=" & Request.QueryString("xItem") & "&ctNode=" & Request.QueryString("ctNode") & "&mp=" & Request.QueryString("mp") & "&kpi=0")

  End Sub

End Class
