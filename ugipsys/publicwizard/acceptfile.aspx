<%@ Page Language="vb" AutoEventWireup="false" %>
<%@ Import Namespace = "System.Data.SqlClient" %>
<%@ Import Namespace = "System.Data" %>
<%@ Import Namespace = "System.Web" %>
<%@ Import Namespace = "System.Web.Security" %>

<script language="VB" runat="server">
	Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load	
	
		Dim MyPostedMember As HttpPostedFile = Request.Files("myfile")		

		if not MyPostedMember is Nothing then
			Dim MyFileName As String = MyPostedMember.FileName
			if MyFileName <> "" Then
				MyFileName = System.IO.Path.Combine(server.MapPath(".") & "/upload" , System.IO.Path.GetFileName(MyFileName))
				'response.write(MyFileName)
				MyPostedMember.SaveAs(MyFileName)
			end if
		end If	
	End Sub	
</script>