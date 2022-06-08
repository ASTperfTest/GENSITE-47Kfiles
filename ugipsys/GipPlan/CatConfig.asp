<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "GE1T21" %>
<!--#include virtual = "/inc/index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<%
	DataParent = Request.Form("DataParent")
	DataLevel = Request.Form("DataLevel")
	CatID = Request.Form("CatID")

	If Request.QueryString("ConfigCk") = "Add" Then
	 SQL = "Select * from CatTreeNode Where CatName = "& NoHTMLCode(Request.Form("CatName")) _
	 	& " AND CtRootID=" & pkStr(ItemID,"") _
	 	&" And DataParent="& pkstr(DataParent,"")
	 Set RS = Conn.execute(SQL)
	  If rs.EOF Then
		 ShowOrder = 1
		 if isNumeric(request.Form("CatShowOrder")) then	ShowOrder = request.Form("CatShowOrder")

		 SQL = "INSERT INTO CatTreeNode "&_
	      	   "(CtRootID, CatName, CatShowOrder, EditDate, EditUserID, DataLevel, "&_
	      	   "DataParent) VALUES (N'"& ItemID &"',"& NoHTMLCode(Request.Form("CatName")) &","& ShowOrder &_
	      	   ",N'"& Date() &"',N'"& session("UserID") &"',"& DataLevel &","& DataParent &")"
		 Set RS = Conn.execute(SQL)

		 SQL = "UpDate CatTreeNode Set ChildCount = ChildCount + 1 Where CtNodeID = "& DataParent
		 Set RS = Conn.execute(SQL)
		 alertmsg = "新增成功！" %>
<script language=VBScript>
  alert("<%=alertmsg%>")
  document.location.href="CatList.asp?ItemID=<%=ItemID%>&CatID=<%=DataParent%>"
  window.parent.Catalogue.location.reload
</script>
<%	  Else %>
        <script language=VBS>
          alert "類別名稱重覆，請重新輸入！"
		  history.back
        </script>
<%	  End IF
	ElseIf Request.QueryString("ConfigCk") = "Edit" Then
	 SQL = "Select * from CatTreeNode Where CatName = "& NoHTMLCode(Request.Form("CatName")) &" And CtRootID = N'"& ItemID &"' And CtNodeID <> "& CatID
	 Set RS = Conn.execute(SQL)
	  If rs.EOF Then
		 ShowOrder = 1
		 if isNumeric(request.Form("CatShowOrder")) then	ShowOrder = request.Form("CatShowOrder")

		 SQL = "UpDate CatTreeNode Set CatName = "& NoHTMLCode(Request.Form("CatName")) &", EditDate = "& pkstr(Date(),",") _
		 	& " EditUserID=" & pkstr(session("UserID"),",") _
		 	& " CatShowOrder=" & ShowOrder _
		 	& " Where CtNodeID = "& CatID
		 Set RS = Conn.execute(SQL)
		 alertmsg = "編修成功！" %>
<script language=VBScript>
  alert("<%=alertmsg%>")
  document.location.href="CatView.asp?ItemID=<%=ItemID%>&CatID=<%=CatID%>"
  window.parent.Catalogue.location.reload
</script>
<%	  Else %>
        <script language=VBS>
          alert "類別名稱重覆，請重新輸入！"
		  history.back
        </script>
<%	  End IF
	ElseIf Request.QueryString("ConfigCk") = "Del" Then
	 SQL = "DELETE FROM CatTreeNode Where CtNodeID = "& CatID
	 Set RS = Conn.execute(SQL)
		 SQL = "UpDate CatTreeNode Set ChildCount = ChildCount - 1 Where CtNodeID = "& DataParent
		 Set RS = Conn.execute(SQL)
		 alertmsg = "刪除成功！" %>
<script language=VBScript>
  alert("<%=alertmsg%>")
  document.location.href="CatList.asp?ItemID=<%=ItemID%>&CatID=<%=DataParent%>"
  window.parent.Catalogue.location.reload
</script>
<%	ElseIf Request.QueryString("ConfigCk") = "Order" Then
	 UnitIDArrar = Split(CatID ,",")
	 NowShowOrder = 0
	  For xno = 0 to UBound(UnitIDArrar)
	   NowShowOrder = NowShowOrder + 1
		sql2 = "Update CatTreeNode set CatShowOrder =" & NowShowOrder & " Where CtNodeID ="& UnitIDArrar(xno)
		set rs2 = conn.execute(sql2)
	  Next
	  alertmsg = "排序完成！" %>
<script language=VBScript>
  alert("<%=alertmsg%>")
  document.location.href="CatList.asp?ItemID=<%=ItemID%>&CatID=<%=DataParent%>"
  window.parent.Catalogue.location.reload
</script>
<%	End If

	Function NoHTMLCode(DataCode)
	  NewData = ""
	    NewDate = Replace(DataCode,"'","''")
	  If DataCode <> "" Then
	   If HTMLDecide = "N" Then
	    NewDate = Replace(NewDate,"<","<")
	    NewDate = Replace(NewDate,">",">")
	   End IF
	  End If
	  NoHTMLCode = "'" & NewDate & "'"
	End Function
%>
