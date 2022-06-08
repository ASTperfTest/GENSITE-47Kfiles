<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "Pn03M04" %>
<!--#include virtual = "/inc/index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<%
	CatID = Request.Form("CatID")
	ForumID = Request.Form("ForumID")

	If Request.QueryString("ConfigCk") = "Add" Then
	 SQL = "Select * from Forum Where ForumName = "& NoHTMLCode(Request.Form("ForumName")) &" And ItemID = N'"& ItemID &"' And CatID ="& CatID
	 Set RS = Conn.execute(SQL)
	  If rs.EOF Then
		 ShowOrder = 1
		 SQLCom = "Select * from Forum Where ItemID = N'"& ItemID &"' And CatID = "& CatID &" Order by ForumShowOrder"
		 Set RSCom = Conn.execute(SQLCom)
		  If Not RSCom.EOF Then
		   NowShowOrder = 2
		   Do while not RSCom.EOF
		    SQL = "UpDate Forum Set ForumShowOrder = "& NowShowOrder &" Where ForumID = "& RSCom("ForumID")
		    Set RS = Conn.execute(SQL)
		    NowShowOrder = NowShowOrder + 1
	       RSCom.movenext
	       Loop
		  End If
		 SQL = "INSERT INTO Forum "&_
	      	   "(ItemID, ForumName, ForumShowOrder, EditDate, EditUserID, CatID, "&_
	      	   "ForumRemark, ForumMaster) VALUES (N'"& ItemID &"',"& NoHTMLCode(Request.Form("ForumName")) &","& ShowOrder &_
	      	   ",N'"& Date() &"',N'"& session("UserID") &"',"& CatID &","& NoHTMLCode(Request.Form("ForumRemark")) &","& Request.Form("ForumMaster") &")"
		 Set RS = Conn.execute(SQL)
		 SQL = "UpDate Catalogue Set DataCount = DataCount + 1 Where CatID = "& CatID
		 Set RS = Conn.execute(SQL) %>
<script language=VBScript>
  alert("新增完成！")
  document.location.href="CatView.asp?ItemID=<%=ItemID%>&CatID=<%=CatID%>"
  window.parent.Catalogue.location.reload
</script>
<%	  Else %>
        <script language=VBS>
          alert "討論板名稱重覆，請重新輸入！"
		  history.back
        </script>
<%	  End IF
	ElseIf Request.QueryString("ConfigCk") = "Edit" Then
	 SQL = "Select * from Forum Where ForumName = "& NoHTMLCode(Request.Form("ForumName")) &" And ItemID = N'"& ItemID &"' And CatID ="& CatID &" And ForumID <> "& ForumID
	 Set RS = Conn.execute(SQL)
	  If rs.EOF Then
		 SQL = "UpDate Forum Set ForumRemark = "& NoHTMLCode(Request.Form("ForumRemark")) &",ForumName = "& NoHTMLCode(Request.Form("ForumName")) &", EditDate = N'"& Date() &"', EditUserID = N'"& session("UserID") &"' Where ForumID = "& ForumID
		 Set RS = Conn.execute(SQL) %>
<script language=VBScript>
  alert("編修成功！")
  document.location.href="ForumView.asp?ItemID=<%=ItemID%>&ForumID=<%=ForumID%>"
  window.parent.Catalogue.location.reload
</script>
<%	  Else %>
        <script language=VBS>
          alert "討論板名稱重覆，請重新輸入！"
		  history.back
        </script>
<%	  End IF
	ElseIf Request.QueryString("ConfigCk") = "Del" Then
	 SQL = "DELETE FROM Forum Where ForumID = "& ForumID
	 Set RS = Conn.execute(SQL)
		 SQLCom = "Select * from ForumArticle Where ForumID = "& ForumID
		 Set RSCom = Conn.execute(SQLCom)
		  If Not RSCom.EOF Then
		   Do while not RSCom.EOF
		    SQLCom1 = "Select * from FileUp Where ItemID = N'"& ItemID &"' And ParentID = "& RSCom("ArticleID")
		    Set RSCom1 = Conn.execute(SQLCom1)
		     If Not RSCom1.EOF Then
		      Do while not RSCom1.EOF
				 set dfile = server.createobject("scripting.filesystemobject")
				 ifile = server.mappath(session("Public") & RSCom1("NFileName"))
				 dfile.deletefile(ifile)
			     xSQL = "DELETE FROM FileUp Where ItemID = N'"& ItemID &"' And ParentID = "& RSCom("ArticleID")
			 	 Set xRS = Conn.execute(xSQL)
		      RSCom1.movenext
		      Loop
		     End IF
	       RSCom.movenext
	       Loop
		  End If
	     SQL = "DELETE FROM ForumArticle Where ForumID = "& ForumID
	 	 Set RS = Conn.execute(SQL)
	     SQL = "DELETE FROM ForumEssenceArticle Where ForumID = "& ForumID
	 	 Set RS = Conn.execute(SQL)
		 SQLCom = "Select * from Forum Where ItemID = N'"& ItemID &"' And CatID = "& CatID &" Order by ForumShowOrder"
		 Set RSCom = Conn.execute(SQLCom)
		  If Not RSCom.EOF Then
		   NowShowOrder = 1
		   Do while not RSCom.EOF
		    SQL = "UpDate Forum Set ForumShowOrder = "& NowShowOrder &" Where ForumID = "& RSCom("ForumID")
		    Set RS = Conn.execute(SQL)
		    NowShowOrder = NowShowOrder + 1
	       RSCom.movenext
	       Loop
		  End If
		 SQL = "UpDate Catalogue Set DataCount = DataCount - 1 Where CatID = "& CatID
		 Set RS = Conn.execute(SQL) %>
<script language=VBScript>
  alert("刪除成功！")
  document.location.href="page.asp"
  window.parent.Catalogue.location.reload
</script>
<%	ElseIf Request.QueryString("ConfigCk") = "Order" Then
	 UnitIDArrar = Split(ForumID,",")
	 CatID = Request.QueryString("CatID")
	 NowShowOrder = 0
	  For xno = 0 to UBound(UnitIDArrar)
	   NowShowOrder = NowShowOrder + 1
		sql2 = "Update Forum set ForumShowOrder =" & NowShowOrder & " Where ForumID ="& UnitIDArrar(xno)
		set rs2 = conn.execute(sql2)
	  Next %>
<script language=VBScript>
  alert("排序成功！")
  document.location.href="CatView.asp?ItemID=<%=ItemID%>&CatID=<%=CatID%>&DataParent=<%=CatID%>"
  window.parent.Catalogue.location.reload
</script>
<%	ElseIf Request.QueryString("ConfigCk") = "DelA" Then
	 ForumID = Request.QueryString("ForumID")
	 ArticleID = Request.Form("ArticleID")
	 UnitIDArrar = Split(ArticleID,",")
	  For xno = 0 to UBound(UnitIDArrar)
		    SQLCom1 = "Select * from FileUp Where ItemID = N'"& ItemID &"' And ParentID ="& UnitIDArrar(xno)
		    Set RSCom1 = Conn.execute(SQLCom1)
		     If Not RSCom1.EOF Then
		      Do while not RSCom1.EOF
				 set dfile = server.createobject("scripting.filesystemobject")
				 ifile = server.mappath(session("Public") & RSCom1("NFileName"))
				 dfile.deletefile(ifile)
			     xSQL = "DELETE FROM FileUp Where ItemID = N'"& ItemID &"' And ParentID ="& UnitIDArrar(xno)
			 	 Set xRS = Conn.execute(xSQL)
				 xxxSQL = "DELETE FROM ForumEssenceArticle Where ArticleID = "& UnitIDArrar(xno)
				 Set xxxRS = Conn.execute(xxxSQL)
		      RSCom1.movenext
		      Loop
		     End IF
	     xxSQL = "DELETE FROM ForumArticle Where ArticleID ="& UnitIDArrar(xno)
	 	 Set xxRS = Conn.execute(xxSQL)
	  Next %>
<script language=VBScript>
  alert("文章刪除成功！")
  window.parent.parent.ForumToc.location.href="ArticleView.asp?ItemID=<%=ItemID%>&ForumID=<%=ForumID%>"
  window.parent.parent.Catalogue.location.reload
</script>
<%	ElseIf Request.QueryString("ConfigCk") = "ForumEs" Then
	 ForumID = Request.QueryString("ForumID")
	 ArticleID = Request.Form("ArticleID")
	 ForumEs = Request.QueryString("ForumEs")
	 SQL = "DELETE FROM ForumEssenceArticle Where ForumID = "& ForumID
	 Set RS = Conn.execute(SQL)
	 UnitIDArrar = Split(ArticleID,",")
	  For xno = 0 to UBound(UnitIDArrar)
	 	 SQLCom = "INSERT INTO ForumEssenceArticle (ArticleID, ForumID) VALUES ("& UnitIDArrar(xno) &","& ForumID &")"
	 	 Set RS = Conn.execute(SQLCom)
	  Next %>
<script language=VBScript>
  alert("精華區編修成功！")
  window.parent.parent.ForumToc.location.href="ArticleView.asp?ItemID=<%=ItemID%>&ForumID=<%=ForumID%>&ForumEs=<%=ForumEs%>"
  window.parent.parent.Catalogue.location.reload
</script>
<%	End IF

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
