<%@ CodePage = 65001 %>
<%
Response.Expires = 0
HTProgCap=""
HTProgFunc=""
HTProgCode="KForumlist"
HTProgPrefix=""
Dim iBaseDSD : iBaseDSD = "39"
Dim iCTUnit : iCTUnit = "934"
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/selectTree.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<link href="/css/list.css" rel="stylesheet" type="text/css">
	<link href="/css/layout.css" rel="stylesheet" type="text/css">
	<title>資料管理／資料上稿</title>
</head>
<% Sub showCantFindBox(lMsg) %>
  <html>
		<head>
			<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
			<meta name="GENERATOR" content="Hometown Code Generator 1.0">			
			<title>編修表單</title>
    </head>
    <body>
			<script language=vbs>
			  alert("<%=lMsg%>")	    				
				history.back
			</script>
    </body>
  </html>
<% End sub %>
<%
	'---if request("submitTask")="submit"時, 寄送Email給專家---
	'---else 秀出表單頁---
	Dim submitTask : submitTask = request("submitTask")	
	if submitTask = "submit" then 
		HandleMail
	else	
		ShowForm
	end if
%>
<% sub ShowForm %>
<body>
<div id="FuncName">
	<h1>資料管理／資料上稿</h1>
	<font size=2>【目錄樹節點: 知識問題】</font>
	<div id="Nav">
	  <a href="/KnowledgeForum/KnowledgeQuestionMail.asp?questionId=<%=request.querystring("questionId")%>&nowPage=<%=request.querystring("nowPage")%>&pagesize=<%=request.querystring("pagesize")%>">回前頁</a>
	</div>
	<div id="ClearFloat"></div>
</div>
<div id="FormName">	
	單元資料維護&nbsp;
	<font size=2>【轉寄問題－轉寄清單一覽】
</div>
<%
	Dim questionId
	questionId = request.querystring("questionId")
	sql = "SELECT Member.account, Member.realname, KnowledgeForum.expertSendDate, KnowledgeForum.expertReplyDate, KnowledgeForum.Status, KnowledgeForum.expertId, "
	sql = sql & "CuDTGeneric.xBody, KnowledgeForum.expertRand, KnowledgeForum.gicuitem FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem "
	sql = sql & "INNER JOIN Member ON KnowledgeForum.expertId = Member.account "
	sql = sql & "WHERE KnowledgeForum.ParentIcuitem = " & questionId & " AND CuDTGeneric.iCTUnit = " & iCTUnit & " ORDER BY KnowledgeForum.expertSendDate DESC"
	Set RSreg = conn.execute(sql)
	if RSreg.eof then
		showCantFindBox "無資料"				
	else
%>
	<Form id="Form2" name="reg" method="POST" >
		<INPUT TYPE="hidden" name="submitTask" value="" />
		<!-- 列表 -->
		<table cellspacing="0" id="ListTable">
		<tr>
		  <th scope="col">專家帳號</th>
			<th scope="col">真實姓名</th>
			<th scope="col">轉寄日期</th>
			<th scope="col">回覆狀態</th>
			<th scope="col">回覆日期</th>
			<th>專家補充內容(摘要)</th>
			<th>&nbsp;</th>				
	  </tr>                  
		<% While not RSreg.eof %>
		<tr>
			<td align="center"><a href="/member/newMemberEdit.asp?account=<%=trim(RSreg("account"))%>&nowPage=<%=request.querystring("nowPage")%>&pagesize=<%=request.querystring("pagesize")%>"><%=trim(RSreg("account"))%></a></td>
			<td align="center"><%=trim(RSreg("realname"))%></td>
			<TD class="eTableContent"><font size=2><%=RSreg("expertSendDate")%>&nbsp;</font></td>
			<td align="center">						
				<% if RSreg("Status") = "W" then response.write "&nbsp;" %>
				<% if RSreg("Status") <> "W" then response.write "V" %>
			</td>
			<TD class="eTableContent"><font size=2><%=RSreg("expertReplyDate")%>&nbsp;</font></td>
			<TD class="eTableContent"><font size=2><A href="KnowledgeExpertContent.asp?questionid=<%=questionId%>&expertaddid=<%=RSreg("gicuitem")%>&nowPage=<%=request.querystring("nowPage")%>&pagesize=<%=request.querystring("pagesize")%>"><%=RSreg("xBody")%>&nbsp;</A></font></td>
			<TD class="eTableContent">
				<span class="mailQ">
				<% if RSreg("Status") = "W" then %>
					<input name="replyButton" type="submit" value="代替回覆" class="cbutton" onclick="RedirectReply('<%=questionId%>','<%=RSreg("gicuitem")%>','<%=RSreg("expertRand")%>','<%=RSreg("expertId")%>')" />
				<% else %>
					<input name="replyButton" type="submit" value="代替回覆" class="cbutton" disabled="disabled" />
				<% end if %>
				</span>
			</td>
		</tr>
		<%
				RSreg.MoveNext
			Wend
		%>
		</TABLE>
		<div align="center">
			<span class="mailQ">
        <input name="back" type="button" value="回前頁" class="cbutton" onClick="RedirectToMail()" />
      </span>
		</div>
	</form>  
<script language="javascript">       
	function RedirectToMail()
	{
		window.location.href = "/KnowledgeForum/KnowledgeQuestionMail.asp?questionId=<%=request.querystring("questionId")%>&nowPage=<%=request.querystring("nowPage")%>&pagesize=<%=request.querystring("pagesize")%>";
	}
	function RedirectReply( qid, did, rid, eid)
	{
		window.open("<%=session("myWWWSiteURL")%>/Knowledge/KnowledgeExpertReply.aspx?ArticleId=" + qid + "&DArticleId=" + did + "&rand=" + rid + "&expertId=" + eid);
	}
</script>
<%
	end if
	RSreg.close
	Set RSreg = nothing
%> 
</body>
</html>
<% end sub %>