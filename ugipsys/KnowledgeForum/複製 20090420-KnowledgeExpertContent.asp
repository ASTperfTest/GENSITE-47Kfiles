<%
Response.Expires = 0
HTProgCap=""
HTProgFunc=""
HTProgCode="KForumlist"
HTProgPrefix=""
%>
<!--#include virtual = "/inc/server.inc" -->

<%
Dim questionid : questionid = request.querystring("questionid")
Dim expertaddid : expertaddid = request.querystring("expertaddid")
	
if request("submitTask") = "edit" then

	Dim stitle : stitle = pkStr(request("stitle"), "")
	Dim xbody : xbody = pkStr(request("xbody"), "")
	Dim fctupublic : fctupublic = pkStr(request("fctupublic"), "")
	
	sql = "UPDATE CuDTGeneric SET stitle = " & stitle & ", xbody = " & xbody & ", fctupublic = " & fctupublic & " WHERE iCUItem = " & expertaddid
	conn.execute(sql)
	
	showDoneBox "資料更新成功！"
elseif request("submitTask") = "delete" then

	
	sql = "UPDATE KnowledgeForum SET Status = 'D' WHERE gicuitem = " & expertaddid
	conn.execute(sql)
	
	showDoneBox "資料刪除成功！"
end if
function pkStr (s, endchar)
	if s = "" then
		pkStr = "''" & endchar		
	else
		pos = InStr(s, "'")
		While pos > 0
			s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)
			pos = InStr(pos + 2, s, "'")
		Wend										
		pkStr = "N'" & s & "'" & endchar		
	end if
END FUNCTION
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<link href="../css/form.css" rel="stylesheet" type="text/css">
	<link href="../css/layout.css" rel="stylesheet" type="text/css">
	<title>資料管理／資料上稿</title>
</head>

<body>
	<div id="FuncName">
		<h1>資料管理／資料上稿</h1>
		<font size=2>【目錄樹節點: 知識問題】</font>
		<div id="Nav">	    
	    <a href="javascript:history.go(-1)">回前列</a>	
		</div>
		<div id="ClearFloat"></div>
	</div>
	<div id="FormName">
	 	單元資料維護&nbsp;
    <font size=2>【知識問題 / 專家補充】
	</div>
<% Sub showDoneBox(lMsg) %>
  <html>
		<head>
			<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
			<meta name="GENERATOR" content="Hometown Code Generator 1.0">
			<meta name="ProgId" content="<%=HTprogPrefix%>Edit.asp">
			<title>編修表單</title>
    </head>
    <body>
			<script language=vbs>
			  alert("<%=lMsg%>")	  
			  document.location.href="KnowledgeExpertAdditionalList.asp?questionId=<%=questionId%>"
				//history.back()				
			</script>
    </body>
  </html>
<% End sub %>
<%
	
	
	sql = "SELECT * FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem "
	sql = sql & "WHERE iCUItem = " & expertaddid
	
	Set RSreg = conn.execute(sql)
	if RSreg.eof then
		'response.write "<script>alert('找不到資料');history.back();</script>"
		showDoneBox "無專家補充資料"		
	else
%>
		<form id="Form1" method="POST" name="reg" action="KnowledgeExpertContent.asp?questionid=<%=request.querystring("questionid")%>&expertaddid=<%=request.querystring("expertaddid")%>" >
			<INPUT TYPE=hidden name=submitTask value="">
			<INPUT TYPE=hidden name=showType value="">
			<INPUT TYPE=hidden name=refID value="">
			<input type="hidden" name="icuitem">		
			<table cellspacing="0">
			<TR><TD colspan="6" class="eTableContent"></TD></TR>			
			<TR>
				<TD class="Label" align="right"><span class="Must">*</span>標題</TD>
				<TD colspan="6" class="eTableContent"><input name="stitle" type="text" size="50" value="<%=RSreg("sTitle")%>" ></TD>
			</TR>
			<TR>
				<TD class="Label" align="right">內文</TD>
				<TD colspan="6" class="eTableContent"><textarea name="xbody" rows="8" cols="60" ><%=trim(RSreg("xbody"))%></textarea></TD>
			</TR>			
			<TR>
				<TD class="Label" align="right"><span class="Must">*</span>是否公開(草稿)</TD>
				<TD colspan="6" class="eTableContent">
					<Select name="fctupublic" size=1>
						<option value="" <%if RSreg("fctupublic")="" then%>  selected <%end if%>>請選擇</option>
						<option value="Y" <%if RSreg("fctupublic")="Y" then%>  selected <%end if%>>公開</option>
						<option value="N" <%if RSreg("fctupublic")="N" then%>  selected <%end if%>>不公開</option>
					</select>
				</TD>
			</TR>
		<%	
			sql = "SELECT account, realname, nickname FROM Member WHERE account = '" & Rsreg("expertId") & "'"
			Set RS2 = conn.execute(sql)
			if not RS2.eof then
		%>
			<TR>
				<TD class="Label" align="right">發表專家</TD>
				<TD colspan="6" class="eTableContent">
			    <input name="account" class="rdonly" value="<%=trim(RS2("account"))%>" size="10" readonly="true">
					<input name="realname" class="rdonly" value="<%=trim(RS2("realname"))%>" size="10" readonly="true">
					<input name="nickname" class="rdonly" value="<%=trim(RS2("nickname"))%>" size="10" readonly="true">
				</TD>
			</TR>
		<%
			end if
			rs2.close
			set rs2 = nothing
			sql = "SELECT DiscussCount, CommandCount, BrowseCount, TraceCount, GradeCount, GradePersonCount " & _
						"FROM KnowledgeForum WHERE gicuitem = " & questionid
			Set RS3 = conn.execute(sql)
			if not rs3.eof then
			%>
			<TR>
				<TD rowspan="2" align="right" class="Label">統計</TD>
				<TD nowrap class="eTableContent"><span class="Label">討論數</span></TD>
				<TD class="eTableContent"><input name="htx_DiscussCount" size="10" value="<%=trim(RS3("DiscussCount"))%>"readonly="true" class="rdonly"></TD>
				<td nowrap class="eTableContent"><span class="Label">意見數</span></td>
				<TD class="eTableContent"><input name="htx_CommandCount" size="10" value="<%=trim(RS3("CommandCount"))%>" readonly="true" class="rdonly"></TD>
				<TD nowrap class="eTableContent"><span class="Label">瀏覽數</span></TD>
				<TD class="eTableContent"><input name="htx_BrowseCount" size="10" value="<%=trim(RS3("BrowseCount"))%>" readonly="true" class="rdonly"></TD>
			</TR>
			<TR>
				<TD nowrap class="eTableContent"><span class="Label">追蹤數</span></TD>
				<TD class="eTableContent"><input name="htx_TraceCount" size="10" value="<%=trim(RS3("TraceCount"))%>" readonly="true" class="rdonly"></TD>
				<td nowrap class="eTableContent"><span class="Label">評價數</span></td>
				<TD class="eTableContent"><input name="htx_GradeCount" size="10" value="<%=trim(RS3("GradeCount"))%>" readonly="true" class="rdonly"></TD>
				<TD nowrap class="eTableContent"><span class="Label">評價人數</span></TD>
				<TD class="eTableContent"><input name="htx_GradePersonCount" size="10" value="<%=trim(RS3("GradePersonCount"))%>" readonly="true" class="rdonly"></TD>
			</TR>
			<%
			end if
			rs3.close
			set rs3 = nothing
			%>	
			<TR>
				<TD class="Label" align="right">狀態</TD>
				<TD colspan="6" class="eTableContent">
					<input name="htx_Status" size="10"  class="rdonly" readonly="true" value="<%=RSreg("STATUS")%>" >
					<INPUT TYPE="hidden" name="status" value="<%=RSreg("status")%>">
				</TD>
			</TR>			
			</TABLE>			
      <input type=button value="編修存檔" class="cbutton" onclick="formUpdateSubmit();">
      <input type=button value ="刪　除" class="cbutton" onClick="formDelSubmit();">   
      <input type=button value ="重　填" class="cbutton" onClick="returnToEdit();">
      <input type=button value ="回前頁" class="cbutton" onClick="formBack();">       
		</form>          
<%
	end if
	RSreg.close
	Set RSreg = nothing
%> 
</body>
</html>
<script language="javascript">
	function formUpdateSubmit() {
		document.getElementById("submitTask").value = "edit";
		document.getElementById("reg").submit();
	}
	function formDelSubmit() {
		document.getElementById("submitTask").value = "delete";
		document.getElementById("reg").submit();
	}
	function returnToEdit() {
		document.getElementById("submitTask").value = "reset";
		document.getElementById("reg").submit();
	}
	function formBack() {
		window.location.href = "/KnowledgeForum/KnowledgeExpertAdditionalList.asp?questionId=<%=questionId%>";
	}
</script>