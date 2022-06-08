<%@ CodePage = 65001 %>
<%
Response.Expires = 0
HTProgCap=""
HTProgFunc=""
HTProgCode="KForumlist"
HTProgPrefix = "KnowledgeCommand" 
Dim iCtUnitId : iCtUnitId = "935"

%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/selectTree.inc" -->
<!-- #INCLUDE FILE="KpiFunction.inc" -->
<%
Dim questionId : questionId = request.querystring("questionId")
Dim discussId : discussId = request.querystring("discussId")
Dim commandid : commandid = request.querystring("commandid")

if request("submitTask") = "reset" then
	showForm
elseif request("submitTask") = "edit" then

	Dim xBody : xBody = pkStr(request("xBody"), "")
	Dim xURL : xURL = pkStr(request("xURL"), "")
	Dim xNewWindow : xNewWindow = pkStr(request("xNewWindow"), "")
	Dim fCTUPublic : fCTUPublic = pkStr(request("fCTUPublic"), "")
	Dim vGroup : vGroup = pkStr(request("vGroup"), "")
	Dim xKeyword : xKeyword = pkStr(request("xKeyword"), "")
	Dim xImportant : xImportant = request("xImportant")
	Dim iDept : iDept = pkStr(request("iDept"), "")
	Dim KStatus : KStatus = pkStr(request("KStatus"), "")
	if xImportant = "" then xImportant = 0
	
	'
	sql_ch = "SELECT fCTUPublic FROM CuDTGeneric WHERE iCUItem = " & commandid
	set rs_ch = conn.execute(sql_ch)
	ch_key = rs_ch("fCTUPublic")
	if ch_key <> request("fCTUPublic") then
		if request("fCTUPublic") = "Y" then
			sql_up = "UPDATE KnowledgeForum SET CommandCount = CommandCount + 1 " & _
					"WHERE gicuitem = " & discussId			
			conn.execute(sql_up)
		elseif request("fCTUPublic") = "N" then
			sql_up = "UPDATE KnowledgeForum SET CommandCount = CommandCount - 1 " & _
					"WHERE gicuitem = " & discussId			
			conn.execute(sql_up)
		end if
	end if
	'
	
	sql = "UPDATE CuDTGeneric SET xBody = " & xBody & ", xURL = " & xURL & ", xNewWindow = " & xNewWindow & ", " & _
				"fCTUPublic = " & fCTUPublic & ", vGroup = " & vGroup & ", xKeyword = " & xKeyword & ", xImportant = " & xImportant & ", " & _
				"iDept = " & iDept & " WHERE iCUItem = " & commandid
	conn.execute(sql)
	'刪除救回
	sql2_1 = "SELECT * FROM KnowledgeForum WHERE gicuitem = " & commandid
	set rs2_1 = conn.execute(sql2_1)
	if request("KStatus") = "N" and rs2_1("Status") <> KStatus then
		sql3 = "UPDATE KnowledgeForum SET CommandCount = CommandCount +1 WHERE gicuitem = " & questionId
		conn.execute(sql3)
		sql3_2 = "UPDATE KnowledgeForum SET CommandCount = CommandCount +1 WHERE gicuitem = " & discussId
		conn.execute(sql3_2)
	elseif request("KStatus") = "D" and rs2_1("Status") <> KStatus then
		sql3 = "UPDATE KnowledgeForum SET CommandCount = CommandCount -1 WHERE gicuitem = " & questionId
		conn.execute(sql3)
		sql3_2 = "UPDATE KnowledgeForum SET CommandCount = CommandCount -1 WHERE gicuitem = " & discussId
		conn.execute(sql3_2)
	end if
	sql2 = "UPDATE KnowledgeForum SET Status = "& KStatus &" WHERE gicuitem = " & commandid
	conn.execute(sql2)
	'刪除救回end
	showDoneBox "資料更新成功！"
	
elseif request("submitTask") = "delete" then

	sql = "SELECT Status FROM KnowledgeForum WHERE gicuitem = " & discussId	
	set rs = conn.execute(sql)
	if not rs.eof then
		if rs("Status") = "D" then 
			response.write "<script>alert('article can not be deleted!!');window.location.href='KnowledgeForumlist.asp';</script>"
			response.end
		end if
	end if
	sql = "SELECT Status FROM KnowledgeForum WHERE gicuitem = " & commandid	
	set rs = conn.execute(sql)
	if not rs.eof then
		if rs("Status") = "D" then 
			response.write "<script>alert('article can not be deleted!!');window.location.href='KnowledgeForumlist.asp';</script>"
			response.end
		end if
	end if
	
	DeleteCommend commandid '---刪除評價---	
	DeleteOpinion commandid '---刪除意見---			
	DeleteDiscuss commandid '---刪除討論---		
		
	'---更新parent的count---				
	sql = "SELECT * FROM KnowledgeForum WHERE gicuitem = " & commandid
	set delrs = conn.execute(sql)
	while not delrs.eof 
		CommandCount = CInt(delrs("CommandCount"))
		GradeCount = CInt(delrs("GradeCount"))
		GradePersonCount = CInt(delrs("GradePersonCount"))
		ParentIcuitem = delrs("ParentIcuitem")					
		sql = "UPDATE KnowledgeForum SET DiscussCount = DiscussCount - 1, CommandCount = CommandCount - 1, " & _
					"GradeCount = GradeCount - " & GradeCount & ", GradePersonCount = GradePersonCount - " & GradePersonCount & " " & _
					"WHERE gicuitem = " & discussId			
		conn.execute(sql)	
response.write sql		
		delrs.movenext					
	wend				
	delrs.close
	set delrs = nothing
	'刪除
	sql2 = "UPDATE KnowledgeForum SET Status = 'D' WHERE gicuitem = " & commandid
	conn.execute(sql2)
	'刪除end
	showDoneBox "刪除完成！"
else
	showForm
end if
%>
<%
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
				document.location.href="<%=HTprogPrefix%>List.asp?questionId=<%=questionId%>&discussId=<%=discussId%>"
			</script>
    </body>
  </html>
<% End sub %>
<%
sub showForm

	'-------------------------------------------------------------
	Dim questionTitle
	sql = "SELECT sTitle FROM CuDTGeneric WHERE iCUItem = " & questionId
	set rs = conn.execute(sql)
	if not rs.eof then
		questionTitle = rs("sTitle")
	end if
	rs.close
	set rs = nothing
	'-------------------------------------------------------------

	sql = "SELECT *, KnowledgeForum.Status AS KStatus FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem " & _
				"INNER JOIN Member ON CuDTGeneric.iEditor = Member.account  " & _
				"WHERE (CuDTGeneric.iCUItem = " & commandid & ") AND (CuDTGeneric.iCTUnit = " & iCtUnitId & ") "	
	set rs = conn.execute(sql)			
	if rs.eof then
		response.write "<script>alert('找不到資料');history.back();</script>"
		response.end
	else		
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="/css/form.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<title></title>
</head>

<body>
<div id="FuncName">
	<h1>資料管理／資料上稿</h1>
	<font size=2>【目錄樹節點: 知識意見】</font>
	<div id="Nav">
    <!-- a href="qp_questionmail.htm">轉寄問題</a>&nbsp; -->
	<a href="KnowledgeCommandList.asp?questionId=<%=request.querystring("questionId")%>&discussId=<%=request.querystring("discussId")%>&nowPage=<%=request.querystring("nowPage")%>&pagesize=<%=request.querystring("pagesize")%>">意見管理</a>&nbsp;
    <a href="/KnowledgeForum/KnowledgeForumList.asp?nowPage=<%=request.querystring("nowPage")%>&pagesize=<%=request.querystring("pagesize")%>">回條列</a>	
	</div>
	<div id="ClearFloat"></div>
</div>
<div id="FormName">
	單元資料維護&nbsp;
  <font size=2>【編輯(<font color=red>一般資料式</font>)--主題單元:(知識家)-知識意見 / 單元資料:知識家討論區】
</div>

	<form id="Form1" method="POST" name="reg" action="KnowledgeCommandContent.asp?questionId=<%=questionId%>&discussId=<%=discussId%>&commandid=<%=commandid%>" >
		<INPUT TYPE="hidden" name="submitTask" value="">			
		<table cellspacing="0">			
		<TR>
			<TD class="Label" align="right">
				<span class="Must">*</span>標題</TD>
				<TD class="eTableContent"><input name="sTitle" size="50" value="<%=questionTitle%>" readonly="true" class="rdonly">
			</TD>
		</TR>
		<TR>
			<TD class="Label" align="right">內文</TD>
			<TD class="eTableContent">
				<textarea name="xBody" rows="8" cols="60"" title="內容、本文、網頁、說明""><%=rs("xBody")%></textarea>
			</TD>
		</TR>
		<TR>
			<TD class="Label" align="right">網址</TD>
			<TD class="eTableContent"><input name="xURL" size="50" title="連結" value="<%=rs("xURL")%>"></TD>
		</TR>
		<TR>
			<TD class="Label" align="right">討論關閉</TD>
			<TD class="eTableContent">
				<select name="xNewWindow" size="1">
					<option value="">請選擇</option>
					<option value="Y" <% if rs("xNewWindow") = "Y" Then %> selected <% end if %>>是</option>
					<option value="N" <% if rs("xNewWindow") = "N" Then %> selected <% end if %>>否</option>
				</select>
			</TD>
		</TR>
		<TR>
			<TD class="Label" align="right"><span class="Must">*</span>是否公開(草稿)</TD>
			<TD class="eTableContent">
				<select name="fCTUPublic" size="1">
					<option value="">請選擇</option>
					<option value="Y" <% if rs("fCTUPublic") = "Y" Then %> selected <% end if %>>公開</option>
					<option value="N" <% if rs("fCTUPublic") = "N" Then %> selected <% end if %>>不公開</option>
				</select>		
			</TD>
		</TR>
		<TR>
			<TD class="Label" align="right">發問主題</TD>
			<TD class="eTableContent">
				<table width="100%">
				<tr>
					<td><input type="checkbox" name="vGroup" value="A" <% if rs("vGroup") = "A" then %> checked <% end if %> ><font size=2>疑難知識</font></td>					
				</tr>
				</table>
			</TD>
		</TR>
		<TR>
			<TD class="Label" align="right">關鍵字詞</TD>
			<TD class="eTableContent"><input name="xKeyword" size="50" title="請以,分隔" value="<%=rs("xKeyword")%>"></TD>
		</TR>
		<TR>
			<TD class="Label" align="right">重要性</TD>
			<TD class="eTableContent"><input name="xImportant" size="2" title="(不重要) 0~99 (重要)" value="<%=rs("xImportant")%>"></TD>
		</TR>
		<TR>
			<TD class="Label" align="right"><span class="Must">*</span>單位</TD>
			<TD class="eTableContent">
			<%				 		      		       
				response.write "<Select name=""iDept"" size=""1"">" & vbCRLF	
				sqlCom = "SELECT D.deptId,D.deptName,D.parent,len(D.deptId)-1,D.seq," & _
								 "(Select Count(*) from Dept WHERE parent=D.deptId AND nodeKind='D') " & _
								 "FROM Dept AS D Where D.nodeKind='D' " & _
								 "AND D.deptId LIKE '" & session("deptId") & "%'" & _
								 "ORDER BY len(D.deptId), D.parent, D.seq "				
				set RSS = conn.execute(sqlCom)
				if not RSS.EOF then
					ARYDept = RSS.getrows(300)
					glastmsglevel = 0
					genlist 0, 0, 1, 0
			    expandfrom ARYDept(cid, 0), 0, 0
				end if
				response.write "</select>"			
			%>			
			</TD>
		</TR>	
		<TR>
			<TD class="Label" align="right">發表會員</TD>
			<TD class="eTableContent">
				<input name="iEditor" class="rdonly" value="<%=rs("iEditor")%>" size="10" readonly="true">
				<input name="realname" class="rdonly" value="<%=rs("realname")%>" size="10" readonly="true">
				<input name="nickname" class="rdonly" value="<%=rs("nickname")%>" size="10" readonly="true">
			</TD>
		</TR>
		<!--TR>
			<TD rowspan="2" align="right" class="Label">統計</TD>
			<TD nowrap class="eTableContent">
				<span class="Label">討論數</span>
				<input name="DiscussCount" size="10" readonly="true" class="rdonly" value="<%=rs("DiscussCount")%>">			
				<span class="Label">意見數</span>
				<input name="CommandCount" size="10" readonly="true" class="rdonly" value="<%=rs("CommandCount")%>">			
				<span class="Label">瀏覽數</span>
				<input name="BrowseCount" size="10" readonly="true" class="rdonly" value="<%=rs("BrowseCount")%>">
			</TD>
		</TR>
		<TR>
			<TD nowrap class="eTableContent">
				<span class="Label">追蹤數</span>
				<input name="TraceCount" size="10" readonly="true" class="rdonly" value="<%=rs("TraceCount")%>">
				<span class="Label">評價數</span>
				<input name="GradeCount" size="10" readonly="true" class="rdonly" value="<%=rs("GradeCount")%>">
				<span class="Label">評價人數</span>
				<input name="GradePersonCount" size="10" readonly="true" class="rdonly" value="<%=rs("GradePersonCount")%>">
			</TD>
		</TR-->
		<%
			sql = "SELECT picStatus FROM KnowledgePicInfo WHERE parentIcuitem = " & discussId
				Set RS5 = conn.execute(sql)
				Dim counter : counter = 1
		%>
			<!--TR>
				<TD class="Label" align="right">上傳圖片</TD>
				<TD colspan="6" class="eTableContent">
				<% while not RS5.eof %>
				<% if (counter mod 6) = 0 then response.write "<br>" %>
				<%=counter%><input class="rdonly" value="<%if trim(RS5("picStatus"))="W" then%> 待審核 <%end if%><%if trim(RS5("picStatus"))="N" then%> 不通過 <%end if%><%if trim(RS5("picStatus"))="Y" then%> 通過 <%end if%>" size="10" readonly="true">
				<% 
						counter=counter + 1
						RS5.MoveNext 
					wend
					RS5.close
					set RS5 = nothing
				%>
				</TD>
			</TR-->
		<TR>
			<TD class="Label" align="right">狀態</TD>
			<TD colspan="6" class="eTableContent">
			<!--input name="Status" size="10" readonly="true" value="<%=rs("KStatus")%>" class="rdonly"-->
			<Select name="KStatus">
			<option value="N" <%if rs("KStatus") = "N" then response.write "selected" end if%>>正常</option>
			<option value="D" <%if rs("KStatus") = "D" then response.write "selected" end if%>>刪除</option>
			</select>
			</TD>
		</TR>		
		</TABLE>
		<input type=button value ="編修存檔" class="cbutton" onClick="formEdit()">
    <input type=button value ="刪　除" class="cbutton" onClick="formDelete()">   
    <input type=button value ="重　填" class="cbutton" onClick="formReset()">
    <input type=button value ="回前頁" class="cbutton" onClick="formBack()">      
	</form>     
</body>
</html>
<script language="javascript">
	function formEdit() {
		document.getElementById("submitTask").value = "edit";
		if ( document.getElementById("xImportant").value = "" ) {
			document.getElementById("xImportant").value = "0";
		}
		document.getElementById("reg").submit();
	}
	function formDelete() {
		document.getElementById("submitTask").value = "delete";
		document.getElementById("reg").submit();
	}
	function formReset() {
		document.getElementById("submitTask").value = "reset";
		document.getElementById("reg").submit();
	}
	function formBack() {
		window.location.href = "/KnowledgeForum/KnowledgeCommandList.asp?questionId=<%=questionId%>&discussId=<%=discussId%>";
	}
</script>
<%
	end if
	rs.close
	set rs = nothing
end sub
%>