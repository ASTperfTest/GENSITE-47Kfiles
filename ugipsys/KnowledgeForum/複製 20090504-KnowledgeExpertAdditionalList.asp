<%@ CodePage = 65001 %>
<%
Response.Expires = 0
HTProgCap=""
HTProgFunc=""
HTProgCode="KForumlist"
HTProgPrefix = "" 
Dim iCtUnitId : iCtUnitId = "934"
%>
<!--#include virtual = "/inc/server.inc" -->
<%
Dim questionId : questionId = request.querystring("questionId")
if request("submitTask") = "delete" then

	Dim selecteditems : selecteditems = request("selectedItems")
	Dim items : items = split(selecteditems, ";")
	Dim count : count = 0
	for i = 0 to ubound(items)
		if items(i) <> "" then
			sql = "UPDATE KnowledgeForum SET Status = 'D' WHERE gicuitem = " & items(i) 
			conn.execute(sql)
			count = count + 1
		end if
	next
	showDoneBox "刪除成功,共刪除 " & count & " 筆", "2"
else
	showForm
end if
%>

<% Sub showDoneBox(lMsg, atype) %>
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
				<% if atype = "1" then %>
					//history.back()
					document.location.href="cp_question.asp?iCUItem=<%=request.querystring("questionId")%>"
				<% else %>
					document.location.href="cp_question.asp?iCUItem=<%=request.querystring("questionId")%>&nowPage=<%=request.querystring("nowPage")%>&pagesize=<%=request.querystring("pagesize")%>"
				<% end if %>
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

	sql = "SELECT * FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.icuitem = KnowledgeForum.gicuitem " & _
				"WHERE CuDTGeneric.iCTUnit = " & iCtUnitId & " AND KnowledgeForum.Status <> 'W' " & _
				"AND KnowledgeForum.ParentIcuitem = " & questionId & " ORDER BY KnowledgeForum.expertReplyDate DESC"
	set rs = conn.execute(sql)			
	if rs.eof then		
		showDoneBox "本資料無專家補充", "1"
		response.end
	else
		Dim recordCount
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="/css/list.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<title></title>
</head>

<body>
	<div id="FuncName">
		<h1>資料管理／資料上稿</h1>
		<font size=2>【目錄樹節點: 知識討論】
		<div id="Nav">
			<a href="/KnowledgeForum/cp_question.asp?iCUItem=<%=questionId%>&nowPage=<%=request.querystring("nowPage")%>&pagesize=<%=request.querystring("pagesize")%>">回問題</a>
		</div>
		<div id="ClearFloat"></div>
	</div>
	<div id="FormName">
		問題標題：<%=questionTitle%>
		<font size=2>【主題單元:(2008)知識討論 / 單元資料:純網頁】</div>
	
		<Form id="Form2" name="reg" method="POST" action="KnowledgeExpertAdditionalList.asp?questionId=<%=questionId%>&nowPage=<%=request.querystring("nowPage")%>&pagesize=<%=request.querystring("pagesize")%>">
		<INPUT TYPE="hidden" name="submitTask" value="">
		<INPUT TYPE="hidden" name="selectedItems" value="">
		<table cellspacing="0" id="ListTable">
		<tr>
			<th class="First" scope="col">&nbsp;</th>
			<th class=eTableLable>專家補充內容</th>
			<th class=eTableLable>發佈專家</th>
			<th class=eTableLable>狀態</th>
			<th class=eTableLable>補充發佈日</th>
			<th class=eTableLable>是否轉寄</th>
			<th class=eTableLable>轉寄日期</th>			
	  </tr>         
		<% 
			while not rs.eof 
				recordCount = recordCount + 1
		%>
		<tr>
			<TD class=eTableContent><input type="checkbox" name="ckbox" value="<%=rs("iCUItem")%>"></td>                  	   
			<TD class=eTableContent>
				<font size=2>
					<A href="KnowledgeExpertContent.asp?questionid=<%=questionId%>&expertaddid=<%=rs("gicuitem")%>&nowPage=<%=request.querystring("nowPage")%>&pagesize=<%=request.querystring("pagesize")%>">
					<%
						if len(rs("xBody")) > 75 then
							response.write mid(rs("xBody"), 1, 75) & "..."
						else
							response.write rs("xBody")
						end if
					%>
					</A> 
				</font>
			</td>
			<TD class=eTableContent><%=rs("iEditor")%>|<%=rs("expertId")%>&nbsp;</td>
			<TD class=eTableContent><font size=2><%=rs("Status")%>&nbsp;</font></td>
			<TD class=eTableContent><font size=2><%=rs("expertReplyDate")%>&nbsp;</font></td>	
			<TD class=eTableContent><font size=2>
				<%
					if rs("expertSend") = "Y" then 
						response.write "V"
					else
						response.write "&nbsp;"
					end if
				%>
			</font></td>	
			<TD class=eTableContent><font size=2><%=rs("expertSendDate")%>&nbsp;</font></td>	
		</tr>
		<% 
				rs.movenext
			wend 
			rs.close
			set rs = nothing
		%>
		</TABLE>
    <div align="center">
			<input type="button" name="SubmitBtn" value="整批刪除" onclick="formDelete()" id="SubmitBtn" class="cbutton" />      
      <input type="submit" name="SubmitBtn3" value="回前頁" onclick="formBack()" id="SubmitBtn3" class="cbutton" />
    </div>
	</form>  
</body>
</html>                                 
<script language="javascript">
	function formDelete()
	{
		var checkitems = "";
		var count = <%=recordCount%>;
		if ( count == 1 ) {
			if( document.reg.ckbox.checked ) {
				checkitems += document.reg.ckbox.value + ";";	
			}									
		}
		else {
			for( var i = 0; i < <%=recordCount%>; i++) {
				if( document.reg.ckbox[i].checked ) {
					checkitems += document.reg.ckbox[i].value + ";";	
				}									
			}		
		}
		document.getElementById("submitTask").value = "delete";		
		document.getElementById("selectedItems").value = checkitems;				
		//alert(checkitems);
		document.reg.submit();
	}	
</script>
<%
	end if
end sub
%>
