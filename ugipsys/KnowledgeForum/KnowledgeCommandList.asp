<%@ CodePage = 65001 %>
<%
Response.Expires = 0
HTProgCap=""
HTProgFunc=""
HTProgCode="KForumlist"
HTProgPrefix = "KnowledgeCommand" 
Response.charset = "utf-8"
Dim iCtUnitId : iCtUnitId = "935"
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include file = "KpiFunction.inc" -->
<%
Dim questionId : questionId = request.querystring("questionId")
Dim discussId : discussId = request.querystring("discussId")
if request("submitTask") = "delete" then

	sql = "SELECT Status FROM KnowledgeForum WHERE gicuitem = " & discussId	
	set rs = conn.execute(sql)
	if not rs.eof then
		if rs("Status") = "D" then 
			response.write "<script>alert('article can not be deleted!!');window.location.href='KnowledgeForumlist.asp';</script>"
			response.end
		end if
	end if
	rs.close
	set rs = nothing
	
	Dim selectedItems : selectedItems = request("selectedItems")
	Dim items : items = split(selectedItems, ";")
	Dim counter : counter = 0
	for each item in items
		if item <> "" then 		
			sql = "SELECT Status FROM KnowledgeForum WHERE gicuitem = " & item	
			set rs = conn.execute(sql)
			if not rs.eof then
				if rs("Status") = "D" then 
				else
					DeleteCommend item '---刪除評價---	
					DeleteOpinion item '---刪除意見---			
					DeleteDiscuss item '---刪除討論---		
					'---更新parent的count---				
					sql = "SELECT * FROM KnowledgeForum WHERE gicuitem = " & item
					set delrs = conn.execute(sql)
					while not delrs.eof 
						CommandCount = CInt(delrs("CommandCount"))
						GradeCount = CInt(delrs("GradeCount"))
						GradePersonCount = CInt(delrs("GradePersonCount"))
						ParentIcuitem = delrs("ParentIcuitem")					
						sql = "UPDATE KnowledgeForum SET DiscussCount = DiscussCount - 1, CommandCount = CommandCount - " & CommandCount & ", " & _
									"GradeCount = GradeCount - " & GradeCount & ", GradePersonCount = GradePersonCount - " & GradePersonCount & " " & _
									"WHERE gicuitem = " & discussId			
						conn.execute(sql)					
						delrs.movenext					
					wend				
					delrs.close
					set delrs = nothing
					counter = counter + 1						
				end if			
			end if
			rs.close
			set rs = nothing
		end if
	next
	showDoneBox "刪除完成，共刪除 " & counter & " 筆"
else
	showForm
end if
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
				document.location.href="<%=HTprogPrefix%>List.asp?questionId=<%=request.querystring("questionId")%>"
			</script>
    </body>
  </html>
<% End sub %>
<%
sub showForm
	'-------------------------------------------------------------
	Dim questionTitle
	sql = "SELECT sTitle FROM CuDTGeneric WHERE iCUItem = " & discussId
	set rs = conn.execute(sql)
	if not rs.eof then
		questionTitle = rs("sTitle")
	end if
	rs.close
	set rs = nothing
	'-------------------------------------------------------------

	sql = "SELECT *,KnowledgeForum.Status AS KStatus,KnowledgeForum.CommandCount FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem " & _
				"INNER JOIN Member ON CuDTGeneric.iEditor = Member.account  " & _
				"WHERE (KnowledgeForum.ParentIcuitem = " & discussId & ") AND (CuDTGeneric.iCTUnit = " & iCtUnitId & ") " & _
				"ORDER BY xPostDate DESC"
	set rs = conn.execute(sql)			
	if rs.eof then		
		response.write "<script>alert('本討論無意見');history.back();</script>"
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
		<font size=2>【目錄樹節點: 知識意見】
		<div id="Nav">
			<a href="/KnowledgeForum/KnowledgeDiscussContent.asp?questionId=<%=questionId%>&discussId=<%=discussId%>&nowPage=<%=request.querystring("nowPage")%>&pagesize=<%=request.querystring("pagesize")%>">回討論</a>
		</div>
		<div id="ClearFloat"></div>
	</div>
	<div id="FormName">
		問題標題：<%=questionTitle%>
		<font size=2>【主題單元:(2008)知識意見 / 單元資料:純網頁】</div>
	
		<Form id="Form2" name="reg" method="POST" action="KnowledgeCommandList.asp?discussId=<%=discussId%>">
		<INPUT TYPE=hidden name=submitTask value="">
		<INPUT TYPE="hidden" name="selectedItems" value="">
		<table cellspacing="0" id="ListTable">
		<tr>
			<th class="First" scope="col">&nbsp;</th>
			<th class=eTableLable>意見內容</th>
			<th class=eTableLable>發佈者</th>
			<th class=eTableLable>狀態</th>
			<th class=eTableLable>發佈日</th>			
	  </tr>       
		<% 
			while not rs.eof 
				recordCount = recordCount + 1
		%>
		<tr>
			<TD class=eTableContent><input type="checkbox" name="ckbox" value="<%=rs("iCUItem")%>"></td>                  	   
			<TD class=eTableContent>
				<font size=2>
					<A href="/KnowledgeForum/KnowledgeCommandContent.asp?questionId=<%=questionId%>&discussId=<%=discussId%>&commandid=<%=rs("iCUItem")%>&nowPage=<%=request.querystring("nowPage")%>&pagesize=<%=request.querystring("pagesize")%>">
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
			<TD class=eTableContent><%=rs("iEditor")%>|<%=rs("realname")%>|<%=rs("nickname")%></td>
			<TD class=eTableContent><font size=2><%=rs("KStatus")%></font></td>
			<TD class=eTableContent><font size=2><%=rs("xPostDate")%></font></td>	
		</tr>
		<% 
				rs.movenext
			wend 
			rs.close
			set rs = nothing
		%>
		</TABLE>
    <div align="center">
			<!--input type="button" name="SubmitBtn" value="整批刪除" onclick="formDelete()" id="SubmitBtn" class="cbutton" /-->      
      <input type="button" name="SubmitBtn3" value="回前頁" onclick="formBack()" id="SubmitBtn3" class="cbutton" />
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
	function formBack() {
		window.location.href = "/KnowledgeForum/KnowledgeDiscussContent.asp?questionId=<%=questionId%>&discussId=<%=discussId%>&nowPage=<%=request.querystring("nowPage")%>&pagesize=<%=request.querystring("pagesize")%>";
	}	
</script>
<%
	end if
end sub
%>
