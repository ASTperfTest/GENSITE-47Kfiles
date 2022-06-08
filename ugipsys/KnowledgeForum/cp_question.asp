<%
Response.Expires = 0
HTProgCap=""
HTProgFunc=""
HTProgCode="KForumlist"
HTProgPrefix=""
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/selectTree.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><head>
<script language="JavaScript" type="text/javascript">
<!--
	function formDelSubmit()
	{
		var f = document.getElementById("disCount");
		if( f.value != "0")
		{
			if(confirm("此發問下含有其他討論,是否確定刪除？")) {
				
				document.getElementById("submitTask").value = "delete";
				document.reg.submit();
			}
		}
		else
		{
			document.getElementById("submitTask").value = "delete";
			document.reg.submit();
		}
		
		return false;
	
		
	}
	function returnToEdit() {
		window.location.href = "cp_question.asp?iCUItem=<%=request.querystring("iCUItem")%>&nowPage=<%=request.querystring("nowPage")%>&pagesize=<%=cint(request.querystring("pagesize"))%>";
	}
	function formBack() {
		var backActiveFlag = "<%=request.querystring("activity")%>";
		if (backActiveFlag.value != ""){
			//window.location.href = session("myWWWSiteURL") & "/knowledge/ManagerMemberKnowledge_Question_Lp.aspx?MemberId=<%=request.querystring("MemberId")%>&nowPage=<%=request.querystring("nowPage")%>&pagesize=<%=cint(request.querystring("pagesize"))%>";
			window.location.href = "javascript:history.go(-1)"
		}
		else{
			window.location.href = "KnowledgeForumList.asp?iCUItem=<%=request.querystring("iCUItem")%>&nowPage=<%=request.querystring("nowPage")%>&pagesize=<%=cint(request.querystring("pagesize"))%>";
		}
	}
	function formUpdateSubmit()
	{
		var rege = "/^([a-zA-Z0-9_\.\-])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9])+$/";
		if (document.Form1.stitle.value == ""  ) 
		{
			alert("請輸入標題！"); 
			document.Form1.stitle.focus(); 
		} 
    else if (document.Form1.htx_xnewWindow.value == ""  ) 
		{
			alert("請選擇是否關閉討論！"); 
			document.Form1.htx_xnewWindow.focus(); 
		} 
		else if (document.Form1.htx_fctupublic.value == "" ) 
		{
			alert("請選擇是否公開！"); 
			document.Form1.htx_fctupublic.focus(); 
		} 
		else 
		{
			document.getElementById("submitTask").value = "UPDATE";
			document.getElementById("Form1").submit();
		}
	}
	
	//知識家活動問題狀態修改連動;上架:預設不公開；下架:預設討論關閉
	function knowledgeSelection()
	{
		var vGroupObj = document.getElementById('htx_vgroup')
		if (vGroupObj != null && document.getElementById("htx_fctupublic") != null && document.getElementById("htx_xnewWindow") != null)
		{
			if(vGroupObj.value == "A")
			{
				document.getElementById("htx_fctupublic").value = "N"
			}
			else if (vGroupObj.value == "OFF")
			{
				document.getElementById("htx_xnewWindow").value = "Y";
			}
		}
	}
	//-->	
</script>

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
	    <a href="KnowledgeQuestionMail.asp?questionId=<%=request.querystring("iCUItem")%>&nowPage=<%=request.querystring("nowPage")%>&pagesize=<%=request.querystring("pagesize")%>">轉寄問題</a>&nbsp;
	    <a href="knowledgepiclist.asp?iCUItem=<%=request.querystring("iCUItem")%>">圖片管理</a>&nbsp;
	    <a href="KnowledgeExpertAdditionalList.asp?questionId=<%=request.querystring("iCUItem")%>&nowPage=<%=request.querystring("nowPage")%>&pagesize=<%=request.querystring("pagesize")%>">專家補充管理</a>&nbsp;
	    <a href="KnowledgeDiscussList.asp?questionId=<%=request.querystring("iCUItem")%>&nowPage=<%=request.querystring("nowPage")%>&pagesize=<%=request.querystring("pagesize")%>">討論管理</a>&nbsp;
	    <a href="publishset.asp?questionId=<%=request.querystring("iCUItem")%>&nowPage=<%=request.querystring("nowPage")%>&pagesize=<%=request.querystring("pagesize")%>">發佈至主題館</a>&nbsp;
	    <a href="KnowledgeForumlist.asp">回條列</a>	
		</div>
		<div id="ClearFloat"></div>
	</div>
	<div id="FormName">
	 	單元資料維護&nbsp;
    <font size=2>【編輯(<font color=red>一般資料式</font>)--主題單元:(.知識家)-知識問題 / 單元資料:知識家討論區】
	</div>
<% Sub showCantFindBox(lMsg) %>
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
				history.back
			</script>
    </body>
  </html>
<% End sub %>
<%
	dim iCUItem
	iCUItem = request.querystring("iCUItem")
	sql = "SELECT * FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem "
	sql = sql & "WHERE iCUItem = " & iCUItem 
		
	Set RSreg = conn.execute(sql)
	if RSreg.eof then
		showCantFindBox "無資料"		
	else
	
%>
		<form id="Form1" method="POST" name="reg" action="cp_question_Act.asp?iCUItem=<%=trim(RSreg("iCUItem"))%>" >
			<INPUT TYPE="hidden" name="submitTask" id="submitTask" value="">
			<INPUT TYPE="hidden" name="showType" id="showType" value="">
			<INPUT TYPE="hidden" name="refID" id="refID" value="">
			<input type="hidden" name="htx_icuitem" id="htx_icuitem">
			<input type="hidden" name="htx_ibaseDsd" id="htx_ibaseDsd">
			<input type="hidden" name="htx_ictunit" id="htx_ictunit" >
			<input type="hidden" name="htx_ieditor" id="htx_ieditor" >
			<input type="hidden" name="htx_deditDate" id="htx_deditDate" >
			<input type="hidden" name="htx_Created_Date" id="htx_Created_Date" >
			<input type="hidden" name="htx_siteId" id="htx_siteId" >
			<input type="hidden" name="htx_ParentIcuitem" id="htx_ParentIcuitem" >
			<table cellspacing="0">
			<TR><TD colspan="6" class="eTableContent"></TD></TR>			
			<TR>
				<TD class="Label" align="right"><span class="Must">*</span>標題</TD>
				<TD colspan="6" class="eTableContent"><input name="stitle" type="text" size="50" value="<%=RSreg("sTitle")%>" ></TD>
			</TR>
			<TR>
				<TD class="Label" align="right">內文</TD>
				<TD colspan="6" class="eTableContent"><textarea name="htx_xbody" rows="8" cols="60" ><%=trim(RSreg("xbody"))%></textarea></TD>
			</TR>			
			<TR>
				<TD class="Label" align="right">網址</TD>
				<TD colspan="6" class="eTableContent"><input name="htx_xurl" size="50" value="<%=trim(RSreg("xurl"))%>" ></TD>
			</TR>
			<TR>
				<TD class="Label" align="right">討論關閉</TD>
				<TD colspan="6" class="eTableContent">
					<Select name="htx_xnewWindow" size=1>
						<option value="" <%if RSreg("xNewWindow")="" then%>  selected <%end if%>>請選擇</option>
						<option value="Y" <%if RSreg("xNewWindow")="Y" then%>  selected <%end if%>>是</option>
						<option value="N" <%if RSreg("xNewWindow")="N" then%>  selected <%end if%>>否</option>
					</select>
				</TD>
			</TR>
			<TR>
				<TD class="Label" align="right"><span class="Must">*</span>是否公開(草稿)</TD>
				<TD colspan="6" class="eTableContent">
					<Select name="htx_fctupublic" size=1>
						<option value="" <%if RSreg("fctupublic")="" then%>  selected <%end if%>>請選擇</option>
						<option value="Y" <%if RSreg("fctupublic")="Y" then%>  selected <%end if%>>公開</option>
						<option value="N" <%if RSreg("fctupublic")="N" then%>  selected <%end if%>>不公開</option>
					</select>
				</TD>
			</TR>
			<TR>
				<TD class="Label" align="right">資料大類</TD>
				<TD colspan="6" class="eTableContent">
					<Select name="htx_topCat" size=1>
					<%
						fsql = "SELECT mValue,mCode FROM CodeMain WHERE (codeMetaID = 'KnowledgeType') ORDER BY mSortValue"
						Set RS = conn.execute(fsql)
						while not RS.eof
					%>
							<option <%if RSreg("topCat")=trim(RS("mCode")) then%> selected <%end if%> value="<%=trim(RS("mCode"))%>"> <% response.write trim(RS("mValue"))%></option>
					<%	
							RS.MoveNext
							wend
					%> 
					</select>
				</TD>
			</TR>
			<TR>
				<TD class="Label" align="right">活動問題</TD>
				<TD colspan="6" class="eTableContent">
					<Select name="htx_vgroup" size=1 onChange=javascript:knowledgeSelection();>
						<option value="A" <%if trim(RSreg("vGroup"))="A" then%>  selected <%end if%>>是</option>
						<option value="" <%if trim(RSreg("vGroup"))="" or isnull(RSreg("vGroup")) then%>  selected <%end if%>>否</option>
						<option value="OFF" <%if trim(RSreg("vGroup"))="OFF" then%>  selected <%end if%> >已下架</option>
					</select>
				</TD>
			</TR>
			<TR>
				<TD class="Label" align="right">關鍵字詞</TD>
				<TD colspan="6" class="eTableContent"><input name="htx_xkeyword" size="50" value="<%=trim(RSreg("xkeyword"))%>"></TD>
			</TR>
			<TR>
				<TD class="Label" align="right">重要性</TD>
				<TD colspan="6" class="eTableContent"><input name="htx_ximportant" size="2" value="<%=trim(RSreg("ximportant"))%>"></TD>
			</TR>
			<TR>
				<TD class="Label" align="right"><span class="Must">*</span>單位</TD>
				<%
					response.write "<TD colspan='6' class='eTableContent'>"		    	
					response.write "<Select name=""htx_idept"" size=1>" & vbCRLF	
					sqlCom ="SELECT D.deptId,D.deptName,D.parent,len(D.deptId)-1,D.seq," & _
									"(Select Count(*) from Dept WHERE parent=D.deptId AND nodeKind='D') " & _
									"  FROM Dept AS D Where D.nodeKind='D' " _
									& " AND D.deptId LIKE '" & session("deptId") & "%'" _
									& " ORDER BY len(D.deptId), D.parent, D.seq"	
					set RSS = conn.execute(sqlCom)
					if not RSS.EOF then
						ARYDept = RSS.getrows(300)
						glastmsglevel = 0
						genlist 0, 0, 1, 0
						expandfrom ARYDept(cid, 0), 0, 0
					end if
					response.write "</select>"
					response.write "</TD></TR></TD></TR>"	
			
					sql="SELECT  Member.account, Member.realname, Member.nickname FROM  CuDTGeneric INNER JOIN Member ON CuDTGeneric.iEditor = Member.account WHERE  CuDTGeneric.iCUItem ="& "'"  & iCUItem & "'"
					Set RS2 = conn.execute(sql)
					if  not RS2.eof then
			    response.write "<TR>"
				response.write "<TD class=""Label"" align=""right"">發表會員</TD>"
				response.write "<TD colspan=""6"" class=""eTableContent"">"
			    response.write "<input name=""account"" class=""rdonly"" value=""" & trim(RS2("account"))& """ size=""10"" readonly=""true"">"
				response.write "<input name=""realname"" class=""rdonly"" value=""" & trim(RS2("realname"))& """ size=""10"" readonly=""true"">"
				response.write "<input name=""nickname"" class=""rdonly"" value=""" & trim(RS2("nickname"))& """ size=""10"" readonly=""true"">"
				response.write "</TD>"
			    response.write "</TR>"
			end if
				sql="SELECT  KnowledgeForum.DiscussCount, KnowledgeForum.CommandCount, KnowledgeForum.BrowseCount, KnowledgeForum.TraceCount, KnowledgeForum.GradeCount, KnowledgeForum.GradePersonCount FROM  KnowledgeForum INNER JOIN CuDTGeneric ON KnowledgeForum.gicuitem = CuDTGeneric.iCUItem WHERE CuDTGeneric.iCUItem = "& "'"  & iCUItem & "'"
				Set RS3 = conn.execute(sql)
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
				sql="SELECT picStatus FROM  KnowledgePicInfo WHERE  parentIcuitem = "& "'"  & iCUItem & "'"
				Set RS5 = conn.execute(sql)
				Dim counter : counter = 1
			
				if not RS5.eof then
				response.write "<TR>"
				response.write "<TD class=""Label"" align=""right"">上傳圖片</TD>"
				response.write "<TD colspan=""6"" class=""eTableContent"">"
				while not RS5.eof 
				 if (counter mod 6)=0 then response.write "<br>" %>
				<% response.write counter %><input class="rdonly" value="<%if trim(RS5("picStatus"))="W" then%> 待審核 <%end if%><%if trim(RS5("picStatus"))="N" then%> 不通過 <%end if%><%if trim(RS5("picStatus"))="Y" then%> 通過 <%end if%>" size="10" readonly="true">
				<% 
						counter=counter+1
						RS5.MoveNext 
					wend
				end if 
				%>
				</TD>
			</TR>
			<%
				sql="SELECT  KnowledgeForum.status FROM  KnowledgeForum INNER JOIN CuDTGeneric ON KnowledgeForum.gicuitem = CuDTGeneric.iCUItem WHERE CuDTGeneric.iCUItem = "& "'"  & iCUItem & "'"
				Set RS4 = conn.execute(sql)
			%>
			<TR>
				<TD class="Label" align="right">狀態</TD>
				<TD colspan="6" class="eTableContent">
					<!--input name="htx_Status" size="10"  class="rdonly" readonly="true" value="<%=RS4("STATUS")%>" -->
					<INPUT TYPE="hidden" name="status1" value="<%=RS4("status")%>">
					<Select name="status">
					<option value="N" <%if RS4("status") = "N" then response.write "selected" end if%>>正常</option>
					<option value="D" <%if RS4("status") = "D" then response.write "selected" end if%>>刪除</option>
					</select>
				</TD>
			</TR>
			<TR>
				<TD class="Label" align="right">發佈至主題館</TD>
				<TD colspan="6" class="eTableContent">
					<%
					subjectSql = "SELECT DISTINCT K2.subjectName FROM dbo.KnowledgeToSubject K1 " & _
							"INNER JOIN KnowledgeToSubject K2 ON K1.subjectId = K2.subjectId " & _
							"WHERE K1.iCUItem='" & iCUItem & "' " & _
							"AND K1.subjectId = K2.subjectId " & _
							"AND K2.subjectName IS NOT NULL " 
					set subjectRS = conn.execute(subjectSql)
					dim publishSubject,sizeNum
					sizeNum = 10
					while not subjectRS.eof 
						sizeNum = sizeNum + 10
						if publishSubject <> "" then
							publishSubject = publishSubject & "," & subjectRS("subjectName") 
						else
							publishSubject = subjectRS("subjectName")
						end if
					subjectRS.MoveNext
					wend
					
					
					
					%>
					<input name="publishSetCol" size="<%=sizeNum%>" value="<%=publishSubject%>"readonly="true" class="rdonly">
				</TD>
			</TR>
			<TR>
				<TD>
				<%
				sqlDiscuss = "SELECT Count(1) as disCount FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem "
				sqlDiscuss = sqlDiscuss & "WHERE (KnowledgeForum.ParentIcuitem = " & iCUItem & ") AND (KnowledgeForum.Status = 'N') AND (CuDTGeneric.iCtUnit = 933)"
				set rsDiscuss = conn.execute(sqlDiscuss)
				if not rsDiscuss.eof then
					disCount = rsDiscuss("disCount")
				end if
				
				%>
				<INPUT TYPE="hidden" name="disCount" id="disCount" value="<%=disCount%>">
				</TD>
			</TR>
			</TABLE>
			
			<object data="../../inc/calendar.asp" id=calendar1 type=text/x-scriptlet width=245 height=160 style='position:absolute;top:0;left:230;visibility:hidden'></object>
			<INPUT TYPE=hidden name=CalendarTarget>
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
