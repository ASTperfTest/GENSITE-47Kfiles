<%@ CodePage = 65001 %>
<% 
Response.Expires = 0
HTProgCap="單元資料維護"
HTProgFunc="編修"
HTUploadPath=session("Public")+"data/"
HTProgCode="PEDIA01"
HTProgPrefix="PediaAdditional" 

Dim iCtUnitId : iCtUnitId = session("PediaAdditionalUnitId")
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/selectTree.inc" -->
<!-- #INCLUDE FILE="../inc/dbFunc.inc" -->
<%
	Function Send_email (S_email,R_email,Re_Sbj,Re_Body)
		Set objNewMail = CreateObject("CDONTS.NewMail") 
		objNewMail.MailFormat = 0
		objNewMail.BodyFormat = 0 
		call objNewMail.Send(S_email,R_email,Re_Sbj,Re_Body)
		Set objNewMail = Nothing	
	End Function	
	
	FUNCTION mypkStr (s, cname, endchar)
		if s = "" then
			if cname = "xPostDate" then
				mypkStr = "GETDATE()"
			else
				mypkStr = "''" & endchar
			end if
		else
			pos = InStr(s, "'")
			While pos > 0
				s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)
				pos = InStr(pos + 2, s, "'")
			Wend
			if cname = "xPostDate" then
				mypkStr = "'" & s & "'" & endchar
			else
				mypkStr = "N'" & s & "'" & endchar
			end if
		end if
	END FUNCTION
	
%>
<% Sub showErrBox() %>
		<script language=VBScript>
			alert "<%=errMsg%>"
			'window.history.back
		</script>
<% End sub %>
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
				document.location.href="<%=HTprogPrefix%>List.asp?PediaAdditionalEdit.asp?icuitem=<%=request.querystring("picuitem")%>&keep=Y&nowpage=<%=request.querystring("nowpage")%>&pagesize=<%=request.querystring("pagesize")%>"				
			</script>
    </body>
  </html>
<% End sub %>
<%
	Sub doUpdateDB()
	
		for each form in xup.Form		
			if form.Name = "xBody" Then
				sql = "UPDATE CuDtGeneric SET xBody = " & pkStr(form, "")	 & " WHERE icuitem = " & request.querystring("icuitem")				
				conn.execute(sql)
				exit for
			end if
		next		
	End sub 
%>	
<%
	apath = server.mappath(HTUploadPath) & "\"

	if request.querystring("phase")="edit" or session("BatchDphase") = "edit" then
		Set xup = Server.CreateObject("TABS.Upload")
	else
		Set xup = Server.CreateObject("TABS.Upload")
		xup.codepage=65001
		xup.Start apath
	end if

	function xUpForm(xvar)
		xUpForm = xup.form(xvar)
	end function
	    
	if xUpForm("submitTask") = "UPDATE" then
		
		errMsg = ""		
		if errMsg <> "" then
			EditInBothCase()
		else
			doUpdateDB()
			showDoneBox("資料更新成功！")
		end if
			
	elseif xUpForm("submitTask") = "DELETE" then
					
		icuitem = request.queryString("iCuItem")		
		'----刪除主表---將狀態由Y改成D---
		sql = "UPDATE Pedia SET xStatus = 'D' WHERE gicuitem = " & pkStr(icuitem, "")		
		conn.execute(sql)
		
		Dim flag : flag = false
		sql = "SELECT * FROM Activity WHERE ActivityId = 'pedia' AND GETDATE() BETWEEN ActivityStartTime AND ActivityEndTime"
		set rs = conn.execute(sql)
		if not rs.eof then
			flag = true
		end if
		
		if flag then
			sql = "UPDATE ActivityPediaMember SET commendAdditionalCount = commendAdditionalCount - 1 WHERE memberId = '" & xUpForm("memberId") & "' AND commendAdditionalCount > 0"
			conn.execute(sql)
		end if
		
		'---刪除補充解釋文章---且將分數扣除---
		
		showDoneBox("資料刪除成功！")	
		
	else		
		if errMsg <> "" then	showErrBox()		
		showForm()		
	end if
%>
<%
	Function GetNameAndDate(id)
	
		if id = "0" then
			GetNameAndDate = ""
		else
			sql = "SELECT Member.realname, Pedia.commendTime FROM CuDTGeneric INNER JOIN Pedia ON CuDTGeneric.iCUItem = Pedia.gicuitem "
			sql = sql & " INNER JOIN Member ON Pedia.memberId = Member.account WHERE CuDTGeneric.iCUItem = " & id
			set grs = conn.execute(sql)
			if not grs.eof then
				GetNameAndDate = grs("realname") & " | " & grs("commendTime")
			end if
			set grs = nothing
		end if
		
	End Function
%>
<% 
Sub showForm() 

	sql = "SELECT * FROM CuDTGeneric INNER JOIN Pedia ON CuDTGeneric.iCUItem = Pedia.gicuitem "
	sql = sql & " INNER JOIN Member ON Pedia.memberId = Member.account WHERE CuDTGeneric.iCUItem = " & request.queryString("iCuItem")

	Set rs = conn.execute(sql)
	if rs.eof then
		showDoneBox("無此資料!")
	else
	
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="/css/form.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
</head>
<body>
<div id="FuncName">
	<h1>資料管理／資料上稿</h1>
	<font size=2>【知識小百科－補充解釋】</font>
	<div id="Nav">
	  <a href="/Pedia/PediaEdit.asp?icuitem=<%=request.querystring("picuitem")%>&phase=edit">回詞目頁</a>&nbsp;
	  <a href="/Pedia/PediaAdditionalList.asp?icuitem=<%=request.querystring("picuitem")%>">回補充管理頁</a>&nbsp;
	  <a href="/Pedia/PediaAdditionalList.asp">回條列</a>	
	</div>
	<div id="ClearFloat"></div>
</div>
<div id="FormName">	
  單元資料維護&nbsp;
  <font size=2>【知識小百科－補充解釋】</font>
</div>

<form id="Form1" method="POST" name="reg" action="PediaAdditionalEdit.asp?icuitem=<%=request.querystring("icuitem")%>&picuitem=<%=request.querystring("picuitem")%>&nowpage=<%=request.querystring("nowpage")%>&pagesize=<%=request.querystring("pagesize")%>" ENCTYPE="MULTIPART/FORM-DATA">
	<INPUT TYPE="hidden" name="submitTask" value="">
	<table cellspacing="0">
	<TR>
	  <TD class="Label" align="right"><span class="Must">*</span>詞目</TD>
	  <TD class="eTableContent"><input name="sTitle" value="<%=rs("sTitle")%>" size="50" class="rdonly" readonly="true"></TD>
	</TR>
	<TR>
	  <TD class="Label" align="right">引用</TD>
	  <TD class="eTableContent">
			<% if rs("quoteIcuitem") = "0" Then %>
				本篇為第一篇補充｜無引用資料
			<% else %>
				<a href="/Pedia/PediaAdditionalEdit.asp?icuitem=<%=rs("quoteIcuitem")%>&picuitem=<%=rs("parentIcuitem")%>&phase=edit">補充引用原文</a>｜
			<% end if%>
			<%=GetNameAndDate(rs("quoteIcuitem"))%>
		</TD>
	</TR>
	<TR>
	  <TD class="Label" align="right">補充解釋</TD>
	  <TD class="eTableContent"><textarea name="xBody" rows="8" cols="60"" title="內容、本文、網頁、說明""><%=rs("xBody")%></textarea></TD>
	</TR>
	<TR>
		<TD class="Label" align="right">發表會員</TD>
	  <TD class="eTableContent">
			<input name="memberId" class="rdonly" value="<%=rs("memberId")%>" size="10" readonly="true">
	    <input name="realname" class="rdonly" value="<%=rs("realname")%>" size="10" readonly="true">
	    <input name="nickname" class="rdonly" value="<%=rs("nickname")%>" size="10" readonly="true">
		</TD>
	</TR>
	<TR>
		<TD class="Label" align="right">發布日期</TD>
		<TD class="eTableContent">
			<input name="xPostDate" size="20" type="text" value="<%=rs("commendTime")%>" readonly="true" class="rdonly" >			
		</TD>
	</TR>
	</TABLE>
	<input name="button" type=button class="cbutton" onClick="formModSubmit()" value ="編修存檔">
	<input type=button value ="刪　除" class="cbutton" onClick="formDelSubmit()">   
	<input type=button value ="重　填" class="cbutton" onClick="returnToEdit()">
	<input type=button value ="回前頁" class="cbutton" onClick="Vbs:history.back">	       
</form>     
</body>
</html>
<script language="javascript">
	function formModSubmit() {
		document.getElementById("submitTask").value = "UPDATE";
		document.getElementById("reg").submit();
	}
	function formDelSubmit() {
		document.getElementById("submitTask").value = "DELETE";
		document.getElementById("reg").submit();
	}
	function returnToEdit() {
		window.location.href = "/Pedia/PediaAdditionalEdit.asp?icuitem=<%=request.querystring("icuitem")%>&picuitem=<%=request.querystring("picuitem")%>&phase=edit&nowpage=<%=request.querystring("nowPage")%>&pagesize=<%=request.querystring("pagesize")%>";
	}
</script>
<% 		
	end if 
	set rs = nothing
%>
<% End sub %>


 
