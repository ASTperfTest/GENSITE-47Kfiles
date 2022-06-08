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
	<link href="../css/form.css" rel="stylesheet" type="text/css">
	<link href="../css/layout.css" rel="stylesheet" type="text/css">
	<title>資料管理／資料上稿</title>
</head>
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
<!--#include file = "ExpertMail.inc" -->
<!--#include file = "ExpertCommon.inc" -->
<%
	'---在專家回覆的單元中新增一筆資料---ictunit = 934---knowledgeforum.status = 'W'---
	'---一個專家就新增一筆---
	sub HandleMail
		Dim questionId : questionId = request.querystring("questionId")
		Dim experts : experts = request("experts")
		Dim sTitle : sTitle = GetsTitle(questionId, "1")
		Dim items
		Dim icuitem
		if instr(experts, ",") > 0 then
			experts = experts & ","
			items = split(experts, ",")
		else
			redim items(1)
			items(0) = experts
		end if		
		Dim myrand
		for i = 0 to ubound(items) - 1
			icuitem = ""
			sql = "INSERT INTO CuDTGeneric(iBaseDSD, iCTUnit, fCTUPublic, sTitle, iEditor, iDept, showType, siteId, xBody) " & _
						"VALUES(" & iBaseDSD & ", " & iCTUnit & ", 'Y', '" & sTitle & "', 'expert', '0', '1', '3', '')"
			sql = "set nocount on;" & sql & "; select @@IDENTITY as NewID"	
			'response.write sql & "<hr>"			
			Set rs = conn.Execute(sql)
			icuitem = rs(0)
			myrand = GetRand()
			sql = "INSERT INTO KnowledgeForum(gicuitem, ParentIcuitem, Status, expertId, expertRand, expertSendDate, expertSend) " & _
						"VALUES(" & icuitem & ", " & questionId & ", 'W', '" & items(i) & "', '" & myrand & "', GETDATE(), 'Y')"
			response.write sql & "<hr>"			
			conn.execute(sql)
			SendEmail questionId, items(i), request("mailBody"), myrand, icuitem
		next
		showDoneBox "轉寄成功"
	end sub
	
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
				document.location.href="/KnowledgeForum/cp_question.asp?iCUItem=<%=request.querystring("questionId")%>&nowPage=<%=request.querystring("nowPage")%>&pagesize=<%=request.querystring("pagesize")%>"
			</script>
    </body>
  </html>
<% End sub %>
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
<% sub ShowForm %>
<body>
	<div id="FuncName">
		<h1>資料管理／資料上稿</h1>
		<font size=2>【目錄樹節點: 知識問題】</font>
		<div id="Nav"><a href="/KnowledgeForum/cp_question.asp?iCUItem=<%=request.querystring("questionId")%>&nowPage=<%=request.querystring("nowPage")%>&pagesize=<%=request.querystring("pagesize")%>">回前頁</a></div>
		<div id="ClearFloat"></div>
	</div>
	<div id="FormName">單元資料維護&nbsp;<font size=2>【轉寄問題】</font></div>
<%
	Dim questionId
	questionId = request.querystring("questionId")
	sql = "SELECT * FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem "
	sql = sql & "WHERE iCUItem = " & questionId 
	Set RSreg = conn.execute(sql)
	if RSreg.eof then
		showCantFindBox "無資料"		
	else
%>
		<form id="Form1" method="POST" name="reg" action="/KnowledgeForum/KnowledgeQuestionMail.asp?questionId=<%=questionId%>&nowPage=<%=request.querystring("nowPage")%>&pagesize=<%=request.querystring("pagesize")%>" >
			<INPUT TYPE="hidden" name="submitTask" value="">			
			<table cellspacing="0">
			<TR><TD colspan="6" class="eTableContent"></TD></TR>			
			<TR>
				<TD class="Label" align="right">信件內文</TD>
				<TD colspan="6" class="eTableContent"><textarea name="mailBody" rows="8" cols="60" ></textarea></TD>
			</TR>
			<TR>
				<TD class="Label" align="right">選擇專家</TD>
				<TD class="eTableContent">
					<input name="experts" title="請以,分隔" value="" size="50" readonly="Y" >
					<input type="hidden" name="expertSelected" />
					<br />
					<span class="mailQ">
						<BUTTON id="ExpertSelect" type="button" class="cbutton">選擇專家</BUTTON>						
					</span>
				</TD>
			</TR>
			<TR>
				<TD class="Label" align="right">發送紀錄</TD>
				<TD class="eTableContent">
					<input name="button2" type="button" class="cbutton" onClick="RedirectToMailList()" value ="已寄送清單瀏覽">
				</TD>
			</TR>
			<TR>
				<TD class="Label" align="right">問題標題</TD>
				<TD colspan="6" class="eTableContent"><input name="stitle" type="text" size="50" value="<%=RSreg("sTitle")%>" readonly="Y" class="rdonly" ></TD>
			</TR>
			<TR>
				<TD class="Label" align="right">問題內文</TD>
				<TD colspan="6" class="eTableContent"><textarea name="htx_xbody" rows="8" cols="60" readonly="Y" class="rdonly" ><%=trim(RSreg("xbody"))%></textarea></TD>
			</TR>
			</TABLE>						
      <input type="button" value="確定發送" class="cbutton" onclick="CheckSubmit();">      
      <input type="button" value="重　  填" class="cbutton" onClick="returnToEdit();">
      <input type="button" value="取    消" class="cbutton" onClick="formBack();">       
		</form>    
		<script language="vbs">                     			
			sub ExpertSelect_onClick
				window.open "/KnowledgeForum/ExpertSelect.asp",null,"height=430,width=750,status=no,scrollbars=yes,toolbar=no,menubar=no,location=no"
			end sub
		</script>		
		<script language="javascript">    
			function CheckSubmit() 
			{
				if ( document.getElementById("mailBody").value == "" ) {
					alert("請填寫信件內文");
					document.getElementById("mailBody").focus();
				}
				else if ( document.getElementById("experts").value == "" ) {
					alert("請選擇專家");
					document.getElementById("ExpertSelect").focus();					
				}
				else {
					document.getElementById("submitTask").value = "submit";
					document.getElementById("Form1").submit();
				}
			}
			function RedirectToMailList()
			{
				window.location.href = "/KnowledgeForum/KnowledgeQuestionMailList.asp?questionId=<%=request.querystring("questionId")%>&nowPage=<%=request.querystring("nowPage")%>&pagesize=<%=request.querystring("pagesize")%>";
			}
			function formBack()
			{
				window.location.href = "/KnowledgeForum/cp_question.asp?iCUItem=<%=request.querystring("questionId")%>&nowPage=<%=request.querystring("nowPage")%>&pagesize=<%=request.querystring("pagesize")%>";
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