<%@ CodePage = 65001 %>
<% 
Response.Expires = 0
HTProgCap="單元資料維護"
HTProgFunc="編修"
HTUploadPath=session("Public")+"data/"
HTProgCode="PEDIA01"
HTProgPrefix="Pedia" 

Dim iBaseDSDId : iBaseDSDId = "40"
Dim iCtUnitId : iCtUnitId = "1507"
const keywordCount = 10
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/selectTree.inc" -->
<!-- #INCLUDE FILE="../inc/dbFunc.inc" -->
<%
	Sub doUpdateDB()		
		keyword = ""
		for each form in xup.Form	
			if form.Name <> "submitTask" and form.Name <> "sTitle" Then				
				if form <> "" then
					keyword = keyword & form & ";"
				end if				
			end if
		next
		sql = "UPDATE CuDtGeneric SET xKeyword = '" & keyword & "' WHERE icuitem = " & request.querystring("icuitem")				
		conn.execute(sql)
	End sub 
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
				document.location.href="<%=HTprogPrefix%>List.asp?keep=Y"				
			</script>
    </body>
  </html>
<% End sub %>
<%
	apath = server.mappath(HTUploadPath) & "\"

	if request.querystring("phase")="edit" then
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
	else		
		if errMsg <> "" then	showErrBox()		
		showForm()		
	end if
%>
<% 
sub showForm() 
		
	sql = "SELECT sTitle, xKeyword FROM CuDtGeneric WHERE icuitem = " & request.querystring("icuitem")
	Set rs = conn.execute(sql)
	if rs.eof then
		showDoneBox("無此資料!")
	else	
		Dim keywordItem()
		sTitle = rs("sTitle")
		keywords = rs("xKeyword")		
		if keywords = "" then
			for i = 0 to keywordCount - 1
				ReDim Preserve keywordItem(i)
				keywordItem(i) = ""
			next
		else
			items = split(keywords, ";")
			for i = 0 to keywordCount - 1
				ReDim Preserve keywordItem(i)
				if i <= uBound(items) then					
					if items(i) <> "" then					
						keywordItem(i) = items(i)
					else
						keywordItem(i) = ""
					end if
				else
					keywordItem(i) = ""
				end if
			next			
		end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<link href="/css/form.css" rel="stylesheet" type="text/css">
	<link href="/css/layout.css" rel="stylesheet" type="text/css">
	<title>資料管理</title>
</head>
<body>
<div id="FuncName">
	<h1></h1><font size=2>【目錄樹節點: 知識小百科】</font>
	<div id="Nav">
		<a href="/Pedia/PediaList.asp">回條列</a>	
	</div>
	<div id="ClearFloat"></div>
</div>
<div id="FormName">
 	單元資料維護&nbsp;
  <font size=2>)【知識小百科 / 單元資料:純網頁】</font>
</div>

<form id="reg" method="POST" name="reg" action="PediaKeyword.asp?icuitem=<%=request.querystring("icuitem")%>" ENCTYPE="MULTIPART/FORM-DATA">
	<INPUT TYPE="hidden" name="submitTask" id ="submitTask" value="">
	<table cellspacing="0">
	<TR>
		<TD class="Label" align="right"><span class="Must">*</span>詞目</TD>
		<TD class="eTableContent"><input name="sTitle" value="<%=sTitle%>" size="50" readonly="true" class="rdonly"></TD>
	</TR>
	<% for i = 0 to keywordCount - 1 %>
		<TR>	
		<% if i = 0 then %>
			<TD class="Label" align="right">相關詞</TD>
		<% else %>
			<TD class="Label" align="right">&nbsp;</TD>
		<% end if %>
			<TD class="eTableContent">
				<input name="keyword<%=i%>" id="keyword<%=i%>" value="<%=keywordItem(i)%>" size="30" readonly="true" class="rdonly">
				<input name="button<%=i%>" id="button<%=i%>" type="button" class="cbutton" value ="新增" onclick="doPageLink('keyword<%=i%>')"> 
				<input name="button<%=i%>" id="button<%=i%>" type="button" class="cbutton" value ="清空" onclick="clearText('keyword<%=i%>')">  
			</TD>
		</TR>
	<% next%>	
	</TABLE>
	<input name="button" type="button" class="cbutton" onClick="formModSubmit()" value ="編修存檔">
	<input type="button" value ="回前頁" class="cbutton" onClick="returnToEdit()">       
</form>     
</body>
</html>
<% 
	end if
	Set rs = nothing
end sub 
%>
<script language="javascript">
	function formModSubmit() {
		document.getElementById("submitTask").value = "UPDATE";
		document.getElementById("reg").submit();
	}
	function clearText(item) {
		document.getElementById(item).value = "";
	}
	function doPageLink(item){	
		var fcolor=showModalDialog("keywordLink.asp?icuitem=<%=request.querystring("icuitem")%>",false,"dialogWidth:500px;dialogHeight:200px;status:0;");		
		document.getElementById(item).value = fcolor;
	}
	function returnToEdit() {
		window.location.href = "/Pedia/PediaEdit.asp?icuitem=<%=request.querystring("icuitem")%>&phase=edit";
	}
</script>
