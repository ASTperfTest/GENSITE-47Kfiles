<%@ CodePage = 65001 %>
<%
Response.Expires = 0
HTProgCap=""
HTProgFunc=""
HTProgCode="KForumlist"
HTProgPrefix=""
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
				history.back
			</script>
    </body>
  </html>
<% End sub %>

<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/client.inc" -->

<%
	dim iCUItem
  iCUItem = request.querystring("iCUItem")
	picId = request("picId")
	
	if request("submitTask") = "Pass" then
		sql = "UPDATE KnowledgePicInfo SET picStatus = 'Y' WHERE parentIcuitem = " & iCUItem & " AND picId= " & picId 
		conn.execute(sql)
	elseif request("submitTask") = "NoPass" then
    sql = "UPDATE KnowledgePicInfo SET picStatus = 'N' WHERE parentIcuitem = " & iCUItem & " AND picId= " & picId 
		conn.execute(sql)
	end if

	sql = "SELECT * FROM CuDTGeneric INNER JOIN KnowledgePicInfo ON CuDTGeneric.iCUItem = KnowledgePicInfo.parentIcuitem "
	sql = sql & "where parentIcuitem = " & iCUItem & " ORDER BY xPostDate DESC"
	Set RSreg = conn.execute(sql)

	if RSreg.eof then
		showDoneBox "本資料無圖片"
	else
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="../css/list.css" rel="stylesheet" type="text/css">
<link href="../css/layout.css" rel="stylesheet" type="text/css">
<title>資料管理／資料上稿</title>
</head>
<body>
<div id="FuncName">
	<h1>資料管理／問題圖片管理</h1>
	<font size=2>【目錄樹節點: 知識問題】
	<div id="Nav">
	    <a href="lp_knowledgepiclist.asp">回圖片整合管理條列頁</a>
	    <a href="cp_question.asp?iCUItem=<%=request.querystring("iCUItem")%>">回資料內容頁</a>	</div>
	<div id="ClearFloat"></div>
</div>
<div id="FormName">問題標題：<%=RSreg("sTitle")%>
	<font size=2>【主題單元:(2008)知識問題圖片 / 單元資料:純網頁】</div>

<Form id="Form2" name=reg method="POST" action="knowledgepiclist.asp?iCUItem=<%=request.querystring("iCUItem")%>">
 <INPUT TYPE=hidden name=submitTask value="">
 <INPUT TYPE=hidden name=picId value="">
	
	<table cellspacing="0" id="ListTable">
		<tr>
		<th class=eTableLable>圖片標題</th>
					<th class=eTableLable>圖片</th>
					<th scope="col">審核</th>
					<th class=eTableLable>狀態</th>					
	  </tr> 
<% while not RSreg.eof %>	  
<tr>
  <TD class=eTableContent><font size=3 ><%=RSreg("picTitle")%></font></td>
	<td nowrap><span class="eTableContent"><a href="<%=trim(RSreg("picPath"))%>" target="_blank"><img src="<%=trim(RSreg("picPath"))%>" alt="圖片" width="120" height="70" border="0"></a></span></td>
	<TD nowrap class=eTableContent>
	<INPUT type="button" value="通過"  onClick="PassValidate('<%=RSreg("picId")%>')">
	<INPUT type="button" value="不通過" onClick="NoPassValidate('<%=RSreg("picId")%>')">
	</td>
	<TD nowrap class=eTableContent>
	<%if  trim(RSreg("picStatus"))="W" then%>
	<span class="style1"> 待審核</span>
	<% end if%>
	<%if  trim(RSreg("picStatus"))="Y" then%>
	 通過
	<% end if%>
	<%if  trim(RSreg("picStatus"))="N" then%>
    不通過
	<% end if%>
	</td>
	</tr>
               
 <%
			RSreg.MoveNext
		wend
 %>                
</TABLE>
</form>  
</body>
</html>                                 
<script language="javascript">
	function PassValidate(id) 
	{
	    document.getElementById("picId").value = id;
		document.getElementById("submitTask").value = "Pass";
		document.getElementById("reg").submit();
	}
	function NoPassValidate(id) {
	    document.getElementById("picId").value = id;
		document.getElementById("submitTask").value = "NoPass";
		document.getElementById("reg").submit();
	}
</script>
<%
end if
%>
