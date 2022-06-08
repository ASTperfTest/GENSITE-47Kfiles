<%@ CodePage = 65001 %>
<%
	Response.Expires = 0
	HTProgCap="單元資料維護"
	HTProgFunc="編修"
	HTUploadPath=session("Public")+"data/"
	HTProgCode="CW01"
	HTProgPrefix="CW01" 
%>
<!--#include virtual = "/inc/server.inc" -->
<!-- #INCLUDE FILE="../inc/dbFunc.inc" -->
<%
	if request("actionType") = "edit" then		
		doUpdateDB()
		showDoneBox("資料更新成功！")		
	else	
		showForm()
	end if	

	sub doUpdateDB()
		
		Dim inter : inter = "0"
		Dim tank : tank = "0"
		Dim topic : topic = "0"
		Dim knowledge : knowledge = "0"		
		
		inter = request("inter")
		tank = request("tank")
		topic = request("topic")
		knowledge = request("knowledge")
				
		sql = "UPDATE CodeMain SET mValue = '" & inter & "' WHERE codeMetaID = 'commendWord' AND mCode = 'inter';"
		sql = sql & "UPDATE CodeMain SET mValue = '" & tank & "' WHERE codeMetaID = 'commendWord' AND mCode = 'tank';"
		sql = sql & "UPDATE CodeMain SET mValue = '" & topic & "' WHERE codeMetaID = 'commendWord' AND mCode = 'topic';"
		sql = sql & "UPDATE CodeMain SET mValue = '" & knowledge & "' WHERE codeMetaID = 'commendWord' AND mCode = 'knowledge';"		
		conn.execute(sql)
		
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
	    	document.location.href="CommendWordSetting.asp"	    
			</script>
    </body>
  </html>
<% End sub %> 

<%	
sub showForm()
	
	sql = "SELECT * FROM CodeMain WHERE codeMetaID = 'commendWord' ORDER BY mSortValue"
	Set rs = conn.execute(sql)
	If rs.eof Then
		response.write "<script>alert('CodeMain找不到設定值');history.back();</script>"
	Else
		Dim inter : inter = "0"
		Dim tank : tank = "0"
		Dim topic : topic = "0"
		Dim knowledge : knowledge = "0"		
		While Not rs.EOF 		
			if rs("mCode") = "inter" then inter = rs("mValue")
			if rs("mCode") = "tank" then tank = rs("mValue")
			if rs("mCode") = "topic" then topic = rs("mValue")
			if rs("mCode") = "knowledge" then knowledge = rs("mValue")
			rs.MoveNext
		Wend	
		Set rs = nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
	<link rel="stylesheet" href="/inc/setstyle.css">
	<link type="text/css" rel="stylesheet" href="/css/list.css">
	<link type="text/css" rel="stylesheet" href="/css/layout.css">
	<link type="text/css" rel="stylesheet" href="/css/setstyle.css">
<title></title>
</head>
<body>
	<table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
	<tr><td width="50%" align="left" nowrap class="FormName">前台開放推薦詞彙設定&nbsp;</td></tr>
	<tr><td width="100%" colspan="2"><hr noshade size="1" color="#000080"></td></tr>
	<tr><td class="Formtext" colspan="2" height="15"></td></tr>  
  <tr>
    <td width="95%" colspan="2" height=230 valign=top>
		<Form name="reg" method="POST" action="CommendWordSetting.asp">  
		<input name="actionType" type="hidden" value="edit" />
    <CENTER>
	  <table width=100% cellpadding="0" cellspacing="1" id="ListTable">
    <tr align=left>
			<th>項目</th>
      <th width="15%">是否開放推薦詞彙</th>
    </tr>	
    <tr>
      <td class="eTableContent">入口網</td>
      <td align="right" class="eTableContent">
				<select name="inter">
					<option value="1" <% if inter = "1" Then %> selected <% end if %> >開放</option>
					<option value="0" <% if inter = "0" Then %> selected <% end if %> >不開放</option>
				</select>
			</td>
    </tr>
    <tr>
			<td class="eTableContent">知識庫</td>
      <td align="right" class="eTableContent">
				<select name="tank">
					<option value="1" <% if tank = "1" Then %> selected <% end if %> >開放</option>
					<option value="0" <% if tank = "0" Then %> selected <% end if %> >不開放</option>
				</select>
			</td>
    </tr>
    <tr>
      <td class="eTableContent">主題館</td>
      <td align="right" class="eTableContent">
				<select name="topic">
					<option value="1" <% if topic = "1" Then %> selected <% end if %> >開放</option>
					<option value="0" <% if topic = "0" Then %> selected <% end if %> >不開放</option>
				</select>
			</td>
    </tr>
    <tr>
			<td class="eTableContent">知識家</td>
      <td align="right" class="eTableContent">
				<select name="knowledge">
					<option value="1" <% if knowledge = "1" Then %> selected <% end if %> >開放</option>
					<option value="0" <% if knowledge = "0" Then %> selected <% end if %> >不開放</option>
				</select>
			</td>
    </tr>
    </table>
    </CENTER>
		<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center"><input name="SubmitBtn" type="submit" class="cbutton" value="編修存檔"></td>
		</tr>
		</table>
		</form>  
		</td>
  </tr>  
	<tr><td width="100%" colspan="2" align="center"></td></tr>
	</table> 
</body>
</html>                                 
<%
	end if 	
End Sub
%>

