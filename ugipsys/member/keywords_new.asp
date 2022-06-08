<%@ CodePage = 65001 %>
<% 
Response.Expires = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.CacheControl = "Private"
HTProgCap="單元資料維護"
HTProgFunc="編修"
HTUploadPath=session("Public")+"data/"
HTProgCode="newmember"
HTProgPrefix = "newMember"
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/client.inc" -->
<!--#include virtual = "/inc/dbUtil.inc" -->
<%
	Dim memberId : memberId = request.querystring("memberId")
	Dim add : add = request.querystring("add")
	
	if add = "Y" then
		addKeyword
	else
		showForm
	end if	
%>
<% 
Sub addKeyword 

	Dim keyword : keyword = ""
	Dim newkeyword : newkeyword = request("keyword")
	
	sql = "SELECT ISNULL(keyword, '') AS keyword FROM Member WHERE account = " & pkstr(memberId, "")
	set rs = conn.execute(sql)
	if not rs.eof then
		keyword = rs("keyword")
	end if
	rs.close
	set rs = nothing
	
	keyword = keyword & newkeyword & ","
	sql = "UPDATE Member SET keyword = " & pkstr(keyword, "") & " WHERE account = " & pkstr(memberId, "")
	conn.execute(sql)
	showDoneBox "新增完成"
End Sub
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
				document.location.href="keywords_set.asp?memberId=<%=memberId%>&nowPage=<%=request.querystring("nowPage")%>&pagesize=<%=request.querystring("pagesize")%>"				
			</script>
    </body>
  </html>
<% End sub %>
<% Sub showForm %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html> 
  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <meta name="GENERATOR" content="Hometown Code Generator 1.0">
    <meta name="ProgId" content="kpiQuery.asp">
    <title>查詢表單</title>
		<link type="text/css" rel="stylesheet" href="../css/list.css">
		<link type="text/css" rel="stylesheet" href="../css/layout.css">
		<link type="text/css" rel="stylesheet" href="../css/setstyle.css">
    <script src="../Scripts/AC_ActiveX.js" type="text/javascript"></script>
    <script src="../Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
    <style type="text/css">
		<!--
			.style1 {color: #000000}
		-->
    </style>    
  </head>
	<body>
		<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
	    <td width="50%" align="left" nowrap class="FormName">專家專長關鍵字設定：自訂關鍵字&nbsp;		</td>
			<td width="50%" class="FormLink" align="right"><A href="Javascript:window.history.back();" title="回前頁">回前頁</A></td>	
		</tr>
	  <tr>
	    <td width="100%" colspan="2"><hr noshade size="1" color="#000080"></td>
	  </tr>
		<tr>
			<td class="Formtext" colspan="2" height="15"></td>
		</tr>
	  <tr>
	    <td align=center colspan=2 width=80% height=230 valign=top>   
			<form name="iForm" method="post" action="keywords_new.asp?add=Y&memberId=<%=memberId%>&nowPage=<%=request("nowPage")%>&pagesize=<%=request("pagesize")%>">
			<CENTER>
        <TABLE width="100%" id="ListTable">
        <TR><th>關鍵字</th></TR>
        <TR>
          <td class="eTableContent"><input name="keyword" value="請輸入關鍵字" size="50"></td>
        </TR>
        </TABLE>
			</CENTER>
			<table border="0" width="100%" cellspacing="0" cellpadding="0">
			<tr>
				<td align="center">     
					<input name="button4" type="button" class="cbutton" onClick="formSubmit()" value="新　增">
					<input type=button value ="重　設" class="cbutton" onClick="formReset()">
					<input type=button value ="取　消" class="cbutton" onClick="formCancel()">
				</td>
			</tr>
			</table>
			</form>
		</tr>
		</table>
	</body>
</html>
<script language="javascript">	
	function formSubmit() 
	{
		if(document.getElementById("keyword").value == "" || document.getElementById("keyword").value == "請輸入關鍵字" ) {
			alert("請輸入關鍵字");
		}
		else {
			document.getElementById("iForm").submit();
		}		
  }
	function formReset() 
	{
		window.location.href = "keywords_new.asp?memberId=<%=memberId%>&nowPage=<%=request("nowPage")%>&pagesize=<%=request("pagesize")%>";
  }
	function formCancel() 
	{
		window.location.href = "keywords_set.asp?memberId=<%=memberId%>&nowPage=<%=request("nowPage")%>&pagesize=<%=request("pagesize")%>";
  }
</script>	
<% end sub %>




