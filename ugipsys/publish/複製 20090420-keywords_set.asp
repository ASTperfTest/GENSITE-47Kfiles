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
HTProgCode="publish"
HTProgPrefix = "publish"
%>
<!--#include virtual = "/inc/server.inc" -->
<Html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
	<link rel="stylesheet" href="/inc/setstyle.css">
	<link type="text/css" rel="stylesheet" href="../css/list.css">
	<link type="text/css" rel="stylesheet" href="../css/layout.css">
	<link type="text/css" rel="stylesheet" href="../css/setstyle.css">
	<title></title>  
</head>
<body>
	<table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
  <tr>
	  <td width="50%" align="left" nowrap class="FormName">關連發佈單元設定&nbsp;</td>
		<td width="50%" class="FormLink" align="right"></td>	
  </tr>
	<tr>
	  <td width="100%" colspan="2"><hr noshade size="1" color="#000080"></td>
	</tr>
	<tr>
	  <td class="Formtext" colspan="2" height="15"></td>
	</tr>  
  <tr>
    <td width="95%" colspan="2" height=230 valign=top>		
			<CENTER>
			<table width=100% cellpadding="0" cellspacing="1" id="ListTable">
      <tr align=left>
        <th>發佈名稱</th>
        <th width="15%">單元設定</th>
      </tr>
      <tr>
        <td class=eTableContent>知識家問題發佈</td>
        <td class=eTableContent><input name="button4" type="button" class="cbutton" onClick="redirect('knowledge')" value="新增修改"></td>
      </tr>
      <tr>
        <td class=eTableContent>新聞發佈</td>
        <td class=eTableContent><input name="button42" type="button" class="cbutton" onClick="redirect('news')" value="新增修改"></td>
      </tr>      
      </table>
			</CENTER>      
			<table border="0" width="100%" cellspacing="0" cellpadding="0">	</table>     
    </td>
  </tr>  
	<tr>
		<td width="100%" colspan="2" align="center"></td>
	</tr>
	</table> 
</body>
</html>                                 
<script language="javascript">	
	function redirect(type) 
	{
		window.location.href = "publish_ctNodeset.asp?type=" + type;
  }	
</script>	

