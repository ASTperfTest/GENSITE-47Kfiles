<%Response.Expires = 0
HTProgCap="主題館績效統計"
HTProgFunc="查詢清單"
HTProgCode="KForumlist"
HTProgPrefix="KForumlist" %>
<!--#include virtual = "/inc/server.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="/css/list.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<title></title>
</head>

<Form id="reg" name="reg" method="POST" action="knowledgeQuery_Act.asp">
 <INPUT TYPE=hidden name=submitTask value="">
 
<body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
	  <td width="50%" align="left" nowrap class="FormName">圖片查詢&nbsp;</td>
		<td width="50%" class="FormLink" align="right">
			<A href="Javascript:window.history.back();" title="回前頁">回前頁</A>			
		</td>	
  </tr>
	<tr>
	  <td width="100%" colspan="2"><hr noshade size="1" color="#000080"></td>
	</tr>
	<tr>
		<td class="Formtext" colspan="2" height="15"></td>
	</tr>  
	<tr>
		<td align=center colspan=2 width=80% height=230 valign=top>    
		
		<CENTER>
		<TABLE border="0" id="ListTable" cellspacing="1" cellpadding="2">
		<TR>
			<Th>審核狀態</Th>
			<TD class="eTableContent">
				<select name="validate" id="validate" >
  		
    <option value="" <%if validate="" then%>selected<%END IF%>>請選擇</option>
    <option value="" <%if validate="all" then%>selected<%END IF%>>全部</option>
    <option value="W" <%if validate="W" then%>selected<%END IF%>>待審核</option>
    <option value="Y" <%if validate="Y" then%>selected<%END IF%>>審核通過</option>
    <option value="N" <%if validate="N" then%>selected<%END IF%>>審核不通過</option>
  </select>
			</TD>
		</TR>
		<TR>
		<Th>上傳者(帳號、姓名、暱稱)</Th>
			<TD class="eTableContent"><input name="memberId" size="50" value="" ></TD>
		</TR>
			<TR>
			<Th>圖片標題</Th>
			<TD class="eTableContent"><input name="picTitle" size="50" value=""></TD>
		</TR>
		<TR>
		<Th>上傳資料ID</Th>
			<TD class="eTableContent"><input name="PicidTerms" size="50" value=""></TD>
		</TR>
		</TABLE>
		</td>
	</tr>
	</table>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td width="100%">     
			<p align="center">
				<input name="button" type="button" class="cbutton" onclick="QuerySubmit()" value ="查 詢">
                <input type="button" value ="重　填" class="cbutton" onClick="returnToEdit()">        
		</td>
	</tr>
	</table>

</TABLE>
         
       <!-- 程式結束 ---------------------------------->  
</form>  
</body>
</html>     

<script language="javascript">
	function QuerySubmit() 
	{
	    //document.getElementById("picId").value = id;
		//document.getElementById("parentIcuitem").value = parentIcuitem;
		
		document.getElementById("reg").submit();
	}
	function returnToEdit() {
		//window.location.href = "cp_question.asp?iCUItem=<%=request.querystring("iCUItem")%>&nowPage=<%=request.querystring("nowPage")%>&pagesize=<%=cint(request.querystring("pagesize"))%>";
	     window.location.href = "knowledgeQuery.asp";
	
	}
	
</script>     
                                                   

</body>
</html>
