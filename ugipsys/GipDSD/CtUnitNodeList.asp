<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="主題單元管理"
HTProgFunc="項目節點清單"
HTProgCode="GE1T11"
HTProgPrefix="CtUnit" %>
<!--#include virtual = "/inc/server.inc" -->
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="/css/list.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<title><%=Title%></title>
</head>
<!--#include virtual = "/inc/dbutil.inc" -->
<%

 Set RSreg = Server.CreateObject("ADODB.RecordSet")
 fSql = "SELECT htx.*, r.ctRootName FROM CatTreeNode AS htx JOIN CatTreeRoot AS r ON r.ctRootId=htx.ctRootId" _
 	& " WHERE ctUnitId=" & pkStr(request("ctUnitId"),"")
	SET RSreg = conn.execute(fSql)
%>
<body>
<div id="FuncName">
	<h1><%=Title%></h1>
	<div id="Nav">
			<A href="#" onClick="VBS: window.close()" title="關閉">關閉</A> 
	</div>
	<div id="ClearFloat"></div>
</div>
<div id="FormName"><%=HTProgFunc%></div>
	<table cellspacing="0" id="ListTable">
     <tr>    
	<th>目錄樹</th>
	<th>項目名稱</th>
    </tr>	                                
<%                   
    while not RSreg.eof                  
%>                  
	<TD class="Center"><font size=2>
<%=RSreg("ctRootName")%>
</font></td>
	<TD class="Center"><font size=2>
<%=RSreg("CatName")%>
</font></td>
    </tr>
    <%
         RSreg.moveNext
    wend 
   %>
    </TABLE>
    </CENTER>
       <!-- 程式結束 ---------------------------------->  
     </form>  
    </td>
  </tr>  
	<tr>
		<td width="100%" colspan="2" align="center">
	</td></tr>
</table> 
</form>
</body>
</html>                                 

