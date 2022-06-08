<%@ CodePage = 65001 %>
<% Response.Expires = 0 
response.charset = "utf-8"
HTProgCode = "webgeb1"
   %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/inc/selectTree.inc" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8; no-caches">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../inc/setstyle.css">
<title>主題館權限管理</title>
</head>
<body bgcolor="#FFFFFF" background="../images/management/gridbg.gif">
<table border="0" width="100%" cellspacing="1" cellpadding="0">
  <tr>
    <td width="100%" class="FormName" colspan="2">系統管理／主題館權限設定</td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
    <td class="Formtext" width="60%">【使用者查詢引擎】不設定任何條件可查詢全部使用者資料</td>
  </tr>
  <tr>
    <td class="Formtext" colspan="2" height="15"></td>
  </tr>
  <tr>
    <td width="95%" colspan="2">
      <form method="POST" name="reg">
       <input type=hidden name="CanlendarTarget" value="">    
          <center>    
          <table border="0" width="90%" cellspacing="1" cellpadding="3" class="bluetable">    
            <tr>    
              <td class="lightbluetable" align="right">使用帳號</td>    
              <td class="whitetablebg"><input type="text" name="tfx_userId" size="15"></td>    
            </tr>
            <tr>          	  
              <td class="lightbluetable" align="right">姓     名</td>
              <td class="whitetablebg"><input type="text" name="tfx_userName" size="15"></td>
            </tr>
            <tr>          	  
              <td class="lightbluetable" align="right">單     位</td>
              <td class="whitetablebg">
              <Select name="sfx_deptId" size=1>
<%
	sqlCom ="SELECT D.deptId,D.deptName,D.parent,len(D.deptId)-1,D.seq," & _
		"(Select Count(*) from Dept WHERE parent=D.deptId AND nodeKind='D') " & _
		"  FROM Dept AS D Where D.nodeKind='D' " _
		& " AND D.deptId LIKE N'" & session("deptId") & "%'" _
		& " ORDER BY len(D.deptId), D.parent, D.seq"	
'response.write SqlCom
	set RSS = conn.execute(sqlCom)
	if not RSS.EOF then
		ARYDept = RSS.getrows(300)
		glastmsglevel = 0
		genlist 0, 0, 1, 0
	        expandfrom ARYDept(cid, 0), 0, 0
	end if
%>
</select><BR/>
              </td>
            </tr>
          </table>
          </center>
            <p align="center">  
            <img src="../images/management/titlehr.gif" width="570" height="1" vspace="3"><br>
            <% if (HTProgRight and 2)=2 then %><input type="button" value="確定查詢" name="task" class="cbutton" OnClick="formSearch()">&nbsp;
            <input type="reset" value="清除重填" class="cbutton"><% End If %>         
      </form>          
    </td>
  </tr>  
  <tr>                               
    <td width="100%" colspan="2" align="center" class="English">                               
      </td>                                         
  </tr> 
</table>
</body>
</html> 
<script language="vbscript">
Sub formSearch()
    reg.action="UserList.asp?id=<%=request("id")%>&type=<%=request("type")%>"
    reg.Submit
End Sub
</script> 
