<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "UUC" %>
<!--#include virtual = "/inc/index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"; cache=no-caches;>
<meta http-equiv="pragma: nocache">
<%
	FtypeName = "特殊單元設定"
%>
<script language=javascript>
 var OpenFolder = "/imageFolder/ftv2folderclosed.gif"
 var CloseFolder = "/imageFolder/ftv2folderopen.gif"
 var DocImage = "/imageFolder/openfold.gif"
</script>
<script src="/imageFolder/FolderTree.js"></script>
<title></title>
<style>
.FolderStyle { text-decoration: none; color:#000000}
</style>
</head>
<body leftmargin="3" topmargin="3">
<script language="javascript">
function checkform()
{
    if (document.form1.userId.value == "")
    {
        alert('請填寫可讀取下列單元之帳號！');
        document.form1.userId.focus();
        return false;
    }
    return true;
}
</script>
特殊單元設定<font color="#FFFFFF">(CtUnitUser)</font><HR>

<%
  userId = request("userId")
%>
<form name="form1" method="post" action="CtUnitControl_act.asp" onSubmit="return checkform();">
<input type="hidden" name="userId" value="<%=userId %>" >
<%
    SQLCom = "select * from CtUnit Where 1=1 "
    SQLCom = SQLCom & " AND (deptId IS NULL OR deptId LIKE N'" & session("deptId") & "%')"
    SQLCom = SQLCom & " Order By CtUnitName, CtUnitId"
    set RSUnit = conn.execute(SQLCom)
	while not RSUnit.eof

        ForumCheck = "<input type=checkbox name='ctUnitId' value='" & RSUnit("ctUnitId") & "' >" & RSUnit("CtUnitName")
        
        SQLCom = "select * from CtUnitUser Where (userId='"& userId &"')"
        SQLCom = SQLCom & " AND (CtUnitId ="& RSUnit("ctUnitId") &")"
        set RSUserUnit = conn.execute(SQLCom)
        if not RSUserUnit.eof then
	     if cint(RSUserUnit("rights")) = 0 then
	     else
	       ForumCheck = "<input type=checkbox name='ctUnitId' value='" & RSUnit("ctUnitId") & "' checked>" & RSUnit("CtUnitName")
	     end if
        end if
        Response.Write "<font style='font-size:80%'>"& ForumCheck & "</font><br />"
		RSUnit.moveNext
	wend
%>
<HR>
<input name="submit" type="submit" value="確定"  onClick="return confirm('確定嗎？');">
</form>
</body></html>

