<!--#Include virtual = "/inc/dbFunc.inc" -->
<!--#Include virtual = "/inc/client.inc" -->
<!--#include virtual = "/inc/web.sqlInjection.inc" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<?xml version="1.0"  encoding="utf-8" ?>
<html xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:hyweb="urn:gip-hyweb-com" xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title></title>
<script type="text/javascript">
        function setHref(url) {
            var isIE = (navigator.appName.indexOf("Microsoft") != -1) ? true : false;
	    if (!isIE) {
                parent.location.href = url;
            } else {
                var lha = document.getElementById('_lha');
                if (!lha) {
                    lha = document.createElement('a');
                    lha.id = '_lha';
                    lha.target = '_parent';  // Set target: for IE
                    document.body.appendChild(lha);
                }
                lha.href = url;
                lha.click();
            }
        }
    </script>

</head>
<body>
<%
url = Request.QueryString("sourceUrl")

userId = session("memID")

if (userId = "") then 
	Response.Write "<script>alert('請先登入會員!');setHref('/login/login.asp');</script>"
else
	guid = GetGuid()
	sql = "INSERT INTO SSO (GUID , USER_ID , CREATION_DATETIME ) VALUES ( '" & guid & "','" & userId & "',getdate())"
	set rs = conn.execute(sql)
	'rs.close
	'set rs = nothing
	if ( InStr(url,"?") = 0 ) then
	
				Response.Redirect  url + "?guid=" + guId 
	else
				Response.Redirect  url + "&guid=" + guId 	
	end if
end if

Function GetGuid()
    Set TypeLib = Server.CreateObject("Scriptlet.TypeLib")
    tg = TypeLib.Guid
	 GetGuid = mid(tg, 2 ,len(tg)-4)
    Set TypeLib = Nothing
End Function
%>
</body>
</html>
