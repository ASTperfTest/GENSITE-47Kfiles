<%@ CodePage = 65001 %>
<% Response.Expires = 0 
   HTProgCode = "Pn50M06"
%>
<!--#include virtual = "/inc/server.inc" -->
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../inc/setstyle.css">
<title>代碼權限群組查詢引擎</title></head><body bgcolor="#FFFFFF" background="../images/management/gridbg.gif">
<table border="0" width="95%" cellspacing="1" cellpadding="0">
  <tr>
    <td width="100%" class="FormName" colspan="2"><%=Title%></td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
    <td class="Formtext" width="60%">【代碼權限群組查詢引擎】</td>
    <td class="FormLink" valign="top" width="40%" align="right">
    </td>
  </tr>
  <tr>
    <td width="100%" colspan="2" class="FormRtext">&nbsp;不設定任何條件可查詢全部代碼權限群組資料</td>
  </tr>  
  <tr>
    <td class="Formtext" colspan="2" height="15"></td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
	<form name=fmGrpQuery method="POST" action="ListGroup.asp"> 
    <center>
	 <table border="0" cellspacing="1" cellpadding="3" width="70%" class="bluetable">
     <tr>    
      <td align="center" class="lightbluetable">群組ID</td>     
      <td class="whitetablebg"><input name="fgugrpId" size="20"> </td>     
     </tr>
     <tr>    
      <td align="center" class="lightbluetable">群組名稱</td>     
      <td class="whitetablebg"><input name="fgugrpName" size="20"> </td>     
     </tr>     
     <tr>       
      <td align="center" class="lightbluetable">說明</td>      
      <td colspan=3 class="whitetablebg"><input name="fgremark" size="20"></td>      
    </tr>      
	</table>      
   </center>     
   <p align="center"> 
   <img src="../images/management/titlehr.gif" width="400" height="1" vspace="3"><br>     
   <% if (HTProgRight and 2)=2 then %><input type="submit" value="查詢" name="task" class="cbutton"><% End IF %>    
   </form>      
    </td>      
  </tr>      
  <tr>                               
    <td width="100%" colspan="2" align="center" class="English">                               
      <!--#include virtual = "/inc/Footer.inc" -->
      </td>                                         
  </tr>           
</table>            
</body></html>            
