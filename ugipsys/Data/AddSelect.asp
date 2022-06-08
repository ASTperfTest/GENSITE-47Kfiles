<%@ CodePage = 65001 %>
<!--#Include file = "index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
<title>新增類別</title>
</head>
<body topmargin="0" leftmargin="0">
<form name="reg" method="POST">
<table border="0" width="245" cellspacing="1" cellpadding="3" id="SetOrderView" class="bluetable" height="100%">
  <tr class="lightbluetable">
    <td>
      <table border="0" width="100%" cellspacing="0" cellpadding="0">
        <tr class="lightbluetable">
          <td>新增類別</td>
          <td align="right"></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr class="whitetablebg">
    <td align="center">

			<input name="AddSelectItem" type="hidden" value="">
                <table border="0" cellpadding="3" cellspacing="1">
      	    	  <tr class="whitetablebg">
      	    	     <td align="right">類別名稱</td><TD><input name="CatName" type="text" size="10"></TD>
      	    	  </tr>
      	    	</table>
                <p align="right">
				<% if (HTProgRight and 4) = 4 then %><input type="button" value="確定" class="cbutton" OnClick="AddForm()"><% End if %>
		      	<input type="button" value="取消" class="cbutton" OnClick="VBS: window.close"></p>

    </td>
  </tr>
</table>
</form> 
</body></html>
<script language="VBscript">
 Sub AddForm()

  msg1 = "請輸入分類名稱！"
  If trim(reg.CatName.value) = Empty Then
     MsgBox msg1, 64, "Sorry!"
     reg.CatName.focus
     Exit Sub
  End if

<% If Language = "C" Then %>
  msg = "抱歉!"& vbcrlf & vbcrlf &"類別名稱字數過多，請勿超過10個中文字。"& vbcrlf & vbcrlf &"請重新輸入 !"               
<% ElseIf Language = "E" Then %>
  msg = "抱歉!"& vbcrlf & vbcrlf &"類別名稱字數過多，請勿超過20個字元。"& vbcrlf & vbcrlf &"請重新輸入 !"           
<% End IF %>
  If blen(reg.CatName.value) > 20 Then
   MsgBox msg, 64, "Sorry!"          
   reg.CatName.focus          
   Exit Sub
  End IF

  CatName = reg.CatName.value
  window.ReturnValue = "AddSelectConfig.asp?CatName="& CatName &"&DataType=<%=DataType%>&Language=<%=Language%>" 
  window.close 
 End Sub
 
function blen(xs)
 xl = len(xs)
 for i=1 to len(xs)
  if asc(mid(xs,i,1))<0 then xl = xl + 1
 next
 blen = xl
end function
 
</script>