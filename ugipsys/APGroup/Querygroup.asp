<%@ CodePage = 65001 %>
<% Response.Expires = 0 
   HTProgCode = "HT001"
%>
<!--#include virtual = "/inc/server.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="/css/form.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<title><%=Title%></title>
</head>

<body>
<div id="FuncName">
	<h1><%=Title%></h1>
	<div id="Nav">
 <% if (HTProgRight and 4)=4 then %><a href="../APGroup/Addgroup.asp">新增</a><% End IF %>
	</div>
	<div id="ClearFloat"></div>
</div>
<div id="FormName">權限群組查詢</div>

<form id="Form1" name=fmGrpQuery method="POST" action="ListGroup.asp"> 
  <table cellspacing="0">
     <tr>    
      <td align="right" class="Label">群組ID</td>     
      <td><input name="fgugrpId" class="InputText" size="20"> </td>     
     </tr>
     <tr>    
      <td align="right" class="Label">群組名稱</td>     
      <td class="whitetablebg"><input name="fgugrpName" size="20"> </td>     
     </tr>     
     <tr>       
      <td align="right" class="Label">說明</td>      
      <td colspan=3 class="whitetablebg"><input name="fgremark" size="20"></td>      
    </tr>      
	<TR><TD class="Label" align="right">是否公開</TD>
	<TD class="whitetablebg"><Select name="fgisPublic" size=1>
	<option value="">請選擇</option>
				<%SQL="Select mcode,mvalue from CodeMain where codeMetaId=N'boolYN' Order by msortValue"
				SET RSS=conn.execute(SQL)
				While not RSS.EOF%>
	
				<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
				<%	RSS.movenext
				wend%>
	</select></TD></TR>
	</table>      
   <% if (HTProgRight and 2)=2 then %>
   	<input type="submit" value="查詢" name="task" class="cbutton">
   					<input type=reset value="清除重填" class="cbutton">
   <% End IF %>    
    </td>      
  </tr>             
  </table>            
</form>    

<div id="Explain">
	<h1>說明</h1>
	<ul>
		<li>不設定任何條件可查詢全部權限群組資料</li> 
	</ul>
</div>
  
</body></html>            
