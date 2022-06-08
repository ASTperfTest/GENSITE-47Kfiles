<%@ CodePage = 65001 %>
<% Response.Expires = 0 
   HTProgCode = "HT001"

' step 1: 使用 CreateObject 產生 Object. DLL name: hysdk, class name: hyft  
Set obj = Server.CreateObject("hysdk.hyft.1")

' step 2: 呼叫 hyft_connect 連結 hyftd	hyft_connect() 傳回值為連線 id, 存放於 obj.ret_val 中
x = obj.hyft_connect("localhost", 2816, "gg", "", "")
id = obj.ret_val

if id = -1 then
	'若是連線失敗, 則錯誤訊息存放於 obj.errmsg
      	response.write "連線 hyftd 發生錯誤, 錯誤訊息: " & obj.errmsg & "<br><br>"
      	response.end
end if

query="今天我們要一刀兩斷真的一刀兩斷"

' step 3: 啟用hyft_textkeys斷詞功能,如果成功,會得到一個key_id,
x=obj.hyft_textkeys(id, query,10)
key_id = obj.ret_val
if key_id = -1 then
	'若是連線失敗, 則錯誤訊息存放於 obj.errmsg
      	response.write "連線 hyftd 發生錯誤, 錯誤訊息: " & obj.errmsg & "<br><br>"
      	response.end
end if

' step 4: 算出query中斷詞的數目
x=obj.hyft_num_keys(id,key_id)
num =obj.ret_val
response.write "共有" & num & "個斷詞<br><br>"

for i = 0 to num-1
	x=obj.hyft_keylen(id,key_id,i)
	key_len=obj.ret_val
			
	response.write "key_len:" & key_len & "<br><br>"
	x=obj.hyft_onekey(id,key_id,i,key_len)
	'response.write "錯誤訊息: " & obj.errmsg & "<br><br>"
	key=obj.key
	response.write "key:" & key & "<br><br>"
		
next 
	
	
x=obj.hyft_free_key(id,key_id)
	
set obj=Nothing	
response.write "<br>Done5!"   
   
%>
<!--#Include file = "../inc/server.inc" -->
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../inc/setstyle.css">
<title>權限群組查詢引擎</title></head><body bgcolor="#FFFFFF" background="../images/management/gridbg.gif">
<table border="0" width="100%" cellspacing="1" cellpadding="0">
  <tr>
    <td width="100%" class="FormName" colspan="2"><%=Title%></td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
    <td class="Formtext" width="60%">【權限群組查詢引擎】</td>
    <td class="FormLink" valign="top" width="40%" align="right">
 <% if (HTProgRight and 4)=4 then %><a href="../APGroup/Addgroup.asp">新增</a><% End IF %>
    </td>
  </tr>
  <tr>
    <td width="100%" colspan="2" class="FormRtext">&nbsp;不設定任何條件可查詢全部權限群組資料</td>
  </tr>  
  <tr>
    <td class="Formtext" colspan="2" height="15"></td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
	<form name=fmGrpQuery method="POST" action="ListGroup.asp"> 
    <center>
	 <table border="0" cellspacing="1" cellpadding="3" width="400" class="bluetable">
     <tr>    
      <td align="right" class="lightbluetable">群組ID</td>     
      <td class="whitetablebg"><input name="fgugrpID" size="20"> </td>     
     </tr>
     <tr>    
      <td align="right" class="lightbluetable">群組名稱</td>     
      <td class="whitetablebg"><input name="fgugrpName" size="20"> </td>     
     </tr>     
     <tr>       
      <td align="right" class="lightbluetable">說明</td>      
      <td colspan=3 class="whitetablebg"><input name="fgRemark" size="20"></td>      
    </tr>      
	</table>      
   </center>     
   <p align="center"> 
   <img src="../images/management/titlehr.gif" width="400" height="1" vspace="3"><br>     
   <% if (HTProgRight and 2)=2 then %><input type="submit" value="查詢" name="task" class="cbutton">
   					<input type=reset value="重設" class="cbutton">
   <% End IF %>    
   </form>      
    </td>      
  </tr>             
</table>            
</body></html>            
