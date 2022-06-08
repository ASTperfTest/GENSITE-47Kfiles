
<html>
<head>
<meta http-equiv=Content-Type content="text/html; charset=utf-8">
<title>
Hyftd_asp
</title>
</head>

<body>
<h1>
Hyftd asp_demo
</h1>
<form method="post" action="show_key.asp">
<br>Enter new value   :  <input name="query" type="text"><br>
<br><br>
<input type="submit" name="Submit" value="Submit">
<input type="reset" value="Reset">
<br>

<%
   
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

   query = Request("query")
 	
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
%>
</form>
</body>
</html>
