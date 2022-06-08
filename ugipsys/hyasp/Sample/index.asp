
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
<form method="post">
<br>Enter new value   :  <input name="query" type="text"><br>
<br><br>
<input type="submit" name="Submit" value="Submit">
<input type="reset" value="Reset">
<br>

<%
   
   gg="gg"
   ' step 1: 使用 CreateObject 產生 Object. DLL name: hysdk, class name: hyft
   Set obj = Server.CreateObject("hysdk.hyft.1")

   ' step 2: 呼叫 hyft_connect 連結 hyftd
   x = obj.hyft_connect("localhost", 2816, "gg", "", "")

   ' hyft_connect() 傳回值為連線 id, 存放於 obj.ret_val 中
   id = obj.ret_val
  
   if id = -1 then
      '若是連線失敗, 則錯誤訊息存放於 obj.errmsg
      response.write "連線 hyftd 發生錯誤, 錯誤訊息: " & obj.errmsg & "<br><br>"
      response.end
   else
      'response.write "連線 hyftd 成功! id = " & id & "<br><br>"
   end if

   ' step 3: 初始化查詢環境
   x = obj.hyft_initquery(id)
   ' 傳回值為 obj.ret_val
   ret = obj.ret_val 
   if ret = -1 then
      '若是初始化失敗, 則錯誤訊息存放於 obj.errmsg
      response.write "initquery 發生錯誤, 錯誤訊息: " & obj.errmsg
      response.end
   end if
   
   query = Request("query")
   
   x = obj.hyft_addquery(id, "", query, "", 0, 0, 0)   
   ret = obj.ret_val
   if ret = -1 then
      response.write "addquery 發生錯誤, 錯誤訊息: " & obj.errmsg
      response.end
   end if

   'step 5: 呼叫 hyft_query, 開始查詢(筆數為1至1000)
    x = obj.hyft_query(id, 0,1000)
 	res_id= obj.ret_val
 	'response.write "res_id = " & res_id & "<br><br>"
   	
   	' step 6: 取得筆數,筆數傳回值存放於 obj.ret_val
   	x = obj.hyft_num_sysid(id,res_id)
   	num = obj.ret_val
   	
   	' step 7: 取得相關資料的id
   	x = obj.hyft_get_assdata_by_resultset(id,res_id,"gg")
   	assdata_id = obj.ret_val 
   	
   	' step 8: 取得索引及相關資料內容
   	if num >0 then
   		for i = 0 to num-1
			
			'抓出每一筆資料的SYSID
			x = obj.hyft_fetch_sysid(id, res_id, i)
        	sysid = obj.sysid
        	'response.write "sysid=***" & sysid & "***<br>"
        	
        	'抓出每一筆資料的相關資料長度
        	x=obj.hyft_one_assdata_len(id,gg, sysid)
        	data_len=obj.ret_val
        	'response.write "data_len=***" & data_len& "***<br>"
        	
        	'抓出每一筆資料的相關資料內容
        	x=obj.hyft_one_assdata(id,data_len)
       		data =obj.data   
       		response.write "data=***" &  data & "***<br>"
    		
    	next 
    	
    	'釋放相關資料id
    	x=obj.hyft_free_assdata(id, assdata_id)
   	else
   		response.write "查無資料"
   	end if
   	
   set obj=Nothing	
%>
</form>
</body>
</html>
