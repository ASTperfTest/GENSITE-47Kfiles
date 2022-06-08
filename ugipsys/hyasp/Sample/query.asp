<html>
<title>Hysearch 全文檢索引擎 ASP 測試範例</title>
<body>
<center>
<br>
<h3>HySearch 全文檢索系統  ASP 版本測試網頁</h3><br>
查詢結果<br><br>
<%
   ' step 1: 使用 CreateObject 產生 Object. DLL name: hysdk, class name: hyft
   Set obj = Server.CreateObject("hysdk.hyft.1")

   ' step 2: 呼叫 hyft_connect 連結 hyftd
   x = obj.hyft_connect("localhost", 2816, "WRITER", "", "")

   ' hyft_connect() 傳回值為連線 id, 存放於 obj.ret_val 中
   id = obj.ret_val
   if id = -1 then
      '若是連線失敗, 則錯誤訊息存放於 obj.errmsg
      response.write "連線 hyftd 發生錯誤, 錯誤訊息: " & obj.errmsg & "<br><br>"
      response.end
   else
      response.write "連線 hyftd 成功! id = " & id & "<br><br>"
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

   ' step 4: 開始呼叫 hyft_addquery, 將要查詢之欄位資料填入
   response.write "查詢條件：au=法蘭克林<br><br>"
   x = obj.hyft_addquery(id, "au", "法蘭克林", "", 0, 0, 0)
   if ret = -1 then
      response.write "addquery 發生錯誤, 錯誤訊息: " & obj.errmsg
      response.end
   end if

   ' step 5: 呼叫 hyft_query, 開始查詢
   x = obj.hyft_query(id, 0, 1000)
   res_id = obj.ret_val
   if res_id = -1 then
      response.write "query 發生錯誤, 錯誤訊息: " & obj.errmsg
      response.end
   end if 

   ' step 6: 取得筆數
   x = obj.hyft_num_sysid(id, res_id)
   ' 筆數傳回值存放於 obj.ret_val
   num = obj.ret_val
   if num < 0 then
      response.write "num_sysid 發生錯誤, 錯誤訊息: " & obj.errmsg
      response.end
   end if 
   ' num > 0 表示有查到資料
   if num > 0 then
      ' step 6-1: 一次抓取一筆 sysid
      for i = 0 to num-1
	x = obj.hyft_fetch_sysid(id, res_id, i)
        '抓取出來之 sysid 存放於 obj.sysid
        sysid = obj.sysid
        response.write "sysid=***" & sysid & "***<br><br>"
      next
      response.write "完成!"
   else
      response.write "num = 0, no data found."
   end if
   set obj=Nothing	
%>
<br><br>
<a href="javascript:history.go(-1)">回上頁</a>
</center>
</body>
</html>
