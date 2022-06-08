<html>
<title>Hysearch 全文檢索引擎 ASP 測試範例</title>
<body>
<center>
<br>
<h3>HySearch 全文檢索系統  ASP 版本測試網頁</h3><br><br>
<%
   ' step 1: 使用 CreateObject 產生 Object. DLL name: hysdk, class name: hyft
   Set obj = Server.CreateObject("hysdk.hyft.1")

   ' step 2: 呼叫 hyft_connect 連結 hyftd
   x = obj.hyft_connect("10.10.5.77", 2816, "hykix", "", "")

   ' hyft_connect() 傳回值為連線 id, 存放於 obj.ret_val 中
   id = obj.ret_val
   if id = -1 then
      '若是連線失敗, 則錯誤訊息存放於 obj.errmsg
      response.write "連線 hyftd 發生錯誤, 錯誤訊息: " & obj.errmsg & "<br><br>"
      response.end
   else
      response.write "連線 hyftd 成功! id = " & id & "<br><br>"
   end if

   ' step 3: 初始化索引建立環境
   x = obj.hyft_initbuild(id)
   ' 傳回值為 obj.ret_val
   ret = obj.ret_val
   if ret = -1 then
      '若是初始化失敗, 則錯誤訊息存放於 obj.errmsg
      response.write "initbuild 發生錯誤, 錯誤訊息: " & obj.errmsg
      response.end
   end if

   response.write "建索引中......<br><br>"
   ' step 4: 開始呼叫 hyft_addbuild, 將要建索引之欄位資料填入
   response.write "subject = 主題館專案<br>"
   x = obj.hyft_addbuild(id, "subject", "主題館專案")
   ret = obj.ret_val
   if ret = -1 then
      response.write "填入 subject 欄位時發生錯誤, 錯誤訊息: " & obj.errmsg
      response.end
   end if

   response.write "keywords = 主題館<br>"
   x = obj.hyft_addbuild(id, "keywords", "主題館")
   ret = obj.ret_val
   if ret = -1 then
      response.write "填入 keywords 欄位時發生錯誤, 錯誤訊息: " & obj.errmsg
      response.end
   end if

   response.write "author = Hello<br>"
   x = obj.hyft_addbuild(id, "author", "hello")
   ret = obj.ret_val
   if ret = -1 then
      response.write "填入 author 欄位時發生錯誤, 錯誤訊息: " & obj.errmsg
      response.end
   end if
   
   ' step 5: 呼叫 hyft_build 並傳入 sysid 進行建索引動作
   x = obj.hyft_build(id, "AC210")
   ret = obj.ret_val
   if ret = -1 then
      response.write "建索引時發生錯誤, 錯誤訊息: " & obj.errmsg
      response.end
   end if

   ' step 6: 關閉與 hyftd 的連線
   x = obj.hyft_close(id)
   ret = obj.ret_val
   if ret = -1 then
      response.write "結束連線時發生錯誤, 錯誤訊息: " & obj.errmsg
      response.end
   end if
   response.write "<br>建立索引完成!<br>"
   set obj=Nothing	
%>

<br><br>
</center>
</body>
</html>
<script language=VBS>
alert "OK!"
</script>
