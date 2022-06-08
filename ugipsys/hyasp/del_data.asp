<html>
<title>Hysearch 全文檢索引擎 ASP 測試範例</title>
<body>
<center>
<br>
<h3>HySearch 全文檢索系統  ASP 版本測試網頁</h3><br>
刪除某筆資料索引<br><br>
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


    	
	'----刪除
    	x = obj.hyft_delete(id, "AA?REPORT_ID=1001")

   	
  set obj=Nothing	
%>
<br><br>
<a href="javascript:history.go(-1)">回上頁</a>
</center>
</body>
</html>
