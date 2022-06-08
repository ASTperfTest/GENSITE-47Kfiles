<%
function hyftdGIP(xfunc,xiCuItem)
   on error resume next
   '----hyftd開始
   '----step 1: 使用 CreateObject 產生 Object. DLL name: hysdk, class name: hyft
   Set obj = Server.CreateObject("hysdk.hyft.1")

   '----step 2: 呼叫 hyft_connect 連結 hyftd
   x = obj.hyft_connect("10.10.5.77", 2816, "hykix", "", "")

   ' hyft_connect() 傳回值為連線 id, 存放於 obj.ret_val 中
   id = obj.ret_val
   if id = -1 then
      '若是連線失敗, 則錯誤訊息存放於 obj.errmsg
      hyftdGIP = "連線 hyftd 發生錯誤, 錯誤訊息: " & obj.errmsg
      exit function
   end if
   xsysid="AB?CuItem="&xiCuItem
   if xfunc="add" or xfunc="update" then   '----新增或更新索引    
        '----更新時先刪除索引
        if xfunc="update" then x = obj.hyft_delete(id, xsysid) 
        '----開始處理新增索引
        '----準備索引欄位陣列
   	SQLK="Select * from hykixGIPRelation"
   	SET RSK=conn.execute(SQLK)
   	if not RSK.eof then hykixArray=RSK.getrows()
        '----讀SQL資料
        SQLI="Select * from CuDTGeneric where iCuItem=" & xiCuItem
        SET RSI=conn.execute(SQLI)

	'----step 3: 初始化索引建立環境
	x = obj.hyft_initbuild(id)
	' 傳回值為 obj.ret_val
	ret = obj.ret_val
	if ret = -1 then
		'若是初始化失敗, 則錯誤訊息存放於 obj.errmsg
		hyftdGIP = "initbuild 發生錯誤, 錯誤訊息: " & obj.errmsg
		exit function
	end if
	
	'----step 4: 開始呼叫 hyft_addbuild, 將要建索引之欄位資料填入
	for i=0 to ubound(hykixArray,2)
	    if not isNull(RSI(hykixArray(1,i))) then
	    	x = obj.hyft_addbuild(id, hykixArray(0,i), RSI(hykixArray(1,i)))
		ret = obj.ret_val
		if ret = -1 then
			hyftdGIP = "填入 " & hykixArray(0,i)& " 欄位時發生錯誤, 錯誤訊息: " & obj.errmsg
			exit function
		end if	    	
	    end if
   	next
	x = obj.hyft_addbuild(id, "kixid", "AB")
	ret = obj.ret_val
	if ret = -1 then
		hyftdGIP = "填入 kixid 欄位時發生錯誤, 錯誤訊息: " & obj.errmsg
		exit function
	end if
'	x = obj.hyft_addbuild(id, "sysid", xsysid)
'	ret = obj.ret_val
'	if ret = -1 then
'		xxxxx = ""
'		hyftdGIP = "填入 sysid 欄位時發生錯誤, 錯誤訊息: " & obj.errmsg
'		exit function
'	end if		
	'----step 5: 呼叫 hyft_build 並傳入 sysid 進行建索引動作
	x = obj.hyft_build(id, xsysid)
	ret = obj.ret_val
	if ret = -1 then
		hyftdGIP = "建索引時發生錯誤, 錯誤訊息: " & obj.errmsg
		exit function
	end if
	
   else		'----刪除索引
   	x = obj.hyft_delete(id, xsysid)   
   end if

   '----step 6: 關閉與 hyftd 的連線
   x = obj.hyft_close(id)
   ret = obj.ret_val
   if ret = -1 then
      hyftdGIP = "結束連線時發生錯誤, 錯誤訊息: " & obj.errmsg
      exit function
   end if
      
   set obj = Nothing
   hyftdGIP = ""
end function
'response.write hyftdGIP("update","13279")

%>

