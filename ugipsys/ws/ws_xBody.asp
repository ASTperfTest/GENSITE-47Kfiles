<%
response.expires = -1
Function wordFreq(xBody,xWord)
	xTimes = 0
	xStr = xBody
	xPos = instr(xStr,xWord)
	while xPos <> 0
		xTimes = xTimes + 1
		xStr = mid(xStr, xPos+Len(xWord))
		xPos = instr(xStr,xWord)
	wend
	wordFreq = xTimes
end Function  

Set conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'conn.Open session("ODBCDSN")
Set conn = Server.CreateObject("HyWebDB3.dbExecute")
conn.ConnectionString = session("ODBCDSN")
'----------HyWeb GIP DB CONNECTION PATCH----------


on error resume next

	'Get XML string
	set oXMLDoc = server.createObject("Microsoft.XMLDOM")
	oXmlDoc.async = false
	xv = oXmlDoc.load( request)
	
	xkeyword = oXmlDoc.selectSingleNode("//xBody").text
	xReturnValue = ""

		'----叫用hysearch, 若xStr不存在串入字串中
		' step 1: 使用 CreateObject 產生 Object. DLL name: hysdk, class name: hyft  
		Set obj = Server.CreateObject("hysdk.hyft.1")

		' step 2: 呼叫 hyft_connect 連結 hyftd	hyft_connect() 傳回值為連線 id, 存放於 obj.ret_val 中
				idxStr = session("mySiteID")
				x = obj.hyft_connect("localhost", 2816, idxStr , "", "")
		id = obj.ret_val
   		if id = -1 then
      			'若是連線失敗, 則錯誤訊息存放於 obj.errmsg
      			response.write "連線 hyftd 發生錯誤, 錯誤訊息: " & obj.errmsg & "<br><br>"
      			response.end
   		end if

		query=xkeyword

		' step 3: 啟用hyft_textkeys斷詞功能,如果成功,會得到一個key_id,
		x=obj.hyft_textkeys(id, query,5)
		key_id = obj.ret_val
      		'若是連線失敗, 則錯誤訊息存放於 obj.errmsg
		if key_id = -1 then
      			response.write "連線 hyftd 發生錯誤, 錯誤訊息: " & obj.errmsg & "<br><br>"
      			response.end
    		end if


		' step 4: 算出query中斷詞的數目
		x=obj.hyft_num_keys(id,key_id)
		num =obj.ret_val
	
'		if num > 5 then num = 5
		
		for i = 0 to num-1
			x=obj.hyft_keylen(id,key_id,i)
			key_len=obj.ret_val
			
			x=obj.hyft_onekey(id,key_id,i,key_len)
			key=obj.key
			xTimes = wordFreq(xkeyword,key)
			if xTimes <>1 then
				xReturnValue=xReturnValue+key+"*"+cStr(xTimes)+","
			else
				xReturnValue=xReturnValue+key+","
			end if
		next 	
		x=obj.hyft_free_key(id,key_id)

		set obj=Nothing			
		'----hysearch完成
			
	if xReturnValue<>"" then xReturnValue=Left(xReturnValue,Len(xReturnValue)-1)
'response.write 	xReturnValue
'response.end
	oXmlDoc.selectSingleNode("//xBody").text = xReturnValue
'	response.write "<?xml version=""1.0"" encoding=""utf-8"" ?>" & vbCRLF
'	response.write "<xBodyList><xBody><![CDATA["+xReturnValue+"]]></xBody></xBodyList>"
'	response.write oXMlDoc.xml
	response.ContentType = "text/xml"
	oXMlDoc.save(Response)	
%>
