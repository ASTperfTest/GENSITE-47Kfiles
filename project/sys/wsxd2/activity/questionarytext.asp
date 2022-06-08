<%
'#####(AutoGen) 開始：此段程式碼為自動產生，註解部份請勿刪除 #####
'下列有設定部份如有需要可自行設定
'此版本為  Ver.0.2
'此段程式碼產生日期為： 2009/6/10 上午 09:59:37
'(可修改)未來是否自動更新此程式中的 Pattern (Y/N) : Y

'(可修改)此程式是否記錄 Log 檔
activeLog4U=true

'(可修改)錯誤發生時要到那個頁面,未填時直接結束，1. 可以是 / :首頁, 2. http://xxx.xxx.xxx.xxx/ooo.htm
onErrorPath="/"

'目前程式位置在
progPath="D:\hyweb\GENSITE\project\sys\wsxd2\activity\questionarytext.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("mp")
genParamsPattern=array("<", ">", "%3C", "%3E", ";", "%27", "'", "=", "+", "-", "*", "/", "%", "@", "~", "!", "#", "$", "&", "\", "(", ")", "{", "}", "[", "]")	'#### 要檢查的 pattern(程式會自動更新):GenPat ####
genParamsMessage=now() & vbTab & "Error(1):REQUEST變數含特殊字元"

'-------- (可修改)只檢查單引號，如 Request 變數未來將要放入資料庫，請一定要設定(防 SQL Injection) --------
sqlInjParamsArray=array()
sqlInjParamsPattern=array("'")	'#### 要檢查的 pattern(程式會自動更新):DB ####
sqlInjParamsMessage=now() & vbTab & "Error(2):REQUEST變數含單引號"

'-------- (可修改)只檢查 HTML標籤符號，如 Request 變數未來將要輸出，請設定 (防 Cross Site Scripting)--------
xssParamsArray=array()
xssParamsPattern=array("<", ">", "%3C", "%3E")	'#### 要檢查的 pattern(程式會自動更新):HTML ####
xssParamsMessage=now() & vbTab & "Error(3):REQUEST變數含HTML標籤符號"

'-------- (可修改)檢查數字格式 --------
chkNumericArray=array()
chkNumericMessage=now() & vbTab & "Error(4):REQUEST變數不為數字格式"

'-------- (可修改)檢察日期格式 --------
chkDateArray=array()
chkDateMessage=now() & vbTab & "Error(5):REQUEST變數不為日期格式"

'##########################################
ChkPattern genParamsArray, genParamsPattern, genParamsMessage
ChkPattern sqlInjParamsArray, sqlInjParamsPattern, sqlInjParamsMessage
ChkPattern xssParamsArray, xssParamsPattern, xssParamsMessage
ChkNumeric chkNumericArray, chkNumericMessage
ChkDate chkDateArray, chkDateMessage
'--------- 檢查 request 變數名稱 --------
Sub ChkPattern(pArray, patern, message)
	for each str in pArray	
		p=request(str)
		for each ptn in patern
			if (Instr(p, ptn) >0) then
				message = message & vbTab & progPath & vbTab & "request(" & str & ")=" & p & vbTab & request.serverVariables("REMOTE_ADDR") & vbTab & Request.QueryString
				Log4U(message) '寫入到 log
				OnErrorAction
			end if
		next
	next
End Sub

'-------- 檢查數字格式 --------
Sub ChkNumeric(pArray, message)
	for each str in pArray
		p=request(str)
		if not isNumeric(p) then
			message = message & vbTab & progPath & vbTab & "request(" & str & ")=" & p & vbTab & request.serverVariables("REMOTE_ADDR") & vbTab & Request.QueryString
			Log4U(message) '寫入到 log
			OnErrorAction
		end if
	next
End Sub

'--------檢察日期格式 --------
Sub ChkDate(pArray, message)
	for each str in pArray
		p=request(str)
		if not IsDate(p) then
			message = message & vbTab & progPath & vbTab & "request(" & str & ")=" & p & vbTab & request.serverVariables("REMOTE_ADDR") & vbTab & Request.QueryString
			Log4U(message) '寫入到 log
			OnErrorAction
		end if
	next
End Sub

'onError
Sub OnErrorAction()
	if (onErrorPath<>"") then response.redirect(onErrorPath)
	response.end
End Sub

'Log 放在網站根目錄下的 /Logs，檔名： YYYYMMDD_log4U.txt
Function Log4U(strLog)
	if (activeLog4U) then
		fldr=Server.mapPath("/") & "/Logs"
		filename=Year(Date()) & Right("0"&Month(Date()), 2) & Right("0"&Day(Date()),2)
		
		filename = filename & "_log4U.txt"
		
		Dim fso, f
		Set fso = CreateObject("Scripting.FileSystemObject")
		
		'產生新的目錄
		If (Not fso.FolderExists(fldr)) Then
			Set f = fso.CreateFolder(fldr)
		Else
			ShowAbsolutePath = fso.GetAbsolutePathName(fldr)
		End If
		
		Const ForReading = 1, ForWriting = 2, ForAppending = 8
		'開啟檔案
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set f = fso.OpenTextFile( fldr & "\" & filename , ForAppending, True, -1)
		f.Write strLog  & vbCrLf
	end if
End Function
'##### 結束：此段程式碼為自動產生，註解部份請勿刪除 #####
%><!--#Include virtual = "/inc/client.inc" -->
<%
	ActivityId = Application("ActivityId")
	sql = "SELECT * FROM m011 WHERE GETDATE() BETWEEN m011_bdate AND m011_edate AND m011_subjectid = " & ActivityId
	set rs = conn.execute(sql)
	If rs.EOF Then
		response.write "<script>window.location.href='/'</script>"
	Else
		mp = request("mp")
		if not isnumeric(mp) then mp = "1"
%>
		<td class="center">
			<div class="path">目前位置：<a href="/">首頁</a>&gt;<a href="#">問卷調查</a></div>
							
			<div class="actblock">
				<h3>農業知識入口網－97年網站使用者需求(滿意度)問卷調查－</h3>
		    <ul class="function">
		    	<li><a href="/sp.asp?xdURL=activity/qactivity-1.asp&mp=<%=mp%>">活動簡介</a></li>
		      <li><a href="/sp.asp?xdURL=activity/qactivity-2.asp&mp=<%=mp%>">活動辦法</a></li>
		      <!--li><a href="/sp.asp?xdURL=activity/qactivity-3.asp&mp=<%=mp%>">活動獎項</a></li-->
		      <li><a href="/sp.asp?xdURL=activity/qactivity-4.asp&mp=<%=mp%>">活動說明</a></li>
		      <li class="now"><a href="/sp.asp?xdURL=activity/qactivity-5.asp&mp=<%=mp%>">登入參加問卷調查</a></li>
		    </ul>
		    
		  	<table class="main" border="0" align="center" summary="排版表格" style="margin-top:5px">
		    <tr>
		      <td>
			  		<br/><p>感謝您參與本次農業知識入口網－網站滿意度調查活動－，由於本次調查項目中有包含對<strong class="red01">農業主題館喜好度的調查</strong>，請您在參與調查前，點選下方的各個主題館連結進行瀏覽。</p>
			  	</td>
				</tr>
		  	</table>
		  	<br/>
				<h4 class="title3">參與調查前，請您點選並瀏覽下列農業知識主題館，謝謝您的配合。</h4>
				<table class="style">
		  	<tr>
		    	<td>
		    		<h5><a href="/subject/dp.asp?mp=141">鴨主題館</a></h5>
		    		<a href="/subject/dp.asp?mp=141"><img src="xslgip/style1/images2/s_L.jpg" title="鴨主題館" alt="鴨主題館" /></a>
		    	</td>
		    	<td>
		    		<h5><a href="/subject/dp.asp?mp=139">文心蘭主題館</a></h5>
		    		<a href="/subject/dp.asp?mp=139"><img src="xslgip/style1/images2/s_N.jpg" title="文心蘭主題館" alt="文心蘭主題館" /></a>
		    	</td>
		    	<td>
		    		<h5><a href="/subject/dp.asp?mp=7">鳳梨主題館</a></h5>
		    		<a href="/subject/dp.asp?mp=7"><img src="xslgip/style1/images2/s_Z.jpg" title="鳳梨主題館" alt="鳳梨主題館" /></a>    
		    	</td>
		    	<td>
		    		<h5><a href="/subject/dp.asp?mp=154">銀柳主題館</a></h5>
		    		<a href="/subject/dp.asp?mp=154"><img src="xslgip/style1/images2/s_T.jpg" title="銀柳主題館" alt="銀柳主題館" /></a>    
		    	</td>
		  	</tr>
		  	<tr>
		    	<td>
		    		<h5><a href="/subject/dp.asp?mp=175">芋主題館</a></h5>
		    		<a href="/subject/dp.asp?mp=175"><img src="xslgip/style1/images2/s_H.jpg" title="芋主題館" alt="芋主題館" /></a>    
		    	</td>
		    	<td>
		    		<h5><a href="/subject/dp.asp?mp=146">花壇植物主題館</a></h5>
		    		<a href="/subject/dp.asp?mp=146"><img src="xslgip/style1/images2/s_I.jpg" title="花壇植物主題館" alt="花壇植物主題館" /></a>    
		    	</td>
		    	<td>
		    		<h5><a href="/subject/dp.asp?mp=176">茄主題館</a></h5>
		    		<a href="/subject/dp.asp?mp=176"><img src="xslgip/style1/images2/s_J.jpg" title="茄主題館" alt="茄主題館" /></a>
		    	</td>
		    	<td>
		    		<h5><a href="/subject/dp.asp?mp=164">聖誕紅主題館</a></h5>
		    		<a href="/subject/dp.asp?mp=164"><img src="xslgip/style1/images2/s_K.jpg" title="聖誕紅主題館" alt="聖誕紅主題館" /></a>
		    	</td>
		  	</tr>
		  	<tr>
		    	<td>
			   		<h5><a href="/subject/dp.asp?mp=86">茶葉主題館</a></h5>
			   		<a href="/subject/dp.asp?mp=86"><img src="xslgip/style1/images2/s_A.jpg" title="茶葉主題館" alt="茶葉主題館" /></a>
		    	</td>
		    	<td>
		    		<h5><a href="/subject/dp.asp?mp=169">保健植物主題館</a></h5>
		    		<a href="/subject/dp.asp?mp=169"><img src="xslgip/style1/images2/s_M.jpg" title="保健植物主題館" alt="保健植物主題館" /></a>
		    	</td>
		    	<td>
		    		<h5><a href="/subject/dp.asp?mp=5">蝴蝶蘭主題館</a></h5>
		    		<a href="/subject/dp.asp?mp=5"><img src="xslgip/style1/images2/s_B.jpg" title="蝴蝶蘭主題館" alt="蝴蝶蘭主題館" /></a>
		    	</td>
			   	<td>
			   		<h5><a href="/subject/dp.asp?mp=164">水生植物主題館</a></h5>
			   		<a href="/subject/dp.asp?mp=164"><img src="xslgip/style1/images2/s_O.jpg" title="水生植物主題館" alt="水生植物主題館" /></a>
			   	</td>
			 	</tr>
			 	<tr>
			    <td>
			 		  <h5><a href="/subject/dp.asp?mp=150">火鶴花主題館</a></h5>
			   		<a href="/subject/dp.asp?mp=150"><img src="xslgip/style1/images2/s_P.jpg" title="火鶴花主題館" alt="火鶴花主題館" /></a>
			   	</td>
			   	<td>
			   		<h5><a href="/subject/dp.asp?mp=153">釋迦主題館</a></h5>
			   		<a href="/subject/dp.asp?mp=153"><img src="xslgip/style1/images2/s_V.jpg" title="釋迦主題館" alt="釋迦主題館" /></a>
			   	</td>
			   	<td>
			   		<h5><a href="/subject/dp.asp?mp=142">蠶桑主題館</a></h5>
			   		<a href="/subject/dp.asp?mp=142"><img src="xslgip/style1/images2/s_W.jpg" title="蠶桑主題館" alt="蠶桑主題館" /></a>
			   	</td>
			   	<td>
			   		<h5><a href="/subject/dp.asp?mp=132">海芋主題館</a></h5>
			   		<a href="/subject/dp.asp?mp=132"><img src="xslgip/style1/images2/s_Y.jpg" title="海芋主題館" alt="海芋主題館" /></a>
			   	</td>
			 	</tr>
			 	<!--tr>
			    <td>
			   		<h5><a href="/subject/dp.asp?mp=86">茶葉主題館</a></h5>
			   		<a href="/subject/dp.asp?mp=86"><img src="xslgip/style1/images2/s_A.jpg" title="茶葉主題館" alt="茶葉主題館" /></a>
			   	</td>
			   	<td>
			   		<h5><a href="#">康乃馨主題館</a></h5>
			   		<a href="#"><img src="xslgip/style1/images2/s_C.jpg" title="康乃馨主題館" alt="康乃馨主題館" /></a>
			   	</td>
			   	<td></td>
			   	<td></td>
			 	</tr-->  
				</table>
				<br/> 
			</div>					
		</td>
<%
	End If
	set rs = nothing
%>   
