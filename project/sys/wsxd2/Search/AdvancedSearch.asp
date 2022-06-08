<%
'#####(AutoGen) 開始：此段程式碼為自動產生，註解部份請勿刪除 #####
'下列有設定部份如有需要可自行設定
'此版本為  Ver.0.2
'此段程式碼產生日期為： 2009/6/10 上午 09:59:38
'(可修改)未來是否自動更新此程式中的 Pattern (Y/N) : Y

'(可修改)此程式是否記錄 Log 檔
activeLog4U=true

'(可修改)錯誤發生時要到那個頁面,未填時直接結束，1. 可以是 / :首頁, 2. http://xxx.xxx.xxx.xxx/ooo.htm
onErrorPath="/"

'目前程式位置在
progPath="D:\hyweb\GENSITE\project\sys\wsxd2\Search\AdvancedSearch.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("Keyword", "FromSiteUnit", "FromKnowledgeTank", "FromKnowledgeHome", "FromTopic","FromPedia","FromVideo")
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
%><%
	Dim Keyword, FromSiteUnit, FromKnowledgeTank, FromKnowledgeHome, FromTopic,FromPedia,FromVideo
	
	'Keyword = Request("Keyword")
	Keyword = ""
	
	FromSiteUnit = Request("FromSiteUnit")
	If FromSiteUnit = "" Then
		FromSiteUnit = "0"
	else
		FromSiteUnit = "1"
	End If	
	
	FromKnowledgeTank = Request("FromKnowledgeTank")
	If FromKnowledgeTank = "" Then 
		FromKnowledgeTank = "0"
	else
		FromKnowledgeTank = "1"
	End If
	
	FromKnowledgeHome = Request("FromKnowledgeHome")
	If FromKnowledgeHome = "" Then
		FromKnowledgeHome = "0"
	else
		FromKnowledgeHome = "1"
	End If
	
	FromTopic= Request("FromTopic")
	If FromTopic = "" Then
		FromTopic = "0"
	else
		FromTopic = "1"
	End If
	'加上【農業小百科】與【影音專區】 by Alyon 
	FromPedia= Request("FromPedia")
	If FromPedia = "" Then
		FromPedia = "0"
	else
		FromPedia = "1"
	End If
	
	FromVideo= Request("FromVideo")
	If FromVideo = "" Then
		FromVideo = "0"
	else
		FromVideo = "1"
	End If
	
	FromTechCD= Request("FromTechCD")
	If FromTechCD = "" Then
		FromTechCD = "0"
	else
		FromTechCD = "1"
	End If
	
%>
<script language="javascript">
	function checkSearchForm(value)
	{				
		if( value == 0 ) {
			if( !document.SearchForm.FromSiteUnit.checked && 
				!document.SearchForm.FromKnowledgeTank.checked &&
				!document.SearchForm.FromKnowledgeHome.checked && 
				!document.SearchForm.FromTopic.checked &&
				!document.SearchForm.FromPedia.checked  &&
				!document.SearchForm.FromVideo.checked ) {
				alert('請至少選擇一項資料來源');
				event.returnValue = false;
			}			
			else if( document.SearchForm.Keyword.value == "" ) {
				alert('請輸入查詢值');
				event.returnValue = false;
			}
			else if( !document.SearchForm.Subject.checked && !document.SearchForm.Description.checked 
				    && !document.SearchForm.Author.checked && !document.SearchForm.Keywords.checked ) {
				alert('請至少選擇一項查詢欄位');
				event.returnValue = false;
			}
			else {
				document.SearchForm.action = "kp.asp?xdURL=Search/SearchResultList.asp&amp;mp=1";
				document.SearchForm.submit();
			}
		}
		else {
			document.SearchForm.Order.selectedIndex = 0;
			document.SearchForm.PageSize.selectedIndex = 0;
			document.SearchForm.Sort.selectedIndex = 0;
			document.SearchForm.StartDate.value = "";
			document.SearchForm.EndDate.value = "";
			document.SearchForm.Keyword.value = "";
			document.SearchForm.FromSiteUnit.checked = true;
			document.SearchForm.FromKnowledgeTank.checked = true;
			document.SearchForm.FromKnowledgeHome.checked = true;
			document.SearchForm.FromTopic.checked = true;
			document.SearchForm.FromPedia.checked = true;
			document.SearchForm.FromVideo.checked = true;
			event.returnValue = false;
		}
	}
</script>

<script language="vbs">
	Dim CanTarget
	Dim followCanTarget

	Sub popCalendar(dateName,followName)        
 		CanTarget = dateName
 		followCanTarget = followName
		xdate = document.all(CanTarget).value
		if not isDate(xdate) then	xdate = date()
		document.all.calendar1.setDate xdate
	
 		If document.all.calendar1.style.visibility="" Then           
   		document.all.calendar1.style.visibility="hidden"        
 		Else        
    	ex = window.event.clientX + 100
			ey = document.body.scrolltop + window.event.clientY + 10
      'ey = window.event.srcElement.parentElement.offsetTop
      if ex > 520 then ex = 520
       document.all.calendar1.style.pixelleft = ex - 80
       document.all.calendar1.style.pixeltop = ey
       document.all.calendar1.style.visibility = ""
 		End If              
	End Sub 

	Sub calendar1_onscriptletevent(n,o) 
    document.all("CalendarTarget").value = n         
    document.all.calendar1.style.visibility = "hidden"    
    if n <> "Cancle" then
      document.all(CanTarget).value = document.all.CalendarTarget.value        
      if followCanTarget <> "" then
        document.all(followCanTarget).value = document.all.CalendarTarget.value        	
      end if
		end if
	End Sub 	
</script>

<form name="SearchForm" method="post" taget="_self" >
	<input name="AdvanceSearch" type="hidden" value="1">
	<table class="layout">
	<tr>
		<td class="center">
			<div class="aligncenter">
				<table class="adlayouttable">
				<tr>
					<th scope="col">資料來源設定</th>
					<td>
						<div>
							<label><input name="FromSiteUnit" type="checkbox" 		value="<%=FromSiteUnit%>" 		<% If FromSiteUnit = "1" Then %> checked <% End If %> 		/>站內單元</label>
							<label><input name="FromKnowledgeTank" type="checkbox"	value="<%=FromKnowledgeTank%>"	<% If FromKnowledgeTank = "1" Then %> checked <% End If %> 	/>知識庫</label>
							<label><input name="FromKnowledgeHome" type="checkbox"	value="<%=FromKnowledgeHome%>" 	<% If FromKnowledgeHome = "1" Then %> checked <% End If %> 	/>知識家</label>
							<label><input name="FromTopic" type="checkbox" 			value="<%=FromTopic%>"			<% If FromTopic = "1" Then %> checked <% End If %> 			/>主題館</label>
							<label><input name="FromPedia" type="checkbox"		value="<%=FromPedia%>"		<% If FromPedia = "1" Then %> checked <% End If %>		/>農業小百科</label>
							<label><input name="FromVideo" type="checkbox" 			value="<%=FromVideo%>"				<% If FromVideo = "1" Then %> checked <% End If %> 			/>影音專區</label>
							<label><input name="FromTechCD" type="checkbox" 		value="<%=FromTechCD%>"		    <% If FromTechCD = "1" Then %> checked <% End If %> 		/>技術光碟 </label>
						</div>
					</td>
				</tr>
				<tr>
					<th scope="col">查詢值設定</th>
					<td>
						<div>
							<label>關鍵字：<input name="Keyword" type="text" size="30" value="<%=Keyword%>" /></label>
							<input type="checkbox" name="Phonetic" value="1">同音                      
							<input type="checkbox" name="Authorize" value="1">同義
							<!--input type="checkbox" name="TypeA3" value="1">容錯-->
						</div>
						<hr/>
						<div>
							查詢欄位：
							<label><input type="checkbox" name="Subject" value="1" checked />標題</label>
							<label><input type="checkbox" name="Description" value="1" checked />內容</label>
							<!--label><input type="checkbox" name="Attachment" value="1" checked />附件</label-->
							<label><input type="checkbox" name="Author" value="1" checked />作者</label>
							<label><input type="checkbox" name="Keywords" value="1" checked />關鍵字</label>
							<!--label><input type="checkbox" name="" value="1" checked />討論</label-->
						</div>
						<hr/>
						<div> 資料發布日期：<label>自
							<input name="StartDate" type="text" value="" size="15" readonly />
							<img src="/xslgip/style1/images3/calendar.gif" alt="設定開始日期" width="16" height="16" border="0" align="absmiddle" onclick="VBS: popCalendar 'StartDate',''">
							起；至
							<input name="EndDate" type="text" value="" size="15" readonly />
							<img src="/xslgip/style1/images3/calendar.gif" alt="設定結束日期" width="16" height="16" border="0" align="absmiddle" onclick="VBS: popCalendar 'EndDate',''">
							<input type="hidden" name="CalendarTarget">
							<OBJECT id="calendar1" style="LEFT: 230px; VISIBILITY: hidden; TOP: 0px; POSITION: ABSOLUTE;" type="text/x-scriptlet" height="160" width="245" data="../inc/calendar.asp"></OBJECT>
							止
							</label>
						</div>
					</td>
				</tr>
				<tr>
					<th scope="col">結果排序設定</th>
					<td>
						<div>排序依：
							<select name="Order">
								<option value="0">關聯度 </option>
								<option value="1">標題 </option>
								<option value="2">發布日期 </option>
								<option value="3">資料來源 </option>								
								<option value="3">點閱率 </option>								
							</select>
						</div>
						<hr/>
						<div>排序方式：
							<select name="Sort">
								<option value="0">遞減 </option>
								<option value="1">遞增 </option>
							</select>
						</div>
						<hr/>
						<div>呈現筆數：
							<select name="PageSize" >
								<option value="10">10</option>
								<option value="30">30</option>
								<option value="50">50</option>
							</select>
						</div>
					</td>
				</tr>
				</table>
				<input name="search" type="image" src="xslgip/style1/images3/searchBtn_1.gif" alt="search" class="btn" onClick="javascript:checkSearchForm(0)"/>
		  	<input name="search" type="image" src="xslgip/style1/images3/searchBtn_3.gif" alt="search" class="btn" onClick="javascript:checkSearchForm(1)"/>
			</div>
		</td>
	</tr>
	</table>
</form>
