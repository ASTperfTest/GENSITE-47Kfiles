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
progPath="D:\hyweb\GENSITE\project\sys\wsxd2\activity\questionary.asp"

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
	Dim flag
	flag = true
	
	ActivityId = Application("ActivityId")
	sql = "SELECT * FROM m011 WHERE GETDATE() BETWEEN m011_bdate AND m011_edate AND m011_subjectid = " & ActivityId
	set rs = conn.execute(sql)
	If rs.EOF Then
		flag = false
		response.write "<script>window.location.href='/'</script>"
	End If
	
	Dim memberId
	If flag Then
		mp = request("mp")
		if not isnumeric(mp) then mp = "1"

		If IsNull(Session("memID")) Or Session("memID") = "" Then
			flag = false
			response.write "<script>window.location.href='/sp.asp?xdURL=activity/qactivity-5.asp&mp=" & mp & "'</script>"
		Else
			memberId = Session("memID")
		End If						
	End If
	
	'---check the user have filled this question---
	If flag Then
		sql = "SELECT * FROM m014 WHERE m014_subjectid = " & ActivityId & " AND m014_name = '" & memberId & "'"
		Set memrs = conn.execute(sql)
		If Not memrs.Eof Then
			flag = false
			response.write "<script>alert('你已做過此調查！');"
			response.write "window.location.href='/sp.asp?xdURL=activity/qactivity-1.asp&mp=" & mp & "'"
			response.write "</script>"			
		End If
	End If
	Set memrs = nothing
	
	If flag Then
											
		Dim questionNo
		sql	= "SELECT m011_questno FROM m011 WHERE m011_subjectid = " & ActivityId
		Set rsm011 = conn.execute(sql)
		If Not rsm011.EOF Then questionNo = rsm011("m011_questno")
		Set rsmo11 = nothing
	
		Dim open_str
		sql = "SELECT m015_questionid, m015_answerid FROM m015 WHERE m015_subjectid = " & ActivityId
		Set rsm015 = conn.execute(sql)
		While Not rsm015.EOF
			open_str = open_str & "*" & rsm015(0) & "*" & rsm015(1) & "*,"
			rsm015.MoveNext
		Wend	
%>
<script language="javascript">
	function send()
	{
		if( document.getElementById("textarea16").value == "" ) 
		{
			alert("請填入第16題的說明原因或其它意見!");
			return false;
		}
		else if( document.getElementById("textarea17").value == "" ) 
		{
			alert("請填入第17題的說明原因或其它意見!");
			return false;
		}
		else
		{
			return true;
		}
	}
</script>
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
		  				<p>感謝您參與本次調查活動，活動時間為2008/6/2中午12時至2008/6/13中午12時。</p>
              <p>本次所有調查項目<strong class="red01">共<%=questionNo%>題</strong>，填寫完成請點選最下方之『完成送出』按鈕，即可完成問卷並具有參加抽獎之資格。</p>
		  				<p>在參加調查前，請您<a href="/sp.asp?xdURL=activity/questionarytext.asp&mp=<%=mp%>" target="_blank">先點選這裡瀏覽本次調查之各個農業知識主題館</a>。</p>
		  				<p>以下各題請圈選您認為符合問題的答案，<strong class="red01">若無特別說明皆為單選</strong>。</p>
		  			</td>
					</tr>
					<tr>
		  			<td valign="top" class="content">
		  			<form method="post" action="/sp.asp?xdURL=activity/questionary_act.asp" name="post" onsubmit="return send()" >      
							<div class="question">
							<dl>
								主題：農業知識入口網－97年網站使用者需求(滿意度)問卷調查－ 
								<%
									sql = "SELECT * FROM m012 WHERE m012_subjectid = " & ActivityId & " ORDER BY m012_questionid"
									Set rsm012 = conn.execute(sql)
									While not rsm012.EOF 
										AnswerType = trim(rsm012("m012_Type"))
										If AnswerType = "1" Or AnswerType = "" Then
											sel = ""
											seltype = "radio"
										Else
											sel = "（本題為複選題）"
											seltype = "checkbox"
										End If
								%>
										<dt><%=rsm012("m012_questionid")%>. 	<%=rsm012("m012_title")%><%=sel%></dt>
										<%
											sql = " SELECT * FROM m013 WHERE m013_subjectid = " & ActivityId & _
														" AND m013_questionid = " & rsm012("m012_questionid") & _
														" ORDER BY m013_answerid "
											Set rsm013 = conn.execute(sql)
											While Not rsm013.Eof 
										%>
												<dd>
													<input type="<%=seltype%>" name="answer<%=rsm012("m012_questionid")%>" value="<%=rsm013("m013_answerid")%>"<% if trim(rsm013("m013_default")) <> "" then %> checked<% end if %>>
				  								<%=trim(rsm013("m013_title")) %>
												</dd>
												<%
												If InStr(open_str, "*" & rsm012("m012_questionid") & "*" & rsm013("m013_answerid") & "*") > 0 Then
													response.write "			<input type=""text"" name=""open_content" & rsm012("m012_questionid") & "_" & rsm013("m013_answerid") & """ size=""30"" maxlength=""512"">"
												End If																						
												
												rsm013.MoveNext
											Wend
											Set rsm013 = nothing
											
											If rsm012("m012_textarea") = "1" Then
													response.write "<br /><br />請說明您的原因或您的其它意見(必填)"
													response.write "<dd><textarea name=""textarea" & rsm012("m012_questionid") & """ cols=""55"" rows=""5""></textarea></dd>"
											End If
												
										rsm012.MoveNext
									Wend
									Set rsm012 = nothing
								%>
								</dl>
							</div>
							<center><input type="submit" name="Submit" id="Submit" Tabindex="34" value="完成送出"></center>	
						</form>
		  		</td>
	    	</tr>
	  		</table>
			</div>
			</td>
<%
	End If
	set rs = nothing
%>    
