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
progPath="D:\hyweb\GENSITE\project\sys\wsxd2\webActivity\QuestionPage.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("id", "mp")
genParamsPattern=array("<", ">", "%3C", "%3E", ";", "%27", "'", "=", "+", "*", "/", "%", "@", "~", "!", "#", "$", "&", "\", "(", ")", "{", "}", "[", "]")	'#### 要檢查的 pattern(程式會自動更新):GenPat ####
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
	Dim activityId : activityId = ""
	Dim flag : flag = true
	Dim mp : mp = ""
	Dim memberRight : memberRight = ""
	Dim memberId : memberId = ""
	
	activityId = request.queryString("id")
	memberId = session("memID")
	mp = request.querystring("mp")
	
	if activityId = "" or isempty(activityId) then  flag = false
	if not isnumeric(activityId) then flag = false
	if isnumeric(activityId) And ( instr(activityId, ".") > 0 or instr(activityId, "-") > 0 ) then flag = false
	if not isnumeric(mp) then mp = "1"		
	
	if not flag then
		response.write "<script>"
		response.write "alert('無相關問卷調查');"
		response.write "window.location.href='/';"
		response.write "</script>"
	end if
	
	if flag then
		sql = "SELECT * FROM m011 WHERE GETDATE() BETWEEN m011_bdate AND m011_edate AND m011_online = '1' AND m011_subjectid = " & activityId
		set rs = conn.execute(sql)
		If rs.EOF Then		
			flag = false
			response.write "<script>"
			response.write "alert('無相關問卷調查');"
			response.write "window.location.href='/';"
			response.write "</script>"			
		'Else
		'	memberRight = rs("m011_memberRight")
		End If
		rs.close
		set rs = nothing		
	end if
	
	if flag then 		
			if session("memID") = "" or IsEmpty(session("memID")) then 
				flag = false
				response.write "<script>alert('此問卷調查需要登入會員');"
				response.write "window.location.href='/sp.asp?xdURL=webActivity/actPage5.asp&mp=" & mp & "&id=" & activityId & "';"
				response.write "</script>"
			end if
	end if
	
	if flag then
		memberId = session("memID")
		sql = "SELECT * FROM m014 WHERE m014_subjectid = " & activityId & " AND m014_name = '" & memberId & "'"
		Set memrs = conn.execute(sql)
		If Not memrs.Eof Then 
			flag = false			
			response.write "<script>alert('你已做過此調查！');"
			response.write "window.location.href='/sp.asp?xdURL=webActivity/actPage1.asp&mp=" & mp & "&id=" & activityId & "'"
			response.write "</script>"		
		end if
		memrs.close
		Set memrs = nothing
	End If
	
	if flag then		
		Dim questionNo : questionNo = 0
		sql	= "SELECT m011_questno FROM m011 WHERE m011_subjectid = " & activityId
		Set rsm011 = conn.execute(sql)
		If Not rsm011.EOF Then questionNo = rsm011("m011_questno")
		rsm011.close
		Set rsm011 = nothing
	
		Dim open_str
		sql = "SELECT m015_questionid, m015_answerid FROM m015 WHERE m015_subjectid = " & ActivityId
		Set rsm015 = conn.execute(sql)
		While Not rsm015.EOF
			open_str = open_str & "*" & rsm015(0) & "*" & rsm015(1) & "*,"
			rsm015.MoveNext
		Wend	
		rsm015.close
		set rsm015 = nothing
%>
		<script language="javascript">
			onload = function()
			   {
				    var answer16s=document.getElementsByName("answer16");
					if(answer16s){
						answer16s[0].checked=true;
						answer16s[0].style.display = "none";
					}
			   }

		
			function CheckForm()
			{
				var answer13s=document.getElementsByName("answer13");
				var answer13counter=0;
				for(i=0;i<answer13s.length;i++){
					if(answer13s[i].checked==true){
						answer13counter=answer13counter+1;
					}
				}
				
				var answer12s=document.getElementsByName("answer12");
				var answer12counter=0;
				for(i=0;i<answer12s.length;i++){
					if(answer12s[i].checked==true){
						answer12counter=answer12counter+1;
					}
				}
				var answer14s=document.getElementsByName("answer14");
				var answer14counter=0;
				for(i=0;i<answer14s.length;i++){
					if(answer14s[i].checked==true){
						answer14counter=answer14counter+1;
					}
				}
				if( document.getElementsByName("answer2")[9].checked==true && document.getElementsByName("open_content2_10")[0].value.replace(/(^\s*)|(\s*$)/g,"") == "" ) {
					alert("請填入第2題的其它意見!");
					document.getElementsByName("open_content2_10")[0].focus();
				}else if( document.getElementsByName("answer4")[4].checked==true && document.getElementsByName("open_content4_5")[0].value.replace(/(^\s*)|(\s*$)/g,"") == "" ) {
					alert("請填入第4題的其它意見!");
					document.getElementsByName("open_content4_5")[0].focus();
				}
				else if( document.getElementsByName("answer8")[4].checked==true && document.getElementsByName("open_content8_5")[0].value.replace(/(^\s*)|(\s*$)/g,"") == "" ) {
					alert("請填入第8題的宣傳廣告(其他網站或新聞)!的內容");
					document.getElementsByName("open_content8_5")[0].focus();
				}
				else if( document.getElementsByName("answer8")[5].checked==true && document.getElementsByName("open_content8_6")[0].value.replace(/(^\s*)|(\s*$)/g,"") == "" ) {
					alert("請填入第8題的其它意見!");
					document.getElementsByName("open_content8_6")[0].focus();
				}
				else if( document.getElementsByName("answer11")[3].checked==true && document.getElementsByName("open_content11_4")[0].value.replace(/(^\s*)|(\s*$)/g,"") == "" ) {
					alert("請填入第7題的其它意見!");
					document.getElementsByName("open_content11_4")[0].focus();
				}
				else if(answer12counter!=5){
						alert("12題請選擇五個主題館!");
						document.getElementsByName("answer12")[0].focus();	
				}else if(answer13counter!=5){
						alert("13題請選擇五個主題館!");
						document.getElementsByName("answer13")[0].focus();	
				}else if(answer14counter!=5){
						alert("14題請選擇五個主題館!");
						document.getElementsByName("answer14")[0].focus();	
				}else if( document.getElementsByName("answer16")[0].checked==true && document.getElementsByName("open_content16_1")[0].value.replace(/(^\s*)|(\s*$)/g,"") == "" ) {
					alert("請填入第16題的意見!");
					document.getElementsByName("open_content16_1")[0].focus();
				}
				else{
					document.getElementById("iForm").submit();
				}
			}
		</script>
		<div class="path">目前位置：<a href="#">首頁</a>&gt;<a href="#">問卷調查</a></div>
			
		<div class="actblock">
			<h3>農業知識入口網－網站使用者需求暨滿意度調查－</h3>
			<ul class="function">
				<!--li><a href="/sp.asp?xdURL=webActivity/actPage1.asp&mp=<%=mp%>&id=<%=activityId%>">活動簡介</a></li-->
			  <!--li><a href="/sp.asp?xdURL=webActivity/actPage2.asp&mp=<%=mp%>&id=<%=activityId%>">活動辦法</a></li-->
			  <!--li><a href="/sp.asp?xdURL=webActivity/actPage3.asp&mp=<%=mp%>&id=<%=activityId%>">活動獎項</a></li-->
			  <!--li><a href="/sp.asp?xdURL=webActivity/actPage4.asp&mp=<%=mp%>&id=<%=activityId%>">活動說明</a></li-->
			  <!--li><a href="/sp.asp?xdURL=webActivity/actPage5.asp&mp=<%=mp%>&id=<%=activityId%>">登入參加問卷調查</a></li-->
			</ul>
			<table class="main" border="0" align="center" summary="排版表格" style="margin-top:5px">
			<tr>
				<td>
					<p>感謝您參與本次調查活動，本次調查項目<strong class="red01">共14題</strong>，請依序圈選最符合的項目，填寫完成請點選最下方之『完成送出』按鈕，即可完成問卷並具有參加抽獎之資格</p>
					<p>請圈選您認為符合問題的答案，<strong class="red01">若無特別說明皆為單選</strong>。</p>	  
				</td>
			</tr>
			<tr>
				<td valign="top" class="content">
				<form method="post" action="/sp.asp?xdURL=webActivity/QuestionPageAct.asp&mp=<%=mp%>&id=<%=activityId%>" name="iForm" id="iForm">
				<div class="question">
					<dl>
						主題：農業知識入口網－網站使用者需求暨滿意度調查－ 
						<%
							sql = "SELECT * FROM m012 WHERE m012_subjectid = " & ActivityId & " ORDER BY m012_questionid"
							Set rsm012 = conn.execute(sql)
							While not rsm012.EOF 
								AnswerType = trim(rsm012("m012_Type"))
								If AnswerType = "1" Or AnswerType = "" Then
									sel = ""
									seltype = "radio"
								Else
									sel = "<font color='red'>（本題為複選題）</font>"
									seltype = "checkbox"
								End If
								Dim titleText 
								titleText= rsm012("m012_title")
								if InStr(titleText,"（") > 0 AND InStr(titleText,"）") > 0 then
									titleText = Left(titleText, InStr(titleText,"（")-1) & _
									"<font color='blue'>" & Mid(titleText,InStr(titleText,"（"),InStr(titleText,"）")-InStr(titleText,"（")+1) & _
									"</font>" & Right(titleText,LEN(titleText)-InStr(titleText,"）"))
								end if 
						%>
								<dt><%=rsm012("m012_questionid")%>. 	<%=titleText%><%=sel%></dt>
						<%
								sql = " SELECT COUNT(*) AS CNT FROM m013 WHERE m013_subjectid = " & ActivityId & _
											" AND m013_questionid = " & rsm012("m012_questionid")
								
								Set rsCount=conn.execute(sql)
								count=0
								While Not rsCount.Eof
									count=rsCount("CNT")
									rsCount.MoveNext
								Wend
								Set rsCount = nothing
								
								sql = " SELECT * FROM m013 WHERE m013_subjectid = " & ActivityId & _
											" AND m013_questionid = " & rsm012("m012_questionid") & _
											" ORDER BY m013_answerid "
								Set rsm013 = conn.execute(sql)
								u=0
								While Not rsm013.Eof
								If count>10 Then
									If u=0 Then
									 response.write "<table>"
									End If
									
									if u mod 3 = 0 then
										response.write "<tr>"
									end if
										
										response.write "<td>"
										response.write "<dd>"
										response.write "<input type='" & seltype & "' name='answer" & rsm012("m012_questionid") & "' value='" & rsm013("m013_answerid") & "'"
										if trim(rsm013("m013_default")) <> "" then
											response.write " checked >" 
										else
											response.write " >"
										end if
										response.write trim(rsm013("m013_title"))
										response.write "</dd>"
										response.write "</td>"
										
									if u mod 3= 2 or u=count-1 then
											response.write "</tr>"
									end if
									
									If u=count-1 Then
									response.write "</table>"
									end If
								else
								
 						%>
									<dd>
										<input type="<%=seltype%>" name="answer<%=rsm012("m012_questionid")%>" value="<%=rsm013("m013_answerid")%>"<% if trim(rsm013("m013_default")) <> "" then %> checked<% end if %>>
		  							<%=trim(rsm013("m013_title")) %>
									</dd>
						<%
								end if	
									If InStr(open_str, "*" & rsm012("m012_questionid") & "*" & rsm013("m013_answerid") & "*") > 0 Then
										response.write "<dd>"
										if rsm012("m012_questionid") = "15" and rsm013("m013_answerid") = "1"  then
										response.write "			<input type=""text"" name=""open_content" & rsm012("m012_questionid") & "_" & rsm013("m013_answerid") & """ size=""60"" maxlength=""512"">"
										else
										response.write "			<input type=""text"" name=""open_content" & rsm012("m012_questionid") & "_" & rsm013("m013_answerid") & """ size=""30"" maxlength=""512"">"
										end if
										response.write "</dd>"
									End If																						
									rsm013.MoveNext
									u=u+1
								Wend
								Set rsm013 = nothing
											
								If rsm012("m012_textarea") = "1" Then
									'response.write "<br /><br />請說明您的原因或您的其它意見(必填)"
									response.write "<dd><textarea name=""textarea" & rsm012("m012_questionid") & """ cols=""55"" rows=""5""></textarea></dd>"
								End If
									
								rsm012.MoveNext
							Wend
							Set rsm012 = nothing
						%>
					</dl>
				</div>
				<center><input type="button"  Tabindex="34" value="完成送出" onclick="CheckForm()"></center>	
				</form>	  
				</td>
			</tr>
			</table>
		</div>	
<%		
	end if
%>