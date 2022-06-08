<%@ CodePage = 65001 %>
<%
	Response.Expires = 0
	'HTProgCap = "功能管理"
	'HTProgFunc = "問卷調查"
	HTProgCode = "GW1_vote01"
	'HTProgPrefix = "mSession"
	response.expires = 0
%>
<!-- #include virtual = "/inc/client.inc" -->
<!-- #include virtual = "/inc/dbutil.inc" -->
<%
	Set RSreg = Server.CreateObject("ADODB.RecordSet")
%>
<%	
	
	subjectid = request("subjectid")
	
	if subjectid = "" then
		response.write "<script language='JavaScript'>history.go(-1);</script>"
		response.end
	end if
	
	
	sql = "" & _
		" select m011_subject, m011_bdate, m011_edate, m011_haveprize from m011 where " & _
		" m011_subjectid = " & subjectid
	set rs = conn.execute(sql)
	if rs.EOF then
		response.write "<script language='JavaScript'>history.go(-1);</script>"
		response.end
	end if
	
	
	set ts = conn.execute("select count(*) from m014 where m014_subjectid = " & subjectid)
	if ts(0) = 0 then
		response.write "<script language='JavaScript'>alert('目前尚無任何答題資料！');history.go(-1);</script>"
		response.end
	end if
	base_answer_no = ts(0)
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="style.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#FFFFFF">
<p>問卷調查</p>
<table border="1" cellspacing="0" cellpadding="2" width="100%" bordercolorlight="#999900" bordercolordark="#FFFFFF">
  <tr> 
    <td colspan="3">統計項目： 
      <input type="button" value="答案累計統計" onClick="javascript:location.replace('02_result.asp?functype=<%= functype %>&subjectid=<%= subjectid %>');">
<%
	'if rs("m011_haveprize") = "1" then
%>
      <input type="button" value="調查對象分析" onClick="javascript:location.replace('02_group.asp?functype=<%= functype %>&subjectid=<%= subjectid %>');">
      <input type="button" value="調查對象明細" onClick="javascript:location.replace('02_namelist.asp?functype=<%= functype %>&subjectid=<%= subjectid %>');">
<%
	'end if
%>
      <input type="button" value="意見回覆" onClick="javascript:location.replace('02_feedback_note.asp?functype=<%= functype %>&subjectid=<%= subjectid %>');">
    </td>
  </tr>
  
  <tr>
    <td colspan="3">
      <input type="button" value="資料匯出" onClick="javascript:location.replace('02_group_exl.asp?subjectid=<%= subjectid %>');">
      <font color="#993300">[請按資料匯出之後,會出現另存新檔,副檔名請存成xls檔即可]</font>
    </td>
  </tr>  
  <tr><td colspan="3">起訖時間：<%= DateValue(rs("m011_bdate")) %> ~ <%= DateValue(rs("m011_edate")) %></td></tr>
  <tr><td colspan="3">調查主題：<%= trim(rs("m011_subject")) %></tr>  

<%
	sql = "select count(*), m014_sex from m014 where m014_subjectid = " & subjectid & " group by m014_sex"
	set ts = conn.execute(sql)
	while not ts.EOF
		if ts("m014_sex") = "M" then
			boy_no = ts(0)
		else
			girl_no = ts(0)
		end if
		
		ts.MoveNext
	wend
	
	boy_ratio = FormatPercent(boy_no / base_answer_no, 1, False)
	girl_ratio = FormatPercent(girl_no / base_answer_no, 1, False)
%>
  <tr> 
    <td nowrap>性別</td>
    <td width="90%"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="3">
        <tr>
          <td>男</td>
          <td width="75%"><img src="images/bar_chart.gif" width="<%= boy_ratio %>" height="12"></td>
          <td width="12%"><%= boy_ratio %>(<%= boy_no %>人)</td>
        </tr>
        <tr>
          <td>女</td>
          <td width="75%"><img src="images/bar_chart.gif" width="<%= girl_ratio %>" height="12"></td>
          <td width="12%"><%= girl_ratio %>(<%= girl_no %>人)</td>
        </tr>
      </table>      
    </td>
  </tr>

  <tr> 
    <td nowrap>年齡</td>
    <td width="90%"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="3">
<%
	dim	age_str(8)
	age_str(1) = "0到11歲"
	age_str(2) = "12到17歲"
	age_str(3) = "18到23歲"
	age_str(4) = "24到29歲"
	age_str(5) = "30到39歲"
	age_str(6) = "40到49歲"
	age_str(7) = "50到59歲"
	age_str(8) = "60歲以上"

	sql = "select count(*), m014_age from m014 where m014_subjectid = " & subjectid & " group by m014_age"
	set ts = conn.execute(sql)
	for i = 1 to 8
		if ts.EOF then
			age_no = 0
		else
			if CInt(i) = CInt(ts("m014_age")) then
				age_no = ts(0)
				ts.MoveNext
			else
				age_no = 0
			end if
		end if
		
		if age_no > 0 then
			age_ratio = FormatPercent(age_no / base_answer_no, 1, False)
%>
        <tr> 
          <td><%= age_str(i) %></td>
          <td width="75%"><img src="images/bar_chart.gif" width="<%= age_ratio %>" height="12"></td>
          <td width="12%"><%= age_ratio %>(<%= age_no %>人)</td>
        </tr>
<%
		end if
	next
%>
      </table>      
    </td>
  </tr>
  
  <tr> 
    <td nowrap>居住縣市</td>
    <td width="90%">
      <table width="100%" border="0" cellspacing="0" cellpadding="3">
<%     
	Dim	city(25)
	city(1) = "臺北縣"
	city(2) = "宜蘭縣"
	city(3) = "桃園縣"
	city(4) = "新竹縣"
	city(5) = "苗栗縣"
	city(6) = "臺中縣"
	city(7) = "彰化縣"
	city(8) = "南投縣"
	city(9) = "雲林縣"
	city(10) = "嘉義縣"
	city(11) = "臺南縣"
	city(12) = "高雄縣"
	city(13) = "屏東縣"
	city(14) = "臺東縣"
	city(15) = "花蓮縣"
	city(16) = "澎湖縣"
	city(17) = "基隆市"
	city(18) = "新竹市"
	city(19) = "臺中市"
	city(20) = "嘉義市"
	city(21) = "臺南市"
	city(22) = "臺北市"
	city(23) = "高雄市"
	city(24) = "金門縣"
	city(25) = "連江縣"
	
	sql = "select count(*), m014_addrarea from m014 where m014_subjectid = " & subjectid & " group by m014_addrarea"
	set ts = conn.execute(sql)
	for i = 1 to 25
		if ts.EOF then
			city_no = 0
		else
			if CInt(i) = CInt(ts("m014_addrarea")) then
				city_no = ts(0)
				ts.MoveNext
			else
				city_no = 0
			end if
		end if
		
		if city_no > 0 then
			city_ratio = FormatPercent(city_no / base_answer_no, 1, False)
%>
        <tr> 
          <td><%= city(i) %></td>
          <td width="75%"><img src="images/bar_chart.gif" width="<%= city_ratio %>" height="12"></td>
          <td width="12%"><%= city_ratio %>(<%= city_no %>人)</td>
        </tr>
<%
		end if
	next
%>
      </table>
    </td>
  </tr>
  
  <tr> 
    <td nowrap>家庭成員人數</td>
    <td width="90%"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="3">
<%     
	Dim	family_str(6)
	family_str(1) = "1人"
	family_str(2) = "2人"
	family_str(3) = "3人"
	family_str(4) = "4人"
	family_str(5) = "5至8人"
	family_str(6) = "8人以上"
	
	sql = "select count(*), m014_familymember from m014 where m014_subjectid = " & subjectid & " group by m014_familymember"
	set ts = conn.execute(sql)
	for i = 1 to 6
		if ts.EOF then
			familymember_no = 0
		else
			if CInt(i) = CInt(ts("m014_familymember")) then
				familymember_no = ts(0)
				ts.MoveNext
			else
				familymember_no = 0
			end if
		end if
		
		if familymember_no > 0 then
			familymember_ratio = FormatPercent(familymember_no / base_answer_no, 1, False)
%>
        <tr> 
          <td><%= family_str(i) %></td>
          <td width="75%"><img src="images/bar_chart.gif" width="<%= familymember_ratio %>" height="12"></td>
          <td width="12%"><%= familymember_ratio %>(<%= familymember_no %>人)</td>
        </tr>
<%
		end if
	next
%>
      </table>
    </td>
  </tr>

  <tr> 
    <td nowrap>收入</td>
    <td width="90%"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="3">
<%
	dim	money_str(6)
	money_str(1) = "5000元以下"
	money_str(2) = "5000到1萬元"
	money_str(3) = "1萬到3萬元"
	money_str(4) = "3萬到5萬元"
	money_str(5) = "5萬到10萬元"
	money_str(6) = "10萬元以上"
	
	sql = "select count(*), m014_money from m014 where m014_subjectid = " & subjectid & " group by m014_money"
	set ts = conn.execute(sql)
	for i = 1 to 6
		if ts.EOF then
			money_no = 0
		else
			if CInt(i) = CInt(ts("m014_money")) then
				money_no = ts(0)
				ts.MoveNext
			else
				money_no = 0
			end if
		end if
		
		if money_no > 0 then
			money_ratio = FormatPercent(money_no / base_answer_no, 1, False)
%>
        <tr> 
          <td><%= money_str(i) %></td>
          <td width="75%"><img src="images/bar_chart.gif" width="<%= money_ratio %>" height="12"></td>
          <td width="12%"><%= money_ratio %>(<%= money_no %>人)</td>
        </tr>
<%
		end if
	next
%>
      </table>
    </td>
  </tr>

  <tr>
    <td nowrap>職業</td>
    <td width="90%">
      <table width="100%" border="0" cellspacing="0" cellpadding="3">
<%
	dim	job_str(7)
	job_str(1) = "軍"
	job_str(2) = "公"
	job_str(3) = "商"
	job_str(4) = "教職"
	job_str(5) = "自由業"
	job_str(6) = "服務業"
	job_str(7) = "其他"
	
	sql = "select count(*), isNull(m014_job, '0') m014_job from m014 where m014_subjectid = " & subjectid & " group by m014_job"
	set ts = conn.execute(sql)
	for i = 1 to 7
		if ts.EOF then
			job_no = 0
		else
			if trim(ts("m014_job")) = "" then
				job_no = 0
			else
				if CInt(i) = CInt(ts("m014_job")) then
					job_no = ts(0)
					ts.MoveNext
				else
					job_no = 0
				end if
			end if
		end if
		
		if job_no > 0 then
			job_ratio = FormatPercent(job_no / base_answer_no, 1, False)
%>
        <tr> 
          <td><%= job_str(i) %></td>
          <td width="75%"><img src="images/bar_chart.gif" width="<%= job_ratio %>" height="12"></td>
          <td width="12%"><%= job_ratio %>(<%= job_no %>人)</td>
        </tr>
<%
		end if
	next
%>
      </table>
    </td>
  </tr>
  
  <tr>
    <td nowrap>學歷</td>
    <td width="90%">
      <table width="100%" border="0" cellspacing="0" cellpadding="3">
<%
	dim	edu_str(7)
	edu_str(1) = "國小"
	edu_str(2) = "國中"
	edu_str(3) = "高中"
	edu_str(4) = "大學"
	edu_str(5) = "碩士"
	edu_str(6) = "博士"
	edu_str(7) = "其他"
	
	sql = "select count(*), m014_edu from m014 where m014_subjectid = " & subjectid & " and isnull(m014_edu, '') <> '' group by m014_edu"
	set ts = conn.execute(sql)
	for i = 1 to 7
		edu_no = 0
		if not ts.eof then
			if int(i) = int(ts("m014_edu")) then
				edu_no = ts(0)
				ts.movenext
			end if
		end if
		
		if edu_no > 0 then
			edu_ratio = FormatPercent(edu_no / base_answer_no, 1, False)
%>
        <tr> 
          <td><%= edu_str(i) %></td>
          <td width="75%"><img src="images/bar_chart.gif" width="<%= edu_ratio %>" height="12"></td>
          <td width="12%"><%= edu_ratio %>(<%= edu_no %>人)</td>
        </tr>
<%
		end if
	next
%>
      </table>
    </td>
  </tr>
  
</table>
<p><a href="javascript:history.go(-1);">Previous Page</a></p>
</body>
</html>
