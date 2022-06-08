<%@ CodePage = 65001 %>
<%
	HTProgCode = "GW1_vote01"
%>
<!-- #include virtual = "/inc/client.inc" -->
<!-- #include virtual = "/inc/dbutil.inc" -->
<%
	subjectid = request("subjectid")
	sql = "select m011_subject, m011_questno, m011_bdate, m011_edate from m011 where m011_subjectid = " & replace(subjectid, "'", "''")
	set rs = conn.execute(sql)
	if rs.eof then	response.end
	
	sex = array("", "男", "女")
	sex_type = array("", "M", "F")
	City = array("", "臺北縣", "宜蘭縣", "桃園縣", "新竹縣", "苗栗縣", "臺中縣", "彰化縣", "南投縣", "雲林縣", "嘉義縣", "臺南縣", "高雄縣", "屏東縣", "臺東縣", "花蓮縣", "澎湖縣", "基隆市", "新竹市", "臺中市", "嘉義市", "臺南市", "臺北市", "高雄市", "金門縣", "連江縣")
	Age = array("", "0到11歲以下", "12到17歲", "18到23歲", "24到29歲", "30到39歲", "40到49歲", "50到59歲", "60歲以上")
	Member = array("", "1人", "2人", "3人", "4人", "5至8人", "8人以上")
	Money = array("", "5000元以下", "5000到1萬元", "1萬到3萬元", "3萬到5萬元", "5萬到10萬元", "10萬元以上")
	Job = array("", "軍", "公", "商", "教職", "自由業", "服務業", "其他")
	edu = array("", "國小", "國中", "高中", "大學", "碩士", "博士", "其他")
	
	response.ContentType = "application/csv"
	response.AddHeader "Content-Disposition", "filename=export.csv;"
	
	response.write """調查對象分析""" & vbcrlf
	response.write """調查主題：" & trim(rs("m011_subject")) & """" & vbcrlf
	response.write """起訖時間：" & datevalue(rs("m011_bdate")) & " ~ " & datevalue(rs("m011_edate")) & """" & vbcrlf
	
	response.write vbcrlf & """性別""" & vbcrlf
	for i = 1 to 2
		sql = "select count(*) from m014 where m014_subjectid = " & replace(subjectid, "'", "''") & " and m014_sex = '" & sex_type(i) & "'"
		set ts = conn.execute(sql)
		response.write """" & sex(i) & """, """ & ts(0) & """" & vbcrlf
	next
	
	response.write vbcrlf & """年齡""" & vbcrlf
	for i = 1 to 8
		sql = "select count(*) from m014 where m014_subjectid = " & replace(subjectid, "'", "''") & " and m014_age = " & i
		set ts = conn.execute(sql)
		response.write """" & age(i) & """, """ & ts(0) & """" & vbcrlf
	next
	
	response.write vbcrlf & """居住縣市""" & vbcrlf
	for i = 1 to 25
		sql = "select count(*) from m014 where m014_subjectid = " & replace(subjectid, "'", "''") & " and m014_addrarea = " & i
		set ts = conn.execute(sql)
		response.write """" & city(i) & """, """ & ts(0) & """" & vbcrlf
	next
	
	response.write vbcrlf & """家庭成員人數""" & vbcrlf
	for i = 1 to 6
		sql = "select count(*) from m014 where m014_subjectid = " & replace(subjectid, "'", "''") & " and m014_familymember = " & i
		set ts = conn.execute(sql)
		response.write """" & Member(i) & """, """ & ts(0) & """" & vbcrlf
	next
	
	response.write vbcrlf & """收入""" & vbcrlf
	for i = 1 to 6
		sql = "select count(*) from m014 where m014_subjectid = " & replace(subjectid, "'", "''") & " and m014_money = " & i
		set ts = conn.execute(sql)
		response.write """" & Money(i) & """, """ & ts(0) & """" & vbcrlf
	next
	
	response.write vbcrlf & """職業""" & vbcrlf
	for i = 1 to 7
		sql = "select count(*) from m014 where m014_subjectid = " & replace(subjectid, "'", "''") & " and m014_job = " & i
		set ts = conn.execute(sql)
		response.write """" & job(i) & """, """ & ts(0) & """" & vbcrlf
	next
	
	response.write vbcrlf & """學歷""" & vbcrlf
	for i = 1 to 7
		sql = "select count(*) from m014 where m014_subjectid = " & replace(subjectid, "'", "''") & " and m014_edu = " & i
		set ts = conn.execute(sql)
		response.write """" & edu(i) & """, """ & ts(0) & """" & vbcrlf
	next
%>