<%@ CodePage = 65001 %>
<!-- #include virtual = "/inc/client.inc" -->
<!-- #include virtual = "/inc/dbFunc.inc" -->
<%
	subjectid = replace(trim(request("subjectid")), "'", "''")


	if subjectid = "" then
		response.write "<script language='javascript'>alert('操作錯誤！');history.go(-1);</script>"
		'response.end
	end if

	subject = trim(request("subject"))
	name = replace(request("name"), "'", "''")
	email= replace(request("email"), "'", "''")
	sex = request("sex")
	age = request("age")
	addrarea = request("addrarea")
	member = request("member")
	money = request("money")
	job = request("job")
	edu = request("edu")
	eid = request("eid")
	hospital = replace(request("hospital"), "'", "''")
	hospitalarea = replace(request("hospitalarea"), "'", "''")

		sql = "select R_name from Registerdata where R_subjectid = " & subjectid & " and R_name = '" & name & "'"
		set rs = conn.execute(sql)
		if not rs.eof then
			response.write "<script language='javascript'>alert('你已做過報名！');history.go(-1);</script>"
			'response.end
		else 

		sql = "" & _
			" insert into Registerdata ( " & _
			" R_subject, " & _
			" R_name, " & _
			" R_sex, " & _
			" R_email, " & _
			" R_age, " & _
			" R_addrarea, " & _
			" R_familymember, " & _
			" R_money, " & _
			" R_job, " & _
			" R_edu, " & _
			" R_subjectid, " & _
			" R_polldate, " & _
			" R_eid, " & _
			" R_hospital, " & _
			" R_hospitalarea " & _
			" ) values ( " & _
			" '" & subject & "', " & _
			" '" & name & "', " & _
			" '" & sex & "', " & _
			" '" & email & "', " & _
			" '" & age & "', " & _
			" '" & addrarea & "', " & _
			" '" & member & "', " & _
			" '" & money & "', " & _
			" '" & job & "', " & _
			" '" & edu & "', " & _
			subjectid & ", " & _
			" getdate(), " & _
			" '" & eid & "', " & _
			" '" & hospital & "', " & _
			" '" & hospitalarea & "' " & _			
			" ); "
		conn.execute(sql)
	response.write "<script language='javascript'>alert('您的報名表已經送出！');location.replace('http://social.hyweb.com.tw/lp.asp?ctNode=170&CtUnit=78&BaseDSD=34');</script>"
	end if
'	response.end
%>