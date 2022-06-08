<%@  codepage="65001" %>
		
<%		
		'過濾字串,數字,strType s 為字串,i 為數字
		Function CheckInput(str,strType)
		Dim strTmp
			strTmp = ""
		If strType = "s" Then
			strTmp = Replace(Trim(str),"'","''")
			strTmp = Replace(Trim(str),"'","''")
			'strTmp = Replace(Trim(strTmp),";","")
			'strTmp = Replace(Trim(strTmp),"'","")
			'strTmp = Replace(Trim(strTmp),"--","")
			'strTmp = Replace(Trim(strTmp),"/*","")
			'strTmp = Replace(Trim(strTmp),"*/","")
			'strTmp = Replace(Trim(strTmp),"*","")
			'strTmp = Replace(Trim(strTmp),"/","")
			'strTmp = Replace(Trim(strTmp),"<","")
			'strTmp = Replace(Trim(strTmp),">","")
		ElseIf strType="i" Then
		If isNumeric(str)=False Then str="0"
			strTmp = str
		Else
			strTmp = str
		End If
			CheckInput = strTmp
		End Function
		
	   		'過濾script
		function DelJs(str)
			Dim  objRegExp
			Set  objRegExp=New  RegExp   
			objRegExp.IgnoreCase=True   
			objRegExp.Global=True   
			objRegExp.Pattern="\<script.+?\<\/script\>"   
			DelJs=objRegExp.Replace(str,"") 
		set  objRegExp=nothing
		end function
		
		function alertAndGoLast(str)
		    %>		    
	        <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
	        <html xmlns="http://www.w3.org/1999/xhtml">
	        <head>
                <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
                <title>資料管理／資料上稿</title>
            </head>
            <body>
		        <script type="text/javascript">
		            alert('<%=str %>')
		            history.go(-1)
		        </script>            
            </body>
            </html>
		    <%		
		    response.End
		end function
		
		
		CtRootID = CheckInput(Request("CtRootID"),"i")
		rowID = CheckInput(Request("rowID"),"i")
		htx_stitle = CheckInput(Request("htx_stitle"),"s")
		htx_stitle = DelJs(htx_stitle)
		htx_xbody = CheckInput(Request("htx_xbody"),"s")
		htx_xbody = DelJs(htx_xbody)
		htx_fctupublic = CheckInput(Request("htx_fctupublic"),"s")
		CreationDT = Date
		Editor = session("UserID")
		
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open session("ODBCDSN")
		
		
		select case request("Mode")
		    case "Modify"
			    
		        CheckPhrase = "Select * From SubjectPhrase Where Phrase='" & htx_stitle & "' And CtRootID='" & CtRootID & "' and rowID != '" & rowID & "'"
		        set rs = conn.execute(CheckPhrase)
	            if rs.eof then
		            '修改詞彙
			        xsql = "UPDATE SubjectPhrase set CtRootID=" & CtRootID & ",Phrase='" & htx_stitle & "',Content='" & htx_xbody & "',CreationDT='" & CreationDT & "',fCTUPublic='" & htx_fctupublic & "',Editor='" & Editor & "' where rowID='" & rowID & "'"
			        conn.execute xsql
			        response.redirect "../Phrase/Phrase.asp?mp=" & request("CtRootID")		        
		        else
			        alertAndGoLast("[" & htx_stitle & "] 已經存在")
		        end if			    
			    
		    case "Create"
		    
		        CheckPhrase = "Select * From SubjectPhrase Where Phrase='" & htx_stitle & "' And CtRootID='" & CtRootID & "'"
		        set rs = conn.execute(CheckPhrase)
	            if rs.eof then
		            '新增詞彙
		            xsql = "Insert into SubjectPhrase([CtRootID],[Phrase],[Content],[CreationDT],[fCTUPublic],[Editor]) "
		            xsql = xsql  & vbcrlf & "values ('" & CtRootID & "','" & htx_stitle & "','"
		            xsql = xsql  & vbcrlf & htx_xbody & "','" & CreationDT & "','" & htx_fctupublic & "','" &  Editor & "')"
		            conn.execute xsql
		            response.redirect "../Phrase/Phrase.asp?mp=" & request("CtRootID")
		        else
			        alertAndGoLast("[" & htx_stitle & "] 已經存在")
		        end if
		    
		    case "del"
			    xsql = "Delete SubjectPhrase where rowID='" & rowID & "'"
			    conn.execute xsql
			    response.redirect "../Phrase/Phrase.asp?mp=" & request("CtRootID")		    
		end select
		

%>
