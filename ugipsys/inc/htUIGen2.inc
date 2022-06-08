<% 
cq = chr(34)

sub dumpTempFile()
    Do While Not xfin.AtEndOfStream
        xinStr = xfin.readline
        response.write xinStr&vbCRLF
    Loop
end sub

sub writeCode(xstr)
	response.write xstr & vbCRLF
end sub

sub writePart(xstr)
	response.write xstr
end sub

sub showUINotes()
  if nullText(htPageDom.selectSingleNode("//pageSpec/UINotes")) <> "" then 
	response.write "<BR><HR><TABLE border=1><TR><TH bgcolor=pink>介 面 說 明</TH></TR><TR><TD>" 
'	response.write htPageDom.selectSingleNode("//pageSpec/UINotes").xml
'	response.write "**" & nullText(htPageDom.selectSingleNode("//pageSpec/UINotes")) & "**"
	recursiveTag htPageDom.selectSingleNode("//pageSpec/UINotes")
	response.write "</TD></TR></TABLE>"
  end if
end sub

sub showAidLinkList
	for each x in htPageDom.selectNodes("//pageSpec/aidLinkList/Anchor")
		ckRight = nullText(x.selectSingleNode("checkRight"))
		if isNumeric(ckRight) then
			ckRight = cint(ckRight)
		else
			ckRight = 0
		end if
		if ckRight = 0 OR (HTProgRight and ckRight)= ckRight then
		  if x.selectSingleNode("AnchorType").text="Back" then
			response.write "<A href=""Javascript:window.history.back();"">" _
				& x.selectSingleNode("AnchorLabel").text & "</A> "
		  else
			anChorURI = x.selectSingleNode("AnchorURI").text
			xpos = inStrRev(anChorURI,".")
			if xpos>0 then	anChorURI = left(anChorURI,xpos-1)
			xPos = inStrRev(anChorURI,"/")
			if xPos>0 then
				xprogPath = left(anChorURI,xpos-1)
				anChorURI = mid(anChorURI,xPos+1)
			else
				xprogPath = request("progPath")
			end if
			response.write "<A href=""uiView.asp?formID=" & anChorURI & "&progPath=" & xprogPath _
				& """ title=""" & nullText(x.selectSingleNode("AnchorDesc")) & """>" _
				& x.selectSingleNode("AnchorLabel").text & "</A> "
		  end if
		end if
	next
end sub

Function HelpEnumerateCodeList(codeID, helpField)
	SQL="Select * from CodeMetaDef where codeID=N'" & codeID & "'"
    SET RSLK=conn.execute(SQL)
	str=ct&ct&ct&cl & "SQL=""SELECT " & RSLK("CodeValueFld") & "," & RSLK("CodeDisplayFld")
	if helpField <> "" then		str = str & "," & helpField
	str = str & " FROM " & RSLK("CodeTblName")
	
	if not RSLK.EOF then
	  if isNull(RSLK("CodeSortFld")) then
		if isNull(RSLK("CodeSrcFld")) then
	    	str= str & """" & VBCRLF & _
		    	ct&ct&ct&"SET RSS=conn.execute(SQL)" & VBCRLF & _
		    	ct&ct&ct&"While not RSS.EOF"&cr& VBCRLF
		else	
	    	str= str & " WHERE " & RSLK("CodeSrcFld") & "=N'" & RSLK("CodeSrcItem") & "'""" & VBCRLF & _
		    	ct&ct&ct&"SET RSS=conn.execute(SQL)" & VBCRLF & _
			    ct&ct&ct&"While not RSS.EOF"&cr& VBCRLF
		end if          
	  else
		if isNull(RSLK("CodeSrcFld")) then
	    	str= str & " WHERE " & RSLK("CodeSortFld") & " IS NOT NULL Order by " & RSLK("CodeSortFld") & """" & VBCRLF & _
		    	ct&ct&ct&"SET RSS=conn.execute(SQL)" & VBCRLF & _
			    ct&ct&ct&"While not RSS.EOF"&cr& VBCRLF
	    else	
	    	str= str & " WHERE " & RSLK("CodeSortFld") & " IS NOT NULL AND " & RSLK("CodeSrcFld") & "=N'" & RSLK("CodeSrcItem") & "' Order by " & RSLK("CodeSortFld") & """" & VBCRLF & _
		    	ct&ct&ct&"SET RSS=conn.execute(SQL)" & VBCRLF & _
			    ct&ct&ct&"While not RSS.EOF"&cr& VBCRLF
		end if
	  end if
	end if
	HelpEnumerateCodeList=str
End Function

Function EnumerateCodeList(codeID)
	SQL="Select * from CodeMetaDef where codeID=N'" & codeID & "'"
        SET RSLK=conn.execute(SQL)
	str=""
	if not RSLK.EOF then
	  if isNull(RSLK("CodeSortFld")) then
		if isNull(RSLK("CodeSrcFld")) then
	    	Sql = "Select " & RSLK("CodeValueFld") & "," & RSLK("CodeDisplayFld") & " from " & RSLK("CodeTblName") 
		else	
	    	Sql = "Select " & RSLK("CodeValueFld") & "," & RSLK("CodeDisplayFld") & " from " & RSLK("CodeTblName") & " where " & RSLK("CodeSrcFld") & "=N'" & RSLK("CodeSrcItem") & "'"
		end if          
	  else
		if isNull(RSLK("CodeSrcFld")) then
	    	Sql = "Select " & RSLK("CodeValueFld") & "," & RSLK("CodeDisplayFld") & " from " & RSLK("CodeTblName") & " where " & RSLK("CodeSortFld") & " IS NOT NULL Order by " & RSLK("CodeSortFld")
	    else	
	    	Sql = "Select " & RSLK("CodeValueFld") & "," & RSLK("CodeDisplayFld") & " from " & RSLK("CodeTblName") & " where " & RSLK("CodeSortFld") & " IS NOT NULL AND " & RSLK("CodeSrcFld") & "=N'" & RSLK("CodeSrcItem") & "' Order by " & RSLK("CodeSortFld") 
		end if
	  end if
	End if
	EnumerateCodeList=Sql
End Function

sub uiEnumerateCodeList(codeID, xoutKind, paramCode)
	SQL="Select * from CodeMetaDef where codeID=N'" & codeID & "'"
        SET RSLK=conn.execute(SQL)
	str=""
	if not RSLK.EOF then
	  if isNull(RSLK("CodeSortFld")) then
		if isNull(RSLK("CodeSrcFld")) then
	    	Sql = "Select " & RSLK("CodeValueFld") & "," & RSLK("CodeDisplayFld") & " from " & RSLK("CodeTblName") 
		else	
	    	Sql = "Select " & RSLK("CodeValueFld") & "," & RSLK("CodeDisplayFld") & " from " & RSLK("CodeTblName") & " where " & RSLK("CodeSrcFld") & "=N'" & RSLK("CodeSrcItem") & "'"
		end if          
	  else
		if isNull(RSLK("CodeSrcFld")) then
	    	Sql = "Select " & RSLK("CodeValueFld") & "," & RSLK("CodeDisplayFld") & " from " & RSLK("CodeTblName") & " where " & RSLK("CodeSortFld") & " IS NOT NULL Order by " & RSLK("CodeSortFld")
	    else	
	    	Sql = "Select " & RSLK("CodeValueFld") & "," & RSLK("CodeDisplayFld") & " from " & RSLK("CodeTblName") & " where " & RSLK("CodeSortFld") & " IS NOT NULL AND " & RSLK("CodeSrcFld") & "=N'" & RSLK("CodeSrcItem") & "' Order by " & RSLK("CodeSortFld") 
		end if
	  end if
'	  response.write sql & "<HR>"
'	  response.end
	  Set RSS = conn.execute(Sql)
	  while not RSS.eof
	  	select case xoutKind
	  	  case "option":
			writeCode ct&ct&ct&"<option value="""& RSS(0) &""">"&RSS(1)&"</option>"
	  	  case "radio":
			writeCode ct&ct&ct&"<input type=""radio"" name=""htx_"&paramCode&""" value="""&RSS(0)& """ " & pdxc & " \>"
	  		writeCode ct&ct&ct&RSS(1)& "&nbsp;&nbsp;"
	  	end select
	  	RSS.moveNext
	  wend
	end if
End sub

sub recursiveTag(xDom)
dim x
	if xDom.nodeName = "refField" then
		processParam(xDom.text)
		exit sub
	end if
	if xDom.nodeName = "refAnchor" then
		processAnchor(xDom.text)
		exit sub
	end if
	if xDom.nodeName = "#comment" then	exit sub
	if xDom.nodeName = "#text" then
		writePart xDom.text
		exit sub
	end if
	writePart "<" & xDom.nodeName
	for xi = 0 to xDom.attributes.length-1
		writePart " " & xDom.attributes.item(xi).nodeName & "=""" _
			& xDom.attributes.item(xi).text & """"
	next
	writePart ">"
	for each x in xDom.childNodes
		recursiveTag x
	next
	writeCode "</" & xDom.nodeName & ">"
end sub

sub recursiveQParam(xDom)
dim x
	if xDom.nodeName = "refField" then
		processQParam(xDom.text)
		exit sub
	end if
	if xDom.nodeName = "refAnchor" then
		processAnchor(xDom.text)
		exit sub
	end if
	if xDom.nodeName = "#comment" then	exit sub
	if xDom.nodeName = "#text" then
		writePart xDom.text
		exit sub
	end if
	writePart "<" & xDom.nodeName
	for xi = 0 to xDom.attributes.length-1
		writePart " " & xDom.attributes.item(xi).nodeName & "=""" _
			& xDom.attributes.item(xi).text & """"
	next
	writePart ">"
	for each x in xDom.childNodes
		recursiveQParam x
	next
	writeCode "</" & xDom.nodeName & ">"
end sub

function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
'	if not isNull(xNode) then
'	  if isObject(xNode) then
'		nullText = 
'	  end if
'	else
'		nullText = "aaa"
'	end if
end function

function ncTabChar(n)
	if cTabchar = "" then
		ncTabChar = ""
	else
		ncTabChar = String(n,cTabchar)
	end if
end function

sub debugPrint(xstr)
	response.write xstr
end sub

sub processInsert(param)
	refField = param.selectSingleNode("fieldName").text
	paramType = param.selectSingleNode("dataType").text

'	if lcase(right(nullText(param.selectSingleNode("inputType")),4)) = "file"	then	exit sub

	if lcase(right(nullText(param.selectSingleNode("inputType")),4)) = "file" then
  	    For each xatt in xup.Attachments
	  	ofname = xatt.FileName
	  	fnExt = ""
	  	if instrRev(ofname, ".")>0 then	fnext=mid(ofname, instrRev(ofname, "."))
	  	tstr = now()
	  	nfname = right(year(tstr),1) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & cint(rnd(second(tstr))*100) & fnext 	    
    		if xatt.Name = "htImg_"+refField then
	  	    sql = sql & mid(xatt.Name,7) & ","
	  	    sqlValue = sqlValue & pkStr(nfname,",")
  	  	    xatt.SaveFile apath & nfname, false
  		elseif xatt.Name = "htFile_"+refField then
	  	    sql = sql & mid(xatt.Name,8) & ","
	  	    sqlValue = sqlValue & pkStr(nfname,",")
  	  	    xatt.SaveFile apath & nfname, false
  	  	    xsql = "INSERT INTO imageFile(newFileName, oldFileName) VALUES(" _
  	  		& pkStr(nfname,",") & pkStr(ofname,")")
  	  	    conn.execute xsql
		end if
  	    Next	
	elseif nullText(param.selectSingleNode("DBSave/type")) <> "" then
		select case param.selectSingleNode("DBSave/type").text
		    case "session"
			sql = sql & refField & ","
			sqlValue = sqlValue & pkStr(session(param.selectSingleNode("DBSave/set").text),",")
		    case "SQLFunc"
			sql = sql & refField & ","
			sqlValue = sqlValue & param.selectSingleNode("DBSave/set").text & ","
		    case "DoNothing"
		end select
	elseif xUpForm("htx_" & refField) <> "" Then
		sql = sql & refField & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_" & refField),",")
	END IF
end sub

sub processUpdate(param)
	refField = param.selectSingleNode("fieldName").text
	if nullText(param.selectSingleNode("mpKey")) = "Y" then	exit sub
	if nullText(param.selectSingleNode("inputType")) = "calc" then	exit sub
	paramType = param.selectSingleNode("dataType").text

	if lcase(right(nullText(param.selectSingleNode("inputType")),4)) = "file" then
	    if nullText(param.selectSingleNode("inputType"))="imgfile" then xType="Img"
	    if nullText(param.selectSingleNode("inputType"))="file" then xType="File"
	    fname = ""
    	    if xUpForm("ht"+xType+"ActCK_"+refField) <> "" then
    	        actCK = xUpForm("ht"+xType+"ActCK_"+refField)
    	        if actCK="editLogo" OR actCK="addLogo" then
	    	    For each xatt in xup.Attachments
		    	ofname = xatt.FileName
	  	    	fnExt = ""
	  	    	if instrRev(ofname, ".")>0 then	fnext=mid(ofname, instrRev(ofname, "."))
	  	    	tstr = now()
	  	    	nfname = right(year(tstr),1) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & cint(rnd(second(tstr))*100) & fnext 	    
	    	    	if xatt.Name="ht"+xType+"_"+refField then    	    	    	    
			    sql = sql & refField & "=" & pkStr(nfname,",")
			    IF xUpForm("ho"+xType+"_"+refField) <> "" Then
			        if xup.IsFileExist( apath & xUpForm("ho"+xType+"_"+refField)) then _
					xup.DeleteFile apath & xUpForm("ho"+xType+"_"+refField)
			        xsql = "DELETE imageFile WHERE newFileName=" & pkStr(xUpForm("ho"+xType+"_"+refField),"")
			        conn.execute xsql
			    END IF
			    xatt.SaveFile apath & nfname, false
	    	    	    if xType = "File" then
			    	xsql = "INSERT INTO imageFile(newFileName, oldFileName) VALUES(" & pkStr(nfname,",") & pkStr(ofname,")")
			    	conn.execute xsql	    	    	
	    	    	    end if   	    
	    	    	end if
	   	    Next    	        
    	        elseif actCK="delLogo" then
		    if xup.IsFileExist( apath & xUpForm("ho"+xType+"_"+refField)) then _
			xup.DeleteFile apath & xUpForm("ho"+xType+"_"+refField)
		    xsql = "DELETE imageFile WHERE newFileName=" & pkStr(xUpForm("ho"+xType+"_"+refField),"")
		    conn.execute xsql	    	    
		    sql = sql & refField & "=null,"    	        
    	        end if
	    end if
	elseif nullText(param.selectSingleNode("DBSave/type")) <> "" then
		select case param.selectSingleNode("DBSave/type").text
		    case "session"
			sql = sql & refField & "="
	  		sql = sql & pkStr(session(param.selectSingleNode("DBSave/set").text),",")
		    case "SQLFunc"
			sql = sql & refField & "="
	  		sql = sql & param.selectSingleNode("DBSave/set").text & ","
		    case "DoNothing"
		end select
	elseif nullText(param.selectSingleNode("inputType"))<>"systemGenerate" then
	    select case paramType
	  	case "calc"
	  	case "integer"
			sql = sql & refField & "=" & drn2("htx_" & refField)
	  	case else
			sql = sql & refField & "=" & pkStr(xUpForm("htx_" & refField),",")
	    end select
	end if
end sub

sub processImgFile(xstr)
'	response.write "<HR>" & xstr & "<HR>"
	inPos = instr(xstr,"/")
	refTable = left(xstr,inPos-1)
	refField = mid(xstr,inPos+1)
	set param = refModel.selectSingleNode("fieldList[tableName='" & refTable & "']/field[fieldName='" & refField & "']")
	inputType = param.selectSingleNode("inputType").text

	response.write ct&"IF xUpForm(""htImgActCK_" & refField & """) <> """" Then"
	response.write ct&"  actCK = xUpForm(""htImgActCK_" & refField & """)"
	response.write ct&"  if actCK=""editLogo"" OR actCK=""addLogo"" then"
	response.write ct&ct&"fname = """""
	response.write ct&ct&"For each xatt in xup.Attachments"
	response.write ct&ct&"  if xatt.Name = ""htImg_" & refField & """ then"
	response.write ct&ct&ct&"ofname = xatt.FileName"
	response.write ct&ct&ct&"fnExt = """""
	response.write ct&ct&ct&"if instrRev(ofname, ""."")>0 then	fnext=mid(ofname, instrRev(ofname, "".""))"
	response.write ct&ct&ct&"tstr = now()"
	response.write ct&ct&ct&"nfname = right(year(tstr),1) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & cint(rnd(second(tstr))*100) & fnext"
	response.write ct&ct&ct&"sql = sql & """ & refField & "="" & " & "pkStr(nfname,"","")"
	response.write ct&ct&ct&"IF xUpForm(""hoImg_" & refField & """) <> """" Then _"
	response.write ct&ct&ct&ct&"xup.DeleteFile apath & xUpForm(""hoImg_" & refField & """)"
	response.write ct&ct&ct&"xatt.SaveFile apath & nfname, false"
	response.write ct&ct&"  end if"
	response.write ct&ct&"Next"
	response.write ct&"  elseif actCK=""delLogo"" then"
	response.write ct&ct&"xup.DeleteFile apath & xUpForm(""hoImg_" & refField & """)"
	response.write ct&ct&"sql = sql & """ & refField & "=null,"""
	response.write ct&"  end if"
	response.write ct&"END IF"

end sub

sub processAttFile(xstr)
'	response.write "<HR>" & xstr & "<HR>"
	inPos = instr(xstr,"/")
	refTable = left(xstr,inPos-1)
	refField = mid(xstr,inPos+1)
	set param = refModel.selectSingleNode("fieldList[tableName='" & refTable & "']/field[fieldName='" & refField & "']")
	inputType = param.selectSingleNode("inputType").text

	response.write ct&"IF xUpForm(""htFileActCK_" & refField & """) <> """" Then"
	response.write ct&"  actCK = xUpForm(""htFileActCK_" & refField & """)"
	response.write ct&"  if actCK=""editLogo"" OR actCK=""addLogo"" then"
	response.write ct&ct&"fname = """""
	response.write ct&ct&"For each xatt in xup.Attachments"
	response.write ct&ct&"  if xatt.Name = ""htFile_" & refField & """ then"
	response.write ct&ct&ct&"ofname = xatt.FileName"
	response.write ct&ct&ct&"fnExt = """""
	response.write ct&ct&ct&"if instrRev(ofname, ""."")>0 then	fnext=mid(ofname, instrRev(ofname, "".""))"
	response.write ct&ct&ct&"tstr = now()"
	response.write ct&ct&ct&"nfname = right(year(tstr),1) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & cint(rnd(second(tstr))*100) & fnext"
	response.write ct&ct&ct&"sql = sql & """ & refField & "="" & " & "pkStr(nfname,"","")"
	response.write ct&ct&ct&"IF xUpForm(""hoFile_" & refField & """) <> """" Then"
	response.write ct&ct&ct&ct&"if xup.IsFileExist( apath & xUpForm(""hoFile_" & refField & """)) then _"
	response.write ct&ct&ct&ct&ct&"xup.DeleteFile apath & xUpForm(""hoFile_" & refField & """)"
	response.write ct&ct&ct&ct&"xsql = ""DELETE imageFile WHERE newFileName="" & pkStr(xUpForm(""hoFile_" & refField & """),"""")"
	response.write ct&ct&ct&ct&"conn.execute xsql"
	response.write ct&ct&ct&"END IF"
	response.write ct&ct&ct&"xatt.SaveFile apath & nfname, false"
	response.write ct&ct&ct&"xsql = ""INSERT INTO imageFile(newFileName, oldFileName) VALUES("" & pkStr(nfname,"","") & pkStr(ofname,"")"")"
	response.write ct&ct&ct&"conn.execute xsql"
	response.write ct&ct&"  end if"
	response.write ct&ct&"Next"
	response.write ct&"  elseif actCK=""delLogo"" then"
	response.write ct&ct&"if xup.IsFileExist( apath & xUpForm(""hoFile_" & refField & """)) then _"
	response.write ct&ct&ct&"xup.DeleteFile apath & xUpForm(""hoFile_" & refField & """)"
	response.write ct&ct&"xsql = ""DELETE imageFile WHERE newFileName="" & pkStr(xUpForm(""hoFile_" & refField & """),"""")"
	response.write ct&ct&"conn.execute xsql"
	response.write ct&ct&"sql = sql & """ & refField & "=null,"""
	response.write ct&"  end if"
	response.write ct&"END IF"

end sub

sub processValid(param)
	refField = param.selectSingleNode("fieldName").text
	it = nullText(param.selectSingleNode("inputType"))
	l  = nullText(param.selectSingleNode("dataLen"))
	if nullText(param.selectSingleNode("canNull")) = "N" then
	 select case nullText(param.selectSingleNode("inputType"))
	  case "refCheckboxOther"
	  case "refCheckbox"
	  case "checkbox"
	  case "refRadioOther"
	  case "refRadio"
	  case "radio"
	  case "imgfile"	  
	  case else
		writeCode ct&"IF reg.htx_" & refField & ".value = Empty Then"
		writeCode ct&ct&"MsgBox replace(nMsg,""{0}""," _
			& cq & param.selectSingleNode("fieldLabel").text & """), 64, ""Sorry!"""
		writeCode ct&ct&"reg.htx_" & refField & ".focus"
		writeCode ct&ct&"exit sub"
		writeCode ct&"END IF"
	 end select
	end if
	dt = nullText(param.selectSingleNode("dataType"))
	if it = "imgfile" then
			writeCode ct&"IF reg.htImg_" & refField & ".value <> """" Then" 
			writeCode ct&ct&"xIMGname = reg.htImg_" & refField & ".value"
			writeCode ct&ct&"xFileType = """""
			writeCode ct&ct&"if instrRev(xIMGname, ""."")>0 then	xFileType=lcase(mid(xIMGname, instrRev(xIMGname, ""."")+1))"
			writeCode ct&ct&"IF xFileType<>""gif"" AND xFileType<>""jpg"" AND xFileType<>""jpeg"" then"
			writeCode ct&ct&ct&"MsgBox replace(pMsg,""{0}""," _
				& cq & param.selectSingleNode("fieldLabel").text & """), 64, ""Sorry!"""
			writeCode ct&ct&ct&"reg.htImg_" & refField & ".focus"
			writeCode ct&ct&ct&"exit sub"
			writeCode ct&ct&"END IF"
			writeCode ct&"END IF"
	elseif it = "MMOfile" then
			writeCode ct&"IF reg.htMMO_" & refField & ".value <> """" Then" 
			writeCode ct&ct&"xMMOname = reg.htMMO_" & refField & ".value"
			writeCode ct&ct&"xFileType = """""
			writeCode ct&ct&"IF instrRev(xMMOname, ""."")=0 then"
			writeCode ct&ct&ct&"MsgBox replace(mMsg,""{0}""," _
				& cq & param.selectSingleNode("fieldLabel").text & """), 64, ""Sorry!"""
			writeCode ct&ct&ct&"reg.htMMO_" & refField & ".focus"
			writeCode ct&ct&ct&"exit sub"
			writeCode ct&ct&"END IF"
			writeCode ct&"END IF"
	elseif lcase(right(dt,4)) = "char" then
		if l<>"" AND (it="" OR it="textbox" OR it="textarea") then
			writeCode ct&"IF blen(reg.htx_" & refField & ".value) > " & l & " Then"
			writeCode ct&ct&"MsgBox replace(replace(lMsg,""{0}""," _
				& cq & param.selectSingleNode("fieldLabel").text & """),""{1}"","""&l&"""), 64, ""Sorry!"""
			writeCode ct&ct&"reg.htx_" & refField & ".focus"
			writeCode ct&ct&"exit sub"
			writeCode ct&"END IF"
		end if
	elseif lcase(right(dt,8)) = "datetime" then
			writeCode ct&"IF (reg.htx_" & refField & ".value <> """") AND (NOT isDate(reg.htx_" & refField & ".value)) Then" 
			writeCode ct&ct&"MsgBox replace(dMsg,""{0}""," _
				& cq & param.selectSingleNode("fieldLabel").text & """), 64, ""Sorry!"""
			writeCode ct&ct&"reg.htx_" & refField & ".focus"
			writeCode ct&ct&"exit sub"
			writeCode ct&"END IF"
	elseif lcase(left(dt,3)) = "int" then
			writeCode ct&"IF (reg.htx_" & refField & ".value <> """") AND (NOT isNumeric(reg.htx_" & refField & ".value)) Then" 
			writeCode ct&ct&"MsgBox replace(iMsg,""{0}""," _
				& cq & param.selectSingleNode("fieldLabel").text & """), 64, ""Sorry!"""
			writeCode ct&ct&"reg.htx_" & refField & ".focus"
			writeCode ct&ct&"exit sub"
			writeCode ct&"END IF"
	end if
end sub


sub processInit(xstr)
	If formFunction = "add" then
		AddProcessInit xstr
	Elseif formFunction = "edit" then
		EditProcessInit xstr
	Elseif formFunction = "query" then
		QueryProcessInit xstr		
	End if
end sub

sub AddProcessInit(param)
	refField = param.selectSingleNode("fieldName").text
	if nullText(param.selectSingleNode("clientDefault/type")) = "" then	exit sub
	select case nullText(param.selectSingleNode("inputType"))
	  case "imgFile"
	  case "file"
	  case "refRadio"
		lhs = ct & "initRadio ""htx_" & refField & ""","
	  case "radio"
		lhs = ct & "initRadio ""htx_" & refField & ""","
	  case "refCheckbox"
		lhs = ct & "initCheckbox ""htx_" & refField & ""","
	  case "checkbox"
		lhs = ct & "initCheckbox ""htx_" & refField & ""","
	  case "popDate"
		lhs = ct&"reg.pcShowhtx_" & refField & ".value= "
	  case else
		lhs = ct&"reg.htx_" & refField & ".value= "
	end select
	select case param.selectSingleNode("clientDefault/type").text
	  case "value"
	  	rhs = cq & param.selectSingleNode("clientDefault/set").text & cq
	  case "clientFunc"
	  	rhs = param.selectSingleNode("clientDefault/set").text
	  case "serverFunc"
	  	rhs = cq & param.selectSingleNode("clientDefault/set").text  & cq
	  case "session"
	  	rhs = cq & session(param.selectSingleNode("clientDefault/set").text) & cq
	end select
  	writeCode lhs & rhs
	if nullText(param.selectSingleNode("inputType"))="popDate" then
		lhs = ct&"reg.htx_" & refField & ".value= "
	  	writeCode lhs & rhs
	end if
	
end sub

sub EditProcessInit(param)
	if nullText(param.selectSingleNode("updateDefault")) = "Y" then
		AddProcessInit(param)
		exit sub
	end if
	refField = param.selectSingleNode("fieldName").text
	if nullText(param.selectSingleNode("calc")) <> "" then _
  		response.write ct & param.selectSingleNode("calc").text
	if nullText(param.selectSingleNode("inputType")) <> "calc" then
	  select case nullText(param.selectSingleNode("inputType"))
	  	case "refRadioOther"
			writeCode ct & "initOtherRadio ""htx_" & refField & """,""" & qqRS(refField) & cq  _
				& ", ""other_" & refField & cq	
	  	case "refRadio"
			writeCode ct & "initRadio ""htx_" & refField & """,""" & qqRS(refField) & cq 
	  	case "radio"
			writeCode ct & "initRadio ""htx_" & refField & """,""" & qqRS(refField) & cq 
	  	case "refCheckboxOther"
			writeCode ct & "initOtherCheckbox ""htx_" & refField & """,""" & qqRS(refField) & cq  _
				& ", ""other_" & refField & cq	
	  	case "refCheckbox"
			writeCode ct & "initCheckbox ""htx_" & refField & """,""" & qqRS(refField) & cq 
	  	case "checkbox"
			writeCode ct & "initCheckbox ""htx_" & refField & """,""" & qqRS(refField) & cq 
	  	case "calc"
	  		writeCode ct & param.selectSingleNode("calc").text
		case "readonlydefault"	  	
			lhs = ct&"reg.htx_" & refField & ".value= "
			select case param.selectSingleNode("clientDefault/type").text
	  		case "value"
	  			rhs = cq & param.selectSingleNode("clientDefault/set").text & cq
	  		case "clientFunc"
	  			rhs = param.selectSingleNode("clientDefault/set").text
	  		case "serverFunc"
	  			rhs = cq & param.selectSingleNode("clientDefault/set").text & cq
	  		case "session"
	  			rhs = cq & session(param.selectSingleNode("clientDefault/set").text) & cq
			end select			
  			writeCode lhs & rhs
  			writeCode	ct&"reg.htx_" & refField & ".readonly=""true"""	
			writeCode	ct&"reg.htx_" & refField & ".classname=""rdonly""" 	  		
	  	case "file"	  	
			writeCode ct & "initAttFile """ & refField & """,""" & qqRS(refField) & cq & ", """ & qqRS("fxr_" & refField) & cq 
'				& """, """ & qqRS("fxr_" & refField) & cq 
'			lhs = ct&"reg.hoFile_" & refField & ".value= "
'			rhs = cq & cl & "=qqRS(""fxr_" & refField & """)" & cr & cq
'  			response.write lhs & rhs
	  	case "imgfile"
			writeCode ct & "initImgFile """ & refField & """,""" & qqRS(refField) & cq 
'			lhs = ct&"document.all(""imgSrc_" & refField & """).src= "
'			rhs = cq & cl & "=HTUploadPath" & cr & cl & "=qqRS(""" & refField & """)" & cr & cq
'  			response.write lhs & rhs
	  	case "MMOfile"
			writeCode ct & "initMMOFile """ & refField & """,""" & qqRS(refField) & cq 
	  	case "popDate"
			lhs = ct&"reg.htx_" & refField & ".value= "
			rhs = cq & qqRS(refField) & cq
  			writeCode lhs & rhs
			lhs = ct&"reg.pcShowhtx_" & refField & ".value= "
			rhs = "d7Date(" & cq & qqRS(refField) & cq & ")"
  			writeCode lhs & rhs
	  	case else
'		response.write "reg.htx_" & kf & ".value= """ & qqRS(kf) & """" & vbCRLF
			lhs = ct&"reg.htx_" & refField & ".value= "
			rhs = cq & qqRS(refField) & cq
  			writeCode lhs & rhs
  	  end select
  	end if
end sub

sub QueryProcessInit(xstr)
	inPos = instr(xstr,"/")
	refTable = left(xstr,inPos-1)
	refField = mid(xstr,inPos+1)
	set param = refModel.selectSingleNode("fieldList[tableName='" & refTable & "']/field[fieldName='" & refField & "']")
	if nullText(param.selectSingleNode("clientDefault/type")) = "" then	exit sub
	lhs = ct&"reg.htx_" & refField & ".value= "
	select case param.selectSingleNode("clientDefault/type").text
	  case "value"
	  	rhs = cq & param.selectSingleNode("clientDefault/set").text & cq
	  case "clientFunc"
	  	rhs = param.selectSingleNode("clientDefault/set").text
	  case "serverFunc"
	  	rhs = cq & cl & "=" & param.selectSingleNode("clientDefault/set").text & cr & cq
	  case "session"
	  	rhs = cq & cl & "=session(" & param.selectSingleNode("clientDefault/set").text & ")" & cr & cq
	end select
  	response.write lhs & rhs
end sub

sub processAnchor(xstr)
'	response.write "anchorList/anchor[funcLabel='" & xstr & "']<HR>"
	set param = refModel.selectSingleNode("anchorList/anchor[funcLabel='" & xstr & "']")
			anChorURI = param.selectSingleNode("url").text
			xpos = inStrRev(anChorURI,".")
			if xpos>0 then	anChorURI = left(anChorURI,xpos-1)
			xPos = inStrRev(anChorURI,"/")
			if xPos>0 then
				xprogPath = left(anChorURI,xpos-1)
				anChorURI = mid(anChorURI,xPos+1)
			else
				xprogPath = request("progPath")
			end if
	select case nullText(param.selectSingleNode("type"))
		case "button"
			response.write ct & "<INPUT TYPE=button VALUE=""" & xstr & """ onClick=""" _
				& param.selectSingleNode("action").text & "('uiView.asp?formID=" _
				& anChorURI & "&progPath=" & xprogPath & "')"">"
	end select
end sub

sub processParam(xstr)
'	response.write xstr
	inPos = instr(xstr,"/")
	refTable = left(xstr,inPos-1)
	refField = mid(xstr,inPos+1)
'	response.write xrefTable & "==>" & xrefField & "<BR>"
'	response.write "fieldList[tableName='" & xrefTable & "']/field[fieldName='" & refField & "']"
'	exit sub

'--------------add or edit function
    	set param = refModel.selectSingleNode("fieldList[tableName='" & refTable & "']/field[fieldName='" & refField & "']")
	processParamField param
end sub

sub processQParam(xstr)
'	response.write xstr
	inPos = instr(xstr,"/")
	refTable = left(xstr,inPos-1)
	refField = mid(xstr,inPos+1)
'	response.write xrefTable & "==>" & xrefField & "<BR>"
'	response.write "fieldList[tableName='" & xrefTable & "']/field[fieldName='" & refField & "']"
'	exit sub

'--------------add or edit function
    	set param = refModel.selectSingleNode("fieldList[tableName='" & refTable & "']/field[fieldName='" & refField & "']")
	processQParamField param
end sub

sub processParamField (param)
		paramCode = param.selectSingleNode("fieldName").text
		paramType = param.selectSingleNode("dataType").text
		paramSize= nullText(param.selectSingleNode("dataLen"))
		fieldDesc= nullText(param.selectSingleNode("fieldDesc"))
		if trimfieldDesc="" then	fieldDesc=nullText(param.selectSingleNode("fieldLabel"))
		if fieldDesc <> "" then _
			fieldDesc = " title=""" & fieldDesc & cq
		if paramSize = "" then paramSize = 50
		if paramSize > 50 then paramSize = 50
'			writeCode nullText(param.selectSingleNode("inputType"))
		select case nullText(param.selectSingleNode("inputType"))
		  case ""
		  	writeCode "<input name=""htx_"&paramCode&""" size="""&paramSize& cq & fieldDesc & ">"
		  case "textbox"
		  	writeCode "<input name=""htx_"&paramCode&""" size="""&paramSize& cq & fieldDesc & " >"
		  case "calc"
		  	writeCode "<input name=""htx_"&paramCode&""" size="""&paramSize&""" readonly=""true"">"
		  case "systemGenerate"
		  	writeCode "<input name=""htx_"&paramCode&""" size="""&paramSize&""" readonly=""true"" class=""rdonly"">"
		  case "readonly"
		  	writeCode "<input name=""htx_"&paramCode&""" size="""&paramSize&""" readonly=""true"" class=""rdonly"">"
		  case "hidden"
		  	writeCode "<input type=""hidden"" name=""htx_"&paramCode&""">"
		  case "varchar"
		  	writeCode "<input name=""htx_"&paramCode&""" size="""&paramSize&""">"
		  case "textarea"
		  	writeCode "<textarea name=""htx_"&paramCode&""" rows="""&nullText(param.selectSingleNode("rows")) _
		  		&""" cols="""&nullText(param.selectSingleNode("cols"))& cq & fieldDesc & ">"
		  	writeCode "</textarea>"
		  case "file"
		   If formFunction = "add" then
		  	writeCode "<input type=""file"" name=""htFile_" & paramCode & """>"
		   Elseif formFunction = "edit" then
			writeCode "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">"
			writeCode ct & "<tr><td colspan=""2"">"
		  	writeCode ct&ct&"<input type=""file"" name=""htFile_" & paramCode & """>"
		  	writeCode ct&ct&"<input type=""hidden"" name=""hoFile_" & paramCode & """>"
		  	writeCode ct&ct&"<input type=""hidden"" name=""htFileActCK_" & paramCode & """>"
			writeCode ct & "</td></tr>"
			writeCode ct & "<tr>"
			writeCode ct & "<td width=""37%""><span id=""logo_" & paramCode & """ class=""rdonly""></span>"
		  	writeCode ct&ct&"<div id=""noLogo_" & paramCode & """ style=""color:red"">無附檔</div></td>"
			writeCode ct & "<td valign=""bottom"">"
		  	writeCode ct&ct&"<div id=""LbtnHide0_" & paramCode & """>"
 		  	writeCode ct&ct&cl&" if (HTProgRight and 4)<>0 then " & cr
		  	writeCode ct&ct&ct&"<img id=""addLogo_" & paramCode & """ class=""hand"" src=""../pagestyle/addimg.gif"" onClick=""VBS: addXFile '" & paramCode & "'"">"
 		  	writeCode ct&ct&cl&" End if " & cr
		  	writeCode ct&ct&"</div>"
		  	writeCode ct&ct&"<div id=""LbtnHide1_" & paramCode & """>"
 		  	writeCode ct&ct&cl&" if (HTProgRight and 8)<>0 then " & cr
		  	writeCode ct&ct&ct&"<img id=""chgLogo_" & paramCode & """ class=""hand"" src=""../pagestyle/chimg.gif"" onClick=""VBS: chgXFile '" & paramCode & "'"">"
 		  	writeCode ct&ct&cl&" End if " & cr
 		  	writeCode ct&ct&cl&" if (HTProgRight and 16)<>0 then " & cr
		  	writeCode ct&ct&ct&"<img id=""delLogo_" & paramCode & """ class=""hand"" src=""../pagestyle/delimg.gif"" onClick=""VBS: delXFile '" & paramCode & "'"">"
 		  	writeCode ct&ct&cl&" End if " & cr
		  	writeCode ct&ct&"</div>"
		  	writeCode ct&ct&"<div id=""LbtnHide2_" & paramCode & """>"
 		  	writeCode ct&ct&cl&" if (HTProgRight and 8)<>0 then " & cr
		  	writeCode ct&ct&ct&"<img id=""orgLogo_" & paramCode & """ class=""hand"" src=""../pagestyle/resetimg.gif"" onClick=""VBS: orgXFile '" & paramCode & "'"">"
 		  	writeCode ct&ct&cl&" End if " & cr
		  	writeCode ct&ct&"</div>"
			writeCode ct & "</td></tr></table>"
		   Elseif formFunction = "xedit" then
			writeCode "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">"
			writeCode ct & "<tr><td colspan=""2"">"
		  	writeCode ct&ct&"<input type=""file"" name=""htFile_" & paramCode & """>"
		  	writeCode ct&ct&"<input type=""hidden"" name=""hoFile_" & paramCode & """>"
		  	writeCode ct&ct&"<input type=""hidden"" name=""htFileActCK_" & paramCode & """>"
			writeCode ct & "</td></tr>"
			writeCode ct & "<tr>"
			writeCode ct & "<td width=""37%""><span id=""logo_" & paramCode & """ class=""rdonly""></span>"
		  	writeCode ct&ct&"<div id=""noLogo_" & paramCode & """ style=""color:red"">無附檔</div></td>"
			writeCode ct & "<td valign=""bottom"">"
		  	writeCode ct&ct&"<div id=""LbtnHide0_" & paramCode & """>"
 		  	writeCode ct&ct&cl&" if (HTProgRight and 4)<>0 then " & cr
		  	writeCode ct&ct&ct&"<img id=""addLogo_" & paramCode & """ class=""hand"" src=""../pagestyle/addimg.gif"" onClick=""VBS: addXFile '" & paramCode & "'"">"
 		  	writeCode ct&ct&cl&" End if " & cr
		  	writeCode ct&ct&"</div>"
		  	writeCode ct&ct&"<div id=""LbtnHide1_" & paramCode & """>"
 		  	writeCode ct&ct&cl&" if (HTProgRight and 8)<>0 then " & cr
		  	writeCode ct&ct&ct&"<img id=""chgLogo_" & paramCode & """ class=""hand"" src=""../pagestyle/chimg.gif"" onClick=""VBS: chgXFile '" & paramCode & "'"">"
 		  	writeCode ct&ct&cl&" End if " & cr
 		  	writeCode ct&ct&cl&" if (HTProgRight and 16)<>0 then " & cr
		  	writeCode ct&ct&ct&"<img id=""delLogo_" & paramCode & """ class=""hand"" src=""../pagestyle/delimg.gif"" onClick=""VBS: delXFile '" & paramCode & "'"">"
 		  	writeCode ct&ct&cl&" End if " & cr
		  	writeCode ct&ct&"</div>"
		  	writeCode ct&ct&"<div id=""LbtnHide2_" & paramCode & """>"
 		  	writeCode ct&ct&cl&" if (HTProgRight and 8)<>0 then " & cr
		  	writeCode ct&ct&ct&"<img id=""orgLogo_" & paramCode & """ class=""hand"" src=""../pagestyle/resetimg.gif"" onClick=""VBS: orgXFile '" & paramCode & "'"">"
 		  	writeCode ct&ct&cl&" End if " & cr
		  	writeCode ct&ct&"</div>"
			writeCode ct & "</td></tr></table>"
		   end if
		  case "MMOfile"
		   If formFunction = "add" then
		  	writeCode "<input type=""file"" name=""htMMO_" & paramCode & """>"
		   Elseif formFunction = "edit" then
			writeCode "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">"
			writeCode ct & "<tr><td colspan=""2"">"
		  	writeCode ct&ct&"<input type=""file"" name=""htMMO_" & paramCode & """>"
		  	writeCode ct&ct&"<input type=""hidden"" name=""hoMMO_" & paramCode & """>"
		  	writeCode ct&ct&"<input type=""hidden"" name=""htMMOActCK_" & paramCode & """>"
			writeCode ct & "</td></tr>"
			writeCode ct & "<tr>"
			writeCode ct & "<td width=""37%""><div><img id=""logo_" & paramCode & """ src=""""></div><span id=""filename_" & paramCode & """ class=""rdonly""></span>"
		  	writeCode ct&ct&"<div id=""noLogo_" & paramCode & """ style=""color:red"">無附檔</div></td>"
			writeCode ct & "<td valign=""bottom"">"
		  	writeCode ct&ct&"<div id=""LbtnHide0_" & paramCode & """>"
 		  	writeCode ct&ct&cl&" if (HTProgRight and 4)<>0 then " & cr
		  	writeCode ct&ct&ct&"<img id=""addLogo_" & paramCode & """ class=""hand"" src=""../pagestyle/but_do_new.gif"" onClick=""VBS: addMMO '" & paramCode & "'"">"
 		  	writeCode ct&ct&cl&" End if " & cr
		  	writeCode ct&ct&"</div>"
		  	writeCode ct&ct&"<div id=""LbtnHide1_" & paramCode & """>"
 		  	writeCode ct&ct&cl&" if (HTProgRight and 8)<>0 then " & cr
		  	writeCode ct&ct&ct&"<img id=""chgLogo_" & paramCode & """ class=""hand"" src=""../pagestyle/but_do_change.gif"" onClick=""VBS: chgMMO '" & paramCode & "'"">"
 		  	writeCode ct&ct&cl&" End if " & cr
 		  	writeCode ct&ct&cl&" if (HTProgRight and 16)<>0 then " & cr
		  	writeCode ct&ct&ct&"<img id=""delLogo_" & paramCode & """ class=""hand"" src=""../pagestyle/but_do_delete.gif"" onClick=""VBS: delMMO '" & paramCode & "'"">"
 		  	writeCode ct&ct&cl&" End if " & cr
		  	writeCode ct&ct&"</div>"
		  	writeCode ct&ct&"<div id=""LbtnHide2_" & paramCode & """>"
 		  	writeCode ct&ct&cl&" if (HTProgRight and 8)<>0 then " & cr
		  	writeCode ct&ct&ct&"<img id=""orgLogo_" & paramCode & """ class=""hand"" src=""../pagestyle/but_do_orginal.gif"" onClick=""VBS: orgMMO '" & paramCode & "'"">"
 		  	writeCode ct&ct&cl&" End if " & cr
		  	writeCode ct&ct&"</div>"
			writeCode ct & "</td></tr></table>"
		   end if		   
		  case "imgfile"
		   If formFunction = "add" then
		  	writeCode "<input type=""file"" name=""htImg_" & paramCode & """>"
		  	if nullText(param.selectSingleNode("fieldDesc"))<>"" then
		  		writeCode "</br><span style=""color:red;"">" & nullText(param.selectSingleNode("fieldDesc")) & "</span>"
		  	end if
		   Elseif formFunction = "edit" then
			writeCode "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">"
			writeCode ct & "<tr><td colspan=""2"">"
		  	writeCode ct&ct&"<input type=""file"" name=""htImg_" & paramCode & """>"
		  	writeCode ct&ct&"<input type=""hidden"" name=""hoImg_" & paramCode & """>"
		  	writeCode ct&ct&"<input type=""hidden"" name=""htImgActCK_" & paramCode & """>"
			writeCode ct & "</td></tr>"
			writeCode ct & "<tr>"
			writeCode ct & "<td width=""37%""><img id=""logo_" & paramCode & """ src="""">"
		  	writeCode ct&ct&"<div id=""noLogo_" & paramCode & """ style=""color:red"">無圖片</div></td>"
			writeCode ct & "<td valign=""bottom"">"
		  	writeCode ct&ct&"<div id=""LbtnHide0_" & paramCode & """>"
 		  	writeCode ct&ct&cl&" if (HTProgRight and 4)<>0 then " & cr
		  	writeCode ct&ct&ct&"<img id=""addLogo_" & paramCode & """ class=""hand"" src=""../pagestyle/addimg.gif"" onClick=""VBS: addLogo '" & paramCode & "'"">"
 		  	writeCode ct&ct&cl&" End if " & cr
		  	writeCode ct&ct&"</div>"
		  	writeCode ct&ct&"<div id=""LbtnHide1_" & paramCode & """>"
 		  	writeCode ct&ct&cl&" if (HTProgRight and 8)<>0 then " & cr
		  	writeCode ct&ct&ct&"<img id=""chgLogo_" & paramCode & """ class=""hand"" src=""../pagestyle/chimg.gif"" onClick=""VBS: chgLogo '" & paramCode & "'"">"
 		  	writeCode ct&ct&cl&" End if " & cr
 		  	writeCode ct&ct&cl&" if (HTProgRight and 16)<>0 then " & cr
		  	writeCode ct&ct&ct&"<img id=""delLogo_" & paramCode & """ class=""hand"" src=""../pagestyle/delimg.gif"" onClick=""VBS: delLogo '" & paramCode & "'"">"
 		  	writeCode ct&ct&cl&" End if " & cr
		  	writeCode ct&ct&"</div>"
		  	writeCode ct&ct&"<div id=""LbtnHide2_" & paramCode & """>"
 		  	writeCode ct&ct&cl&" if (HTProgRight and 8)<>0 then " & cr
		  	writeCode ct&ct&ct&"<img id=""orgLogo_" & paramCode & """ class=""hand"" src=""../pagestyle/resetimg.gif"" onClick=""VBS: orgLogo '" & paramCode & "'"">"
 		  	writeCode ct&ct&cl&" End if " & cr
		  	writeCode ct&ct&"</div>"
			writeCode ct & "</td></tr></table>"
		  	if nullText(param.selectSingleNode("fieldDesc"))<>"" then
		  		writeCode "</br><span style=""color:red;"">" & nullText(param.selectSingleNode("fieldDesc")) & "</span>"
		  	end if
		   end if
		   
		  case "popDate"
		  	calendarFlag = true
	  		writeCode "<input name=""htx_"&paramCode&""" size=""8"" type=""hidden"">"
	  		writeCode "<input name=""pcShowhtx_"&paramCode&""" size=""8"" readonly onclick=""VBS: popCalendar 'htx_"&paramCode&"',''"& cq & fieldDesc & ">"
		  case "refCheckboxOther"
		  	refCode = param.selectSingleNode("refLookup").text
		  	tblRow = nullText(param.selectSingleNode("tblRow"))
		  	if tblRow="" then	tblRow = 6
		  	dynamicCode = ""
		  		for each dc in param.selectNodes("dynamicBehavior")
		  			dynamicCode = " " & dc.selectSingleNode("event").text & "=""" _
		  				& dc.selectSingleNode("pCall").text & """"
		  		next
		  	dynamicCode = replace(dynamicCode,chr(34),chr(34)&chr(34))
			writeCode "<table width=""100%"">"
			pdxc = ""
			tblRowCount = 0
			SQL = EnumerateCodeList(refCode)
'			writeCode HelpEnumerateCodeList(refCode, nullText(param.selectSingleNode("helpTextField")))
			SET RSS=conn.execute(SQL)
			While not RSS.EOF
				xTDtext = "<td"
				if nullText(param.selectSingleNode("helpTextField")) <> "" then _
					xTDtext = xTDtext & " title=""" & RSS(2) & cq
				xTDtext = xTDtext & " " 
				if uCase(RSS(0))="Z" then 	xTDtext = xTDtext & "colspan=2 "
				xTDtext = xTDtext & ">"
				if dynamicCode <> "" then _
					pdxc = replace(dynamicCode,"'mCode'","'"&RSS(0)&"'")
				if (tblRowCount mod tblRow) = 0 then  writeCode "<tr>"
				writeCode ct&xTDtext&"<input type=""checkbox"" name=""htx_"&paramCode&""" value="""&RSS(1)& """ " & pdxc& ">" _
		  			&"<font size=2>"&RSS(1)&"</font>"
				RSS.movenext
				tblRowCount = tblRowCount + 1
	    	wend
			writeCode ct&"<input type=""text"" name=""other_" &paramCode& """ size=10 onChange=""VBS: setOtherCheckbox 'htx_" _
				& paramCode & "',me.value"">"
			if nullText(param.selectSingleNode("btnSelectAll")) <> "" then
				writeCode cl&ct&ct&"tblRowCount = tblRowCount + 1"
				writeCode ct&ct&"if (tblRowCount mod 8) = 0 then  response.write ""<tr>"""
				writeCode ct&ct&"while (tblRowCount mod 8 ) < 7"
				writeCode ct&ct&ct&"response.write ""<TD>"""
				writeCode ct&ct&ct&"tblRowCount = tblRowCount + 1"
				writeCode ct&ct&"wend" & vbCrLf & cr
				writeCode ct&ct&"<TD><input type=button value=""全選"" name=""pickAll" _
					& paramCode & """ class=""cbutton"" onClick=""" _
					& nullText(param.selectSingleNode("btnSelectAll")) & """>"
			end if	    	
	    	writeCode "</table>"
		  case "refCheckbox"
		  	refCode = param.selectSingleNode("refLookup").text
		  	tblRow = nullText(param.selectSingleNode("tblRow"))
		  	if tblRow="" then	tblRow = nullText(param.selectSingleNode("cols"))
		  	if tblRow="" then	tblRow = 6
		  	dynamicCode = ""
		  		for each dc in param.selectNodes("dynamicBehavior")
		  			dynamicCode = " " & dc.selectSingleNode("event").text & "=""" _
		  				& dc.selectSingleNode("pCall").text & """"
		  		next
		  	dynamicCode = replace(dynamicCode,chr(34),chr(34)&chr(34))
			writeCode "<table width=""100%"">"
			pdxc = ""
			tblRowCount = 0
			SQL = EnumerateCodeList(refCode)
			SET RSS=conn.execute(SQL)
			While not RSS.EOF
				if dynamicCode <> "" then _
					pdxc = replace(dynamicCode,"'mCode'","'"&RSS(0)&"'")
				if (tblRowCount mod tblRow) = 0 then  writeCode "<tr>"
				writeCode ct&"<td><input type=""checkbox"" name=""htx_"&paramCode&""" value="""&RSS(0)& """ " & pdxc & ">" _
	  				&"<font size=2>"&RSS(1)&"</font>"
'---- 處理 [ ] XXX ( ___ 個) ----------------
'	  			if isNumeric(RSS(0)) then
'	  				writeCode ct&ct&"<font size=2>( <input name=""mhtx_"&paramCode&"value"&RSS(0)&""" size=3> 個 )</font>"
'	  			end if
'---------------------------------------------
				RSS.movenext
				tblRowCount = tblRowCount + 1
	    	wend
			if nullText(param.selectSingleNode("btnSelectAll")) <> "" then
				writeCode cl&ct&ct&"if (tblRowCount mod 8) = 0 then  response.write ""<tr>"""
				writeCode ct&ct&"while (tblRowCount mod 8 ) < 7"
				writeCode ct&ct&ct&"response.write ""<TD>"""
				writeCode ct&ct&ct&"tblRowCount = tblRowCount + 1"
				writeCode ct&ct&"wend" & vbCrLf & cr
				writeCode ct&ct&"<TD><input type=button value=""全選"" name=""pickAll" _
					& paramCode & """ class=""cbutton"" onClick=""" _
					& nullText(param.selectSingleNode("btnSelectAll")) & """>"
			end if	    	
	    	writeCode "</table>"
		  case "radio"
		  	dynamicCode = ""
		  		for each dc in param.selectNodes("dynamicBehavior")
		  			dynamicCode = " " & dc.selectSingleNode("event").text & "=""" _
		  				& dc.selectSingleNode("pCall").text & """"
		  		next
		  	set optionList = param.selectNodes("item")
			pdxc = ""
		  	for each optItem in optionList 
		  		if dynamicCode <> "" then _
		  			pdxc = replace(dynamicCode, "'mCode'", "'" & optItem.selectSingleNode("mCode").text & "'")
				writeCode "<input type=""radio"" name=""htx_"&paramCode&""" value="""&optItem.selectSingleNode("mCode").text & """ " & pdxc & ">"
		  		writeCode optItem.selectSingleNode("mValue").text & "&nbsp;&nbsp;"
		  	next
		  case "refRadio"
		  	refCode = param.selectSingleNode("refLookup").text
		  	dynamicCode = ""
		  		for each dc in param.selectNodes("dynamicBehavior")
		  			dynamicCode = " " & dc.selectSingleNode("event").text & "=""" _
		  				& dc.selectSingleNode("pCall").text & """"
		  		next
		  	dynamicCode = replace(dynamicCode,chr(34),chr(34)&chr(34))
			writeCode ct&ct&ct&cl&" pdxc = """"" & cr
			if dynamicCode <> "" then _
				writeCode ct&ct&ct&cl&" pdxc = replace("""& dynamicCode & """,""'mCode'"",""'""&RSS(0)&""'"")" & cr
'			writeCode EnumerateCodeList(refCode)
			uiEnumerateCodeList refCode, "radio", paramCode
'			writeCode "<input type=""radio"" name=""htx_"&paramCode&""" value="""&cl&"=RSS(0)"&cr & """ " & cl&"=pdxc"&cr & ">"
'	  		writeCode ct&ct&ct&cl&"=RSS(1)"&cr& "&nbsp;&nbsp;"
'			writeCode ct&ct&ct&cl&ct&"RSS.movenext"& VBCRLF & _
'	    		ct&ct&ct&"wend"& cr

		  case "refRadioOther"
		  	refCode = param.selectSingleNode("refLookup").text
		  	tblRow = param.selectSingleNode("tblRow").text
		  	dynamicCode = ""
		  		for each dc in param.selectNodes("dynamicBehavior")
		  			dynamicCode = " " & dc.selectSingleNode("event").text & "=""" _
		  				& dc.selectSingleNode("pCall").text & """"
		  		next
		  	dynamicCode = replace(dynamicCode,chr(34),chr(34)&chr(34))
			writeCode "<table width=""100%"">"
			writeCode ct&ct&ct&cl&" pdxc = """""
			writeCode ct&ct&ct&ct&" tblRowCount = 0" & cr
			writeCode HelpEnumerateCodeList(refCode, nullText(param.selectSingleNode("helpTextField")))
			xTDtext = "<td"
			if nullText(param.selectSingleNode("helpTextField")) <> "" then _
				xTDtext = xTDtext & " title=""" & cl & "=RSS(2)" & cr & cq
			xTDtext = xTDtext & ">"
			if dynamicCode <> "" then _
				writeCode ct&ct&ct&cl&" pdxc = replace("""& dynamicCode & """,""'mCode'"",""'""&RSS(0)&""'"")" & cr
			writeCode ct&ct&ct&cl& " if (tblRowCount mod " & tblRow & ") = 0 then  response.write ""<tr>""" & cr 
			writeCode ct&xTDtext&"<input type=""radio"" name=""htx_"&paramCode&""" value="""&cl&"=RSS(1)"&cr & """ " & cl&"=pdxc"&cr & ">" _
	  			&"<font size=2>"&cl&"=RSS(1)"&cr&"</font>"
			writeCode ct&ct&cl&ct&"RSS.movenext"& VBCRLF & _
				ct&ct&ct&"tblRowCount = tblRowCount + 1" & vbCrLf & _
	    		ct&ct&"wend"& cr
			writeCode ct&"<input type=""text"" name=""other_" &paramCode& """ size=10 onChange=""VBS: setOtherCheckbox 'htx_" _
				& paramCode & "',me.value"">"
	    	writeCode "</table>"
		  case "selection"
			writeCode "<Select name=""htx_"&paramCode&""" size=1>"
      		writeCode "<option value="""">請選擇</option>"
		  	set optionList = param.selectNodes("item")
		  	for each optItem in optionList 
				writeCode "<option value="""&optItem.selectSingleNode("mCode").text & """>" _
		  			& optItem.selectSingleNode("mValue").text & "</option>"
		  	next
			writeCode "</select>" & vbCRLF

		  case "refSelect"
		  	refCode = param.selectSingleNode("refLookup").text
		  	dynamicCode = ""
		  		for each dc in param.selectNodes("dynamicBehavior")
		  			dynamicCode = " " & dc.selectSingleNode("event").text & "=""" _
		  				& dc.selectSingleNode("pCall").text & """"
		  		next
'		  	response.write refCode
'		  	set optionList = rptXmlDoc.selectNodes("//dsTable[tableName='CodeMain']/instance/row[codeMetaID='"&refCode&"']")
			writeCode "<Select name=""htx_"&paramCode&""" size=1>"
      		writeCode "<option value="""">請選擇</option>"
			uiEnumerateCodeList refCode, "option", paramCode
'			writeCode ct&ct&ct&"<option value="""&cl&"=RSS(0)"&cr&""">"&cl&"=RSS(1)"&cr&"</option>"
'			writeCode ct&ct&ct&cl&ct&"RSS.movenext"& VBCRLF & _
'	    		ct&ct&ct&"wend"& cr
			writeCode "</select>" & vbCRLF
		  case "refSelectOther"
		  	refCode = param.selectSingleNode("refLookup").text
		  	dynamicCode = ""
		  		for each dc in param.selectNodes("dynamicBehavior")
		  			dynamicCode = " " & dc.selectSingleNode("event").text & "=""" _
		  				& dc.selectSingleNode("pCall").text & """"
		  		next
'		  	response.write refCode
'		  	set optionList = rptXmlDoc.selectNodes("//dsTable[tableName='CodeMain']/instance/row[codeMetaID='"&refCode&"']")
			writeCode "<Select name=""htx_"&paramCode&""" size=1>"
      		writeCode "<option value="""">請選擇</option>"
			uiEnumerateCodeList refCode, "option", paramCode
'			writeCode ct&ct&ct&"<option value="""&cl&"=RSS(0)"&cr&""">"&cl&"=RSS(1)"&cr&"</option>"
'			writeCode ct&ct&ct&cl&ct&"RSS.movenext"& VBCRLF & _
'	    		ct&ct&ct&"wend"& cr
			writeCode "</select>" & vbCRLF
		end select
end sub

sub processQParamField(param)
	paramCode = param.selectSingleNode("fieldName").text
	refField = paramCode
		paramType = nullText(param.selectSingleNode("inputType"))
		paramKind = nullText(param.selectSingleNode("paramKind"))
		fieldDesc= nullText(param.selectSingleNode("fieldDesc"))
		if trim(fieldDesc)="" then	fieldDesc=nullText(param.selectSingleNode("fieldLabel"))
		if fieldDesc <> "" then _
			fieldDesc = " title=""" & fieldDesc & cq & " tabindex=""" & xTabIndex & cq 
		paramSize= nullText(param.selectSingleNode("inputLen"))
		if paramSize = "" then paramSize = 50
		if paramSize > 50 then paramSize = 50
		select case paramKind
		 case "range"
		  	if paramType="popDate" then
			  	calendarFlag = true
		  		
				'writeCode "<input name=""pcShowhtx_"&paramCode&"S"" size="""&paramSize&""" xreadonly=""true"" onclick=""VBS: popCalendar 'htx_" & paramCode & "S','htx_"&paramCode&"E'"" onKeyPress=""VBS: popCalendar 'htx_"&paramCode&"S','htx_"&paramCode&"E'"& cq & fieldDesc & "/> ～ "
		  		
				writeCode "<input name=""htx_"&paramCode&"S"" type=""text"" maxlength='10' size='10' id='htx_Id_xpostDateS'/> ～ "
				
				fieldDesc= nullText(param.selectSingleNode("fieldDesc"))
				if trim(fieldDesc)="" then	fieldDesc=nullText(param.selectSingleNode("fieldLabel"))
				xTabIndex = xTabIndex + 1
				
				if fieldDesc <> "" then fieldDesc = " title=""" & fieldDesc &"迄" & cq & " tabindex=""" & xTabIndex & cq
		  		
				'writeCode "<input name=""pcShowhtx_"&paramCode&"E"" size="""&paramSize&""" xreadonly=""true"" onclick=""VBS: popCalendar 'htx_"&paramCode & "E',''"" onKeyPress=""VBS: popCalendar 'htx_"&paramCode&"E',''"& cq & fieldDesc & "/>"

				writeCode "<input name=""htx_" & paramCode & "E"" type=""text"" maxlength='10' size='10' id='htx_Id_xpostDateE'/>"
				
				writeCode "<script type='text/javascript'>"
				writeCode "$(document).ready(function() {"
				writeCode "$('#htx_Id_" & paramCode & "S').datepicker();"
				writeCode "$('#htx_Id_" & paramCode & "E').datepicker();"
				writeCode "})"
				writeCode "</script>"

		  	else
		  		writeCode "<input name=""htx_"&paramCode&"S"" size="""&paramSize& cq & fieldDesc & " /> ～ "
				fieldDesc= nullText(param.selectSingleNode("fieldDesc"))
				if trim(fieldDesc)="" then	fieldDesc=nullText(param.selectSingleNode("fieldLabel"))
				xTabIndex = xTabIndex + 1
				if fieldDesc <> "" then _
					fieldDesc = " title=""" & fieldDesc &"迄" & cq & " tabindex=""" & xTabIndex & cq
		  		writeCode "<input name=""htx_"&paramCode&"E"" size="""&paramSize& cq & fieldDesc & "/>"		  	
		  	end if
		 case else
		  select case paramType
		  case "refSelect"
		  	refCode = param.selectSingleNode("refLookup").text	  	
			writeCode "<select name=""htx_"&paramCode&""" size=""1" & cq & fieldDesc & ">"			
      		writeCode "<option value="""">請選擇</option>"
			uiEnumerateCodeList refCode, "option", paramCode
			writeCode "</select>" & vbCRLF		
		  case "selection"
			writeCode "<select name=""htx_"&paramCode&""" size=""1" & cq & fieldDesc & ">"
      		writeCode "<option value="""">請選擇</option>"
		  	set optionList = param.selectNodes("item")
		  	for each optItem in optionList 
				writeCode "<option value="""&optItem.selectSingleNode("mCode").text & """>" _
		  			& optItem.selectSingleNode("mValue").text & "</option>"
		  	next
			writeCode "</select>" & vbCRLF
		  case "radio"
		  	dynamicCode = ""
		  		for each dc in param.selectNodes("dynamicBehavior")
		  			dynamicCode = " " & dc.selectSingleNode("event").text & "=""" _
		  				& dc.selectSingleNode("pCall").text & """"
		  		next
		  	set optionList = param.selectNodes("item")
			pdxc = ""
		  	for each optItem in optionList 
		  		if dynamicCode <> "" then _
		  			pdxc = replace(dynamicCode, "'mCode'", "'" & optItem.selectSingleNode("mCode").text & "'")
				writeCode "<input type=""radio"" name=""htx_"&paramCode&""" value="""&optItem.selectSingleNode("mCode").text & """ " & pdxc & " />"
		  		writeCode optItem.selectSingleNode("mValue").text & "&amp;nbsp;&amp;nbsp;"
		  	next
		  case "refRadio"
		  	refCode = param.selectSingleNode("refLookup").text
		  	dynamicCode = ""
		  		for each dc in param.selectNodes("dynamicBehavior")
		  			dynamicCode = " " & dc.selectSingleNode("event").text & "=""" _
		  				& dc.selectSingleNode("pCall").text & """"
		  		next
		  	dynamicCode = replace(dynamicCode,chr(34),chr(34)&chr(34))
			pdxc = ""
			if dynamicCode <> "" then _
	  			pdxc = replace(dynamicCode, "'mCode'", "'" & RSS(0) & "'")
			uiEnumerateCodeList refCode, "radio", paramCode

		  case "refCheckbox"
		  	refCode = param.selectSingleNode("refLookup").text
		  	tblRow = nullText(param.selectSingleNode("tblRow"))
		  	if tblRow="" then	tblRow = 6
		  	dynamicCode = ""
		  		for each dc in param.selectNodes("dynamicBehavior")
		  			dynamicCode = " " & dc.selectSingleNode("event").text & "=""" _
		  				& dc.selectSingleNode("pCall").text & """"
		  		next
		  	dynamicCode = replace(dynamicCode,chr(34),chr(34)&chr(34))
			writeCode "<table width=""100%"">"
			pdxc = ""
			tblRowCount = 0
			SQL = EnumerateCodeList(refCode)
			SET RSS=conn.execute(SQL)
			While not RSS.EOF
				if dynamicCode <> "" then _
					pdxc = replace(dynamicCode,"'mCode'","'"&RSS(0)&"'")
				if (tblRowCount mod tblRow) = 0 then  
					if tblRowcount<>0 then		writeCode "</tr>"
					writeCode "<tr>"
				end if
				writeCode ct&"<td><input type=""checkbox"" name=""htx_"&paramCode&""" value="""&RSS(0)& """ " & pdxc & " />" _
	  				& RSS(1)&"</td>"
				RSS.movenext
				tblRowCount = tblRowCount + 1
	    	wend
			if nullText(param.selectSingleNode("btnSelectAll")) <> "" then
				if (tblRowCount mod 8) = 0 then  response.write "<tr>"
				while (tblRowCount mod 8 ) < 7
					response.write "<td>"
					tblRowCount = tblRowCount + 1
				wend
				writeCode "<td><input type=button value=""全選"" name=""pickAll" _
					& paramCode & """ class=""cbutton"" onClick=""" _
					& nullText(param.selectSingleNode("btnSelectAll")) & """ />"
			end if	    	
	    	writeCode "</tr></table>"
		  case "ptypeselect"
		  	writeCode "<Select name=""htx_"&paramCode&""" size=""1"">"
			response.write "<" & "!--#include file=""../xml/pType.inc""--" & ">"
			writeCode "</select>" & vbCRLF		    		
		  case else
		  	writeCode "<input name=""htx_"&paramCode&""" size="""&paramSize& cq & fieldDesc & "/>"
		 end select
		end select	
end sub

%>
