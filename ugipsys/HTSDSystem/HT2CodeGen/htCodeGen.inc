<% 
sub dumpTempFile()
    Do While Not xfin.AtEndOfStream
        xinStr = xfin.readline
        xfout.writeline xinStr
    Loop
end sub

sub writeCode(xstr)
	response.write xstr
	xfout.writeline xstr
end sub

sub writePart(xstr)
	response.write xstr
	xfout.write xstr
end sub

sub showAidLinkList
	for each x in htPageDom.selectNodes("//pageSpec/aidLinkList/Anchor")
		ckRight = nullText(x.selectSingleNode("checkRight"))
		pKeyStr = ""
		if nullText(x.selectSingleNode("AnchorWPK"))="Y" then _
			pKeyStr = "?" & cl & "=pKey" & cr
		if ckRight <> "" then _
			xfout.writeLine ct&ct & cl & "if (HTProgRight and " & ckRight & ")=" & ckRight & " then" & cr
		if x.selectSingleNode("AnchorType").text="Back" then
			xfout.writeLine ct&ct&ct & "<A href=""Javascript:window.history.back();" _
				& """ title=""" & nullText(x.selectSingleNode("AnchorDesc")) & """>" _
				& x.selectSingleNode("AnchorLabel").text & "</A> "
		else
			xfout.writeLine ct&ct&ct & "<A href=""" & x.selectSingleNode("AnchorURI").text _
				& pKeyStr & """ title=""" & nullText(x.selectSingleNode("AnchorDesc")) & """>" _
				& x.selectSingleNode("AnchorLabel").text & "</A>"
		end if
		if ckRight <> "" then _
			xfout.writeLine ct&ct & cl & "end if" & cr
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
	    	str=ct&ct&ct&cl & "SQL=""Select " & RSLK("CodeValueFld") & "," & RSLK("CodeDisplayFld") & " from " & RSLK("CodeTblName") & """" & VBCRLF & _
		    	ct&ct&ct&"SET RSS=conn.execute(SQL)" & VBCRLF & _
		    	ct&ct&ct&"While not RSS.EOF"&cr& VBCRLF
		else	
	    	str=ct&ct&ct&cl & "SQL=""Select " & RSLK("CodeValueFld") & "," & RSLK("CodeDisplayFld") & " from " & RSLK("CodeTblName") & " where " & RSLK("CodeSrcFld") & "=N'" & RSLK("CodeSrcItem") & "'""" & VBCRLF & _
		    	ct&ct&ct&"SET RSS=conn.execute(SQL)" & VBCRLF & _
			    ct&ct&ct&"While not RSS.EOF"&cr& VBCRLF
		end if          
	  else
		if isNull(RSLK("CodeSrcFld")) then
	    	str=ct&ct&ct&cl & "SQL=""Select " & RSLK("CodeValueFld") & "," & RSLK("CodeDisplayFld") & " from " & RSLK("CodeTblName") & " where " & RSLK("CodeSortFld") & " IS NOT NULL Order by " & RSLK("CodeSortFld") & """" & VBCRLF & _
		    	ct&ct&ct&"SET RSS=conn.execute(SQL)" & VBCRLF & _
			    ct&ct&ct&"While not RSS.EOF"&cr& VBCRLF
	    else	
	    	str=ct&ct&ct&cl & "SQL=""Select " & RSLK("CodeValueFld") & "," & RSLK("CodeDisplayFld") & " from " & RSLK("CodeTblName") & " where " & RSLK("CodeSortFld") & " IS NOT NULL AND " & RSLK("CodeSrcFld") & "=N'" & RSLK("CodeSrcItem") & "' Order by " & RSLK("CodeSortFld") & """" & VBCRLF & _
		    	ct&ct&ct&"SET RSS=conn.execute(SQL)" & VBCRLF & _
			    ct&ct&ct&"While not RSS.EOF"&cr& VBCRLF
		end if
	  end if
	end if
	EnumerateCodeList=str
End Function

Function EnumerateCodeRefLookup(codeID)
	SQL="Select * from CodeMetaDef where codeID=N'" & codeID & "'"
        SET RSLK=conn.execute(SQL)
	str=""
	if not RSLK.EOF then
		if isNull(RSLK("CodeSrcFld")) then
	    	str=cl & ct & "SQL=""Select " & RSLK("CodeValueFld") & "," & RSLK("CodeDisplayFld") & " from " & RSLK("CodeTblName") & " WHERE " & RSLK("CodeValueFld") & "= """
		else	
	    	str=cl & ct & "SQL=""Select " & RSLK("CodeValueFld") & "," & RSLK("CodeDisplayFld") & " from " & RSLK("CodeTblName") & " where " & RSLK("CodeSrcFld") & "=N'" & RSLK("CodeSrcItem") & "' AND " & RSLK("CodeValueFld") & "= """
		end if          
	end if
	EnumerateCodeRefLookup=str
End Function

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

sub processInsert(xstr)
	inPos = instr(xstr,"/")
	refTable = left(xstr,inPos-1)
	refField = mid(xstr,inPos+1)
	if nullText(refModel.selectSingleNode("fieldList[tableName='" & refTable & "']/fkLink/asAlias")) <> ""	then	exit sub
	set param = refModel.selectSingleNode("fieldList[tableName='" & refTable & "']/field[fieldName='" & refField & "']")
	paramType = param.selectSingleNode("dataType").text

	if lcase(right(nullText(param.selectSingleNode("inputType")),4)) = "file"	then	exit sub

	xfout.writeLine ct&"IF xUpForm(""htx_" & refField & """) <> """" Then"
	xfout.writeLine ct&ct&"sql = sql & """ & refField & """ & "","""
	xfout.writeLine ct&ct&"sqlValue = sqlValue & pkStr(xUpForm(""htx_" & refField & """),"","")"
	xfout.writeLine ct&"END IF"
end sub

sub processInsertTable (xstr, thisTable)
	inPos = instr(xstr,"/")
	refTable = left(xstr,inPos-1)
	if refTable <> thisTable then	exit sub
	refField = mid(xstr,inPos+1)
'	if nullText(refModel.selectSingleNode("fieldList[tableName='" & refTable & "']/fkLink/asAlias")) <> ""	then	exit sub
	set param = refModel.selectSingleNode("fieldList[tableName='" & refTable & "']/field[fieldName='" & refField & "']")
	paramType = param.selectSingleNode("dataType").text

	if lcase(right(nullText(param.selectSingleNode("inputType")),4)) = "file"	then	exit sub

	xfout.writeLine ct&"IF xUpForm(""htx_" & refField & """) <> """" Then"
	xfout.writeLine ct&ct&"sql = sql & """ & refField & """ & "","""
	xfout.writeLine ct&ct&"sqlValue = sqlValue & pkStr(xUpForm(""htx_" & refField & """),"","")"
	xfout.writeLine ct&"END IF"
end sub

sub processUpdateTable (xstr, thisTable)
	inPos = instr(xstr,"/")
	refTable = left(xstr,inPos-1)
	if refTable <> thisTable then	exit sub
	refField = mid(xstr,inPos+1)
'	if nullText(refModel.selectSingleNode("fieldList[tableName='" & refTable & "']/fkLink/asAlias")) <> ""	then	exit sub

	set param = refModel.selectSingleNode("fieldList[tableName='" & refTable & "']/field[fieldName='" & refField & "']")
	if nullText(param.selectSingleNode("mpKey")) = "Y" then	exit sub
	if nullText(param.selectSingleNode("identity")) = "Y" then	exit sub
	if nullText(param.selectSingleNode("inputType")) = "calc" then	exit sub
	paramType = param.selectSingleNode("dataType").text

	if nullText(param.selectSingleNode("inputType")) = "imgfile"	then
		processImgFile xstr
		exit sub
	elseif nullText(param.selectSingleNode("inputType")) = "file"	then
		processAttFile xstr
		exit sub
	end if

'	xfout.writeLine ct&"IF request(""htx_" & refField & """) <> """" Then"
	select case paramType
	  case "calc"
	  case "integer"
			xfout.writeLine ct&ct&"sql = sql & """ & refField & "="" & " & "drn(""htx_" & refField & """)"
	  case else
			xfout.writeLine ct&ct&"sql = sql & """ & refField & "="" & " & "pkStr(xUpForm(""htx_" & refField & """),"","")"
	end select
'	xfout.writeLine ct&"END IF"
end sub

sub processUpdate(xstr)
	inPos = instr(xstr,"/")
	refTable = left(xstr,inPos-1)
	refField = mid(xstr,inPos+1)
	if nullText(refModel.selectSingleNode("fieldList[tableName='" & refTable & "']/fkLink/asAlias")) <> ""	then	exit sub

	set param = refModel.selectSingleNode("fieldList[tableName='" & refTable & "']/field[fieldName='" & refField & "']")
	if nullText(param.selectSingleNode("mpKey")) = "Y" then	exit sub
	if nullText(param.selectSingleNode("identity")) = "Y" then	exit sub
	if nullText(param.selectSingleNode("inputType")) = "calc" then	exit sub
	paramType = param.selectSingleNode("dataType").text

	if nullText(param.selectSingleNode("inputType")) = "imgfile"	then
		processImgFile xstr
		exit sub
	elseif nullText(param.selectSingleNode("inputType")) = "file"	then
		processAttFile xstr
		exit sub
	end if

'	xfout.writeLine ct&"IF request(""htx_" & refField & """) <> """" Then"
	select case paramType
	  case "calc"
	  case "integer"
			xfout.writeLine ct&ct&"sql = sql & """ & refField & "="" & " & "drn(""htx_" & refField & """)"
	  case else
			xfout.writeLine ct&ct&"sql = sql & """ & refField & "="" & " & "pkStr(xUpForm(""htx_" & refField & """),"","")"
	end select
'	xfout.writeLine ct&"END IF"
end sub

sub processImgFile(xstr)
'	response.write "<HR>" & xstr & "<HR>"
	inPos = instr(xstr,"/")
	refTable = left(xstr,inPos-1)
	refField = mid(xstr,inPos+1)
	set param = refModel.selectSingleNode("fieldList[tableName='" & refTable & "']/field[fieldName='" & refField & "']")
	inputType = param.selectSingleNode("inputType").text

	xfout.writeLine ct&"IF xUpForm(""htImgActCK_" & refField & """) <> """" Then"
	xfout.writeLine ct&"  actCK = xUpForm(""htImgActCK_" & refField & """)"
	xfout.writeLine ct&"  if actCK=""editLogo"" OR actCK=""addLogo"" then"
	xfout.writeLine ct&ct&"fname = """""
	xfout.writeLine ct&ct&"For each xatt in xup.Attachments"
	xfout.writeLine ct&ct&"  if xatt.Name = ""htImg_" & refField & """ then"
	xfout.writeLine ct&ct&ct&"ofname = xatt.FileName"
	xfout.writeLine ct&ct&ct&"fnExt = """""
	xfout.writeLine ct&ct&ct&"if instr(ofname, ""."")>0 then	fnext=mid(ofname, instr(ofname, "".""))"
	xfout.writeLine ct&ct&ct&"tstr = now()"
	xfout.writeLine ct&ct&ct&"nfname = right(year(tstr),1) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & cint(rnd(second(tstr))*100) & fnext"
	xfout.writeLine ct&ct&ct&"sql = sql & """ & refField & "="" & " & "pkStr(nfname,"","")"
	xfout.writeLine ct&ct&ct&"IF xUpForm(""hoImg_" & refField & """) <> """" Then _"
	xfout.writeLine ct&ct&ct&ct&"xup.DeleteFile apath & xUpForm(""hoImg_" & refField & """)"
	xfout.writeLine ct&ct&ct&"xatt.SaveFile apath & nfname, false"
	xfout.writeLine ct&ct&"  end if"
	xfout.writeLine ct&ct&"Next"
	xfout.writeLine ct&"  elseif actCK=""delLogo"" then"
	xfout.writeLine ct&ct&"xup.DeleteFile apath & xUpForm(""hoImg_" & refField & """)"
	xfout.writeLine ct&ct&"sql = sql & """ & refField & "=null,"""
	xfout.writeLine ct&"  end if"
	xfout.writeLine ct&"END IF"

end sub

sub processAttFile(xstr)
'	response.write "<HR>" & xstr & "<HR>"
	inPos = instr(xstr,"/")
	refTable = left(xstr,inPos-1)
	refField = mid(xstr,inPos+1)
	set param = refModel.selectSingleNode("fieldList[tableName='" & refTable & "']/field[fieldName='" & refField & "']")
	inputType = param.selectSingleNode("inputType").text

	xfout.writeLine ct&"IF xUpForm(""htFileActCK_" & refField & """) <> """" Then"
	xfout.writeLine ct&"  actCK = xUpForm(""htFileActCK_" & refField & """)"
	xfout.writeLine ct&"  if actCK=""editLogo"" OR actCK=""addLogo"" then"
	xfout.writeLine ct&ct&"fname = """""
	xfout.writeLine ct&ct&"For each xatt in xup.Attachments"
	xfout.writeLine ct&ct&"  if xatt.Name = ""htFile_" & refField & """ then"
	xfout.writeLine ct&ct&ct&"ofname = xatt.FileName"
	xfout.writeLine ct&ct&ct&"fnExt = """""
	xfout.writeLine ct&ct&ct&"if instr(ofname, ""."")>0 then	fnext=mid(ofname, instr(ofname, "".""))"
	xfout.writeLine ct&ct&ct&"tstr = now()"
	xfout.writeLine ct&ct&ct&"nfname = right(year(tstr),1) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & cint(rnd(second(tstr))*100) & fnext"
	xfout.writeLine ct&ct&ct&"sql = sql & """ & refField & "="" & " & "pkStr(nfname,"","")"
	xfout.writeLine ct&ct&ct&"IF xUpForm(""hoFile_" & refField & """) <> """" Then"
	xfout.writeLine ct&ct&ct&ct&"if xup.IsFileExist( apath & xUpForm(""hoFile_" & refField & """)) then _"
	xfout.writeLine ct&ct&ct&ct&ct&"xup.DeleteFile apath & xUpForm(""hoFile_" & refField & """)"
	xfout.writeLine ct&ct&ct&ct&"xsql = ""DELETE imageFile WHERE newFileName="" & pkStr(xUpForm(""hoFile_" & refField & """),"""")"
	xfout.writeLine ct&ct&ct&ct&"conn.execute xsql"
	xfout.writeLine ct&ct&ct&"END IF"
	xfout.writeLine ct&ct&ct&"xatt.SaveFile apath & nfname, false"
	xfout.writeLine ct&ct&ct&"xsql = ""INSERT INTO imageFile(newFileName, oldFileName) VALUES("" & pkStr(nfname,"","") & pkStr(ofname,"")"")"
	xfout.writeLine ct&ct&ct&"conn.execute xsql"
	xfout.writeLine ct&ct&"  end if"
	xfout.writeLine ct&ct&"Next"
	xfout.writeLine ct&"  elseif actCK=""delLogo"" then"
	xfout.writeLine ct&ct&"if xup.IsFileExist( apath & xUpForm(""hoFile_" & refField & """)) then _"
	xfout.writeLine ct&ct&ct&"xup.DeleteFile apath & xUpForm(""hoFile_" & refField & """)"
	xfout.writeLine ct&ct&"xsql = ""DELETE imageFile WHERE newFileName="" & pkStr(xUpForm(""hoFile_" & refField & """),"""")"
	xfout.writeLine ct&ct&"conn.execute xsql"
	xfout.writeLine ct&ct&"sql = sql & """ & refField & "=null,"""
	xfout.writeLine ct&"  end if"
	xfout.writeLine ct&"END IF"

end sub

sub processValid(xstr)
	inPos = instr(xstr,"/")
	refTable = left(xstr,inPos-1)
	refField = mid(xstr,inPos+1)
	set param = refModel.selectSingleNode("fieldList[tableName='" & refTable & "']/field[fieldName='" & refField & "']")
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
	  case "hidden"
	  case "file"
		xfout.writeLine ct&"IF reg.htfile_" & refField & ".value = Empty Then "
		xfout.writeLine ct&ct&"MsgBox replace(nMsg,""{0}""," _
			& cq & param.selectSingleNode("fieldLabel").text & """), 64, ""Sorry!"""
		xfout.writeLine ct&ct&"reg.htfile_" & refField & ".focus"
		xfout.writeLine ct&ct&"exit sub"
		xfout.writeLine ct&"END IF"
	  case else
		xfout.writeLine ct&"IF reg.htx_" & refField & ".value = Empty Then "
		xfout.writeLine ct&ct&"MsgBox replace(nMsg,""{0}""," _
			& cq & param.selectSingleNode("fieldLabel").text & """), 64, ""Sorry!"""
		xfout.writeLine ct&ct&"reg.htx_" & refField & ".focus"
		xfout.writeLine ct&ct&"exit sub"
		xfout.writeLine ct&"END IF"
	 end select
	end if
	if it="password" then
		xfout.writeLine ct&"IF reg.htx_" & refField & ".value <> " _
			& " reg.htx_" & refField & "2check.value  Then "
		xfout.writeLine ct&ct&"MsgBox ""密碼欄與確認欄的輸入資料需相同"", 64, ""Sorry!"""
		xfout.writeLine ct&ct&"reg.htx_" & refField & ".focus"
		xfout.writeLine ct&ct&"exit sub"
		xfout.writeLine ct&"END IF"
	end if
	dt = nullText(param.selectSingleNode("dataType"))
	if it = "imgfile" then
			xfout.writeLine ct&"IF reg.htImg_" & refField & ".value <> """" Then" 
			xfout.writeLine ct&ct&"xIMGname = reg.htImg_" & refField & ".value"
			xfout.writeLine ct&ct&"xFileType = """""
			xfout.writeLine ct&ct&"if instr(xIMGname, ""."")>0 then	xFileType=lcase(mid(xIMGname, instr(xIMGname, ""."")+1))"
			xfout.writeLine ct&ct&"IF xFileType<>""gif"" AND xFileType<>""jpg"" AND xFileType<>""jpeg"" then"
			xfout.writeLine ct&ct&ct&"MsgBox replace(pMsg,""{0}""," _
				& cq & param.selectSingleNode("fieldLabel").text & """), 64, ""Sorry!"""
			xfout.writeLine ct&ct&ct&"reg.htImg_" & refField & ".focus"
			xfout.writeLine ct&ct&ct&"exit sub"
			xfout.writeLine ct&ct&"END IF"
			xfout.writeLine ct&"END IF"
	elseif lcase(right(dt,4)) = "char" then
		if l<>"" AND (it="" OR it="textbox" OR it="textarea") then
			xfout.writeLine ct&"IF blen(reg.htx_" & refField & ".value) > " & l & " Then"
			xfout.writeLine ct&ct&"MsgBox replace(replace(lMsg,""{0}""," _
				& cq & param.selectSingleNode("fieldLabel").text & """),""{1}"","""&l&"""), 64, ""Sorry!"""
			xfout.writeLine ct&ct&"reg.htx_" & refField & ".focus"
			xfout.writeLine ct&ct&"exit sub"
			xfout.writeLine ct&"END IF"
		end if
	elseif lcase(right(dt,8)) = "datetime" then
			xfout.writeLine ct&"IF (reg.htx_" & refField & ".value <> """") AND (NOT isDate(reg.htx_" & refField & ".value)) Then" 
			xfout.writeLine ct&ct&"MsgBox replace(dMsg,""{0}""," _
				& cq & param.selectSingleNode("fieldLabel").text & """), 64, ""Sorry!"""
			xfout.writeLine ct&ct&"reg.htx_" & refField & ".focus"
			xfout.writeLine ct&ct&"exit sub"
			xfout.writeLine ct&"END IF"
	elseif lcase(left(dt,3)) = "int" OR dt="real" then
			xfout.writeLine ct&"IF (reg.htx_" & refField & ".value <> """") AND (NOT isNumeric(reg.htx_" & refField & ".value)) Then" 
			xfout.writeLine ct&ct&"MsgBox replace(iMsg,""{0}""," _
				& cq & param.selectSingleNode("fieldLabel").text & """), 64, ""Sorry!"""
			xfout.writeLine ct&ct&"reg.htx_" & refField & ".focus"
			xfout.writeLine ct&ct&"exit sub"
			xfout.writeLine ct&"END IF"
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

sub AddProcessInit(xstr)
	inPos = instr(xstr,"/")
	refTable = left(xstr,inPos-1)
	refField = mid(xstr,inPos+1)
	set param = refModel.selectSingleNode("fieldList[tableName='" & refTable & "']/field[fieldName='" & refField & "']")
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
	  	rhs = cq & cl & "=" & param.selectSingleNode("clientDefault/set").text & cr & cq
	  case "session"
	  	rhs = cq & cl & "=session(" & param.selectSingleNode("clientDefault/set").text & ")" & cr & cq
	end select
  	xfout.writeLine lhs & rhs
	if nullText(param.selectSingleNode("inputType"))="popDate" then
		lhs = ct&"reg.htx_" & refField & ".value= "
	  	writeCode lhs & rhs
	end if
end sub

sub EditProcessInit(xstr)
	inPos = instr(xstr,"/")
	refTable = left(xstr,inPos-1)
	refField = mid(xstr,inPos+1)
	set param = refModel.selectSingleNode("fieldList[tableName='" & refTable & "']/field[fieldName='" & refField & "']")
	if nullText(param.selectSingleNode("inputType")) <> "calc" then
	  select case nullText(param.selectSingleNode("inputType"))
	  	case "refRadioOther"
			xfout.writeLine ct & "initOtherRadio ""htx_" & refField & """," & cq & cl & "=qqRS(""" & refField & """)" & cr & cq _
				& ", ""other_" & refField & cq	
	  	case "refRadio"
			xfout.writeLine ct & "initRadio ""htx_" & refField & """," & cq & cl & "=qqRS(""" & refField & """)" & cr & cq 
	  	case "radio"
			xfout.writeLine ct & "initRadio ""htx_" & refField & """," & cq & cl & "=qqRS(""" & refField & """)" & cr & cq 
	  	case "refCheckboxOther"
			xfout.writeLine ct & "initOtherCheckbox ""htx_" & refField & """," & cq & cl & "=qqRS(""" & refField & """)" & cr & cq _
				& ", ""other_" & refField & cq	
	  	case "refCheckbox"
			xfout.writeLine ct & "initCheckbox ""htx_" & refField & """," & cq & cl & "=qqRS(""" & refField & """)" & cr & cq 
	  	case "checkbox"
			xfout.writeLine ct & "initCheckbox ""htx_" & refField & """," & cq & cl & "=qqRS(""" & refField & """)" & cr & cq 
	  	case "calc"
	  		xfout.writeLine ct & param.selectSingleNode("calc").text
		case "readonlydefault"	  	
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
  			xfout.writeLine lhs & rhs
  			xfout.writeLine	ct&"reg.htx_" & refField & ".readonly=""true"""	
			xfout.writeLine	ct&"reg.htx_" & refField & ".classname=""rdonly""" 	  		
	  	case "file"
			xfout.writeLine ct & "initAttFile """ & refField & """, " & cq & cl & "=qqRS(""" & refField & """)" & cr _
				& """, " & cq & cl & "=qqRS(""fxr_" & refField & """)" & cr & cq 
'			lhs = ct&"reg.hoFile_" & refField & ".value= "
'			rhs = cq & cl & "=qqRS(""fxr_" & refField & """)" & cr & cq
'  			xfout.writeLine lhs & rhs
	  	case "imgfile"
			xfout.writeLine ct & "initImgFile """ & refField & """," & cq & cl & "=qqRS(""" & refField & """)" & cr & cq 
'			lhs = ct&"document.all(""imgSrc_" & refField & """).src= "
'			rhs = cq & cl & "=HTUploadPath" & cr & cl & "=qqRS(""" & refField & """)" & cr & cq
'  			xfout.writeLine lhs & rhs
	  	case "password"
			lhs = ct&"reg.htx_" & refField & ".value= "
			rhs = cq & cl & "=qqRS(""" & refField & """)" & cr & cq
  			xfout.writeLine lhs & rhs
			lhs = ct&"reg.htx_" & refField & "2check.value= "
			rhs = cq & cl & "=qqRS(""" & refField & """)" & cr & cq
  			xfout.writeLine lhs & rhs
	  	case "popDate"
			lhs = ct&"reg.htx_" & refField & ".value= "
			rhs = cq & cl & "=qqRS(""" & refField & """)" & cr & cq
  			xfout.writeLine lhs & rhs
			lhs = ct&"reg.pcShowhtx_" & refField & ".value= "
			rhs = "d7Date(" & cq& cl & "=qqRS(""" & refField & """)" & cr & cq & ")"
  			xfout.writeLine lhs & rhs
	  	case else
	  	  if nullText(param.selectSingleNode("dataType")) = "datetime" then
			lhs = ct&"reg.htx_" & refField & ".value= "
			rhs = "d7Date(" & cq & cl & "=qqRS(""" & refField & """)" & cr & cq & ")"
	  	  else
			lhs = ct&"reg.htx_" & refField & ".value= "
			rhs = cq & cl & "=qqRS(""" & refField & """)" & cr & cq
  		  end if
  			xfout.writeLine lhs & rhs
  	  end select
  	end if
	if nullText(param.selectSingleNode("calc")) <> "" then _
  		xfout.writeLine ct & param.selectSingleNode("calc").text
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
  	xfout.writeLine lhs & rhs
end sub

sub processAnchor(xstr)
'	response.write "anchorList/anchor[funcLabel='" & xstr & "']<HR>"
	set param = refModel.selectSingleNode("anchorList/anchor[funcLabel='" & xstr & "']")
	select case nullText(param.selectSingleNode("type"))
		case "button"
			xfout.writeLine ct & "<INPUT TYPE=button VALUE=""" & xstr & """ onClick=""" _
				& param.selectSingleNode("action").text & "('" & param.selectSingleNode("url").text _
				& "?" & cl & "=pKey" & cr & "')"">"
	end select
end sub

sub processParam(xstr)
'	response.write xstr
	inPos = instr(xstr,"/")
	refTable = left(xstr,inPos-1)
	refField = mid(xstr,inPos+1)
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
		if fieldDesc <> "" then _
			fieldDesc = " title=""" & fieldDesc & cq
		if paramSize = "" then paramSize = 10
		if paramSize > 50 then paramSize = 50
'			writeCode nullText(param.selectSingleNode("inputType"))
		select case nullText(param.selectSingleNode("inputType"))
		  case ""
		  	writeCode "<input name=""htx_"&paramCode&""" size="""&paramSize& cq & fieldDesc & ">"
		  case "textbox"
		  	writeCode "<input name=""htx_"&paramCode&""" size="""&paramSize& cq & fieldDesc & ">"
		  case "password"
		  	writeCode "<input name=""htx_"&paramCode&""" type=""password"" size="""&paramSize& cq & fieldDesc & ">"
		  	writeCode " 密碼確認："
			fieldDesc = " title=""再輸入一次密碼供確認檢查" & fieldDesc & cq
		  	writeCode "<input name=""htx_"&paramCode&"2check"" type=""password"" size="""&paramSize& cq & fieldDesc & ">"
		  case "calc"
		  	writeCode "<input name=""htx_"&paramCode&""" size="""&paramSize&""" readonly=""true"">"
		  case "readonly"
		  	refCode = nullText(param.selectSingleNode("refLookup"))
		  	if refCode = "" then 
		  		writeCode "<input name=""htx_"&paramCode&""" size="""&paramSize&""" readonly=""true"" class=""rdonly"">"
		  	else
		  		xRS = "RSMaster"
		  		if formFunction = "edit" then	xRS = "RSreg"
			  	writeCode "<input type=""hidden"" name=""htx_"&paramCode&""">"
				writeCode EnumerateCodeRefLookup(refCode) & " & pkStr(" & xRS & "(""" _
					& paramCode&"""),"""")"
				writeCode ct&"SET RSS=conn.execute(SQL)"
				writeCode ct&"if not RSS.eof then"
				writeCode ct&ct&"response.write RSS(1)"
				writeCode ct&"end if" & cr
		  	end if
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
		   end if
		  case "imgfile"
		   If formFunction = "add" then
		  	writeCode "<input type=""file"" name=""htImg_" & paramCode & """>"
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
		   end if
		  case "popDate"
		  	calendarFlag = true
		  	followStr = nullText(param.selectSingleNode("followChange"))
		  	if followStr <> "" then	followStr = "htx_" & followStr	
		  	followStr = " ,'" & followStr & "'"
	  		writeCode "<input name=""htx_"&paramCode&""" type=hidden>"
	  		writeCode "<input name=""pcShowhtx_"&paramCode&""" size="""&paramSize&""" readonly onclick=""VBS: popCalendar 'htx_"&paramCode&"'"& followStr & cq & fieldDesc & ">"
		  case "refCheckboxOther"
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
			xTDtext = xTDtext & " " & cl & " if uCase(RSS(0))=""Z"" then response.write ""colspan=2 """ & cr
			xTDtext = xTDtext & ">"
			if dynamicCode <> "" then _
				writeCode ct&ct&ct&cl&" pdxc = replace("""& dynamicCode & """,""'mCode'"",""'""&RSS(0)&""'"")" & cr
			writeCode ct&ct&ct&cl& " if (tblRowCount mod " & tblRow & ") = 0 then  response.write ""<tr>""" & cr 
			writeCode ct&xTDtext&"<input type=""checkbox"" name=""htx_"&paramCode&""" value="""&cl&"=RSS(1)"&cr & """ " & cl&"=pdxc"&cr & ">" _
	  			&"<font size=2>"&cl&"=RSS(1)"&cr&"</font>"
			writeCode ct&ct&cl&ct&"RSS.movenext"& VBCRLF & _
				ct&ct&ct&"tblRowCount = tblRowCount + 1" & vbCrLf & _
	    		ct&ct&"wend"& cr
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
			writeCode EnumerateCodeList(refCode)
			if dynamicCode <> "" then _
				writeCode ct&ct&ct&cl&" pdxc = replace("""& dynamicCode & """,""'mCode'"",""'""&RSS(0)&""'"")" & cr
			writeCode ct&ct&ct&cl& " if (tblRowCount mod " & tblRow & ") = 0 then  response.write ""<tr>""" & cr 
			writeCode ct&"<td><input type=""checkbox"" name=""htx_"&paramCode&""" value="""&cl&"=RSS(0)"&cr & """ " & cl&"=pdxc"&cr & ">" _
	  			&"<font size=2>"&cl&"=RSS(1)"&cr&"</font>"
			writeCode ct&ct&cl&ct&"RSS.movenext"& VBCRLF & _
				ct&ct&ct&"tblRowCount = tblRowCount + 1" & vbCrLf & _
	    		ct&ct&"wend"& cr
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
			writeCode EnumerateCodeList(refCode)
			if dynamicCode <> "" then _
				writeCode ct&ct&ct&cl&" pdxc = replace("""& dynamicCode & """,""'mCode'"",""'""&RSS(0)&""'"")" & cr
			writeCode "<input type=""radio"" name=""htx_"&paramCode&""" value="""&cl&"=RSS(0)"&cr & """ " & cl&"=pdxc"&cr & ">"
	  		writeCode ct&ct&ct&cl&"=RSS(1)"&cr& "&nbsp;&nbsp;"
			writeCode ct&ct&ct&cl&ct&"RSS.movenext"& VBCRLF & _
	    		ct&ct&ct&"wend"& cr

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
			writeCode EnumerateCodeList(refCode)
			writeCode ct&ct&ct&"<option value="""&cl&"=RSS(0)"&cr&""">"&cl&"=RSS(1)"&cr&"</option>"
			writeCode ct&ct&ct&cl&ct&"RSS.movenext"& VBCRLF & _
	    		ct&ct&ct&"wend"& cr
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
			writeCode EnumerateCodeList(refCode)
			writeCode ct&ct&ct&"<option value="""&cl&"=RSS(0)"&cr&""">"&cl&"=RSS(1)"&cr&"</option>"
			writeCode ct&ct&ct&cl&ct&"RSS.movenext"& VBCRLF & _
	    		ct&ct&ct&"wend"& cr
			writeCode "</select>" & vbCRLF
		end select
end sub

sub processQParamField(param)
	paramCode = param.selectSingleNode("fieldName").text
	refField = paramCode
		paramType = nullText(param.selectSingleNode("inputType"))
		paramKind = nullText(param.selectSingleNode("paramKind"))
		paramSize= nullText(param.selectSingleNode("inputLen"))
		if paramSize = "" then paramSize = 10
		if paramSize > 50 then paramSize = 50
		select case paramKind
		 case "range"
		  	if paramType="popDate" then
			  	calendarFlag = true
		  		writeCode "<input name=""pcShowhtx_"&paramCode&"S"" size="""&paramSize&""" readonly onclick=""VBS: popCalendar 'htx_"&paramCode&"S','htx_"&paramCode&"E'""> ～ "
		  		writeCode "<input name=""htx_"&paramCode&"S"" type=hidden><input name=""htx_"&paramCode&"E"" type=hidden>"
		  		writeCode "<input name=""pcShowhtx_"&paramCode&"E"" size="""&paramSize&""" readonly onclick=""VBS: popCalendar 'htx_"&paramCode&"E',''"">"
		  	else
		  		writeCode "<input name=""htx_"&paramCode&"S"" size="""&paramSize&"""> ～ "
		  		writeCode "<input name=""htx_"&paramCode&"E"" size="""&paramSize&""">"		  	
		  	end if
		 case else
		  select case paramType
		  case "popDate"
		  		writeCode "<input name=""htx_"&paramCode&""" type=hidden>"
		  		writeCode "<input name=""pcShowhtx_"&paramCode&""" size="""&paramSize&""" readonly onclick=""VBS: popCalendar 'htx_"&paramCode&"',''"">"
		  case "refSelect"
		  	refCode = param.selectSingleNode("refLookup").text	  	
			writeCode "<Select name=""htx_"&paramCode&""" size=1>"			
      		writeCode "<option value="""">請選擇</option>"
			writeCode EnumerateCodeList(refCode)
			writeCode ct&ct&ct&"<option value="""&cl&"=RSS(0)"&cr&""">"&cl&"=RSS(1)"&cr&"</option>"
			writeCode ct&ct&ct&cl&ct&"RSS.movenext"& VBCRLF & _
	    		ct&ct&ct&"wend"& cr
			writeCode "</select>" & vbCRLF		
		  case "selection"
			writeCode "<Select name=""htx_"&paramCode&""" size=1>"
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
			writeCode EnumerateCodeList(refCode)
			if dynamicCode <> "" then _
				writeCode ct&ct&ct&cl&" pdxc = replace("""& dynamicCode & """,""'mCode'"",""'""&RSS(0)&""'"")" & cr
			writeCode "<input type=""radio"" name=""htx_"&paramCode&""" value="""&cl&"=RSS(0)"&cr & """ " & cl&"=pdxc"&cr & ">"
	  		writeCode ct&ct&ct&cl&"=RSS(1)"&cr& "&nbsp;&nbsp;"
			writeCode ct&ct&ct&cl&ct&"RSS.movenext"& VBCRLF & _
	    		ct&ct&ct&"wend"& cr

		  case "refCheckbox"
		  	refCode = param.selectSingleNode("refLookup").text
		  	tblRow = nullText(param.selectSingleNode("tblRow"))
		  	if tblRow = "" then	tblRow = 10
		  	dynamicCode = ""
		  		for each dc in param.selectNodes("dynamicBehavior")
		  			dynamicCode = " " & dc.selectSingleNode("event").text & "=""" _
		  				& dc.selectSingleNode("pCall").text & """"
		  		next
		  	dynamicCode = replace(dynamicCode,chr(34),chr(34)&chr(34))
			writeCode "<table width=""100%"">"
			writeCode ct&ct&ct&cl&" pdxc = """""
			writeCode ct&ct&ct&ct&" tblRowCount = 0" & cr
			writeCode EnumerateCodeList(refCode)
			if dynamicCode <> "" then _
				writeCode ct&ct&ct&cl&" pdxc = replace("""& dynamicCode & """,""'mCode'"",""'""&RSS(0)&""'"")" & cr
			writeCode ct&ct&ct&cl& " if (tblRowCount mod " & tblRow & ") = 0 then  response.write ""<tr>""" & cr 
			writeCode ct&"<td><input type=""checkbox"" name=""htx_"&paramCode&""" value=""'"&cl&"=RSS(0)"&cr & "'"" " & cl&"=pdxc"&cr & ">" _
	  			&"<font size=2>"&cl&"=RSS(1)"&cr&"</font>"
			writeCode ct&ct&cl&ct&"RSS.movenext"& VBCRLF & _
				ct&ct&ct&"tblRowCount = tblRowCount + 1" & vbCrLf & _
	    		ct&ct&"wend"& cr
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
		  case "ptypeselect"
		  	writeCode "<Select name=""htx_"&paramCode&""" size=1>"
			response.write "<" & "!--#include file=""../xml/pType.inc""--" & ">"
			writeCode "</select>" & vbCRLF		    		
		  case else
	  		writeCode "<input name=""htx_"&paramCode&""" size="""&paramSize&""">"
		 end select
		end select	
end sub


sub xprocessQParamField(param)
		paramCode = nullText(param.selectSingleNode("fieldName"))
		paramType = nullText(param.selectSingleNode("inputType"))
		paramKind = nullText(param.selectSingleNode("paramKind"))
		paramSize= nullText(param.selectSingleNode("inputLen"))
		if paramSize = "" then paramSize = 10
		if paramSize > 50 then paramSize = 50
		select case paramKind
		  case "range"
		  	if paramType="popDate" then
			  	calendarFlag = true
		  		writeCode "xx<input name=""pcShowhtx_"&paramCode&"S"" size="""&paramSize&""" readonly onclick=""VBS: popCalendar 'htx_"&paramCode&"S','htx_"&paramCode&"E'""> ～ "
		  		writeCode "xx<input name=""htx_"&paramCode&"S"" type=hidden><input name=""htx_"&paramCode&"E"" type=hidden>"
		  		writeCode "xx<input name=""pcShowhtx_"&paramCode&"E"" size="""&paramSize&""" readonly onclick=""VBS: popCalendar 'htx_"&paramCode&"E',''"">"
		  	else
		  		writeCode "<input name=""htx_"&paramCode&"S"" size="""&paramSize&"""> ～ "
		  		writeCode "<input name=""htx_"&paramCode&"E"" size="""&paramSize&""">"		  	
		  	end if
		  case "value"
		  	if paramType="popDate" then
		  		writeCode "xx<input name=""htx_"&paramCode&""" type=hidden>"
		  		writeCode "xx<input name=""pcShowhtx_"&paramCode&""" size="""&paramSize&""" readonly onclick=""VBS: popCalendar 'htx_"&paramCode&"',''"">"
		  	else
		  		writeCode "xx<input name=""htx_"&paramCode&""" size="""&paramSize&""">"
		  	end if
		  case "refSelect"
		  	refCode = param.selectSingleNode("refLookup").text	  	
			writeCode "<Select name=""htx_"&paramCode&""" size=1>"			
      		writeCode "<option value="""">請選擇</option>"
			writeCode EnumerateCodeList(refCode)
			writeCode ct&ct&ct&"<option value="""&cl&"=RSS(0)"&cr&""">"&cl&"=RSS(1)"&cr&"</option>"
			writeCode ct&ct&ct&cl&ct&"RSS.movenext"& VBCRLF & _
	    		ct&ct&ct&"wend"& cr
			writeCode "</select>" & vbCRLF		
		  case "selection"
			writeCode "<Select name=""htx_"&paramCode&""" size=1>"
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
			writeCode EnumerateCodeList(refCode)
			if dynamicCode <> "" then _
				writeCode ct&ct&ct&cl&" pdxc = replace("""& dynamicCode & """,""'mCode'"",""'""&RSS(0)&""'"")" & cr
			writeCode "<input type=""radio"" name=""htx_"&paramCode&""" value="""&cl&"=RSS(0)"&cr & """ " & cl&"=pdxc"&cr & ">"
	  		writeCode ct&ct&ct&cl&"=RSS(1)"&cr& "&nbsp;&nbsp;"
			writeCode ct&ct&ct&cl&ct&"RSS.movenext"& VBCRLF & _
	    		ct&ct&ct&"wend"& cr
		  case "refCheckbox"
		  	refCode = param.selectSingleNode("refLookup").text
		  	dynamicCode = ""
		  		for each dc in param.selectNodes("dynamicBehavior")
		  			dynamicCode = " " & dc.selectSingleNode("event").text & "=""" _
		  				& dc.selectSingleNode("pCall").text & """"
		  		next
		  	dynamicCode = replace(dynamicCode,chr(34),chr(34)&chr(34))
			writeCode ct&ct&ct&cl&" pdxc = """"" & cr
			writeCode EnumerateCodeList(refCode)
			if dynamicCode <> "" then _
				writeCode ct&ct&ct&cl&" pdxc = replace("""& dynamicCode & """,""'mCode'"",""'""&RSS(0)&""'"")" & cr
			writeCode "<input type=""checkbox"" name=""htx_"&paramCode&""" value="""&cl&"=RSS(0)"&cr & """ " & cl&"=pdxc"&cr & ">"
	  		writeCode ct&ct&ct&cl&"=RSS(1)"&cr& "　"
			writeCode ct&ct&ct&cl&ct&"RSS.movenext"& VBCRLF & _
	    		ct&ct&ct&"wend"& cr
		  case "ptypeselect"
		  	writeCode "<Select name=""htx_"&paramCode&""" size=1>"
			xfout.writeline "<" & "!--#include file=""../xml/pType.inc""--" & ">"
			writeCode "</select>" & vbCRLF		    		
		end select	
end sub

%>
