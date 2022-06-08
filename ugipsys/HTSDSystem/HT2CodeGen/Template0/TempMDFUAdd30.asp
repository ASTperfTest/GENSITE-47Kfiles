
  For each xatt in xup.Attachments
    if left(xatt.Name,6) = "htImg_" then
	  ofname = xatt.FileName
	  fnExt = ""
	  if instr(ofname, ".")>0 then	fnext=mid(ofname, instr(ofname, "."))
	  tstr = now()
	  nfname = right(year(tstr),1) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & cint(rnd(second(tstr))*100) & fnext
	  sql = sql & mid(xatt.Name,7) & ","
	  sqlValue = sqlValue & pkStr(nfname,",")
  	  xatt.SaveFile apath & nfname, false
  	elseif left(xatt.Name,7) = "htFile_" then
	  ofname = xatt.FileName
	  fnExt = ""
	  if instr(ofname, ".")>0 then	fnext=mid(ofname, instr(ofname, "."))
	  tstr = now()
	  nfname = right(year(tstr),1) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & cint(rnd(second(tstr))*100) & fnext
	  sql = sql & mid(xatt.Name,8) & ","
	  sqlValue = sqlValue & pkStr(nfname,",")
  	  xatt.SaveFile apath & nfname, false
  	  xsql = "INSERT INTO imageFile(newFileName, oldFileName) VALUES(" _
  	  	& pkStr(nfname,",") & pkStr(ofname,")")
  	  conn.execute xsql
	end if
  Next

	sql = left(sql,len(sql)-1) & left(sqlValue,len(sqlValue)-1) & ")"

	conn.Execute(SQL)  
