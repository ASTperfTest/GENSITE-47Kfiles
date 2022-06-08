<!--#Include virtual = "/inc/client.inc" -->
<!--#Include virtual = "/inc/dbFunc.inc" -->
<%
'File Name	: counter.inc
'Author		: Will
'Date		: 10/17/1999
'Usage		: Counter(0) is for Graphic mode otherwise is Text mode

'計數檔的絕對路徑

Sub Counter(mode,mp)
	'計數檔
	CounterFilename = server.MapPath("counter.txt")

	Set objFileSystem = CreateObject("Scripting.FileSystemObject")
	On Error Resume Next

	
'	Application("HitCount") = ""
	Hit = Clng(Application("HitCount"))
	if Not isNumeric(Hit) or Hit = 0 or Hit="" then
'		response.write isNumeric(Hit) & "<HR>&&"
'		response.write Hit & "***<HR>"
	
		'讀計數檔
		Set objReadedTextFile = objFileSystem.OpenTextFile(CounterFilename,1,0,0)
		Hit = objReadedTextFile.ReadLine
		objReadedTextFile.Close
		Application("HitCount") = Hit
	end if

	if mp = 1 then
		Hit = Clng(Hit) + 1
		Application.lock
			Application("HitCount") = Application("HitCount") + 1
					
		Application.unlock
	end if
  'add by windjan 2005/11/07
   UpdateDB(Application("HitCount"))


'	Hit = 293876
	'寫入新計數值
'	if Hit mod 20 = 0 then
		Set objWritedTextFile = objFileSystem.CreateTextFile(CounterFilename,-1,0)
		objWritedTextFile.WriteLine Hit
		objWritedTextFile.Close
'	end if

	'圖形或文字模式	
	if mode = 0 then
		for i=1 to Cint(Len(Hit))
			Response.Write "<img src=counter/" & Mid(Hit,i,1) & ".gif alt=" & Mid(Hit,i,1) & ">"
		next
	else
		Response.Write Hit
	End if
End Sub

function UpdateDB(num)
     tmpTimeC = Year(now()) & "/" & Month(now())
  	sqlstring="SELECT  MonthPerson.ID, MonthPerson.TmpTime , MonthPerson.person FROM  MonthPerson  where "
  	sqlstring+=" MonthPerson.TmpTime='" & tmp & "'"
  	SET RS=conn.execute(sqlstring) 
  	dim flag=True 
  
    if RS.Bof=True and RS.eof=True then '無資料存在
      flag=False
	end if
	
	if flag=False then
	  InsertData(num,tmpTimeC)
	 else
	  updateData(num,tmpTimeC)
	 end if  
  
end function

function InsertData(num,tmpTimeC)
   sql="Insert into MonthPerson(TmpTime,Person) Values(tmpTimeC,1)"
   set  cmd=Server.CreateObject("ADODB.Command")
   cmd.CommandText=sql
   cmd.Execute
      
end function



function UpdateData(num,tmpTimeC)
    sql="Update MonthPerson set Person='" &num & "' where tmpTime='" & tmpTime & "'"
   set  cmd=Server.CreateObject("ADODB.Command")
   cmd.CommandText=sql
   cmd.Execute
end function





%>
