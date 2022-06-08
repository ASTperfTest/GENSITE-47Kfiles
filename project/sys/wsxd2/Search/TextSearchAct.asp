<%
	Response.Expires = 0
	Response.Expiresabsolute = Now() - 1 
	Response.AddHeader "pragma","no-cache" 
	Response.AddHeader "cache-control","private" 
	Response.CacheControl = "no-cache"
%>
<!--#Include virtual = "/inc/client.inc" -->
<!--#include virtual = "/inc/HyftdFun.inc" -->
<%		
	on error resume next
	'-----------------------------------------------------------------------  			
	Dim Keyword
	'-------------------------------------------------------------------------------
	'---Keyword Setting----------------------------------------------------------------
	Keyword = Request("searchText") 
	Keyword = Replace(Keyword, "'", "''")
	If IsNull(Keyword) Then Keyword = ""	
	
	HyftdServer = Application("HyftdServer")
  HyftdPort = Application("HyftdPort")
  HyftdUserId	= ""
	HyftdUserPassword = ""
	HyftdGroupName = "term"	
	'-------------------------------------------------------------------------------
	'---initial hyftd parameter-----------------------------------------------------  
 	Set HyftdObj = Server.CreateObject("hysdk.hyft.1")
 	'-------------------------------------------------------------------------------
	'---Set Encoding----------------------------------------------------------------
	Call Hyftd_Set_Encoding(HyftdObj, "big5")			
	'-------------------------------------------------------------------------------
	'---Connect to Hyftd Server and initial query-----------------------------------
	HyftdConnId = Hyftd_Connection( HyftdObj, HyftdServer, HyftdPort, HyftdGroupName, HyftdUserId, HyftdUserPassword )		
	Call Hyftd_Initial_Query(HyftdObj, HyftdConnId)		
	
	Dim keys
	keys = Hyftd_GetKeys( HyftdObj, HyftdConnId, Keyword ) 
	
	HyftdX = HyftdObj.hyft_close(HyftdConnId)    
	HyftdObj = Empty
	keys = SortByCount(keys)
	if keys = "" then
%>
<script type="text/javascript">
	alert("選取無關鍵字，請重新選取！")
	history.go(-1);
</script>
<%
	end if
	function SortByCount( keys )		
		Dim words()		
		Dim wordscount()
		Dim resu
		items = split(keys, ";")
		Dim lens : lens = ubound(items)				
		ReDim words(lens)
		ReDim wordscount(lens)		
		dim atext, btext, count
		dim insertcount : insertcount = 0
		for i = 0 to lens - 1
			flag = true
			count = 0
			atext = items(i)
			for j = 0 to lens - 1
				btext = items(j)
				if j < i and atext = btext then
					flag = false	
				end if 
				if atext = btext then
					count = count + 1
				end if
			next
			if flag = true then
				words(insertcount) = atext
				wordscount(insertcount) = count
				insertcount = insertcount + 1
			end if			
		next		
		'for i = 0 to insertcount	
		'	response.write words(i) & "~" & wordscount(i) & ";"		
		'next		
		for i = 0 to insertcount
			for j = 0 to insertcount - i
				if wordscount(j) < wordscount(j+1) then
					temp = wordscount(j)
					wordscount(j) = wordscount(j+1)
					wordscount(j+1) = temp
					temp = words(j)
					words(j) = words(j+1)
					words(j+1) = temp
				end if 
			next
		next
		for i = 0 to insertcount	
			if words(i) <> "" then
				str = str & words(i) & ";"	
			end if
		next				
		str = left(str, len(str) - 1)
		SortByCount = str
	end function	
%>
<form name="SearchForm" method="post" action="kp.asp?xdURL=Search/SearchResultList.asp&mp=1">
	<input name="Keyword" value="<%=replace(keys, ";" ," ")%>" type="hidden" />
	<input name="FromSiteUnit" value="1" type="hidden" />
	<input name="FromKnowledgeTank" value="1" type="hidden" />
	<input name="FromKnowledgeHome" value="1" type="hidden" />
	<input name="FromTopic" value="1" type="hidden" />	
	<input name="FromPedia" value="1" type="hidden" />	
	<input name="FromVideo" value="1" type="hidden" />	
	<input name="FromTechCD" value="1" type="hidden" />	
</form>
<script type="text/javascript">SearchForm.submit();</script>
