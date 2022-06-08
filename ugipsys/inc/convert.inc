<%

function ChineseMoney(n)
	dim c, d1,d2,d3,d4,d5, w, ret
	if isnull(n) then
		ChineseMoney = ""
		exit function
	end if
	
	c = "壹貳參肆伍陸柒捌玖"
	
		y = cint(n)
	d1 = y \ 10000
	d2 = (y - d1 * 10000) \ 1000
	d3 = (y - d1 * 10000 - d2 * 1000)\100
	d4 = (y - d1 * 10000 - d2 * 1000 - d3 * 100)\10
	d5 = (y - d1 * 10000 - d2 * 1000 - d3 * 100 - d4 * 10)
		
	ret = ""
	if d1 > 0 then
		ret = mid(c, d1, 1) & "萬"
		if d2 = 0 and (d3<>0 or d4<>0 or d5<>0) then
			ret = ret & "零"
		end if
	end if
	if d2 > 0 then
		ret = ret & mid(c, d2, 1) & "仟"
		if d3 = 0 and (d4<>0 or d5<>0) then
			ret = ret & "零"
		end if
	end if
	if d3 > 0 then
		ret = ret & mid(c, d3, 1) & "百"
		if d4 = 0 and  d5<>0 then
			ret = ret & "零"
		end if
	end if	
	if d4 > 0 then
		ret = ret & mid(c, d4, 1) & "十"
	end if	
	if d5 > 0 then
		ret = ret & mid(c, d5, 1)
	end if
	ret = ret & "元整"
	ChineseMoney = ret
	
end function



function WeekName(week)
	dim w
	
	if isnull(week) then
		WeekName = ""
		exit function
	end if
	
	w = "一二三四五六日"
	WeekName = mid(w, cint(week), 1)
end function

function YearName(year)
	dim y, d1, d2, d3, c, ret
	
	if isnull(year) then
		YearName = ""
		exit function
	end if
	
	y = cint(year)
	d1 = y \ 100
	d2 = (y - d1 * 100) \ 10
	d3 = (y - d1 * 100 - d2 * 10)
		
	c = "零一二三四五六七八九"
	ret = ""
	if d1 > 0 then
		ret = mid(c, d1+1, 1) & "百"
	end if
	if d2 > 0 then
		ret = ret & mid(c, d2+1, 1) & "十"
	elseif d3 > 0 then
		ret = ret & mid(c, d2+1, 1)
	end if
	if d3 > 0 then
		ret = ret & mid(c, d3+1, 1)
	end if
	
	YearName = ret	
end function

function SemeName(seme)
	if isnull(seme) then
		SemeName = ""
		exit function
	end if
	
	select case cint(seme)
	case 1
		SemeName = "第一學期"
	case 2
		SemeName = "第二學期"
	case 3
		SemeName = "第一期"
	case 4
		SemeName = "第二期"
	case else
		SemeName = ""
	end select
end function

function SemeName1(seme)
	if isnull(seme) then
		SemeName1 = ""
		exit function
	end if
	
	select case cint(seme)
	case 1
		SemeName1 = "上"
	case 2
		SemeName1 = "下"
	case else
		SemeName1 = ""
	end select
end function

function SeqName(day)
	if isnull(day) then
		SeqName = ""
		exit function
	end if
	
	if cint(day) < 5 then
		SeqName = day
	elseif cint(day) = 5 then
		SeqName = "午"
	else
		SeqName = cint(day) - 1
	end if
end function

function GradeName(grade)
	dim g
	
	if isnull(grade) then
		GradeName = ""
		exit function
	end if
	
	if grade < 1 or grade > 7 then
		GradeName = ""
		exit function
	end if
	
	g = "一二三四五六七"
	GradeName = mid(g, cint(grade), 1) & "年級"
end function

function ChineseNum(n)
	dim c
	
	if isnull(n) then
		ChineseNum = ""
		exit function
	end if
	
	c = "０１２３４５６７８９"
	ChineseNum = mid(c, cint(n) + 1, 1)
end function

function ChineseNum1(n)
	dim c
	
	if isnull(n) then
		ChineseNum1 = ""
		exit function
	end if
	
	c = "零一二三四五六七八九"
	ChineseNum1 = mid(c, cint(n) + 1, 1)
end function

sub AddValue(src, value)
	dim a, found, i
	
	a = split(src, ",", -1, 1)
	found = false
	for i = 0 to ubound(a)
		if a(i) = value then
			found = true
			exit for
		end if
	next
	if not found then
		if src = "" then
			src = value
		else
			src = src & "," & value
		end if
	end if
end sub

sub AddValue1(src, value)
	dim a, found, i
	
	a = split(src, "、", -1, 1)
	found = false
	for i = 0 to ubound(a)
		if a(i) = value then
			found = true
			exit for
		end if
	next
	if not found then
		if src = "" then
			src = value
		else
			src = src & "、" & value
		end if
	end if
end sub

function Schseme(nchyear,nchseme)
	if nchseme =1 then '上學期		
		Schseme=nchyear & "09" & ":" & nchyear & "12"
	else			   '下學期
		schyear=nchyear+1
		if len(schyear)=2 then
		 Tchyear="0" & schyear
		else
		 Tchyear=schyear
		end if  		
		Schseme=Tchyear & "02" & ":" & Tchyear & "06"			
	end if	
end function

function PointName(point)
	dim y, d1, d2, d3, c, ret
	
	if isnull(point) then
		PointName = ""
		exit function
	end if
	
	y = cint(point)
	d1 = y \ 100
	d2 = (y - d1 * 100) \ 10
	d3 = (y - d1 * 100 - d2 * 10)
		
	c = "一二三四五六七八九"
	ret = ""
	if d1 > 1 then
		ret = mid(c, d1, 1)
	end if
	if d2 > 1 then
		ret = ret & mid(c, d2, 1)
	else
		ret = ret & "○"
	end if
	if d3 > 1 then
		ret = ret & mid(c, d3, 1)
	else
		ret = ret & "○"		
	end if
	
	PointName = ret	
end function

function SerialNo(No)
	dim y, d1, d2, d3, c, ret
	
	if isnull(No) then
		SerialNo = ""
		exit function
	end if
		
	c = "○一二三四五六七八九"
	ret = ""
	for i=1 to len(No)
		subNo = cInt(mid(No,i,1))/1	
		ret = ret & mid(c,subNo+1,1)		
	next
	if Left(No,1)="0" then 
		SerialNo = mid(ret,2)
	else
		SerialNo = ret	
	end if
end function
%>