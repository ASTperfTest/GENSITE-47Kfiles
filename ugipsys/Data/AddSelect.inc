<form method="POST" action="" Name="HiddenForm">
    <input name="AddName" type="hidden" value="">
    <input name="xid" type="hidden" value="">
</form>
<script language="VBScript">

	sub xin_CatID_OnChange()
 	 if document.all.xin_CatID.value="Add" then
  	  ShValue = Window.ShowModalDialog("AddSelect.asp?DataType=<%=DataType%>&Language=<%=Language%>", NewCompKey, "dialogHeight=215px;dialogWidth=255px;help=no;status=no")
  	   If ShValue = "" Then
  	    document.all.xin_CatID.options(0).selected = True
  	   Else
 	    window.parent.leftFrame.location.href = ShValue
  	   End If
	 End If
	end sub

	Sub SelectAdd()
	   set xsrc=document.all.xin_CatID
	   pos=xsrc.options.length
	    if document.all("AddName").value <>"" then
	      AddName=document.all("AddName").value
	      valueadd=document.all("xid").value
	      xaddOption xsrc,AddName,valueadd,pos
	    end If
	End Sub

	sub xaddOption(xlist,name,value,pos)
		set xOption = document.createElement("OPTION")
		xOption.text=name
		xOption.value=value
		xoption.selected=true
		xlist.add xOption,pos-1
	end sub
</script>
