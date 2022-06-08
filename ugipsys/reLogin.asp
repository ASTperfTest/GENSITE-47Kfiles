<%@ CodePage = 65001 %><%
'----session out時重導
%>
<script language=VBS>  
sub window_onLoad
	window.top.frames(0).sessionOutReLoad()
end sub
</script>  
