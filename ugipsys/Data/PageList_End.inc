<!-- 瀏覽設定表格開始 -->
<table border="0" width="250" cellspacing="1" cellpadding="3" id="SetView" class="bluetable" style="position:absolute;top:180px;left:255px; width: 249px; height: 146px; visibility: hidden">
  <tr class="lightbluetable">
    <td>
      <table border="0" width="100%" cellspacing="0" cellpadding="0">
        <tr class="lightbluetable">
          <td>排序</td>
          <td align="right"><input type="button" value="Ｘ" class="cbutton" onClick="SetList()"></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr class="whitetablebg">
    <td align="center">

      <form method="POST" name="ListTableForm">
        <table border="0" cellspacing="2" cellpadding="2" class="whitetablebg">
 <% If OrderField <> "" And OrderName <> "" Then %>
         <tr class="whitetablebg">
            <td>依&nbsp;<select size="1" name="OrderList" class="menu">
	<!-- 欄位排序設定 -->
			<% response.write doSortDisplay %>
	<!-- 欄位排序設定結束 -->
              </select>&nbsp;排序</td>
          </tr>
<% End IF %>
          <tr class="whitetablebg">
            <td>每頁&nbsp;<select size="1" name="PageCount" class="menu">
                <option value="10" <% If session("PageCount") = 10 Then %>selected<% End IF %>>10</option>
                <option value="15" <% If session("PageCount") = 15 Then %>selected<% End IF %>>15</option>
                <option value="20" <% If session("PageCount") = 20 Then %>selected<% End IF %>>20</option>
                <option value="30" <% If session("PageCount") = 30 Then %>selected<% End IF %>>30</option>
              </select>&nbsp;筆</td>
          </tr>
        </table>
      </form>
        <table width="90%" border="0" cellspacing="0" cellpadding="0" class="whitetablebg">
          <tr>
            <td align="right"><input type="button" value="確定" class="cbutton" onClick="VBS: ListTableForm.Submit"></td>
          </tr>
         </table>

    </td>
  </tr>
</table>
<!-- 瀏覽設定表格結束 -->
<form method="POST" action="" Name="HiddenForm">
<input type="Hidden" name="SQL" value="<%=session("querySQL")%>">
<input type="Hidden" name="NowPage" value="<%=pno+1%>">
</form>
<script language=vbs>
 Sub SubmitForm(xlink)
  HiddenForm.action = "GroupRightList.asp?APGroupID="& xlink
  HiddenForm.Submit
 End Sub

 tbsize=<%=tbi%>
 dim srAry(<%=tbi%>,2)
 dim srfAry(<%=rowsno%>)

 for i=0 to <%=rowsno%>
	srfAry(i) = 0
 next

 Sub SetList()
  If document.all.SetView.style.visibility="" Then
   document.all.SetView.style.visibility="hidden"
  Else
   document.all.SetView.style.visibility=""
  End If
 End Sub
</script>
<% If OrderByCk = "Y" Then %><script language="vbs" src="../inc/tblSort.inc"></script><% End If %>
