<table width="95%" cellpadding="0" cellspacing="1" border="0">
  <tr>
    <td class="c12-3"><!--<%=ReadmeText%>--></td>
    <td align="right">
      <table height="20" cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td class="whitetablebg" ID="SetListTable" OnClick="SetList()" style="cursor:hand;">瀏覽設定</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
     <table border="0" width="95%" cellspacing="0" cellpadding="0" class="whitetablebg">
      <tr class="c12-1">
        <td width="25%" align="left" valign="bottom">共&nbsp;<%=recno%>&nbsp;筆</td>
        <td width="25%" valign="bottom" align="left"><%=pno+1%>&nbsp;/&nbsp;<%=pgno+1%>&nbsp;頁</td>
        <td width="3%" valign="bottom" title="第一頁"><% if pno>0 then %><span onClick="VBScript: lform.page_no.value='1' : lform.submit" style="cursor: hand;" class="do">9</span><% End If %></td>
        <td width="3%" valign="bottom" title="上一頁"><% if pno>0 then %><span onClick="VBScript: lform.page_no.value='<%=pno%>' : lform.submit" style="cursor: hand;" class="do">3</span><% End If %></td>
        <td width="3%" valign="bottom" title="下一頁"><% if pno<pgno then %><span onClick="VBScript: lform.page_no.value='<%=pno+2%>' : lform.submit" style="cursor: hand;" class="do">4</span><% End If %></td>
        <td width="3%" valign="bottom" title="最後一頁"><% if pno<pgno then %><span onClick="VBScript: lform.page_no.value='<%=pgno+1%>' : lform.submit" style="cursor: hand;" class="do">:</span><% End If %></td>
        <td width="24%" align="right"><% If recno > SetPageCon Then %>到
          <select size="1" name="page_no" onChange="lform.submit()" class="menu">
            <% For k = 1 to pgno+1 %>
            <option value="<%=k%>" <% If pno+1 = k Then %>selected<% End If %>><%=k%></option>
            <% Next %>
          </select>
          頁<% End If %></td>
      </tr>
    </table>
