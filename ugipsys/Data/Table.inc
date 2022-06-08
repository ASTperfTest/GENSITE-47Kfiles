<% If DateDecide = "Y" Then %>
  <object data=../inc/calendar.htm id=calendar type=text/x-scriptlet width=245 height=160 style="position: absolute; top: 30; left: 351; visibility: hidden"></object>
<% End IF %>
<%
 If CatDecide = "Y" Then
  SQL = "Select DataCat.CatName, DataCat.CatID, DataUnit.* From DataUnit, DataCat Where DataCat.CatID = DataUnit.CatID And UnitID = "& UnitID
 Else
  SQL = "Select * From DataUnit Where UnitID = "& UnitID
 End If
  SET RS = conn.execute(SQL)
   EditUser = rs("EditUserID")
   EditDate = rs("EditDate")
%>

  <table border="0" width="580" cellspacing="1" cellpadding="3" class="bluetable">
<% If CatDecide = "Y" Then %>
    <tr ID="CatEdit" title="編修類別" style="cursor:hand;">
      <td width="100" align="right" class="lightbluetable">類別</td>
      <td width="480" class="whitetablebg"><%=rs("CatName")%></td>
    </tr>
<% End IF %>
    <tr ID="SubjectEdit" title="編修<%=Subject%>" style="cursor:hand;">
      <td width="100" align="right" class="lightbluetable"><%=Subject%></td>
      <td width="480" class="whitetablebg"><%=rs("Subject")%></td>
    </tr>
<% If DateDecide = "Y" Then %>
    <tr ID="DateEdit" title="編修公佈時間" style="cursor:hand;">
      <td width="100" align="right" class="lightbluetable">公佈時間</td>
      <td width="480" class="whitetablebg"><%=rs("BeginDate")%>～<%=rs("EndDate")%></td>
    </tr>
<% End IF %>
<% If Not IsNull(Extend_1) Then %>
    <tr ID="Extend_1Edit" title="編修<%=Subject%>" style="cursor:hand;">
      <td width="100" align="right" class="lightbluetable"><%=Extend_1%></td>
      <td width="480" class="whitetablebg">http://<%=rs("Extend_1")%></td>
    </tr>
<% End IF %>
  </table>
<!-- 瀏覽設定表格開始 -->
<table border="0" width="100%" cellspacing="1" cellpadding="3" id="SetView" class="bluetable" style="position:absolute;top:180px;left:255px; width: 300px; height: 146px; visibility: hidden">
  <tr class="lightbluetable">
    <td>
      <table border="0" width="100%" cellspacing="0" cellpadding="0">
        <tr class="lightbluetable">
          <td ID="EditTitle"></td>
          <td align="right"><input type="button" value="Ｘ" class="cbutton" onClick="VBS: SetView.style.visibility = 'hidden'"></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr class="whitetablebg">
    <td align="center">

      <form method="POST" name="ListTableForm" action="" ENCTYPE="MULTIPART/FORM-DATA">
        <DIV Id="CatEditTable" style="Display:none">
        <table width="90%" border="0" cellspacing="0" cellpadding="0" class="whitetablebg">
          <tr>
            <td>類別</td>
            <td><select size="1" name="xin_CatID">
    <% SQLCom = "SELECT * FROM DataCat Where Language = N'"& Language &"' And DataType = N'"& DataType &"' Order By CatShowOrder"
       set rs2 = conn.execute(SQLCom)
        If Not rs2.EOF Then
         Do while not rs2.eof %>
          <option value="<%=rs2("CatID")%>" <% If rs2("CatID") = rs("CatID") Then %>selected<% End if %>><%=rs2("CatName")%></option>
    <%	 rs2.movenext
    	 loop
    	End If %>
          <option value="Add" style="color:red">新增</option>
          </select></td>
          </tr>
         </table>
		 </DIV>

        <DIV id="DataEditTable" style="Display:none">
        <table width="90%" border="0" cellspacing="0" cellpadding="0" class="whitetablebg">
          <tr>
            <td>公佈時間</td>
          <tr><td align="right">
        <input type=text name=Begindate size=10 readonly value="<%=rs("BeginDate")%>" style="cursor:hand;" onclick="VBScript: btdate 1">
      ～<input type=text name=Enddate size=10 readonly value="<%=rs("EndDate")%>" style="cursor:hand;" onclick="VBScript: btdate 2"></td>
          </tr>
        </table>
		</DIV>

        <DIV id="SubjectEditTable" style="Display:none">
        <table width="90%" border="0" cellspacing="0" cellpadding="0" class="whitetablebg">
          <tr>
            <td><%=Subject%></td></tr>
          <tr>
            <td align="right"><input type="text" name="Subject" size="20" value="<%=rs("Subject")%>"></td>
          </tr>
        </table>
		</DIV>

        <DIV id="Extend_1EditTable" style="Display:none">
        <table width="90%" border="0" cellspacing="0" cellpadding="0" class="whitetablebg">
          <tr>
            <td><%=Extend_1%></td></tr>
          <tr>
            <td align="right">http://<input type="text" name="Extend_1" size="20" value="<%=rs("Extend_1")%>"></td>
          </tr>
        </table>
	</DIV>

        <DIV id="AddimgEditTable" style="Display:none">
        <table width="90%" border="0" cellspacing="0" cellpadding="0" class="whitetablebg">
          <tr><td>新增附圖</td></tr>
          <tr><td align="right"><input type="file" name="ImageFile"></td></tr>
          <tr><td align="center"><input type="radio" value="left" name="ImageWay">左　<input type="radio" name="ImageWay" value="right" checked> 右</td></tr>
        </table>
		</DIV>

        <DIV id="ChgimgEditTable" style="Display:none">
        <table width="90%" border="0" cellspacing="0" cellpadding="0" class="whitetablebg">
          <tr><td colspan="2">變更附圖</td>
          </tr>
          <tr><td rowspan="2"><img src="" id="ChgImgsrc" align="absmiddle"></td>
              <td align="right"><input type="file" name="ChgImageFile"></td></tr>
          <tr><td align="center"><input type="radio" value="left" name="ChgImageWay">左　<input type="radio" name="ChgImageWay" value="right" checked> 右</td></tr>
        </table>
		</DIV>

        <DIV id="AddUpStageEditTable" style="Display:none">
        <input type=hidden name="AddUsNowps" value="">
        <table width="90%" border="0" cellspacing="0" class="whitetablebg">
		    <tr>
		      <td width="100" align="right" valign="top"><%=Content%></td>
    		  <td width="380"><textarea rows="6" name="AddUsContent" cols="40"></textarea></td>
 		    </tr>
		    <tr>
		      <td width="100" align="right">附圖</td>
		      <td width="380"><input type=file name="AddUsimagefile" value="">　　<input type="radio" value="left" name="AddUsImageWay"> 左　<input type="radio" name="AddUsImageWay" value="right" checked> 右</td>
		    </tr>
        </table>
		</DIV>

        <DIV id="AddEndStageEditTable" style="Display:none">
        <input type=hidden name="AddEsNowps" value="">
        <table width="90%" border="0" cellspacing="0" class="whitetablebg">
		    <tr>
		      <td width="100" align="right" valign="top"><%=Content%></td>
    		  <td width="380"><textarea rows="6" name="AddEsContent" cols="40"></textarea></td>
 		    </tr>
		    <tr>
		      <td width="100" align="right">附圖</td>
		      <td width="380"><input type=file name="AddEsimagefile" value="">　　<input type="radio" value="left" name="AddEsImageWay"> 左　<input type="radio" name="AddEsImageWay" value="right" checked> 右</td>
		    </tr>
        </table>
		</DIV>


        <DIV id="EditStageEditTable" style="Display:none">
        <table width="90%" border="0" cellspacing="0" class="whitetablebg">
		    <tr>
		      <td width="100" align="right" valign="top"><%=Content%></td>
    		  <td width="380"><textarea rows="6" name="EsContent" cols="40"></textarea></td>
 		    </tr>
        </table>
		</DIV>

		<br>
         <table id="EditTable" width="90%" border="0" cellspacing="0" cellpadding="0" >
          <tr>
            <td align="right"><input type="button" value="確定" class="cbutton" onClick="EditSubmit(NowType)"><input type="button" value="取消" class="cbutton" onClick="VBS: SetView.style.visibility = 'hidden'"></td>
          </tr>
         </table>
	</form>

    </td>
  </tr>
</table>
<!-- 瀏覽設定表格結束 -->


