<%@ page import="java.util.*"%>
<%@ page import="org.jdom.input.*"%>
<%@ page import="org.jdom.output.*"%>
<%@ page import="org.jdom.*"%>
<%@ page import="java.io.*"%>

	<%
		fildLs = elmtRoot.getChildren();
		for (int i = 0; i < fildLs.size(); i++) {
			String fieldSeq = rootText(elmtRoot, "fieldSeq", i);
			String fieldName = rootText(elmtRoot, "fieldName", i);
			String fieldLabel = rootText(elmtRoot, "fieldLabel", i);
			String dataType = rootText(elmtRoot, "dataType", i);
			String dataLen = rootText(elmtRoot, "dataLen", i);
			String inputLen = rootText(elmtRoot, "inputLen", i);
			String canNull = rootText(elmtRoot, "canNull", i);
			String isPrimaryKey = rootText(elmtRoot, "isPrimaryKey", i);
			String inputType = rootText(elmtRoot, "inputType", i);
			String fieldDesc = rootText(elmtRoot, "fieldDesc", i);
			String rows = rootText(elmtRoot, "rows", i);
			String cols = rootText(elmtRoot, "cols", i);
			String formList = rootText(elmtRoot, "formList", i);
			String level = rootText(elmtRoot, "level", i);
			String class_x = rootText(elmtRoot, "class", i);
			String readOnly_e = rootText(elmtRoot, "readOnly", i);
			String defaultValue = rootText(elmtRoot, "defaultValue", i);
			
			if(!"2".equals(level))continue;
	%>
	<tr>
 <td align="right" class="t1c3"><%="N".equals(canNull) ? "<font color=red>*</font>":""%><%=fieldLabel%>：</td>
<td class="t1c3">
			
      
      <%
				if ("textarea".equals(inputType)) {
			%>
			<textarea name="<%=fieldName%>" cols="<%=cols%>" rows="<%=rows%>"
				wrap="VIRTUAL" class="inputtext2"><%
        if("description".equals(fieldName)){
         if (rpt != null) { if (!rpt.getDescription().equals("")) {out.print(rpt.getDescription());} else {out.print("");}} else {out.println("");}
        }else if("foreign_description".equals(fieldName)){
          if (rpt != null) { if (!rpt.getForeign_description().equals("")) {out.print(rpt.getForeign_description());} else {out.print("");}} else {out.println("");}
        } %></textarea>

			<%
				} else if ("".equals(rootText(elmtRoot, "function", i))) {
						List funls = ((Element) fildLs.get(i))
								.getChildren("function");
						Element funElem = (Element) funls.get(0);
						String attributeName = null;
						for (int j = 0; j < funElem.getChildren().size(); j++) {
							attributeName = rootText(funElem, "attributeName", j);
							String attributeValue = rootText(funElem,
									"attributeValue", j);
							String type = rootText(funElem, "type", j);
							String class_f = rootText(funElem, "class", j);
							String event = rootText(funElem, "event", j);
							String action = rootText(funElem, "action", j);
							String actionvalue = rootText(funElem, "actionvalue", j);
							String dataLen_f = rootText(funElem, "dataLen", j);
							String inputLen_f = rootText(funElem, "inputLen", j);
							String readOnly_f = rootText(funElem, "readOnly", j);
							
							if ("a".equals(type) & "popDate".equals(inputType)) {
			%>
			<a href="<%=attributeName%>" <%=("onclick".equals(event))?event+"=\"" + action + "\"" : ""%>><img src="../../images/icon_action03_calendar.gif" alt="日曆" width="15" height="15" align="absbottom"></a>
			<%
				} else if ("a".equals(type)) {
			%>
			<a href="<%=attributeValue%>" <%=("onclick".equals(event))?event + "=\"" + action + "\"" : ""%>><%=attributeName%></a>
			<%
				} else {
				if("popDate".equals(inputType)){
			%>		
			<input name="<%=attributeName%>" type="<%=type%>" <%="".equals(dataLen_f)?"":"size=\""+dataLen_f+"\""%> <%
      if(!"".equals(attributeValue)){%>value="<%=(rpt.getOnline_date().length()>=10?rpt.getOnline_date().substring(0,10):rpt.getOnline_date())%>"<%}%> <%=("onclick".equals(event)) ? event + "=\"" + action + "\"" : ""%> <%="readonly".equals(readOnly_f) ? "readOnly": ""%>>
			<%}else
      {%><input name="<%=attributeName%>" type="<%=type%>" <%="".equals(class_f)?"":"class=\""+class_f+"\"" %> <%="".equals(dataLen_f)?"":"size=\""+dataLen_f+"\""%> <%= "".equals(inputLen_f)?"":"maxlength=\""+inputLen_f+"\"" %> value="<%
  
      if("subject".equals(attributeName)){
      out.print(rpt.getSubject());
    }else if("author".equals(attributeName)){
      out.print(rpt.getAuthor());
    }else if("Submit".equals(attributeName)){
      out.print("帳號瀏覽");
    }else if("catids1".equals(attributeName)){
      out.print(catids1);
    }else if("catnms1".equals(attributeName)){
      out.print(catnms1);
    }else if("catids2".equals(attributeName)){
      out.print(catids2);
    }else if("catnms2".equals(attributeName)){
      out.print(catnms2);
    }else if("catids8".equals(attributeName)){
      out.print(catids8);
    }else if("catnms8".equals(attributeName)){
      out.print(catnms8);
    }else if("online_date".equals(attributeName)){
      out.print(rpt.getOnline_date().length()>=10?rpt.getOnline_date().substring(0,10):rpt.getOnline_date());
    }else if("keywords".equals(attributeName)){
      out.print(rpt.getKeywords());
    }else if("foreign_keywords".equals(attributeName)){
      out.print(rpt.getForeign_keywords());
    }else if("journal".equals(attributeName)){
      out.print(rpt.getJournal());
    }else if("foreign_subject".equals(attributeName)){
      out.print(rpt.getForeign_subject());
    }else if("institution".equals(attributeName)){
      out.print(rpt.getInstitution());
    }else if("domains".equals(attributeName)){
      out.print(rpt.getDomains());
    }else if("inner_author_name".equals(attributeName)){
      out.print(authorNms);
    }else if("inner_author".equals(attributeName)){
      out.print(authorStr);
    }else if("photographer".equals(attributeName)){
      out.print(mediaBean.getPhotographer());
    }else if("photographer_eng".equals(attributeName)){
      out.print(mediaBean.getPhotographer_eng());
    }else if("editor".equals(attributeName)){
      out.print(rpt.getEditor());
    }else if("foreign_editor".equals(attributeName)){
      out.print(rpt.getForeign_editor());
    }else if("foreign_author".equals(attributeName)){
      out.print(rpt.getForeign_author());
    }else if("emails".equals(attributeName)){
      out.print(rpt.getEmails());
    }else if("contact_author".equals(attributeName)){
      out.print(rpt.getContact_author());
    }else if("contact_author_email".equals(attributeName)){
      out.print(rpt.getContact_author_email());
    }else if("shot_place".equals(attributeName)){
      out.print(mediaBean.getShot_place());
    }else if("shot_date".equals(attributeName)){
      out.print(mediaBean.getShot_date().length() > 0 ? (mediaBean.getShot_date().split(" "))[0]:mediaBean.getShot_date());
    }else if("publisher".equals(attributeName)){
      out.print(rpt.getPublisher());
    }else if("published_year_month".equals(attributeName)){
      out.print(rpt.getPublished_year_month());
    }else if("catids0".equals(attributeName)){
      out.print(catids0);
    }else if("catnms0".equals(attributeName)){
      out.print(catnms0);
    }else if("actor_id3".equals(attributeName)){ //尚待check
      out.print("ACTOR007");
    }else if("select_role".equals(attributeName)){
      out.println("角色瀏覽");
    }else if("role_ids".equals(attributeName)){
      out.print(role_ids);
    }else if("role_names".equals(attributeName)){
      out.print(role_nms);
    }else if("catids0".equals(attributeName)){
      out.print(catids0);
    }else if("catids0".equals(attributeName)){
      out.print(catids0);
    }else if("xurl".equals(fieldName)){
      out.print(xurl);
    }else{
      out.print("");
    }%>" <%=("onclick".equals(event)) ? event
									+ "=\"" + action + "\"" : ""%> <%="readonly".equals(readOnly_f) ? "readOnly": ""%>><%}
				}
						}
			%><%=fieldDesc.equals(attributeName) ? ""
							:fieldDesc%>
			<%
				} else {
			%>

              <%if("建檔日期".equals(fieldLabel)){%>
            <%if (rpt != null) {%>
                <%=rpt.getCreate_date()%>
                <input type=hidden name=create_date value="<%=rpt.getCreate_date()%>">
            <%}else {%>
                <%=modify_date%>
                <input type=hidden name=create_date value="<%=modify_date%>">
          <%}%>      
                              
         <%}else if("建檔人員".equals(fieldLabel)){%>
                <%if (rpt != null) {%>
                                <%=rpt.getCreate_user()%>
                                <input type=hidden name=create_user value="<%=rpt.getCreate_user()%>">
                <%} else {%>
                                <%=ekp_login_userid%>
                                <input type=hidden name=create_user value="<%=ekp_login_userid%>">
                  <%}%>
         <%}else if("最後修改人員".equals(fieldLabel)){%>
         <%if (rpt != null) {
			out.println("                                " + rpt.getModify_user());
		} else {
			out.println("                                " + ekp_login_userid);
		}%>
		<input type=hidden name=modify_user value="<%=ekp_login_userid%>">
         <%}else if("最後修改日期".equals(fieldLabel)){%>

    <%if (rpt != null) {
			out.println("                                " + rpt.getModify_date());
		} else {
			out.println("                                " + modify_date);
		}%>
     <input type=hidden name=modify_date value="<%=modify_date%>">
<%}else{%>			<input name="<%=fieldName%>" value="<%
    if("subject".equals(fieldName)){
      out.print(rpt.getSubject());
    }else if("author".equals(fieldName)){
      out.print(rpt.getAuthor());
    }else if("Submit".equals(fieldName)){
      out.print("帳號瀏覽");
    }else if("catids1".equals(fieldName)){
      out.print(catids1);
    }else if("catnms1".equals(fieldName)){
      out.print(catnms1);
    }else if("catids2".equals(fieldName)){
      out.print(catids2);
    }else if("catnms2".equals(fieldName)){
      out.print(catnms2);
    }else if("catids8".equals(fieldName)){
      out.print(catids8);
    }else if("catnms8".equals(fieldName)){
      out.print(catnms8);
    }else if("online_date".equals(fieldName)){
      out.print(rpt.getOnline_date().length()>=10?rpt.getOnline_date().substring(0,10):rpt.getOnline_date());
    }else if("keywords".equals(fieldName)){
      out.print(rpt.getKeywords());
    }else if("foreign_keywords".equals(fieldName)){
      out.print(rpt.getForeign_keywords());
    }else if("journal".equals(fieldName)){
      out.print(rpt.getJournal());
    }else if("foreign_subject".equals(fieldName)){
      out.print(rpt.getForeign_subject());
    }else if("institution".equals(fieldName)){
      out.print(rpt.getInstitution());
    }else if("domains".equals(fieldName)){
      out.print(rpt.getDomains());
    }else if("inner_author_name".equals(fieldName)){
      out.print(authorNms);
    }else if("inner_author".equals(fieldName)){
      out.print(authorStr);
    }else if("photographer".equals(fieldName)){
      out.print(mediaBean.getPhotographer());
    }else if("photographer_eng".equals(fieldName)){
      out.print(mediaBean.getPhotographer_eng());
    }else if("editor".equals(fieldName)){
      out.print(rpt.getEditor());
    }else if("foreign_editor".equals(fieldName)){
      out.print(rpt.getForeign_editor());
    }else if("foreign_author".equals(fieldName)){
      out.print(rpt.getForeign_author());
    }else if("emails".equals(fieldName)){
      out.print(rpt.getEmails());
    }else if("contact_author".equals(fieldName)){
      out.print(rpt.getContact_author());
    }else if("contact_author_email".equals(fieldName)){
      out.print(rpt.getContact_author_email());
    }else if("shot_place".equals(fieldName)){
      out.print(mediaBean.getShot_place());
    }else if("shot_date".equals(fieldName)){
      out.print(mediaBean.getShot_date().length() > 0 ? (mediaBean.getShot_date().split(" "))[0]:mediaBean.getShot_date());
    }else if("publisher".equals(fieldName)){
      out.print(rpt.getPublisher());
    }else if("published_year_month".equals(fieldName)){
      out.print(rpt.getPublished_year_month());
    }else if("catids0".equals(fieldName)){
      out.print(catids0);
    }else if("catnms0".equals(fieldName)){
      out.print(catnms0);
    }else if("actor_id3".equals(fieldName)){ //尚待check
      out.print("ACTOR007");
    }else if("role_ids".equals(fieldName)){
      out.print(role_ids);
    }else if("role_names".equals(fieldName)){
      out.print(role_nms);
    }else if("catids0".equals(fieldName)){
      out.print(catids0);
    }else if("catids0".equals(fieldName)){
      out.print(catids0);
    }else if("xurl".equals(fieldName)){
      out.print(xurl);
    }else{
      out.print("");
    }
         
         
         %>" type=<%=inputType%>
				class="inputtext" size="51" maxlength="<%=inputLen%>"
				<%="readonly".equals(readOnly_e) ? "readOnly"
							: ""%>><%=fieldDesc%><%}%>
         
			<%
				}
			%>
		</td>
	</tr>

			<%
				} 
			%>