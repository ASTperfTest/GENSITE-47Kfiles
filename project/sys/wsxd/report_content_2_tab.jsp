<%@ page language="java" pageEncoding="BIG5"%>
<%@ page import="java.util.*"%>
<%@ page import="org.jdom.input.*"%>
<%@ page import="org.jdom.output.*"%>
<%@ page import="org.jdom.*"%>
<%@ page import="java.io.*"%>
<%
	Document docJDOM;
	//利用SAX建立Document
  
	SAXBuilder bSAX = new SAXBuilder(false);
	

	
		String selUI = "11";
		String xurl = "";   
   if(request.getParameter("report_id")!=null){
	   xurl_ReportUtilBean xurl_rptutil = new xurl_ReportUtilBean();
	   xurl_rptutil.setDataSource(ds);
     String rptidche = request.getParameter("report_id");
     xurl_ReportBean xurl_rpt = xurl_rptutil.getXurl_Report(rptidche);
     if(xurl_rpt!=null){
      selUI = xurl_rpt.getReport_type1_id();
      xurl = xurl_rpt.getXurl();
      }
    }
    if(request.getParameter("report_type1_id")!=null){
    selUI = request.getParameter("report_type1_id");
    } 
  sys = sysUtil.getSysConfig ("REPORT_UI_PATH");
  String xml_path = sys.getParameter_value();
	//String xml_path = new String(application.getInitParameter("reportUI_Path"));
	
	try {
		docJDOM = bSAX.build(new File(xml_path + selUI+ ".xml"));

	} catch (Exception e) {

		e.printStackTrace();

		return;

	}
  
	Element elmtRoot = docJDOM.getRootElement();
%>


                                <tr>
                                    <td align="right" class="t1c3"><font color=red>*</font>文件屬性：</td>
		                            <td>
    <script>
    function sel_report(sel){
      if(sel.options[sel.selectedIndex].value!=''){
        window.location.href='?report_type1_id='+sel.options[sel.selectedIndex].value<%="".equals(report_idInput)?"":"+'&report_id="+report_idInput+"&redirect_url=../personal/my_doc_mtn.jsp'"%>;
      }
    }
    </script>
		
<%
    if (readOnly.equals("")) {
%>
                                <select name=report_type1_id class="sbttn" onchange="sel_report(this)" >
                                  <option value="">請選擇
<%
        String	rpt_type1_id = "";
        if (rpt != null) {
            rpt_type1_id = rpt.getReport_type1_id();
        }

        for (int i=0; i<type1Count; i++) {
            Row row = (Row) type1.elementAt(i);
            String	rpt_type1_idNow = row.getString("REPORT_TYPE1_ID");
           //編號15擬投稿文章與19書籤,兩個屬性不需要因此隱藏
           if("15".equals(rpt_type1_idNow)||"19".equals(rpt_type1_idNow))continue;
           if(request.getParameter("report_type1_id")!=null){%>
           <option value="<%=rpt_type1_idNow%>" <%=rpt_type1_idNow.equals(request.getParameter("report_type1_id"))?"selected":""%>><%=row.getString("REPORT_TYPE1_NAME")%>
           
           <%}else if (rpt_type1_idNow.equals(rpt_type1_id)) {
%>
                                  <option value="<%=rpt_type1_idNow%>" selected><%=row.getString("REPORT_TYPE1_NAME")%>
<%
    		} else {
%>
                                  <option value="<%=rpt_type1_idNow%>" ><%=row.getString("REPORT_TYPE1_NAME")%>
<%
	    	}
	    }
    } else {
        rt1 = type1Util.getReport_type1(rpt.getReport_type1_id());
        if (rt1 != null) {
            out.println("                                " + rt1.getReport_type1_name());
        }
    }
    
%>
                              </select>

		        </td>
	     </tr>
	     <tr>
	</tr>
	<%
		List fildLs = elmtRoot.getChildren();
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
			
			if(!"1".equals(level))continue;
	%>
	
<tr>
<%if("上傳附檔".equals(fieldLabel)){%>

                    <%if (rpt != null){%>
                                <%=rpt.getCreate_date()%>
                                <input type=hidden name=create_date value="<%=rpt.getCreate_date()%>">
                    <%} else {%>
                                <%=modify_date%>
                                <input type=hidden name=create_date value="<%=modify_date%>">
                        <% }%>
         <%}else if("上傳附件".equals(fieldLabel)){%>
                            <%
    // 只有新增文件與版本更新時，可以上傳附檔。
    if ((report_idInput == null) || (report_idInput.equals("")) || (new_version)) {
%>
                            <tr>
                      	    <input type="hidden" name="fname" value="">
                      	    <input type="hidden" name="tfname" value="">
                      	    <input type="hidden" name="tfsize" value="">
                      	    <input type=hidden name=tdescript value="">
                      	    
                      	    
                      	      <td align="right" valign="top" class="t1c3"><font color=red>*</font>上傳附件：
                      	    
                              </td>
                              <td>
                              <% if ( isMedia ) { %>
                              多媒體檔案不可由此處上傳
                            	<% } else {%>
                      		    <a href="#invite" onClick="open_upload_win('../../classss/popup_attache.jsp?form_name=doc_add&focus=submit1','scrollbars=yes,width=450,height=300,screenX=400,screenY=200,top=100,left=400')">上傳附檔</a>
                        	    <span id="attachs">(無)</span>
                        	    <% } %>
                        	    
                              </td>
                            </tr>
<%
    } else {
        if (ras != null) {
%>
                            <tr>
                              <td align="right" valign="top" class="t1c3">文件附件：</td>
                              <td>
<%
	        if (readOnly.equals("")) {
		        String atts = "";
		        for (int j = 0; j < ras.size(); j++) {
                    Row r1 = (Row) ras.elementAt(j);
			        if (atts.length() <= 0) {
				        atts += r1.getString("file_name");
			        } else {
                        atts += ";" + r1.getString("file_name");
			        }
		        }
%>
                      		    <a href="#invite" onClick="MM_openBrWindow('popup_report_attach.jsp?report_id=<%=report_idInput%>&version_no=<%=rpt.getVersion_no()%>','new','scrollbars=yes,width=450,height=300,screenX=400,screenY=200,top=100,left=400')">附檔維護</a>
		                        &nbsp;&nbsp;<%=atts%>
<%
	        } else {
%>
                              	<table>
<%
		        for (int j = 0; j < ras.size(); j++) {
			        Row r1 = (Row) ras.elementAt(j);
%>
                              		<tr>
                              		  <td>
		                                <a href="#invite" onClick="MM_openBrWindow('../../DownloadFileService?file_id=<%=r1.getInt("file_id")%>','new','scrollbars=yes,width=800,height=600,screenX=100,screenY=100,top=100,left=100')"><%=r1.getString("file_name")%></a>
		                              </td>
                              		  <td><%=r1.getInt("file_size")%></td>
                              		</tr>
<%
		        }
%>
                              	</table>
<%
	        }
%>
                              </td>
<%
        }
    }
    
%>
<%}else if("new_media_files".equals(fieldName)){%>
 <Script language="JavaScript">
			MM_openBrWindow('../media/media_readme.htm','new','scrollbars=yes,width=450,height=450');		
</Script>
							<tr>
                               <td align="right" class="t1c3"><font color=red>* </font>影音附件：</td>
                               
                               <td> 
                               <input type=hidden name="new_media_files" value="">
                               <input type=hidden name="new_media_sizes" value="">
                               
                               <%
                               
                               		Vector mda_rows = null;
                               		
                               		if( !report_id.equals("") )
                               		{
                               			//read db
                               			ReportMediaAttachUtil mediaAttachUtil = null;
                               			ReportMediaAttachBean mediaAttachBean = null;
                               			
                               			
                               			
                               			mediaAttachUtil = new ReportMediaAttachUtil();
                               			mediaAttachUtil.setDataSource( ds );
                               			
                               			mda_rows = mediaAttachUtil.getAllByReportId( report_id );
                               			
                               			if( mda_rows != null )
                               			{
                               			
                               			for(int j=0; j< mda_rows.size(); ++j )
                               			{
                               				Row row = (Row) mda_rows.elementAt(j);
                               				String media_filename = null;
                               				String media_url = null;
                               				
                               				media_filename = row.getString("file_name");
                               				media_url = mediaAttachUtil.getMediaUrl( report_id, media_filename );
                               				
                               				if( i!=0 ){
                               					out.print(",");
                               				}
                               	%>
                               				<A href="<%= media_url%>">
                               	<%
                               				
                               					
                               				out.println( media_filename );
                               	%>
                               				</A>
                               	<%
                               	
                               			}
                               			
                               			}
                               			
                               		}
                               		
                               			//維護介面的link
                               			
                               			if( readOnly.equals("") )
                               			{
                               				//String upload_url = null;
                               				//upload_url = sysUtil.getSysConfigValue("MEDIA_WEBSERVER") + "/coa/jupload/upload.jsp" ;
                               				
                               				if( mda_rows!= null ) {
                               	%>
                               				<a href="#invite" onClick="MM_openBrWindow('../media/popup_media_attachment.jsp?form_name=doc_add&field_name_file=new_media_files&field_name_size=new_media_sizes&field_name_text=media_attachs&report_id=<%=report_id%>','new','scrollbars=yes,width=800,height=600,screenX=100,screenY=100,top=100,left=100')">( 刪除檔案 )</a>
                               				<!-- <a href="#invite" onClick="MM_openBrWindow('url','new','scrollbars=yes,width=600,height=500,screenX=100,screenY=100,top=100,left=100')">上傳檔案</a> -->
                               	<%
                               				}
                               				else{
                               					out.println("尚未上傳");
                               				}
                               			}
                               			
                               			else 
                               			{
                               	
                               			}
                               %>
                               <span id="media_attachs"></span>
                                ( 檔案類型：影音檔wmv、wma、wav、asf ，影像檔jpg、gif )</td>
                            </tr>
  
  
<%}else {

%>
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
		}}
	%>
                            <input readonly name=inner_author_name value="<%=authorNms%>" type="hidden" >
                            <input readonly name=inner_author value="<%=authorStr%>" type="hidden" >
                            <input type=hidden name=catids8> <INPUT type=hidden name=catnms8>
<script language="JavaScript">
<!--
function open_br_window(url,size) {
	self.win_child = open(url, "getcode", size)
	self.win_child.win_parent = self;
}
//-->
</script>



<%!public String rootText(Element elem, String rootName, int solt) {

		if (elem.getChildren() != null) {
			List lsroot = elem.getChildren();
			Iterator itroot = lsroot.iterator();
			Element chielm = (Element) lsroot.get(solt);
			Iterator chiItr = chielm.getChildren().iterator();
			while (chiItr.hasNext()) {
				Element chiroot = (Element) chiItr.next();
				if (rootName.equals(chiroot.getName()))
					return chielm.getChildTextTrim(chiroot.getName());
			}
		}
		return elem.getChildTextTrim(elem.getName());
	}
  
  %>
