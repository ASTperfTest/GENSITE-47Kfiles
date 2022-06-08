<%@ include file="../../include/CheckSession.inc" %>
<%@ page import="java.util.Vector,
                 com.hyweb.ekm.bean.*,
                 com.ora.jsp.sql.Row,
                 java.util.Date,
                 java.util.Calendar,
                 com.hyweb.csm"%>
<%@ page contentType="text/html; charset=big5" %>
<%--  include config 從application scope 中取得 DataSource 來源 --%>
<%@ include file="../config.jsp" %>

<%@ include file="../../include/cookie.inc" %>
<%
    request.setCharacterEncoding("Big5");

    String report_idInput = request.getParameter("report_id");
    String readOnly = request.getParameter("readOnly");
    String data_base_idInput = request.getParameter("data_base_id");
	String category_idInput = request.getParameter("category_id");
    String new_versionInput = request.getParameter("new_version");
    String statusInput = request.getParameter("status");
    String isAercMember = request.getParameter("isAercMember");
    boolean isAerc = false;
    boolean isMedia = false;
    boolean isNewReport = false;
    String media_report_type1_id = null;
    ReportMediaBean mediaBean = null;
    ReportMediaUtil mediaUtil = null;
    
    if( isAercMember != null && isAercMember.equalsIgnoreCase("true") )
    {
    	isAerc = true;
    }
    
    if( request.getParameter("isMedia")!=null 
        && request.getParameter("isMedia").equalsIgnoreCase("true") )
    {
    	isMedia = true;
    }
    
    
    String report_id;
    boolean new_version = false;

    if (data_base_idInput == null) {
        data_base_idInput = "no";
    }

    if (category_idInput == null) {
        category_idInput = "no";
    }

    if (readOnly == null) {
        readOnly = "";
    }

    if ((new_versionInput != null) && (new_versionInput.equalsIgnoreCase("true"))) {
        new_version = true;
    }

    if ((statusInput == null) || (statusInput.equals(""))) {
        statusInput = "PUB";
    }

    Calendar create_date = Calendar.getInstance();
	String modify_date = create_date.get(Calendar.YEAR) + "-" + (create_date.get(Calendar.MONTH)+ 1) +
        "-" + create_date.get(Calendar.DAY_OF_MONTH);

    ReportBean rpt = new ReportBean();
    ReportUtilBean rptUtil = new ReportUtilBean();
    rptUtil.setDataSource(ds);
    if ((report_idInput != null) && (!report_idInput.equals(""))) {
        report_id = report_idInput;
        rpt = rptUtil.getReport(report_id);
        request.setAttribute("report", rpt);
        isNewReport = false;
    } else {
        report_id = "";
        rpt.setCreate_date(modify_date);
        rpt.setCreate_user(ekp_login_userid);
        rpt.setModify_date(modify_date);
        rpt.setModify_user(ekp_login_userid);
        rpt.setOnline_date( modify_date );
        isNewReport = true;
    }

    SysConfigBean sys = new SysConfigBean();
    SysConfigUtilBean sysUtil = new SysConfigUtilBean();
    sysUtil.setDataSource(ds);
    
    media_report_type1_id = sysUtil.getSysConfigValue("MEDIA_REPORT_TYPE1_ID");
    if( 
    		!isMedia 
    		&& !isNewReport 
    		&& media_report_type1_id.equals( rpt.getReport_type1_id() )
      )
     {
      	isMedia = true;
      	
      	
      	//Get ReportMediaBean
      	mediaUtil = new ReportMediaUtil();
      	mediaUtil.setDataSource(ds);
      	
      	mediaBean = mediaUtil.getReportMedia( report_id );
     }
     else
     {
     		//Create a empty ReportMediaBean
     		mediaBean = new ReportMediaBean();
    }
    
    if( isMedia )
    {
    	rpt.setReport_type1_id( media_report_type1_id );
    }
          
    
    
    
    sys = sysUtil.getSysConfig("USE_DRM_OR_NOT");
    boolean use_drm_or_not = false;
    if (sys.getParameter_value().equalsIgnoreCase("true")) {
        use_drm_or_not = true;
    }

    ReportType1Bean rt1 = new ReportType1Bean();
    ReportType1UtilBean	type1Util = new ReportType1UtilBean();
	type1Util.setDataSource(ds);
	Vector type1 = type1Util.getAllReport_type1();
	int	type1Count = type1.size();

    ReportType2Bean rt2 = new ReportType2Bean();
    ReportType2UtilBean rt2Util = new ReportType2UtilBean();
    rt2Util.setDataSource(ds);

    ReportAttachUtilBean raUtil = new ReportAttachUtilBean();
    raUtil.setDataSource(ds);
    Vector ras = new Vector();
    if (rpt != null) {
        ras = raUtil.getAllByReport(rpt.getReport_id(), rpt.getVersion_no());
    }

    HisReportUtilBean hisRptUtil = new HisReportUtilBean();
    hisRptUtil.setDataSource(ds);
    Vector hisRpts = new Vector();
    if (rpt != null) {
        hisRpts = hisRptUtil.getAllHisReport(report_id);
    }
%>
<script src="../../include/SS_Popup.js"></script>
<jsp:include page="../../include/report_type.jsp" />
                          <table width="100%" border="0">



<%
    AuthorUtilBean aUtil = new AuthorUtilBean();
	aUtil.setDataSource(ds);
	csm csm1 = new csm();
	csm1.setDataSource(ds);
    Vector authors = null;
	String authorStr = "", authorNms = "";
	if ((report_idInput !=  null) && (!report_idInput.equals(""))) {
		authors = aUtil.getAllAuthor(report_idInput, "");
	}
	if (authors != null) {
		for (int i=0; i<authors.size(); i++) {
			Row r26 = (Row) authors.elementAt(i);
            if (authorStr.length() > 0) {
	            authorStr += ";";
	            authorNms += ";";
            }
			authorStr += r26.getString("USER_ID");
			authorNms += csm1.getActorDetail("ACTOR004", r26.getString("USER_ID"));
		}
	}
    
    if (readOnly.equals("")) {
	    MyCategoryBean myCat = new MyCategoryBean();
	    MyCategoryUtilBean myCatUtil = new MyCategoryUtilBean();
	    myCatUtil.setDataSource(ds);
	    MyCat2rptUtilBean myC2rUtil = new MyCat2rptUtilBean();
	    myC2rUtil.setDataSource(ds);
	    CategoryBean cat = new CategoryBean();
	    CategoryUtilBean catUtil = new CategoryUtilBean();
	    catUtil.setDataSource(ds);
	    Cat2rptUtilBean c2rUtil = new Cat2rptUtilBean();
	    c2rUtil.setDataSource(ds);

	    Vector v1 = null;
	    String catids0 = "", catnms0 = "", catids1 = "", catnms1 = "", catids2 = "", catnms2 = "";
	    String catids3 = "", catnms3 = "", catids4 = "", catnms4 = "", catids5 = "", catnms5 = "";
	    String catids6 = "", catnms6 = "", catids7 = "", catnms7 = "", catids8 = "", catnms8 = "";
	    if (!rpt.getReport_id().equals("")) {
		    String dbid = ekp_login_userid;
	        v1 = myC2rUtil.getAllCat2rpt(rpt.getReport_id(), dbid, "");
		    if (v1 != null) {
			    for (int i=0; i<v1.size(); i++) {
				    Row r1 = (Row) v1.elementAt(i);
				    myCat = myCatUtil.getCategory(ekp_login_userid, r1.getString("CATEGORY_ID"));
				    if (myCat != null) {
					    if (catids0.length() <= 0) {
						    catids0 = ekp_login_userid + "@" + r1.getString("CATEGORY_ID");
						    catnms0 = myCat.getCategory_name();
					    } else {
						    catids0 += "," + ekp_login_userid + "@" + r1.getString("CATEGORY_ID");
						    catnms0 += "," + myCat.getCategory_name();
					    }
				    }
			    }
			    catids0 += ",";
		    }
		    
        dbid = "DB020";
		    v1 = c2rUtil.getAllCat2rpt(rpt.getReport_id(), dbid, "");
		    if (v1 != null) {
			    for (int i=0; i<v1.size(); i++) {
				    Row r1 = (Row) v1.elementAt(i);
				    cat = catUtil.getCategory(dbid, r1.getString("CATEGORY_ID"));
				    if (cat != null) {
					    if (catids1.length() <= 0) {
						    catids1 = dbid + "@" + r1.getString("CATEGORY_ID");
						    catnms1 = cat.getCategory_name();
					    } else {
						    catids1 += "," + dbid + "@" + r1.getString("CATEGORY_ID");
						    catnms1 += "," + cat.getCategory_name();
					    }
				    }
			    }
			    catids1 += ",";
		    }
		    
		    dbid = "DB001";
		    v1 = c2rUtil.getAllCat2rpt(rpt.getReport_id(), dbid, "");
		    if (v1 != null) {
			    for (int i=0; i<v1.size(); i++) {
				    Row r1 = (Row) v1.elementAt(i);
				    cat = catUtil.getCategory(dbid, r1.getString("CATEGORY_ID"));
				    if (cat != null) {
					    if (catids1.length() <= 0) {
						    catids1 = dbid + "@" + r1.getString("CATEGORY_ID");
						    catnms1 = cat.getCategory_name();
					    } else {
						    catids1 += "," + dbid + "@" + r1.getString("CATEGORY_ID");
						    catnms1 += "," + cat.getCategory_name();
					    }
				    }
			    }
			    catids1 += ",";
		    }

		    dbid = "DB002";
		    v1 = c2rUtil.getAllCat2rpt(rpt.getReport_id(), dbid, "");
		    if (v1 != null) {
			    for (int i=0; i<v1.size(); i++) {
				    Row r1 = (Row) v1.elementAt(i);
				    cat = catUtil.getCategory(dbid, r1.getString("CATEGORY_ID"));
				    if (cat != null) {
					    if (catids2.length() <= 0) {
						    catids2 = dbid + "@" + r1.getString("CATEGORY_ID");
						    catnms2 = cat.getCategory_name();
					    } else {
						    catids2 += "," + dbid + "@" + r1.getString("CATEGORY_ID");
						    catnms2 += "," + cat.getCategory_name();
					    }
				    }
			    }
			    catids2 += ",";
		    }

		    dbid = "DB003";
		    v1 = c2rUtil.getAllCat2rpt(rpt.getReport_id(), dbid, "");
		    if (v1 != null) {
			    for (int i=0; i<v1.size(); i++) {
				    Row r1 = (Row) v1.elementAt(i);
				    cat = catUtil.getCategory(dbid, r1.getString("CATEGORY_ID"));
				    if (cat != null) {
					    if (catids3.length() <= 0) {
						    catids3 = dbid + "@" + r1.getString("CATEGORY_ID");
						    catnms3 = cat.getCategory_name();
					    } else {
						    catids3 += "," + dbid + "@" + r1.getString("CATEGORY_ID");
						    catnms3 += "," + cat.getCategory_name();
					    }
				    }
			    }
			    catids3 += ",";
		    }

		    dbid = "DB004";
		    v1 = c2rUtil.getAllCat2rpt(rpt.getReport_id(), dbid, "");
		    if (v1 != null) {
			    for (int i=0; i<v1.size(); i++) {
				    Row r1 = (Row) v1.elementAt(i);
				    cat = catUtil.getCategory(dbid, r1.getString("CATEGORY_ID"));
				    if (cat != null) {
					    if (catids4.length() <= 0) {
						    catids4 = dbid + "@" + r1.getString("CATEGORY_ID");
						    catnms4 = cat.getCategory_name();
					    } else {
						    catids4 += "," + dbid + "@" + r1.getString("CATEGORY_ID");
						    catnms4 += "," + cat.getCategory_name();
					    }
				    }
			    }
			    catids4 += ",";
		    }

		    dbid = "DB005";
		    v1 = c2rUtil.getAllCat2rpt(rpt.getReport_id(), dbid, "");
		    if (v1 != null) {
			    for (int i=0; i<v1.size(); i++) {
				    Row r1 = (Row) v1.elementAt(i);
				    cat = catUtil.getCategory(dbid, r1.getString("CATEGORY_ID"));
				    if (cat != null) {
					    if (catids5.length() <= 0) {
						    catids5 = dbid + "@" + r1.getString("CATEGORY_ID");
						    catnms5 = cat.getCategory_name();
					    } else {
						    catids5 += "," + dbid + "@" + r1.getString("CATEGORY_ID");
						    catnms5 += "," + cat.getCategory_name();
					    }
				    }
			    }
			    catids5 += ",";
		    }

		    dbid = "DB006";
		    v1 = c2rUtil.getAllCat2rpt(rpt.getReport_id(), dbid, "");
		    if (v1 != null) {
			    for (int i=0; i<v1.size(); i++) {
				    Row r1 = (Row) v1.elementAt(i);
				    cat = catUtil.getCategory(dbid, r1.getString("CATEGORY_ID"));
				    if (cat != null) {
					    if (catids6.length() <= 0) {
						    catids6 = dbid + "@" + r1.getString("CATEGORY_ID");
						    catnms6 = cat.getCategory_name();
					    } else {
						    catids6 += "," + dbid + "@" + r1.getString("CATEGORY_ID");
						    catnms6 += "," + cat.getCategory_name();
					    }
				    }
			    }
			    catids6 += ",";
		    }

		    dbid = "DB007";
		    v1 = c2rUtil.getAllCat2rpt(rpt.getReport_id(), dbid, "");
		    if (v1 != null) {
			    for (int i=0; i<v1.size(); i++) {
				    Row r1 = (Row) v1.elementAt(i);
				    cat = catUtil.getCategory(dbid, r1.getString("CATEGORY_ID"));
				    if (cat != null) {
					    if (catids7.length() <= 0) {
						    catids7 = dbid + "@" + r1.getString("CATEGORY_ID");
						    catnms7 = cat.getCategory_name();
					    } else {
						    catids7 += "," + dbid + "@" + r1.getString("CATEGORY_ID");
						    catnms7 += "," + cat.getCategory_name();
					    }
				    }
			    }
		    }
		    catids7 += ",";
		    
		    dbid = "DB010";
		    v1 = c2rUtil.getAllCat2rpt(rpt.getReport_id(), dbid, "");
		    if (v1 != null) {
			    for (int i=0; i<v1.size(); i++) {
				    Row r1 = (Row) v1.elementAt(i);
				    cat = catUtil.getCategory(dbid, r1.getString("CATEGORY_ID"));
				    if (cat != null) {
					    if (catids8.length() <= 0) {
						    catids8 = dbid + "@" + r1.getString("CATEGORY_ID");
						    catnms8 = cat.getCategory_name();
					    } else {
						    catids8 += "," + dbid + "@" + r1.getString("CATEGORY_ID");
						    catnms8 += "," + cat.getCategory_name();
					    }
				    }
			    }
		    }
		    catids8 += ",";
	    }

	    String resource_id = "";
	    sys = sysUtil.getSysConfig("RIGHT_TO_RPT_OR_TO_CAT");
		boolean TO_CAT = true;
	    String role_ids="", role_nms="", read = "1";
	    String actor_info_id[] = null;
		if (sys.getParameter_value().equalsIgnoreCase("REPORT")) {
			TO_CAT = false;
		}
	    if (TO_CAT) {
			if ((data_base_idInput != null) && (category_idInput != null)) {
				resource_id = data_base_idInput + "@" + category_idInput;
			}
		} else {
			if (report_idInput != null) {
				resource_id = "report@" + report_idInput;
			}
		}

	    actor_info_id = csm1.getActorInfoId(resource_id, "ACTOR007", read);
		if (actor_info_id != null) {
			for (int i=0; i < actor_info_id.length; i++) {
				if (actor_info_id[i] != null) {
					if (i == 0) {
						role_ids += actor_info_id[i];
						role_nms += csm1.getActorDetail("ACTOR007", actor_info_id[i]);
					} else {
						role_ids += ";" + actor_info_id[i];
						role_nms += ";" + csm1.getActorDetail("ACTOR007", actor_info_id[i]);
					}
				}
			}
		}
%>





<%
    if ((rpt != null) && (!rpt.getReport_id().equals(""))) {
%>
                          <input type=hidden name=report_id value="<%=rpt.getReport_id()%>">
                          <input type="hidden" name="status" value="<%=rpt.getStatus()%>">
                          <input type="hidden" name="click_count" value="<%=rpt.getClick_count()%>">
<%
    } else {
%>
                          <input type="hidden" name="status" value="<%=statusInput%>">
<%
    }
%>
                            <tr>
                              <td align="right" class="t1c3">版本：</td>
                              <td>
<%
	if (((hisRpts != null) && (hisRpts.size() >= 0)) && (!new_version)) {
		for (int i=0; i<hisRpts.size(); i++) {
			Row r1 = (Row) hisRpts.elementAt(i);
			out.print("                                <a href='../browse_db/his_report_detail.jsp?report_id=" +r1.getString("REPORT_ID")+
                    "&version_no=" +r1.getString("VERSION_NO")+
                    "&data_base_id=" +data_base_idInput+
                    "&category_id="+category_idInput+"'>" +r1.getString("VERSION_NO")+ "</a>");
		}
	}

    if (rpt != null) {
        if (new_version) {
//            out.print("                                <a href='../browse_db/report_detail.jsp?report_id=" +rpt.getReport_id()+
//                    "&data_base_id=" +data_base_idInput+
//                    "&category_id="+category_idInput+"'>" +rpt.getVersion_no()+ "</a>");
%>
                                <b><font color=red><%=rpt.getVersion_no() + 1%></font></b>
                                <input name=version_no value="<%=rpt.getVersion_no() + 1%>" type=hidden>
<%
        } else {
%>
                                <b><font color=red><%=rpt.getVersion_no()%></font></b>
                                <input name=version_no value="<%=rpt.getVersion_no()%>" type=hidden>
<%
        }
    } else {
%>
                                <b><font color=red>1</font></b>
                                <input name=version_no value="1" type=hidden>
<%
    }
%>
                              </td>
                            </tr>
                            
                            
  <%@ include file="report_content_2_tab.jsp" %>                     
     </table>
  <script>
     function showClassify(){
        var classifyTag =  document.getElementById("classify");
        if(classifyTag.style.display=='none'){
          classifyTag.style.display='';
        }else{
          classifyTag.style.display='none';
        }
     }
  </script>
  <div align="right"><a href="javascript:showClassify()">進階設定</a></div>
  <br/> 
<div id="classify" style="display:none">
  <table>
    <%@ include file="report_content_2_tab2.jsp" %>
  </table>   
</div>                   
<%
	}
%>
                          
