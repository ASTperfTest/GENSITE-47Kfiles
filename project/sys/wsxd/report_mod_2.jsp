<%@ page import="com.hyweb.ekm.bean.ReportBean,
                 com.hyweb.ekm.bean.ReportUtilBean,
                 com.hyweb.ekm.bean.SysConfigBean,
                 com.hyweb.ekm.bean.SysConfigUtilBean,
                 com.hyweb.ekm.bean.Cat2rptUtilBean"%>
<%@ page contentType="text/html; charset=big5" %>
<%--  include config �qapplication scope �����o DataSource �ӷ� --%>
<%@ include file="../config.jsp" %>

<%@ include file="../../include/cookie.inc" %>
<%@ include file="../../include/action_ctl.inc" %>

<%
    request.setCharacterEncoding("Big5");
    String data_base_idInput = request.getParameter("data_base_id");
	String category_idInput = request.getParameter("category_id");
    String report_idInput = request.getParameter("report_id");
    String isAercMember = request.getParameter("isAercMember");
    String isMedia = request.getParameter("isMedia");
    boolean isAerc = false;
    String title = "";
    String state = "NEW";
    Cat2rptUtilBean c2rUtil = null;
    
    c2rUtil = new Cat2rptUtilBean();
    c2rUtil.setDataSource( ds );

	if( isAercMember != null && isAercMember.equalsIgnoreCase("true") )
    {
    	isAerc = true;
    }    
    else if( report_idInput != null )
    {
    	String dbid = "DB010";
    	Vector v1 = null;
		v1 = c2rUtil.getAllCat2rpt(report_idInput, dbid, "");
		
		if( v1!= null && v1.size() > 0 )
		{
			isAerc = true;
			isAercMember="true";
		}
    }    
    
    if (report_idInput == null) {
        report_idInput = "";
        title = "�s�W";
    } else {
        title = "�ק�";
    }

    if ((data_base_idInput == null) || (data_base_idInput.equals(""))) {
        data_base_idInput="no";
    }
    if ((category_idInput == null) || (category_idInput.equals(""))) {
        category_idInput="no";
    }

    String redirect_url = request.getParameter("redirect_url");
    if (redirect_url == null) {
        redirect_url = "../manage_doc/report_mod.jsp";
    }

    SysConfigBean sys = new SysConfigBean();
	SysConfigUtilBean sysUtil = new SysConfigUtilBean();
	sysUtil.setDataSource(ds);
	sys = sysUtil.getSysConfig("RIGHT_TO_RPT_OR_TO_CAT");
	boolean TO_CAT = true;
	if (sys.getParameter_value().equalsIgnoreCase("REPORT")) {
		TO_CAT = false;
	}
%>

<%	//�ˬd�ϥΪ̬O�_���ìP�ϥΪ̡A�Y�O���ܻݦ^�ìP�W�Z	97-09-17
	Connection conn = null;
	try{
		conn = ds.getConnection();
		String sqlStr = "select * from login_user "+
			"where site_user_id = ? and substring(dept_id,1,10) in "+
			"( select dept_id from ekm_satellite_site where dept_id is not null ) ";
		PreparedStatement pstm = conn.prepareStatement(sqlStr);
		pstm.setString(1,ekp_login_userid);
		ResultSet rs = pstm.executeQuery();
		if ( rs.next() ){	//�O�ìP���b��
			out.println("<script>");
			out.println("alert('�ìP�ϥΪ̽Ц^�ìP�t�ζi��峹���@�A���¡C');");
			out.println("location.href='../personal/my_doc_mtn.jsp'");
			out.println("</script>");
		}
	}catch(Exception e){
		e.printStackTrace();
	}finally{
		try{
			if( conn != null ) conn.close();
		}catch(Exception e){}
	}
%>

<html>
<head>
<title>�����@ -- <%=title%>���</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../../css/css.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
	window.open(theURL,winName,features);
}

function Checkinput(form) {
	
		

        if (document.doc_add.subject.value=="") {
            alert('�ж�g�����D,����')
            document.doc_add.focus()
            return false

<%
	if (state.equals("VER")) {
%>
	    } else if (document.doc_add.to_update_report_id.value=="") {
            alert('�ж�ҭn��s�����i�N��,����')
            document.doc_add.to_update_report_id.focus()
            return false
<%
	}
%>
//		} else if (document.doc_add.keywords.value=="") {
//			alert('�ж�����r,����')
//			document.doc_add.keywords.focus()
//			return false
		}
		
		if (document.doc_add.keywords.value=="") {
			alert('�ж񤤤�����r,����')
			document.doc_add.keywords.focus()
			return false
		}
		
		if (document.doc_add.foreign_keywords.value=="") {
			alert('�ж�^������r,����')
			document.doc_add.foreign_keywords.focus()
			return false
		}
    
    if (document.doc_add.tfname.value=="") {
			alert('�ФW�Ǫ���,����')
			return false
		}
		
		var cat = new Array()
		cat[0] = document.doc_add.catnms0.value;
		cat[1] = document.doc_add.catnms1.value;
		cat[2] = document.doc_add.catnms2.value;
		/*
		cat[3] = <%= (isAerc)? "" : "document.doc_add.catnms3.value;" %>
		cat[4] = <%= (isAerc)? "" : "document.doc_add.catnms4.value;" %>
		cat[5] = <%= (isAerc)? "" : "document.doc_add.catnms5.value;" %>
		cat[6] = <%= (isAerc)? "" : "document.doc_add.catnms6.value;" %>
		cat[7] = <%= (isAerc)? "" : "document.doc_add.catnms7.value;" %>
		*/
		cat[8] = document.doc_add.catnms8.value;

		<% if( !isAerc ) 
			{
		%>
				if((cat[1]=="") || (cat[2]==""))
				{
					alert('�п�ܹA�~�P��쪾�Ѿ𪺤�����,����')
					document.doc_add.keywords.focus()
					return false
				}
		<%
			}
		    else
			{
		%>
				if((cat[8]=="") || (cat[2]==""))
				{
					alert('�п�ܹA�Ф��Q�P��쪾�Ѿ𪺤�����,����')
					document.doc_add.keywords.focus()
					return false
				}
		<%
			}
		%>

		if (document.doc_add.report_type1_id.value=="")
		{
			alert('�п�ܤ���ݩ�,����')
			document.doc_add.focus()
			return false
		}

		numdays=new Array(31,28,31,30,31,30,31,31,30,31,30,31)
		//�U������Ѽ�
		var i=0;
		var k=0;

		i=0;
		k=0;
		var c=document.doc_add.online_date.value.length;
		var Odate1=document.doc_add.online_date.value;
		for(i=0;i<c;i++)
		{
			if(Odate1.charAt(i)=="-")
			{
				k++
			}
		}
		if (k!=2)
		{
			alert('���G����榡����,�ʤ�-�Φh�F-,�Э��s��J,����')
			document.doc_add.focus()
			return false
		}
		//�p��"-"���ƥ�  ���T���ɶ��榡 �p:2000-2-2 �������"-"

		var Odate2=Odate1.split("-");
		if(Odate2[0]<1800 || Odate2[0]>2050)
		{
			alert('���G����~���榡����,�Э��s��J,����')
			document.doc_add.focus()
			return false
		}

		numdays[1]=28
		checkyear=Odate2[0]%4
		if(checkyear==0){numdays[1]=29}//�ˬd�O�_���|�~

		if(Odate2[1]<1 || Odate2[1]>12)
		{
			alert('���G�������榡����,�Э��s��J,����')
			document.doc_add.focus()
			return false
		}
		//����b1��12��

		if(Odate2[2]<1 || Odate2[2]>numdays[Odate2[1]-1] )
		{
			alert('���G�������榡����,�Э��s��J,����')
			document.doc_add.focus()
			return false
		}
		//�ˬd�O�_�b�U����Ѽ�

		<% if( isAerc ) { %>
		
		role = new Array()
		role[0] = "�A�������~��";
		role[1] = "���O��";
		role[2] = "���~�t��";
		role[3] = "��s�H��";
		role[4] = "�M���H��";

		var rlength = document.doc_add.role_names.value.length;
		var ridarr = document.doc_add.role_names.value;
		var rid = ridarr.split(";");
		if(rlength == 1)
		{
			for(var j=0;j<5;j++){
			  if((role[j]==rid[0]) && (cat[j+3]==""))
			  {
			        alert('�п�ܻP����۹�����������O,����')
			        document.doc_add.focus()
			        return false
			  }
			}
		} else if(rlength > 1)
		{
		  for(var i=0;i<rlength;i++){
		    for(var j=0;j<5;j++){
			  if((role[j]==rid[i]) && (cat[j+3]==""))
			  {
			        alert('�п�ܻP����۹�����������O,����')
			        document.doc_add.focus()
			        return false
			  }
		    }
		  }
		}
		
		<% } %>
		
		return true
}

function open_upload_win(url,size)
{
	var fnamearr = document.doc_add.fname.value.split(", ");
	var tdescriptarr = document.doc_add.tdescript.value.split(", ");
	var tfnamearr = document.doc_add.tfname.value.split(", ");
	var tfsizearr = document.doc_add.tfsize.value.split(", ");
	for(var i=0;i<fnamearr.length;i++){
		url = url + "&fname=" + fnamearr[i];
		url = url + "&tdescript=" + tdescriptarr[i];
		url = url + "&tfname=" + tfnamearr[i];
		url = url + "&tfsize=" + tfsizearr[i];
	}
	url = url + "&codeset=no"
	//alert(url);
	self.win_child = open(url, "getcode", size)
	self.win_child.win_parent = self;

//alert('abc123')
	
//	var fnamearr = document.doc_add.fname.value.split(", ");
//	var tdescriptarr = document.doc_add.tdescript.value.split(", ");
//	var tfnamearr = document.doc_add.tfname.value.split(", ");
//	var tfsizearr = document.doc_add.tfsize.value.split(", ");
//	for(var i=0;i<fnamearr.length;i++){
//		url = url + "&fname=" + fnamearr[i];
//		url = url + "&tdescript=" + tdescriptarr[i];
//		url = url + "&tfname=" + tfnamearr[i];
//		url = url + "&tfsize=" + tfsizearr[i];
//	}
//	url = url + "&codeset=no"
//	
//	
//	
//	// The Javascript escape and unescape functions do not correspond
// // with what browsers actually do...
// var SAFECHARS = "0123456789" +     // Numeric
//     "ABCDEFGHIJKLMNOPQRSTUVWXYZ" + // Alphabetic
//     "abcdefghijklmnopqrstuvwxyz" +
//     "-_.!~*'()&?=/";     // RFC2396 Mark characters
// var HEX = "0123456789ABCDEF";
// 
// alert( url )
// 
// var plaintext = url;
// var encoded = "";
// for (var i = 0; i < plaintext.length; i++ ) {
//  var ch = plaintext.charAt(i);
//     if (ch == " ") {
//      encoded += "+";    // x-www-urlencoded, rather than %20
//  } else if (SAFECHARS.indexOf(ch) != -1) {
//      encoded += ch;
//  } else {
//      var charCode = ch.charCodeAt(0);
//      var charCode2 = ch.charCodeAt(1);
//      
//   if (charCode > 255) {
//       alert( "Unicode Character '" 
//                        + ch 
//                        + "' cannot be encoded using standard URL encoding.\n" +
//              "(URL encoding only supports 8-bit characters.)\n" +
//        "A space (+) will be substituted." );
//    encoded += "+";
//   } else {
//    encoded += "%";
//    encoded += HEX.charAt((charCode >> 4) & 0xF);
//    encoded += HEX.charAt(charCode & 0xF);
//   }
//   
//   
//   
//   if (charCode2 > 255) {
//       alert( "Unicode Character '" 
//                        + ch 
//                        + "' cannot be encoded using standard URL encoding.\n" +
//              "(URL encoding only supports 8-bit characters.)\n" +
//        "A space (+) will be substituted." );
//    encoded += "+";
//   } else {
//    encoded += "%";
//    encoded += HEX.charAt((charCode2 >> 4) & 0xF);
//    encoded += HEX.charAt(charCode2 & 0xF);
//   }
//   
//   
//  }
// } // for
//
//	
//	
//	
//	
//	
//	
//	
//	
//	
//	
//	
//	
//	
//	
//	
//	
//	
//	
//	alert(encoded);
//	
//	//self.win_child = open(url, "getcode", size)
//	self.win_child = open( encoded, "getcode", size)
//	self.win_child.win_parent = self;
}
//-->
</script>
<script src="../../include/SS_Popup.js"></script>
</head>
<body bgcolor="#ffffff" alink="black" vlink="gray" link="black" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" cellspacing="0" cellpadding="0" height="100%" class="tab1" border="0">
  <tr>
    <td valign="top">
      <table border="0" cellpadding="0" cellspacing="0" width="100%">
        <tr>
          <td class="headBkg"><img src="../../images/main_title_ekm.gif" name="index_20" border="0"></td>
        </tr>
        <tr>
          <td><table width="95%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="33">&nbsp;</td>
                <td>
                  <table width="100%" border="0" cellspacing="0" cellpadding="0" style="table-layout:fixed; word-break:break-all">
                  <form name="doc_add" method="post" action="report_mod_act.jsp" onsubmit="return Checkinput(this);">
                    <input type="hidden" name="data_base_id" value="<%=data_base_idInput%>">
                    <input type="hidden" name="category_id" value="<%=category_idInput%>">
                    <input type="hidden" name="redirect_url" value="<%=redirect_url%>">
                    <input type="hidden" name="isAercMember" value="<%=request.getParameter("isAercMember")%>">
                    <input type="hidden" name="ui_type" value="<%= ( request.getAttribute("ui_type")==null) ? "" : (String) request.getAttribute("ui_type") %>" >
                    <tr>
                      <td class="t1"> <img src="../../images/main_titleicon01_pfm.gif" align="absbottom">���޲z(��s������)</td>
                    </tr>
                    <tr>
                      <td background="../../images/main_line_01.gif" height="10"><spacer height="1" width="1" type="block"></td>
                    </tr>
                    <tr>
                      <td class="t1c1"><p><img src="../../images/arrow01_03.gif"><%=title%>���(<font color=red>*</font>��ܥ������)</td>
                    </tr>
                    <tr valign="top" bgcolor="#CCCCCC">
                      <td height="1" bgcolor="#CCCCCC"><spacer height="1" width="1" type="block"></td>
                    </tr>
                    <tr>
                      <td>

<jsp:include page="../manage_doc/report_content_2.jsp">
<jsp:param name="report_id" value="<%=report_idInput%>" />
<jsp:param name="isAercMember" value="<%=isAercMember%>" />
<jsp:param name="isMedia" value="<%= isMedia %>" />
</jsp:include>

                      </td>
                    </tr>
                    <tr>
                      <td>&nbsp;</td>
                    </tr>
                    <tr>
                      <td>&nbsp;</td>
                    </tr>
                  </table>
                  <table width="100%" border="0" cellspacing="0" cellpadding="0" style="table-layout:fixed; word-break:break-all">
<%
    if (!TO_CAT) {
	    String forAllRead = "";
	    csm csm2 = new csm();
		csm2.setDataSource(ds);
	    String rid = "report@" + report_idInput;
	    String[] r = csm2.getResourceRight(rid, "ACTOR004", "forAllRead");
	    // default �s�W���ɬ��w�_��F�ק�ɡA�h���ˬd�O�_���e���Q�_��C
	    if ( /*(title.equals("�s�W")) ||*/ (r[0] != null)) {
            forAllRead = "checked";
            
	    }
%>
                    <tr>
                      <td class="t1c1" width=20%><img src="../../images/arrow01_03.gif">�\Ū�v���]�w</td>
                      <td><input type=checkbox name=forAllRead value="SysAllUser" <%=forAllRead%>>�����}��\Ū</td>
                    </tr>
                    <tr valign="top" bgcolor="#CCCCCC">
                      <td colspan=2 height="1" bgcolor="#CCCCCC"><spacer height="1" width="1" type="block"></td>
                    </tr>
                    <tr>
                      <td colspan=2>

<jsp:include page="../manage_doc/set_read_rights.jsp">
<jsp:param name="report_id" value="<%=report_idInput%>" />
</jsp:include>

                      </td>
                    </tr>
<%
    }
%>
                    <tr valign="top" bgcolor="#CCCCCC">
                      <td colspan=2 height="1" bgcolor="#CCCCCC"><spacer height="1" width="1" type="block"></td>
                    </tr>
                    <tr>
                      <td>&nbsp;</td>
                    </tr>
                    <tr>
                      <td colspan=2 align="center">
                        <input name="state" type="submit" class="sbttn" value="�T�w">
                      </td>
                    </tr>
                    <tr>
                      <td>&nbsp;</td>
                    </tr>
                  </form>
                  </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>

<Script language="JavaScript">
<!--
<%
	if ( "true".equalsIgnoreCase(isMedia) )
	{
		%>
			MM_openBrWindow('../media/media_readme.htm','new','scrollbars=yes,width=450,height=450');		<%
	}
%>
//
//	 select_cate(document.doc_add.report_type1_id);
-->
</Script>