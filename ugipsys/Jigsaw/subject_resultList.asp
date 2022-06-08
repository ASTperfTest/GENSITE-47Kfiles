<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="主題館績效統計"
HTProgFunc="查詢清單"
HTProgCode="jigsaw"
HTProgPrefix="newkpi" %>

<!--#include virtual = "/inc/server.inc" -->
<html> 
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
    <meta name="GENERATOR" content="Hometown Code Generator 1.0">
    <meta name="ProgId" content="kpiQuery.asp">
    <title>查詢結果</title>
	<link type="text/css" rel="stylesheet" href="../css/list.css">
	<link type="text/css" rel="stylesheet" href="../css/layout.css">
	<link type="text/css" rel="stylesheet" href="../css/setstyle.css">
    <script src="../Scripts/AC_ActiveX.js" type="text/javascript"></script>
    <script src="../Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
    <style type="text/css">
<!--
.style1 {color: #000000}
-->
    </style>
</head>
<body>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
<td width="50%" align="left" nowrap class="FormName">農業推薦單元知識拼圖內容管理&nbsp;
<font size=2>【內容條例清單--合理化施肥主題專區 最新議題】</font>
</td>
<td width="50%" class="FormLink" align="right">
<a href="subject_query.htm" target="">重新查詢</a>
<A href="Javascript:window.history.back();" title="回前頁">回前頁</A>
</td>
</tr>
<tr>
<td width="100%" colspan="2"><hr noshade size="1" color="#000080"></td>
</tr>
<tr><td class="Formtext" colspan="2" height="15"></td></tr>
  <tr>
    <td align=center colspan=2 width=80% height=230 valign=top>
  <p align="center">  
     <font size="2" color="rgb(63,142,186)"> 第
     <font size="2" color="#FF0000">1/8</font>頁|                      
        共<font color="#FF0000">130</font>
       <font size="2" color="rgb(63,142,186)">筆| 跳至第       
         <select id=GoPage size="1" style="color:#FF0000">
                   <option value="1"selected>1</option>          
         </select>      
         頁</font>           
        | 每頁筆數:
       <select id=PerPage size="1" style="color:#FF0000">
             <option value="10" selected>15</option>
             <option value="30">30</option>
             <option value="50">50</option>
        </select>     
     </font>     
  </p>
<script language="vbscript">
    sub formSearchSubmit()
         reg.action="GipApproveList.asp"
         reg.Submit
    end Sub
   </script>
        <form method="POST" name="reg" action="../GipApproveList.asp">
          <INPUT TYPE=hidden name=submitTask2 value="">
          <INPUT TYPE=hidden name=CalendarTarget>
                    <CENTER>
                        <TABLE width="100%" id="ListTable">
                          <TBODY>
                            <TR align="left">
                              <th align="middle" width="7%"><div align="center">
                                <INPUT name="ckall" type="button" class="cbutton" onClick="ChkAll" value="全選">
                              </div></th>
                              <th>預覽</th>
                              <th>目錄樹</th>
                              <th>單元名稱</th>
                              <th>資料標題</th>
                              <th>帳號</th>
                              <th>單位</th>
                              <th>編修日期</th>
                            </TR>
                            <TR>
                              <TD align="middle"><div align="center">
                                <INPUT type="checkbox" value="Y" name="ckbox">
                                <INPUT type="hidden" value="314" name="xphKeyicuitem">
                              </div></TD>
                              <TD><A href="#" target="_wMof">View</A></TD>
                              <TD>農業知識入口網（消費者）</TD>
                              <TD>Speeches</TD>
                              <TD>The Current Development and Challenges of   Higher Education in Taiwan </TD>
                              <TD>管理員</TD>
                              <TD>農委會</TD>
                              <TD>097/01/29 18:50</TD>
                            </TR>
                            <TR>
                              <TD align="middle"><div align="center">
                                <INPUT type="checkbox" value="Y" name="ckbox">
                                <INPUT type="hidden" value="1190" name="xphKeyicuitem">
                              </div></TD>
                              <TD><A href="#" target="_wMof">View</A></TD>
                              <TD>農業知識入口網（消費者）</TD>
                              <TD>Speeches</TD>
                              <TD>Taiwan's Sustainable Schools &amp;   Eco-campuses </TD>
                              <TD>管理員</TD>
                              <TD>農委會</TD>
                              <TD>096/08/11 15:06</TD>
                            </TR>
                            <TR>
                              <TD align="middle"><div align="center">
                                <INPUT type="checkbox" value="Y" name="ckbox">
                                <INPUT type="hidden" value="1184" name="xphKeyicuitem">
                              </div></TD>
                              <TD><A href="#" target="_wMof">View</A></TD>
                              <TD>農業知識入口網（消費者）</TD>
                              <TD>Speeches</TD>
                              <TD>A Tour of Taiwan's Sustainable Schools   &amp; Eco-campuses </TD>
                              <TD>管理員</TD>
                              <TD>農委會</TD>
                              <TD>096/08/11 15:04</TD>
                            </TR>
                            <TR>
                              <TD align="middle"><div align="center">
                                <INPUT type="checkbox" value="Y" name="ckbox">
                                <INPUT type="hidden" value="8570" name="xphKeyicuitem">
                              </div></TD>
                              <TD><A href="#" target="_wMof">View</A></TD>
                              <TD>蝴蝶蘭主題館</TD>
                              <TD>News</TD>
                              <TD>Erich Thies, Secretary General of the   German Conference of Ministers of Culture and Education Visits Taiwan </TD>
                              <TD>管理員</TD>
                              <TD>林務局</TD>
                              <TD>096/08/11 15:04</TD>
                            </TR>
                            <TR>
                              <TD align="middle"><div align="center">
                                <INPUT type="checkbox" value="Y" name="ckbox">
                                <INPUT type="hidden" value="8577" name="xphKeyicuitem">
                              </div></TD>
                              <TD><A href="#" target="_wMof">View</A></TD>
                              <TD>蝴蝶蘭主題館</TD>
                              <TD>Video (wmv) </TD>
                              <TD>Welcome to Fo Guang University </TD>
                              <TD>管理員</TD>
                              <TD>林務局</TD>
                              <TD>096/12/19 15:44</TD>
                            </TR>
                            <TR>
                              <TD align="middle"><div align="center">
                                <INPUT type="checkbox" value="Y" name="ckbox">
                                <INPUT type="hidden" value="8574" name="xphKeyicuitem">
                              </div></TD>
                              <TD><A href="#" target="_wMof">View</A></TD>
                              <TD>農業知識入口網2007NEW</TD>
                              <TD>Banner</TD>
                              <TD>Feng Chia University </TD>
                              <TD>adminitsrator</TD>
                              <TD>農業金融局</TD>
                              <TD>096/12/19 15:30</TD>
                            </TR>
                            <TR>
                              <TD align="middle"><div align="center">
                                <INPUT type="checkbox" value="Y" name="ckbox">
                                <INPUT type="hidden" value="8572" name="xphKeyicuitem">
                              </div></TD>
                              <TD><A href="#" target="_wMof">View</A></TD>
                              <TD>農業知識入口網2007NEW</TD>
                              <TD>Banner</TD>
                              <TD>FCU </TD>
                              <TD>蕭尚文</TD>
                              <TD>農業金融局</TD>
                              <TD>096/12/19 15:21</TD>
                            </TR>
                            <TR>
                              <TD align="middle"><div align="center">
                                <INPUT type="checkbox" value="Y" name="ckbox">
                                <INPUT type="hidden" value="8565" name="xphKeyicuitem">
                              </div></TD>
                              <TD><A href="#" target="_wMof">View</A></TD>
                              <TD>知識家</TD>
                              <TD>News</TD>
                              <TD>2008興光閃耀跨年晚會 </TD>
                              <TD>計資中心</TD>
                              <TD>農委會</TD>
                              <TD>096/12/19 10:50</TD>
                            </TR>
                            <TR>
                              <TD align="middle"><div align="center">
                                <INPUT type="checkbox" value="Y" name="ckbox">
                                <INPUT type="hidden" value="5437" name="xphKeyicuitem">
                              </div></TD>
                              <TD><A href="#" target="_wMof">View</A></TD>
                              <TD>知識家</TD>
                              <TD>Video (wmv) </TD>
                              <TD>IEM </TD>
                              <TD>ISU管理員</TD>
                              <TD>農委會</TD>
                              <TD>096/11/26 14:30</TD>
                            </TR>
                            <TR>
                              <TD align="middle"><div align="center">
                                <INPUT type="checkbox" value="Y" name="ckbox">
                                <INPUT type="hidden" value="6959" name="xphKeyicuitem">
                              </div></TD>
                              <TD><A href="#" target="_wMof">View</A></TD>
                              <TD>吳郭魚主題館</TD>
                              <TD>News</TD>
                              <TD>starting date for next semester </TD>
                              <TD>管理員</TD>
                              <TD>林務局</TD>
                              <TD>096/06/07 10:45</TD>
                            </TR>
                            <TR>
                              <TD align="middle"><div align="center">
                                <INPUT type="checkbox" value="Y" name="ckbox">
                                <INPUT type="hidden" value="6958" name="xphKeyicuitem">
                              </div></TD>
                              <TD><A href="#" target="_wMof">View</A></TD>
                              <TD>農業知識入口網2007NEW</TD>
                              <TD>News</TD>
                              <TD>English song competition </TD>
                              <TD>管理員</TD>
                              <TD>林務局</TD>
                              <TD>096/06/07 10:44</TD>
                            </TR>
                            <TR>
                              <TD align="middle"><div align="center">
                                <INPUT type="checkbox" value="Y" name="ckbox">
                                <INPUT type="hidden" value="6678" name="xphKeyicuitem">
                              </div></TD>
                              <TD><A href="#" target="_wMof">View</A></TD>
                              <TD>知識家</TD>
                              <TD>News</TD>
                              <TD>News </TD>
                              <TD>管理員</TD>
                              <TD>農委會</TD>
                              <TD>096/01/07 14:03</TD>
                            </TR>
                            <TR>
                              <TD align="middle"><div align="center">
                                <INPUT type="checkbox" value="Y" name="ckbox">
                                <INPUT type="hidden" value="6412" name="xphKeyicuitem">
                              </div></TD>
                              <TD><A href="#" target="_wMof">View</A></TD>
                              <TD>知識家</TD>
                              <TD>News</TD>
                              <TD>Vice President </TD>
                              <TD>管理員</TD>
                              <TD>農委會</TD>
                              <TD>095/12/13 16:42</TD>
                            </TR>
                            <TR>
                              <TD align="middle"><div align="center">
                                <INPUT type="checkbox" value="Y" name="ckbox">
                                <INPUT type="hidden" value="6088" name="xphKeyicuitem">
                              </div></TD>
                              <TD><A href="#" target="_wMof">View</A></TD>
                              <TD>知識家</TD>
                              <TD>News</TD>
                              <TD>List of Sister Schools </TD>
                              <TD>knuuser</TD>
                              <TD>農委會</TD>
                              <TD>095/12/11 19:26</TD>
                            </TR>
                            <TR>
                              <TD align="middle"><div align="center">
                                <INPUT type="checkbox" value="Y" name="ckbox">
                                <INPUT type="hidden" value="6355" name="xphKeyicuitem">
                              </div></TD>
                              <TD><A href="#" target="_wMof">View</A></TD>
                              <TD>農業知識入口網2007NEW</TD>
                              <TD>News</TD>
                              <TD>test </TD>
                              <TD>管理員</TD>
                              <TD>農委會</TD>
                              <TD>095/12/11 16:38</TD>
                            </TR>
                          </TBODY>
                        </TABLE>
                    </CENTER>
        </form>

    <input name="button" type="button" class="cbutton" onClick="location.href='subject_setList(2).htm'" value="新增存檔">
    <input name="button" type="button" class="cbutton" onClick="location.href='subject_resultList.htm'" value="儲存本頁選擇">
    <INPUT type="button" class="cbutton" onClick="DsdXMLEdit '6355','P' " value="重設" )?></td>
  </tr>
</table>
</body>
</html>