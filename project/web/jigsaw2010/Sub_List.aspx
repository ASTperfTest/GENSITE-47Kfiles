<%@ Page Title="" Language="C#" MasterPageFile="~/4SideMasterPage.master" AutoEventWireup="true" CodeFile="Sub_List.aspx.cs" Inherits="Sub_List" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <link href="css/style.css" rel="stylesheet" type="text/css" />
    <script src="Scripts/swfobject_modified.js" type="text/javascript"></script>
	<script type="text/javascript">
		window.onload = function ()
		{
			var searchkey = "<%= searchKey%>" ;
			if (searchkey != ""){
				document.getElementById("county").style.display = 'none';
			}
			document.getElementById("searchbox").value = searchkey;
			document.getElementById("countylist").value = "<%= county%>";
		}
		function RedirectSearch(){
			if(document.getElementById("searchbox").value != ""){
				var params = "searchkey=" + escape(document.getElementById("searchbox").value);
				params = params + "&county=" + escape(document.getElementById("countylist").value);
				window.location = "./sub_list.aspx?" + params ;
			}else{
				alert('請輸入關鍵字');
			}
		}
	</script>
    <div class="path" style="float: left; padding-top: 10px">
        目前位置：
    </div>
    <div style="float: left; width: 80%">
        <asp:Literal ID="nav" runat="server"></asp:Literal>
    </div>
    <div id="content">
        <div id="flashmap">
            <object id="FlashID" classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" width="500"
                height="360">
                <param name="movie" value="map.swf" />
                <param name="quality" value="high" />
                <param name="wmode" value="opaque" />
                <param name="swfversion" value="6.0.65.0" />
                <!-- This param tag prompts users with Flash Player 6.0 r65 and higher to download the latest version of Flash Player. Delete it if you don’t want users to see the prompt. -->
                <param name="expressinstall" value="Scripts/expressInstall.swf" />
                <!-- Next object tag is for non-IE browsers. So hide it from IE using IECC. -->
                <!--[if !IE]>-->
                <object type="application/x-shockwave-flash" data="map.swf" width="500" height="360">
                    <!--<![endif]-->
                    <param name="quality" value="high" />
                    <param name="wmode" value="opaque" />
                    <param name="swfversion" value="6.0.65.0" />
                    <param name="expressinstall" value="Scripts/expressInstall.swf" />
                    <!-- The browser displays the following alternative content for users with Flash Player 6.0 and older. -->
                    <div>
                        <h4>
                            Content on this page requires a newer version of Adobe Flash Player.</h4>
                        <p>
                            <a href="http://www.adobe.com/go/getflashplayer">
                                <img src="http://www.adobe.com/images/shared/download_buttons/get_flash_player.gif"
                                    alt="Get Adobe Flash player" width="112" height="33" /></a></p>
                    </div>
                    <!--[if !IE]>-->
                </object>
                <!--<![endif]-->
            </object>
        </div>
    </div>
    		<div class="queryBox"> 
		<b>作物/魚種查詢：</b><input id="searchbox" type="text" value="" size="20" /> <input id="searchbutton" type="button" value="查詢" onclick="RedirectSearch()" />
		<b>地區：</b>
		<select id="countylist" >
		<option value='all'>全部</option>
		<option value='基隆市'>基隆市</option>
		<option value='新北市'>新北市</option>
		<option value='台北市'>台北市</option>
		<option value='桃園縣'>桃園縣</option>
		<option value='新竹縣'>新竹縣</option>
		<option value='新竹市'>新竹市</option>
		<option value='苗栗縣'>苗栗縣</option>
		<option value='台中市'>台中市</option>
		<option value='彰化縣'>彰化縣</option>
		<option value='南投縣'>南投縣</option>
		<option value='雲林縣'>雲林縣</option>
		<option value='嘉義縣'>嘉義縣</option>
		<option value='嘉義市'>嘉義市</option>
		<option value='台南市'>台南市</option>
		<option value='高雄市'>高雄市</option>
		<option value='屏東縣'>屏東縣</option>
		<option value='宜蘭縣'>宜蘭縣</option>
		<option value='花蓮縣'>花蓮縣</option>
		<option value='台東縣'>台東縣</option>
		<option value='澎湖縣'>澎湖縣</option>
		<option value='馬祖'>馬祖</option>
		<option value='金門'>金門</option>
		</select>
		</div>
		<div style="clear:both;"></div>	

    <h3 id="county">
        <asp:Literal ID="sub_title" runat="server"></asp:Literal>
	</h3>
    <span>
        <ul class="group">
            <li class="<%=(types=="0")? "activity": ""%>"><a href="Sub_List.aspx?county=<%=Server.UrlEncode(county)%>&searchkey=<%=Server.UrlEncode(searchKey)%>&types=0&pagesize=<%=pager.PageSize%>#county">作物</a></li>
            <li class="<%=(types=="1")? "activity": ""%>"><a href="Sub_List.aspx?county=<%=Server.UrlEncode(county)%>&searchkey=<%=Server.UrlEncode(searchKey)%>&types=1&pagesize=<%=pager.PageSize%>#county">魚種</a></li>
        </ul>
    </span>
    <script type="text/javascript">
        function pageChange(p) {
            window.location.href = 'Sub_List.aspx?county=<%=Server.UrlEncode(county)%>&searchkey=<%=Server.UrlEncode(searchKey)%>&types=<%=types%>&pagesize=<%=pager.PageSize%>&page=' + (p - 1);
        }
        function perPageChange(ps) {
            window.location.href = 'Sub_List.aspx?county=<%=Server.UrlEncode(county)%>&searchkey=<%=Server.UrlEncode(searchKey)%>&types=<%=types%>&pagesize=' + ps;
        }
    </script>
    <asp:panel id="listdata" runat="server">
    <asp:Repeater ID="rptList" runat="server">
        <HeaderTemplate>
            <div class="Page">
                第 <span class="Number">
                    <%=pager.PageIndex+1 %>/<%=pager.TotalPages%></span> 頁，共 <span class="Number">
                        <%=pager.TotalCount %></span> 筆
                <% if (pager.HasPreviousPage)
                   { %>
                <a title="上一頁" href="Sub_List.aspx?county=<%=Server.UrlEncode(county)%>&searchkey=<%=Server.UrlEncode(searchKey)%>&types=<%=types%>&pagesize=<%=pager.PageSize%>&page=<%=pager.PageIndex-1 %>">
                    <img alt="上一頁" src="xslgip/style3/images/arrow_left.gif">上一頁</a>
                <% } %>
                &nbsp;，到第
                <select class="inputtext" onchange="pageChange(this.value)">
                    <%=pager.PageOptions %>
                </select><label for="select">頁</label>，
                <% if (pager.HasNextPage)
                   { %>
                <a title="下一頁" href="Sub_List.aspx?county=<%=Server.UrlEncode(county)%>&searchkey=<%=Server.UrlEncode(searchKey)%>&types=<%=types%>&pagesize=<%=pager.PageSize%>&page=<%=pager.PageIndex+1 %>">
                    <img alt="下一頁" src="xslgip/style3/images/arrow_right.gif">下一頁 </a>
                <% } %>
                &nbsp; ，每頁
                <select class="inputtext" onchange="perPageChange(this.value)">
                    <%=pager.PageSizeOptions%>
                </select><label for="select">筆</label>資料
            </div>
            <div class="lplist">
                <ul>
        </HeaderTemplate>
        <ItemTemplate>
            <li>
                <table>
                    <tr>
                        <td width="50%">
                            <img src="<%# Eval("xImgFile")%>" alt="<%# Eval("xImgFile")%>" title="<%# Eval("sTitle")%>">
                            <a class="title" href="Detail.aspx?county=<%=Server.UrlEncode(county)%>&item=<%# Eval("iCUItem")%>"><%# PhoneticNotation.GetNotation(Eval("sTitle").ToString())%></a><span class="date"></span>
                                <span class="abstract"><%# Eval("feature")%></span>
                        </td>
        </ItemTemplate>
        <AlternatingItemTemplate>
                        <td width="50%">
                            <img src="<%# Eval("xImgFile")%>" alt="<%# Eval("xImgFile")%>" title="<%# Eval("sTitle")%>">
                            <a class="title" href="Detail.aspx?county=<%=Server.UrlEncode(county)%>&item=<%# Eval("iCUItem")%>"><%# PhoneticNotation.GetNotation(Eval("sTitle").ToString())%></a><span class="date"></span><span
                                    class="abstract"><%# Eval("feature")%></span>
                        </td>
                    </tr>
                </table>
            </li>
        </AlternatingItemTemplate>
        <FooterTemplate>
            <%=(pager.HasNextPage || pager.TotalCount % 2 == 0) ? "" : "<td></td></tr></table></li>"%>
            </ul> </div>
        </FooterTemplate>
    </asp:Repeater>
    </asp:panel>
    <asp:panel id="nonedata" runat="server" visible="false">
        查無資料
    </asp:panel>
</asp:Content>

