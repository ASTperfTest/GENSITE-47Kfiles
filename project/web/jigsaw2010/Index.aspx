<%@ Page Title="" Language="C#" MasterPageFile="~/4SideMasterPage.master" AutoEventWireup="true" CodeFile="Index.aspx.cs" Inherits="Index" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <link href="css/style.css" rel="stylesheet" type="text/css" />
    <script src="Scripts/swfobject_modified.js" type="text/javascript"></script>
	<script type="text/javascript">
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
        <asp:panel id="season2010spring" runat="server" class="season2010 spring" style="margin-top:14px">
            <div class="season_lable">
                <h1>
                    <img src="image/season_lb_spring.gif" width="84" height="72" /></h1>
            </div>
            <div class="season_block">
                <h2>
                    代表作物</h2>
                <p>
                    <asp:Repeater ID="rptList10" runat="server">
                        <ItemTemplate>
                            <a href="Detail.aspx?item=<%# Eval("iCUItem")%>"><%# Eval("sTitle")%></a></ItemTemplate>
                    </asp:Repeater>
                </p>
                <h2>
                    代表漁產
                </h2>
                <p>
                    <asp:Repeater ID="rptList11" runat="server">
                        <ItemTemplate>
                            <a href="Detail.aspx?item=<%# Eval("iCUItem")%>"><%# Eval("sTitle")%></a></ItemTemplate>
                    </asp:Repeater>
                </p>
                 <asp:panel id="morespring"  runat="server" align="right"><a href="Index.aspx?&mp=1&season=spring" target="_blank">more..</a> </asp:panel>
            </div>
         </asp:panel>
         <asp:panel id="season2010summer" runat="server" class="season2010 summer">
            <div class="season_lable">
                <h1>
                    <img src="image/season_lb_summer.gif" width="84" height="72" /></h1>
            </div>
            <div class="season_block">
                <h2>
                    代表作物</h2>
                <p>
                    <asp:Repeater ID="rptList20" runat="server">
                        <ItemTemplate>
                            <a href="Detail.aspx?item=<%# Eval("iCUItem")%>"><%# Eval("sTitle")%></a></ItemTemplate>
                    </asp:Repeater>
                </p>
                <h2>
                    代表漁產
                </h2>
                <p>
                   <asp:Repeater ID="rptList21" runat="server">
                        <ItemTemplate>
                            <a href="Detail.aspx?item=<%# Eval("iCUItem")%>"><%# Eval("sTitle")%></a></ItemTemplate>
                   </asp:Repeater>
                </p>
                <asp:panel id="moresummer"  runat="server" align="right"><a href="Index.aspx?&mp=1&season=summer" target="_blank">more..</a> </asp:panel>
            </div>
        </asp:panel>
        <asp:panel id="season2010autumn" runat="server" class="season2010 autumn">
            <div class="season_lable">
                <h1>
                    <img src="image/season_lb_autumn.gif" width="84" height="72" /></h1>
            </div>
            <div class="season_block">
                <h2>
                    代表作物</h2>
                <p>
                    <asp:Repeater ID="rptList30" runat="server">
                        <ItemTemplate>
                            <a href="Detail.aspx?item=<%# Eval("iCUItem")%>"><%# Eval("sTitle")%></a></ItemTemplate>
                    </asp:Repeater>
                </p>
                <h2>
                    代表漁產
                </h2>
                <p>
                    <asp:Repeater ID="rptList31" runat="server">
                        <ItemTemplate>
                            <a href="Detail.aspx?item=<%# Eval("iCUItem")%>"><%# Eval("sTitle")%></a></ItemTemplate>
                    </asp:Repeater>
                </p>
               <asp:panel id="moreautumn"  runat="server" align="right"><a href="Index.aspx?&mp=1&season=autumn" target="_blank">more..</a></asp:panel>
            </div>
        </asp:panel>
       <asp:panel id="season2010winter" runat="server" class="season2010 winter">
            <div class="season_lable">
                <h1>
                    <img src="image/season_lb_winter.gif" width="84" height="72" /></h1>
            </div>
            <div class="season_block">
                <h2>
                    代表作物</h2>
                <p>
                    <asp:Repeater ID="rptList40" runat="server">
                        <ItemTemplate>
                            <a href="Detail.aspx?item=<%# Eval("iCUItem")%>"><%# Eval("sTitle")%></a></ItemTemplate>
                    </asp:Repeater>
                </p>
                <h2>
                    代表漁產
                </h2>
                <p>
                    <asp:Repeater ID="rptList41" runat="server">
                        <ItemTemplate>
                            <a href="Detail.aspx?item=<%# Eval("iCUItem")%>"><%# Eval("sTitle")%></a></ItemTemplate>
                    </asp:Repeater>
                </p>
                 <asp:panel id="moreWinter"  runat="server" align="right" ><a href="Index.aspx?&mp=1&season=winter" target="_blank">more..</a></asp:panel>
            </div>
         </asp:panel>
    </div>
</asp:Content>
