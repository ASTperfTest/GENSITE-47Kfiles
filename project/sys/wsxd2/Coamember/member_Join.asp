<!--Added By Leo 2011-09-21     增加網址參數  Start-->
<!--#Include virtual = "/inc/InvitationCode.inc" -->
<!--Added By Leo 2011-09-21     增加網址參數   End -->

<script language="javascript">
<!--
    function CheckSelect() {
        //if  (forma.checkboxC.checked !=true &&  forma.checkboxP.checked !=true &&  forma.checkboxS.checked !=true )
        if (forma.checkboxC.checked != true && forma.checkboxS.checked != true) {
            alert("您尚未選擇會員類型!");
            return false;
        }
        return true;
    }
    function goToHome() {
        window.location = "default.asp";
        //window.location= "index.htm";
    }		
//-->
</script>

<div class="path">
    目前位置： <a title="首頁" href="mp.asp">首頁</a> &gt; 加入會員
</div>
<h3>
    會員可享有的權限與功能說明</h3>
<div id="Magazine">
    <div class="Event">
        <form action="sp.asp?xdURL=coamember/member_Agree.asp<%=setInvitationCodeURL() %>" method="post" class="FormA" name="forma">
        <div class="experts">
            <p>農業知識入口網的會員可享有分眾知識樹、農業知識家問答討論、對網站單元發表評價與意見、以及訂閱農業知識入口網電子報等網站服務，當網站不定期辦理線上活動時，您亦可憑此會員身份參加每次的活動。</p>
            <p>一般的網站瀏覽皆不需登入即有閱讀權限，但知識家的發問、討論等功能，皆需登入後才享有發表的權限（一般會員或學者會員皆可）。</p>
            <br />
            <p>會員身分說明：</p>
            <p>一般會員：可瀏覽「消費者知識樹」、「生產者知識樹」，並具有使用知識家功能的權限。申請完成即可開始使用。</p>
            <p>學者會員：須具有學術的背景身分，如研究員、教職人員、或相關科系學生。學者會員可瀏覽「消費者知識樹」、「生產者知識樹」、以及「學者知識樹」。學者會員申請後，需待系統管理員審核通過後方可開通學者權限；若審核未通過，您仍將具有一般會員權限。</p>
            <hr />
            <br />
            <h5>
                請您選擇您想申請的會員身份：</h5>
            <p>
                <input name="memberradio" type="radio" value="1" checked />一般會員<br />
                <input name="memberradio" type="radio" value="2" />學者會員<br />
                <br />
                <center>
    <%        
        invitationCode = request("Fcode") & request("Ecode") & request("Ucode")
        kpi=0
        
        if invitationCode<> "" then
            rspMsg="您的邀請碼不正確，但您還是可以繼續完成會員的註冊動作"
        
            if isnumeric(invitationCode) then            
                Set Conn = Server.CreateObject("ADODB.Connection")
                Conn.Open session("ODBCDSN")                
                
                sqlstr="select * from InviteFriends_Head where invitationCode=" & invitationCode
                set rs=conn.execute (sqlstr)
                if not rs.eof then      '<= 查詢邀請碼
                
                    sqlstr= "select rank0_1 from kpi_set_score where rank0 = 'st_3' and rank0_2='st_316'"                
                    set rs=conn.execute (sqlstr)
                    if not rs.eof then  '<= 取回KPI分數
                        kpi=rs(0)
                        rspMsg=""
                    end if
                end if
            end if
        end if	        
        
        if rspMsg <> "" then response.Write "<font color='red'>" & rspMsg & "</font>"        
        if kpi > 0 then response.Write "<font color='red'>歡迎您的註冊，當您成功註冊後您及您的邀請者都可以獲得KPI分數喔!!</font>"        
    %>
                    <br />
                    <input name="button" type="submit" value="確定" class="Button" onclick="return CheckSelect()">
                    <input type="button" value="取消" class="Button2" onclick="goToHome()">
                </center>
            </p>
        </form>
    </div>
</div>
</div> 