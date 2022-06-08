<%@ Page Language="VB" MasterPageFile="~/4SideMasterPage.master" AutoEventWireup="false" CodeFile="PediaInfoDetail.aspx.vb" Inherits="Pedia_PediaInfoDetail" Title="Untitled Page" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <a href="#" title="網頁內容資料" class="Accesskey" accesskey="C">:::</a>
    <div class="path">
        目前位置：<a href="/">首頁</a>&gt;<a href="/Pedia/PediaList.aspx">農業知識小百科</a></div>
    <div class="pantology">
        <div class="head">
        </div>
        <div class="body">
            <h2>
                ~小百科規則~</h2>
            <div class="sublayout">
                <div class="Function2">
                    <ul>
                        <li><a href="javascript:history.go(-1);" class="Back" title="回上一頁">回上一頁</a><noscript>本網頁使用SCRIPT編碼方式執行回上一頁的動作，如果您的瀏覽器不支援SCRIPT，請直接使用「Alt+方向鍵向左按鍵」來返回上一頁</noscript>
                        </li>
                    </ul>
                </div>
                <div style="font-size: 15px; font-family: sans-serif; font-style: normal; color: #000000; letter-spacing:1px; line-height:22.5px">
                    <img src="Icon.png">
                    <h4 style="font-size: 16px; font-family: sans-serif; font-style: normal; color: #000000">農業知識小百科功能公告</h4>
                    <p>
                        將農業知識庫和農業主題館中看到的「農業詞彙」推薦給小百科，例如提出對「摘心」、「組識培養」有疑惑，可使用「推薦詞彙」發表詢問，而網站的會員可使用「我來補充」的編輯行列，到小百科裡一一來解釋農業詞彙的意義。</p>
                    <br />
                    <p>
                        <strong>推薦詞彙範例</strong><br />
                        小百科推薦詞彙限定『農業知識』類別，如：栽培技術、防治方法、病蟲害防制技術...等。請勿推薦非農業知識類詞彙。 另，生物或設備等名詞詞彙亦不符合本單元範圍，如蝴蝶蘭、荷蘭牛...等。<br />
                        <br />
                        合格詞彙範例：隧道式栽培、設施栽培、扦插法、摘心、高效輪作制<br />
                        不符合詞彙範例：金棗、藥用植物、網室、有機蔬果、雄蕊
                    </p>
                    <p>
                        <strong>補充解釋</strong><br />
                        1.邀請農業知識達人的您，來幫忙網友解釋農業知識詞彙的意義。<br />
                        2.每個詞目皆會呈現目前的補充狀況，請勿重複補充解釋，情況嚴重者管理員可關閉其會員權限。<br />
                        3.補充解釋可引用前文，您可以補充您的說明，讓解釋更完整。<br />
                    </p>
                </div>
            </div>
        </div>
        <div class="foot">
        </div>
    </div>
</asp:Content>
