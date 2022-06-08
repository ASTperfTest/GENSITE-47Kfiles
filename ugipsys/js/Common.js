﻿//序列化Form將Form中所有控件當作參數
function SerializeFormWorkAjax(url, excfunction) {
    var dataArr = $("form").serialize();
    $.ajax({
        type: "post",
        url: url,
        dataType:"text",
        data: dataArr,
        success: function(html) {
            switch (html) {
                case "1": //表示操作成功
                    //alert("操作成功！");
                    break;
                case "0": //表示操作失敗
                    //alert("操作失敗！");
                    return;
                    break;
                default: //彈出提示信息
                    alert(html);
                    return;
                    break;
            }
            if (html == "1") {
                if (typeof excfunction == "function") {
                    excfunction();
                }
                else if (typeof excfunction == "string" && excfunction != "") {
                    eval(excfunction);
                }
            }
        }
    });
}

function NewAjax(id, url, Eventjs, dataArr) {
    $.ajax({
        type: 'post',
        url: url,
        dataType:"text",
        data: dataArr,
        cache: false,
        success: function(html) {
            if (id != "") {
                $("#" + id).html(html);
            }
            if (typeof Eventjs == "function") {
                Eventjs();
            }
            else if (typeof Eventjs == "string") {
                eval(Eventjs);
            }
        }
    });
}


}
 