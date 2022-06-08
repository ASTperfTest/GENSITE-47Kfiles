var msg="請勿擅自複製本網頁的內容！";
function disableIE() {if (document.all) {alert(msg);return false;}
}
function disableNS(e) {
  if (document.layers||(document.getElementById&&!document.all)) {
    if (e.which==2||e.which==3) {alert(msg);return false;}
  }
}
if (document.layers) {
  document.captureEvents(Event.MOUSEDOWN);document.onmousedown=disableNS;
} else {
  document.onmouseup=disableNS;document.oncontextmenu=disableIE;
}
document.oncontextmenu=new Function("alert(msg);return false")

//----------------------------------
//禁止使用Ctrl+C鍵
　　var travel=true
　　var hotkey=17　 /* hotkey即是ASII碼,17代表Ctrl鍵 */
　　if (document.layers)
　　 document.captureEvents(Event.KEYDOWN)
　　function onKeydownFunction(e)
　　{ var forbiddenKeys = new Array('c');
        var key;
        var isCtrl;
        if(window.event)
        {
                key = window.event.keyCode;     //IE
                if(window.event.ctrlKey)
                        isCtrl = true;
                else
                        isCtrl = false;
        }
        else
        {
                key = e.which;     //firefox
                if(e.ctrlKey)
                        isCtrl = true;
                else
                        isCtrl = false;
        }
        //if ctrl is pressed check if other key is in forbidenKeys array
        if(isCtrl)
        {
                for(i=0; i<forbiddenKeys.length; i++)
                {
                        //case-insensitive comparation
                        if(forbiddenKeys[i].toLowerCase() == String.fromCharCode(key).toLowerCase())
                        {
                                alert(msg);
                                return false;
                        }
                }
        }
        return true;
}
　　document.onkeydown=onKeydownFunction 

