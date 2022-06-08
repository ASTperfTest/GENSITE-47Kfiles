<%
'Call <IMG SRC="count.asp?ID=MYPAGE1"> and change the
'MYPAGE1 to whatever you want to create another counters
function Barra(Pasta)
Barra = Pasta
if mid(Barra,len(Barra),1) <> "\" then Barra = Barra + "\"
end function

Set Sistema = CreateObject("Scripting.FileSystemObject")
set ILIB = server.createobject("Overpower.ImageLib")
BaseFolder = barra(server.mappath("/"))
CounterFile = BaseFolder + "count"+request.querystring("ID") + ".COU"
contagem = 0
if Sistema.FileExists(counterfile) then
Set arquivo = Sistema.GetFile(counterfile)
Set texto = arquivo.OpenAsTextStream(1, -2)
contagem = texto.readline
texto.close
end if
contagem = contagem + 1

Set texto = sistema.CreateTextFile(counterfile, true, false)
texto.writeline contagem
texto.close

texto = contagem & " visit"
if contagem > 1 then texto = texto + "s"
ILIB.FontColor = "ClRed"
ILIB.BrushColor = "clBlack"
ILIB.FontFace = "Comic Sans MS"
ILIB.FontSize = 12
ILIB.FontBold = false
ILIB.width = ILIB.GetTextwidth(texto)+5
ILIB.height = ILIB.GetTextHeight(texto)+5
ILIB.PenColor = "clRed"
ILIB.fBox 1,1,ILIB.WIDTH,ILIB.HEIGHT
ILIB.Textout texto,3,-3
ILIB.FontColor = "ClWhite"
ILIB.FontSize = 7
tamanho = ILIB.GetTextWidth("Example")
ILIB.Textout "Example",ILIB.width/2 - tamanho/2,10
ILIB.PictureBinaryWrite 2, 0, ""
%>