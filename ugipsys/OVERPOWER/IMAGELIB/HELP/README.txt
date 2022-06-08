Overpower ASP Image Library Readme
==================================

Version: 1.35
Creator: Guba
EMAIL: gubah@sti.com.br
Homepage: http://www.overpower.com.br/ImageLib

The ASP Image Library is a useful asp component to create dynamic 
images using Active Server Pages. For example: You can create an
ASP counter using this component, also thumbnails (see thumbnail asp include on Imagelib homepage)
For now, ASP Image Library supports GIF (Non animated), JPG and BMP
formats.

Installing and registering
==========================

Just copy the Overpower.dll to your Windows\system directory (Different on NT)
and run C:\WINDOWS\SYSTEM\REGSVR32.EXE C:\WINDOWS\SYSTEM\OVERPOWER.DLL

Heres a example on how to use it:
set ILIB = server.createobject("Overpower.ImageLib")
ILIB.width = 100
ILIB.height = 20
ILIB.TextColor = "clBlack" 'Or ILIB.TextColor = "#000000"
ILIB.Textout "Here it is",0,0
ILIB.PictureBinaryWrite 3,50,""

==================
     NOTES
==================

When you receive an error about one of the functions of this new version are not avaliable, it's probable that your webserver is reading from a old version library, to avoid this, unregister your old library (C:\WINDOWS\SYSTEM\regsvr32.exe /D C:\WINDOWS\SYSTEM\OVERPOWER.DLL), delete the DLL, copy the new DLL and then register again...

Any doubt, please email me: gubah@sti.com.br