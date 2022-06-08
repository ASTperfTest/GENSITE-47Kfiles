Overpower ASP Image Library Readme
==================================

Version: 1.512
Date of release: August 25, 1999
Creator: Guba
EMAIL: gubah@sti.com.br
Homepage: http://www.overpower.com.br/ImageLib
Forum: http://www.overpower.com.br/imagelib/forum/forumscript.asp

The ASP Image Library is a useful asp component to create dynamic 
images using Active Server Pages. For example: You can create an
ASP counter using this component, also thumbnails (see thumbnail asp include
 on Imagelib homepage), another example is a custom advertise rotator 
(see rotator include on Imagelib too).
For now, ASP Image Library supports GIF (Non animated), JPG and BMP
formats for input and output and WMF for input.

Installing
==========

If you're using Winzip just open the zip and press Install, otherwise,
just extract all files to a temporary directory and run setup.exe, 
it will do the rest for you :)

Heres a example on how to use it:
set ILIB = server.createobject("Overpower.ImageLib")
ILIB.width = 100
ILIB.height = 20
ILIB.FontColor = "clBlack" 'Or ILIB.FontColor = "#000000"
ILIB.Textout "Here it is",0,0
ILIB.PictureBinaryWrite 3,50,""

==================
     NOTES
==================

By default, Overpower Imagelib autoset the contenttype when you use 
PictureBinaryWrite, but in some cases, it may not (very rare), so, 
just add 
response.contenttype = 'image/gif' or 
response.contenttype = 'image/jpg' or 
response.contenttype = 'image/bmp' 
on the end of the ASP file. Another thing, you should set your system
color to 24bit or it might not work right..

Any doubt, please email me: gubah@sti.com.br