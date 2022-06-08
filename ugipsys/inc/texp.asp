<%@ CodePage = 65001 %>
<%
Function ReplaceTest(patrn, replStr, str1)
  Dim regEx               ' Create variables.
'  str1 = "The quick brown fox jumped over the lazy dog."
  Set regEx = New RegExp            ' Create regular expression.
  regEx.Pattern = patrn            ' Set pattern.
  regEx.IgnoreCase = True            ' Make case insensitive.
  regEx.Global = True   ' Set global applicability.
  ReplaceTest = regEx.Replace(str1, replStr)   ' Make replacement.
End Function

	xstr = "<table x=123>"
	response.write ReplaceTest("<table([^>]*)>", "<table$1 summary=""列表資料"">", xstr)	'-- mso-bidi-font-size: 10.0pt;
%>