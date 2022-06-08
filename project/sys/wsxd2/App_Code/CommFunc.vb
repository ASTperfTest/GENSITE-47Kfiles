Imports Microsoft.VisualBasic
Imports System.Xml

Public Class CommFunc
    Shared Function Counter(ByVal mode As String, ByVal mp As String, ByVal filename As String, ByVal Application As System.Web.HttpApplicationState) As String
        '璸计郎
        'Dim CounterFilename As String = filename

        'Dim objFileSystem As Object = CreateObject("Scripting.FileSystemObject")

        '	Application("HitCount") = ""
        'Dim Hit As Long = CLng(Application("HitCount"))
        'If Not IsNumeric(Hit) Or Hit = 0 Then
        '弄璸计郎
        'Dim objReadedTextFile As Object = objFileSystem.OpenTextFile(CounterFilename, 1, 0, 0)
        'Hit = objReadedTextFile.ReadLine
        'objReadedTextFile.Close()
        'Application("HitCount") = Hit
        'End If

        'If mp = 1 Then
        '    Hit = CLng(Hit) + 1
        '    Application.Lock()
        '    Application("HitCount") = Application("HitCount") + 1

        '    Application.UnLock()
        'End If
        'If mp = 1 Then
        'UpdateDB()
        'End If

        '	Hit = 293876
        '糶穝璸计
        '	if Hit mod 20 = 0 then
        'Dim objWritedTextFile As Object = objFileSystem.CreateTextFile(CounterFilename, -1, 0)
        'objWritedTextFile.WriteLine(Hit)
        'objWritedTextFile.Close()
        '	end if

        '瓜┪ゅ家Α	
        'If mode = 0 Then
        '    For i = 1 To CInt(Len(Hit))
        '        Response.Write("<img src=counter/" & Mid(Hit, i, 1) & ".gif alt=" & Mid(Hit, i, 1) & ">")
        '    Next
        'Else
        '    Response.Write(Hit)
        'End If
        'Return Hit
    End Function

    Shared Function getValue(ByVal dataValue As Object) As String
        If dataValue Is DBNull.Value Then
            Return ""
        End If
        Return dataValue
    End Function

    Shared Function nullText(ByVal xNode As XmlNode) As String
        If xNode Is Nothing Then
            Return ""
        End If
        Return xNode.InnerText
    End Function

    Shared Function deAmp(ByVal tempstr As String) As String
        If tempstr Is DBNull.Value Then
            Return ""
        End If
        Dim retString As String
        Dim xs As String = tempstr
        If xs = "" Then
            retString = ""
        Else
            retString = Replace(xs, "&", "&amp;")
        End If
        Return retString
    End Function

    Shared Function pkStr(ByVal s As String, ByVal endchar As String) As String
        Dim retString As String
        If s = "" Then
            retString = "null" & endchar
        Else
            Dim pos As Integer = InStr(s, "'")
            While pos > 0
                s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)
                pos = InStr(pos + 2, s, "'")
            End While
            retString = "'" & s & "'" & endchar
        End If
        Return retString
    End Function

    Shared Function deHTML(ByVal tempstr As Object) As String
        If tempstr Is DBNull.Value Then
            Return ""
        End If
        Dim retString As String
        Dim xs As String = tempstr
        If xs = "" Then
            retString = ""
        Else
            xs = ReplaceText("<[^>]*>", "", xs) '-- <.......>
            xs = Replace(xs, vbCrLf & vbCrLf, "<P>")
            xs = Replace(xs, vbCrLf, "<BR>")
            retString = Replace(xs, Chr(10), "<BR>")
        End If
        Return retString
    End Function

    Shared Function ReplaceText(ByVal patrn As String, ByVal replStr As String, ByVal str1 As String)
        Dim regEx As Regex
        regEx = New Regex(patrn, RegexOptions.IgnoreCase)
        Return regEx.Replace(str1, replStr)   ' Make replacement.
    End Function
End Class
