<%
'題目
Class ExamTopic 
	Private Fields
	' 物件初始化
	Private Sub Class_Initialize()
		Set Fields = Server.CreateObject("Scripting.Dictionary")	
        Fields.Add "ct_id", 0
        Fields.Add "et_Id", 0
        Fields.Add "ct_stitle", ""
        Fields.Add "ct_examtype", ""
        Fields.Add "et_tuneId", ""
        Fields.Add "et_correct", NULL        
        
        Set Options = Server.CreateObject("Scripting.Dictionary")
        Fields.Add "Options", Server.CreateObject("Scripting.Dictionary")
    End Sub

	' 物件結束
    Private Sub Class_Terminate()
		Set Fields = Nothing
		Set Options = Nothing
    End Sub

	' 設定指定的屬性
    Public Sub SetField(sName, vValue)
        If Fields.Exists(sName) Then
            If IsObject(vValue) Then
                Set Fields(sName) = vValue
            Else
                Fields(sName) = vValue
            End If
        End If
    End Sub

    ' 取得指定的屬性
    Public Function GetField(sName)
        GetField = Null
        If Fields.Exists(sName) Then
            If IsObject(Fields(sName)) Then
                Set GetField = Fields(sName)
            Else
                GetField = Fields(sName)
            End If
        End If
    End Function
End Class


'選項
Class ExamOption
	Private Fields
	
	' 物件初始化
	Private Sub Class_Initialize()
		Set Fields = Server.CreateObject("Scripting.Dictionary")
			Fields.Add "Optetid", 0        
			Fields.Add "OptTitle", NULL
			Fields.Add "OptAnswer", NULL
			Fields.Add "OptSort", NULL
    End Sub

	' 物件結束
    Private Sub Class_Terminate()
		Set Fields = Nothing
    End Sub

	' 設定指定的屬性
    Public Sub SetField(sName, vValue)
        If Fields.Exists(sName) Then
            If IsObject(vValue) Then
                Set Fields(sName) = vValue
            Else
                Fields(sName) = vValue
            End If
        End If
    End Sub

    ' 取得指定的屬性
    Public Function GetField(sName)
        GetField = Null
        If Fields.Exists(sName) Then
            If IsObject(Fields(sName)) Then
                Set GetField = Fields(sName)
            Else
                GetField = Fields(sName)
            End If
        End If
    End Function
    
End Class
 
%>