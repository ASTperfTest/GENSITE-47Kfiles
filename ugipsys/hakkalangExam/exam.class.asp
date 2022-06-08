<%
' 題目主檔
Class CuDtGeneric
	Private Fields
	
	' 物件初始化
	Private Sub Class_Initialize()
		Set Fields = Server.CreateObject("Scripting.Dictionary")
		Fields.Add "Id", 0
		Fields.Add "Title", NULL
		Fields.Add "TopCat", NULL
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

' 腔調題目類別
Class ExamTopic

	Private Fields

	' 物件初始化
	Private Sub Class_Initialize()
		Set Fields = Server.CreateObject("Scripting.Dictionary")
		Fields.Add "Id", 0
        Fields.Add "CuItemId", 0
        Fields.Add "TuneId", ""
        Fields.Add "Correct", NULL
        
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
    
    ' 檢查資料完整性
    Public Function Validate()
		If Not IsNumeric(Me.GetField("CuItemId")) Then Err.Raise vbObjectError + 1, "", "CuItemId error"
		If Len(Me.GetField("TuneId")) = 0 Then Err.Raise vbObjectError + 1, "", "TuneId error"
    End Function
    
    ' 新增選項
	Public Function AddOption(oExamOption)
		Me.GetField("Options").Add Me.GetField("Options").Count, oExamOption
	End Function
	
End Class

' 題目選項類別
Class ExamOption

	Private Fields

	' 物件初始化
	Private Sub Class_Initialize()
		Set Fields = Server.CreateObject("Scripting.Dictionary")
        Fields.Add "TopicId", 0
        Fields.Add "Title", ""
        Fields.Add "Answer", NULL
        Fields.Add "Sort", NULL
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
    
    ' 檢查資料完整性
    Public Function Validate()
		If Not IsNumeric(Me.GetField("TopicId")) Then Err.Raise vbObjectError + 1, "", "TopicId error"
		If Len(Me.GetField("Title")) = 0 Then Err.Raise vbObjectError + 1, "", "Title error"
		If Not IsNumeric(Me.GetField("Sort")) Then Err.Raise vbObjectError + 1, "", "Sort error"
    End Function

End Class

' 使用者作答主檔類別
Class ExamUserRecord

	Private Fields

	' 物件初始化
	Private Sub Class_Initialize()
		Set Fields = Server.CreateObject("Scripting.Dictionary")
        Fields.Add "Id", 0
        Fields.Add "UserId", 0
        Fields.Add "ExamType", NULL
        Fields.Add "CreateTime", NULL
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

' 使用者回答明細類別
Class ExamUserAnswer

	Private Fields

	' 物件初始化
	Private Sub Class_Initialize()
		Set Fields = Server.CreateObject("Scripting.Dictionary")
        Fields.Add "Id", 0
        Fields.Add "RecordId", 0
        Fields.Add "AnswerId", NULL
        Fields.Add "IsCorrect", False
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