Imports System.Drawing
Imports System.IO

Partial Class UploadImageDialog
    Inherits System.Web.UI.Page

    Dim clipImagePath As String = Server.MapPath(".") + "//clipImage//"
    Dim clipImageFolder As String = "clipImage\"
    Dim originalFileName As String = ""
    Dim clipFileName As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            If Not Directory.Exists(clipImagePath) Then
                Directory.CreateDirectory(clipImagePath)
            End If
        End If
        If CropBox.ImageUrl = "" Then
            CropBox.Visible = False
        Else
            CropBox.Visible = True
        End If
    End Sub

    Protected Sub btnSelect_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSelect.Click
        If SetX.Value <> "" And SetY.Value <> "" And SetW.Value <> "" And SetH.Value <> "" Then
            Dim x As Integer = Convert.ToInt32(SetX.Value)
            Dim y As Integer = Convert.ToInt32(SetY.Value)
            Dim w As Integer = Convert.ToInt32(SetW.Value)
            Dim h As Integer = Convert.ToInt32(SetH.Value)

            Dim image As Image
            clipFileName = "_" + hidFileName.Value
            image = Bitmap.FromFile(clipImagePath + clipFileName)   'HttpContext.Current.Request.PhysicalApplicationPath + "\img\"

            Dim bmp As New Bitmap(w, h, image.PixelFormat)
            Dim g As Graphics = Graphics.FromImage(bmp)
            g.DrawImage(image, New Rectangle(0, 0, w, h), New Rectangle(x, y, w, h), GraphicsUnit.Pixel)
            bmp.Save(clipImagePath + "x" + clipFileName, image.RawFormat)   'Server.MapPath(".")
            image.Dispose()
            bmp.Dispose()
            g.Dispose()

            System.IO.File.Copy(clipImagePath + "x" + clipFileName, clipImagePath + clipFileName, True)
            System.IO.File.Delete(clipImagePath + "x" + clipFileName)

            CropBox.Visible = True
            CropBox.ImageUrl = clipImageFolder + clipFileName + "?temp=" + DateTime.Now.Millisecond.ToString()
        End If

    End Sub

    Protected Sub btnRestore_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRestore.Click
        If hidFileName.Value <> "" Then
            originalFileName = hidFileName.Value
            clipFileName = "_" + originalFileName
            System.IO.File.Copy(clipImagePath + originalFileName, clipImagePath + clipFileName, True)
            CropBox.Visible = True
            CropBox.ImageUrl = clipImageFolder + clipFileName + "?temp=" + DateTime.Now.Millisecond.ToString()
        End If
    End Sub

    Protected Sub btnPreview_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnPreview.Click
        If FileUpload.FileName <> "" Then
            Dim fileExtension As String = FileUpload.FileName.Substring(FileUpload.FileName.LastIndexOf("."))
            If fileExtension.ToLower() <> ".jpg" And fileExtension.ToLower() <> ".jpeg" And fileExtension.ToLower() <> ".gif" Then
                Dim script As String = "<script language='javascript'>alert('請選擇圖片。');history.back();</script>"
                Page.ClientScript.RegisterStartupScript(Me.GetType(), "return", script)
                Return
            End If
            Dim fileSize As Long = FileUpload.PostedFile.ContentLength / 1024
            Dim limitSize As Long = 1024
            If fileSize <= limitSize Then
                hidFileName.Value = System.Guid.NewGuid().ToString() + fileExtension
                originalFileName = hidFileName.Value
                clipFileName = "_" + originalFileName
                FileUpload.SaveAs(clipImagePath + originalFileName)
                FileUpload.SaveAs(clipImagePath + clipFileName)
                CropBox.Visible = True
                CropBox.ImageUrl = clipImageFolder + clipFileName + "?temp=" + DateTime.Now.Millisecond.ToString()
            Else
                Dim Script As String = "<script language='javascript'>alert('您所選擇的檔案大小為" + fileSize.ToString() + "kb，超過了上傳上限" + limitSize.ToString() + "kb！不允許上傳！');history.back();</script>"
                Page.ClientScript.RegisterStartupScript(Me.GetType(), "return", Script)
            End If
        End If
    End Sub

    Protected Sub btnCancel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCancel.Click
        If hidFileName.Value <> "" Then
            originalFileName = hidFileName.Value
            clipFileName = "_" + originalFileName
            System.IO.File.Delete(clipImagePath + originalFileName)
            System.IO.File.Delete(clipImagePath + clipFileName)
        End If
        Page.ClientScript.RegisterStartupScript(Me.GetType(), "close", "window.close();", True)
    End Sub

    Protected Sub btnOK_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnOK.Click
        If hidFileName.Value <> "" Then
            Dim w As Integer = 150, h As Integer = 150
            Dim image As System.Drawing.Image
            originalFileName = hidFileName.Value
            clipFileName = "_" + originalFileName
            image = System.Drawing.Image.FromFile(clipImagePath + clipFileName)

            Dim img As System.Drawing.Bitmap = New System.Drawing.Bitmap(w, h)
            Dim graphic As Graphics = Graphics.FromImage(img)
            graphic.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
            graphic.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality
            graphic.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic
            graphic.DrawImage(image, 0, 0, w, h)
            img.Save(clipImagePath + "x" + clipFileName, image.RawFormat)
            image.Dispose()
            img.Dispose()
            graphic.Dispose()

            System.IO.File.Copy(clipImagePath + "x" + clipFileName, clipImagePath + clipFileName, True)
            System.IO.File.Delete(clipImagePath + "x" + clipFileName)
        End If
        Dim Script As String = "if (navigator.userAgent.indexOf('MSIE')>0) { "
        Script &= " var arg = window.dialogArguments; arg.GetReturnValue('" + clipImageFolder + "/" + clipFileName + "'); window.close(); "
        Script &= " } "
        Script &= " else { "
        Script &= "opener.window.document.getElementById('hidFileName').value='" + clipImageFolder + "/" + clipFileName + "'; window.close();"
        Script &= " } "
        Page.ClientScript.RegisterStartupScript(Me.GetType(), "return", Script, True)

    End Sub
End Class
