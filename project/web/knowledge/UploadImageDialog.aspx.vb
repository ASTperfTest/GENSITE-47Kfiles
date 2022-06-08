Imports System.Drawing
Imports System.IO

Partial Class UploadImageDialog
    Inherits System.Web.UI.Page
    Public SizeLimit As Integer = 1024 * 1024
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
        Dim x As Integer
        Dim y As Integer
        Dim w As Integer
        Dim h As Integer
        If Integer.TryParse(SetX.Value, x) And Integer.TryParse(SetY.Value, y) And Integer.TryParse(SetW.Value, w) And Integer.TryParse(SetH.Value, h) Then
            If (x < 0 Or y < 0 Or w < 0 Or h < 0) Then Exit Sub


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
            System.IO.File.Delete(clipImagePath + clipFileName)
            System.IO.File.Copy(clipImagePath + "x" + clipFileName, clipImagePath + clipFileName, True)
            System.IO.File.Delete(clipImagePath + "x" + clipFileName)

            CropBox.Visible = True
            CropBox.ImageUrl = clipImageFolder + clipFileName + "?temp=" + DateTime.Now.Millisecond.ToString()
        End If

        SetX.Value = ""
        SetY.Value = ""
        SetW.Value = ""
        SetH.Value = ""
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

    Protected Sub ddlImagePixel_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlImagePixel.SelectedIndexChanged
        ReduceImage(True)
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
            If fileSize <= SizeLimit / 1024 Then
                hidFileName.Value = System.Guid.NewGuid().ToString() + fileExtension
                originalFileName = hidFileName.Value
                clipFileName = "_" + originalFileName
                FileUpload.SaveAs(clipImagePath + originalFileName)
                FileUpload.SaveAs(clipImagePath + clipFileName)
                CropBox.Visible = True
                CropBox.ImageUrl = clipImageFolder + clipFileName + "?temp=" + DateTime.Now.Millisecond.ToString()
            Else
                Dim Script As String = "<script language='javascript'>alert('您所選擇的檔案大小為" + fileSize.ToString() + "kb，超過了上傳上限" + (SizeLimit / 1024).ToString() + "kb！不允許上傳！');history.back();</script>"
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
        Dim image As System.Drawing.Image
        Dim width, height As Integer
        If hidFileName.Value <> "" Then
            originalFileName = hidFileName.Value
            clipFileName = "_" + originalFileName
            image = System.Drawing.Image.FromFile(clipImagePath + clipFileName)
            width = image.Width
            height = image.Height
            image.Dispose()
            If width > 800 Or height > 600 Then
                ReduceImage(False)
            End If
        End If
        Dim Script As String = "if (navigator.userAgent.indexOf('MSIE')>0) { "
        Script &= " var arg = window.dialogArguments; arg.GetReturnValue('" + clipImageFolder + "/" + clipFileName + "', '" + txtHint.Text.Trim() + "'); window.close(); "
        Script &= " } "
        Script &= " else { "
        Script &= "opener.window.document.getElementById('ctl00_ContentPlaceHolder1_hidFileName').value+='" + clipImageFolder + "/" + clipFileName + "^';"
        Script &= " opener.window.document.getElementById('ctl00_ContentPlaceHolder1_hidFileContent').value += '" + txtHint.Text.Trim() + "^'; window.close();"
        Script &= " } "
        Page.ClientScript.RegisterStartupScript(Me.GetType(), "return", Script, True)

    End Sub

    Function ReduceImage(ByVal isSelectImagePixel As Boolean)
        'ASP.NET 縮圖功能
        Dim w, h As Integer
        Dim image As System.Drawing.Image
        'Dim anewimage As System.Drawing.Bitmap
        'Dim callb As System.Drawing.Image.GetThumbnailImageAbort
        Dim width, height As Integer
        If hidFileName.Value <> "" Then
            clipFileName = "_" + hidFileName.Value

            image = System.Drawing.Image.FromFile(clipImagePath + clipFileName)
            width = image.Width
            height = image.Height

            If isSelectImagePixel Then
                If ddlImagePixel.SelectedValue <> 0 Then
                    w = Convert.ToInt32(ddlImagePixel.SelectedValue)
                    h = (w / width) * (height)
                Else
                    w = width
                    h = height
                End If

                If Not (width < w And height < h) Then
                    If width > height Then
                        h = w * height / width
                    Else
                        w = h * width / height
                    End If
                End If
            Else
                w = 800
                h = 600
            End If


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

            CropBox.Visible = True
            CropBox.ImageUrl = clipImageFolder + clipFileName + "?temp=" + DateTime.Now.Millisecond.ToString()
        End If
    End Function
End Class
