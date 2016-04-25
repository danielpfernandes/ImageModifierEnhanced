

Public Class frmImageModifier

    Private Sub btnOpen_Click(sender As Object, e As EventArgs) Handles btnOpen.Click
        Dim open As New OpenFileDialog
        open.Title = "Image Location"
        open.Filter = "JPEG Image |*.jpg|All filed (*.*)|*.*"
        If open.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim myImage = New Bitmap(open.FileName, True)
            picOriginal.BorderStyle = BorderStyle.None
            picOriginal.Image = myImage
            picConverted.Image = Nothing
        End If
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If picConverted.Image IsNot Nothing Then
            Dim save As New SaveFileDialog
            save.Title = "Save Folder"
            save.Filter = "JPEG Image |*.jpg|All filed (*.*)|*.*"
            If save.ShowDialog() = Windows.Forms.DialogResult.OK Then
                picConverted.Image.Save(save.FileName)
                MessageBox.Show("The converted image has benn successfully saved!")
            End If
        Else
            MessageBox.Show("There is no image to be saved.", "Error")
        End If
    End Sub

    Private Sub convertMonochrome(image As Bitmap)

        Dim r, g, b, average As Integer
        Dim newColor As Color
        For y = 0 To image.Height - 1
            For x = 0 To image.Width - 1
                Dim pixelColor As Color = image.GetPixel(x, y)
                r = CInt(pixelColor.R)
                g = CInt(pixelColor.G)
                b = CInt(pixelColor.B)
                average = CInt((r + g + b) / 3)
                If (average > 128) Then
                    r = 255
                    g = 255
                    b = 255
                Else
                    r = 0
                    g = 0
                    b = 0
                End If
                newColor = Color.FromArgb(r, g, b)
                image.SetPixel(x, y, newColor)
            Next
        Next
        picConverted.Image = image
        picConverted.BorderStyle = BorderStyle.None
    End Sub

    Private Sub convertGrayscale(image As Bitmap)
        Dim newColor As Color
        Dim r, g, b, average As Integer
        For y = 0 To image.Height - 1
            For x = 0 To image.Width - 1
                Dim pixelColor As Color = image.GetPixel(x, y)
                r = CInt(pixelColor.R)
                g = CInt(pixelColor.G)
                b = CInt(pixelColor.B)
                average = CInt((r + g + b) / 3)
                newColor = Color.FromArgb(average, average, average)
                image.SetPixel(x, y, newColor)
            Next
        Next
        picConverted.Image = image
        picConverted.BorderStyle = BorderStyle.None
    End Sub

    Private Sub txtQuit_Click(sender As Object, e As EventArgs) Handles txtQuit.Click
        Close()
    End Sub

    Private Sub convertSepia(image As Bitmap)
        Dim newColor As Color
        Dim r, g, b, average As Integer
        For y = 0 To image.Height - 1
            For x = 0 To image.Width - 1
                Dim pixelColor As Color = image.GetPixel(x, y)
                r = CInt(pixelColor.R)
                g = CInt(pixelColor.G)
                b = CInt(pixelColor.B)
                average = CInt((r + g + b) / 3)
                If average > 145 Then
                    average = 145
                End If
                newColor = Color.FromArgb(average + 80, average + 40, average)
                image.SetPixel(x, y, newColor)
            Next
        Next
        picConverted.Image = image
        picConverted.BorderStyle = BorderStyle.None
    End Sub

    Private Sub convertFactor(image As Bitmap)
        Try
            Dim newColor As Color
            Dim r, g, b, average As Integer
            Dim shades = InputBox("Enter with the shades of gray", "Conversion Factor", "")
            Dim factor = CInt(255 / (CInt(shades) - 1))
            For y = 0 To image.Height - 1
                For x = 0 To image.Width - 1
                    Dim pixelColor As Color = image.GetPixel(x, y)
                    r = CInt(pixelColor.R)
                    g = CInt(pixelColor.G)
                    b = CInt(pixelColor.B)
                    average = CInt((r + g + b) / 3)
                    average = CInt((average / factor) + 0.5) * factor
                    If average > 255 Then
                        average = 255
                    End If
                    newColor = Color.FromArgb(average, average, average)
                    image.SetPixel(x, y, newColor)
                Next
            Next
            picConverted.Image = image
            picConverted.BorderStyle = BorderStyle.None
        Catch ex As InvalidCastException
            MessageBox.Show("Enter a valid format", "Error")
        Catch ex As OverflowException
            MessageBox.Show("Try a smaller value", "Error")
        Catch ex As Exception
            MessageBox.Show(ex.GetType.ToString)
        End Try
    End Sub

    Private Sub btnConvert_Click(sender As Object, e As EventArgs) Handles btnConvert.Click
        If picOriginal.Image IsNot Nothing Then
            Dim myImage = New Bitmap(picOriginal.Image)
            If cboFilter.SelectedItem IsNot Nothing Then

                Select Case cboFilter.SelectedItem.ToString
                    Case "   Monochrome"
                        convertMonochrome(myImage)
                    Case "   Gray Scale"
                        convertGrayscale(myImage)
                    Case "   Conversion Factor"
                        convertFactor(myImage)
                    Case "   Sepia"
                        convertSepia(myImage)
                End Select
            Else
                MessageBox.Show("Please select a filter to work with.", "Error")
            End If
        Else
            MessageBox.Show("Please select a picture to convert.", "Error")
        End If
    End Sub
End Class
