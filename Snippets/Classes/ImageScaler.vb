Imports System.Drawing

''' <summary>
''' Singleton toolbox class that contains functions to resize images.
''' </summary>
Class ImageScaler
    ''' <summary>
    ''' Resizes the width of a bitmap object.
    ''' </summary>
    Public Shared Function ResizeImageWidth(scalefactor As Double, intType As Integer, bmpsource As System.Drawing.Bitmap) As Bitmap
        Dim intx As Integer = CInt(bmpsource.Width * scalefactor),
            inty As Integer = bmpsource.Height

        Dim bmpout As Bitmap = New Bitmap(intx, inty)

        Using g As Graphics = Graphics.FromImage(bmpout)
            Select Case intType
                Case 0
                    g.InterpolationMode = Drawing.Drawing2D.InterpolationMode.Default
                Case 1
                    g.InterpolationMode = Drawing.Drawing2D.InterpolationMode.High
                Case 2
                    g.InterpolationMode = Drawing.Drawing2D.InterpolationMode.HighQualityBilinear
                Case 3
                    g.InterpolationMode = Drawing.Drawing2D.InterpolationMode.HighQualityBicubic
            End Select

            g.DrawImage(bmpsource, 0, 0, intx, inty)
        End Using


        Return bmpout
    End Function
End Class