Public Class FormI_RLFadePreview
    Dim ini_glb As New IniFileApi
    Dim ini_sys As New IniFileApi

    Dim aspect As Double = 0
    Dim aspect_adjust_height As Integer = 0
    Dim Fade_Base_Resolution_Width As Integer = 0
    Dim Fade_Base_Resolution_Height As Integer = 0

    Dim base_fade_aspect As Double = 1
    Dim l4_current_frame As Integer = 0
    Dim WithEvents l4_timer As New Timer With {.Enabled = False}

    Dim b As Bitmap
    Dim g As Graphics

    Dim l1 As Bitmap = Nothing
    Dim l2 As Bitmap = Nothing
    Dim l3s As Bitmap = Nothing
    Dim l3a As Bitmap = Nothing
    Dim l4() As Bitmap = Nothing
    Dim lt As New List(Of Bitmap)

    Dim layer_params(9) As l_param
    Dim text_params As New List(Of txt_param)
    Structure l_param
        Dim c As Color
        Dim align As String
        Dim bounds As Rectangle
        Dim fps As Integer
    End Structure
    Structure txt_param
        Dim type As String
        Dim txt As String
        Dim f As Font
        Dim c As Color
        Dim h_align As String
        Dim v_align As String
        Dim bounds As Rectangle
    End Structure

    'TODO case 'above' in drawLayer for layer4 additional position
    'TODO maximize, then minimize (or move) - titlebar disappears, because of suspend layout in resizeBegin form

    Public Sub New(sys As String, rom As String, gameData() As String)
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        'Total form height = SystemInformation.CaptionHeight + SystemInformation.Border3DSize.Height + yourFormInstance.Height + SystemInformation.Border3DSize.Height
        aspect = Screen.PrimaryScreen.Bounds.Height / Screen.PrimaryScreen.Bounds.Width
        aspect_adjust_height = SystemInformation.CaptionHeight

        ini_glb.path = Class1.HyperlaunchPath + "\Settings\Global RocketLauncher.ini"
        ini_sys.path = Class1.HyperlaunchPath + "\Settings\" + sys + "\RocketLauncher.ini"

        Fade_Base_Resolution_Width = getFromIni("Fade", "Fade_Base_Resolution_Width", 1920)
        Fade_Base_Resolution_Height = getFromIni("Fade", "Fade_Base_Resolution_Height", 1080)
        layer_params(1).c = ColorTranslator.FromHtml("#" + getFromIni("Fade", "Fade_Layer_1_Color", "FF000000")) 'argb
        Dim Fade_Layer_1_Align_Image = getFromIni("Fade", "Fade_Layer_1_Align_Image", "Align to Top Left")
        Dim Fade_Layer_2_Prefix = getFromIni("Fade", "Fade_Layer_2_Prefix", "Layer 2")
        layer_params(2).align = getFromIni("Fade", "Fade_Layer_2_Alignment", "Bottom Right Corner")
        layer_params(2).bounds.X = getFromIni("Fade", "Fade_Layer_2_X", 300)
        layer_params(2).bounds.Y = getFromIni("Fade", "Fade_Layer_2_Y", 300)
        layer_params(2).bounds.Width = getFromIni("Fade", "Fade_Layer_2_W", 0)
        layer_params(2).bounds.Height = getFromIni("Fade", "Fade_Layer_2_H", 0)
        Dim Fade_Layer_2_Adjust = getFromIni("Fade", "Fade_Layer_2_Adjust", 1)
        Dim Fade_Layer_2_Padding = getFromIni("Fade", "Fade_Layer_2_Padding", 0)
        Dim Fade_Layer_3_Static_Prefix = getFromIni("Fade", "Fade_Layer_3_Static_Prefix", "Layer 3")
        layer_params(3).align = getFromIni("Fade", "Fade_Layer_3_Static_Alignment", "Center")
        layer_params(3).bounds.X = getFromIni("Fade", "Fade_Layer_3_Static_X", 300)
        layer_params(3).bounds.Y = getFromIni("Fade", "Fade_Layer_3_Static_Y", 300)
        layer_params(3).bounds.Width = getFromIni("Fade", "Fade_Layer_3_Static_W", 0)
        layer_params(3).bounds.Height = getFromIni("Fade", "Fade_Layer_3_Static_H", 0)
        Dim Fade_Layer_3_Static_Adjust = getFromIni("Fade", "Fade_Layer_3_Static_Adjust", 0.75)
        Dim Fade_Layer_3_Static_Padding = getFromIni("Fade", "Fade_Layer_3_Static_Padding", 0)
        layer_params(4).align = getFromIni("Fade", "Fade_Layer_3_Alignment", "Center")
        layer_params(4).bounds.X = getFromIni("Fade", "Fade_Layer_3_X", 300)
        layer_params(4).bounds.Y = getFromIni("Fade", "Fade_Layer_3_Y", 300)
        layer_params(4).bounds.Width = getFromIni("Fade", "Fade_Layer_2_W", 0)
        layer_params(4).bounds.Height = getFromIni("Fade", "Fade_Layer_2_H", 0)
        Dim Fade_Layer_3_Adjust = getFromIni("Fade", "Fade_Layer_3_Adjust", 0.75)
        Dim Fade_Layer_3_Padding = getFromIni("Fade", "Fade_Layer_3_Padding", 0)
        Dim Fade_Layer_3_Speed = getFromIni("Fade", "Fade_Layer_3_Speed", 750)
        Dim Fade_Layer_3_Animation = getFromIni("Fade", "Fade_Layer_3_Animation", "DefaultFadeAnimation")
        Dim Fade_Layer_3_7z_Animation = getFromIni("Fade", "Fade_Layer_3_7z_Animation", "DefaultFadeAnimation")
        Dim Fade_Layer_3_Show_7z_Progress = getFromIni("Fade", "Fade_Layer_3_Show_7z_Progress", "true")
        Dim Fade_Layer_3_Type = getFromIni("Fade", "Fade_Layer_3_Type", "imageandbar")
        Dim Fade_Layer_3_Repeat = getFromIni("Fade", "Fade_Layer_3_Repeat", 1)
        layer_params(5).align = getFromIni("Fade", "Fade_Layer_4_Pos", "Above Layer 3 - Left")
        layer_params(5).bounds.X = getFromIni("Fade", "Fade_Layer_4_X", 100)
        layer_params(5).bounds.Y = getFromIni("Fade", "Fade_Layer_4_Y", 100)
        layer_params(5).bounds.Width = getFromIni("Fade", "Fade_Layer_4_W", 0)
        layer_params(5).bounds.Height = getFromIni("Fade", "Fade_Layer_4_H", 0)
        Dim Fade_Layer_4_Adjust = getFromIni("Fade", "Fade_Layer_4_Adjust", 0.75)
        Dim Fade_Layer_4_Padding = getFromIni("Fade", "Fade_Layer_4_Padding", 0)
        Dim Fade_Layer_4_FPS = getFromIni("Fade", "Fade_Layer_4_FPS", 10)
        Dim Fade_Animated_Gif_Transparent_Color = getFromIni("Fade", "Fade_Animated_Gif_Transparent_Color", "FFFFFF")
        base_fade_aspect = Fade_Base_Resolution_Width / Fade_Base_Resolution_Height

        Dim txt_ind = 0
        Dim Fade_Font = getFromIni("Fade", "Fade_Font", "Arial")
        Dim arr_order = getFromIni("Fade", "Fade_Rom_Info_Order", "Description|SystemName|Year|Publisher|Genre|Rating").ToLower.Replace("systemname", "system_name").Split({"|"c}, StringSplitOptions.RemoveEmptyEntries)
        For Each txt_name In arr_order
            txt_ind += 1
            Dim t_type = getFromIni("Fade", "Fade_Rom_Info_" + txt_name, "text").Trim.ToLower
            If t_type = "disabled" Then Continue For

            'opt format - every element is optional: w, h, x, y, c - color, s - size, r - render hint
            Dim param As New txt_param
            Dim t_options = getFromIni("Fade", "Fade_Rom_Info_Text_" + txt_ind.ToString + "_Options", "w310 x165 y960|1665 cFFE1E1E1 r4 s66 Left Regular").Trim.ToLower

            Select Case txt_name.ToLower
                Case "description"
                    param.txt = gameData(0)
                Case "system_name"
                    param.txt = sys
                Case "year"
                    param.txt = gameData(1)
                Case "genre"
                    param.txt = gameData(2)
                Case "publisher"
                    param.txt = gameData(3)
                Case "rating"
                    param.txt = gameData(4)
            End Select

            If t_type = "text with label" Or t_type = "filtered text with label" Then
                If param.txt.Trim <> "" Then
                    param.txt = txt_name.Substring(0, 1).ToUpper + txt_name.Substring(1) + ": " + param.txt
                End If
            End If

            For Each opt In t_options.ToUpper.Split(" "c)
                If opt.StartsWith("X") Then param.bounds.X = CInt(opt.Substring(1).Split({"|"c}, StringSplitOptions.RemoveEmptyEntries)(0))
                If opt.StartsWith("Y") Then param.bounds.Y = CInt(opt.Substring(1).Split({"|"c}, StringSplitOptions.RemoveEmptyEntries)(0))
                If opt.StartsWith("W") Then param.bounds.Width = CInt(opt.Substring(1).Split({"|"c}, StringSplitOptions.RemoveEmptyEntries)(0))
                If opt.StartsWith("H") Then param.bounds.Height = CInt(opt.Substring(1).Split({"|"c}, StringSplitOptions.RemoveEmptyEntries)(0))
                If opt.StartsWith("S") Then param.f = New Font(Fade_Font, CSng(opt.Substring(1)), GraphicsUnit.Pixel)
            Next

            Dim tb As New Bitmap(1, 1)
            If t_type = "image" Then
                Dim subfolder = txt_name
                If subfolder.ToLower = "system_name" Then subfolder = "system"
                tb = getImage(sys, rom, subfolder, param.txt)
                param.type = "IMAGE"
            Else
                Dim tg = Graphics.FromImage(tb)
                Dim ts = tg.MeasureString(param.txt, param.f)
                tb = New Bitmap(CInt(ts.Width), CInt(ts.Height))
                tg = Graphics.FromImage(tb)
                tg.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias
                tg.DrawString(param.txt, param.f, Brushes.White, 1, 1)
                param.type = "TEXT"
            End If

            If tb IsNot Nothing Then
                'tb.Save("D:\Documents\My_Progs\HyperspinChecker\RL_FADE_TEST_" + txt_name + ".png", Imaging.ImageFormat.Png)
                If param.bounds.Width <> 0 Or param.bounds.Height <> 0 Then
                    'Dim w, h As Integer
                    'Dim aspect = tb.Width / tb.Height
                    'If param.bounds.Width <> 0 And param.bounds.Width < tb.Width Then w = param.bounds.Width Else w = tb.Width
                    'If param.bounds.Height <> 0 And param.bounds.Height < tb.Height Then h = param.bounds.Height Else h = tb.Height

                    'Dim ttb As Bitmap
                    'Dim calculated_w = CInt(h * aspect)
                    'Dim calculated_h = CInt(w / aspect)
                    'If calculated_w <= w Then
                    '    ttb = New Bitmap(calculated_w, h)
                    'Else
                    '    ttb = New Bitmap(w, calculated_h)
                    'End If

                    'Dim g = Graphics.FromImage(ttb)
                    ''set high quality resizing
                    'g.CompositingMode = Drawing2D.CompositingMode.SourceCopy
                    'g.CompositingQuality = Drawing2D.CompositingQuality.HighQuality
                    'g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                    'g.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
                    'g.PixelOffsetMode = Drawing2D.PixelOffsetMode.HighQuality
                    'g.DrawImage(tb, New Rectangle(0, 0, ttb.Width, ttb.Height))
                    'tb = ttb
                End If

                'tb.Save("D:\Documents\My_Progs\HyperspinChecker\RL_FADE_TEST_" + txt_name + "_resized.png", Imaging.ImageFormat.Png)
                lt.Add(tb)
                text_params.Add(param)
            End If
        Next

        Dim fade_path As String = ""
        fade_path = Class1.HyperlaunchPath + "\Media\Fade\" + sys + "\" + rom
        If Not FileIO.FileSystem.DirectoryExists(fade_path) Then fade_path = Class1.HyperlaunchPath + "\Media\Fade\" + sys + "\_Default"
        If Not FileIO.FileSystem.DirectoryExists(fade_path) Then fade_path = Class1.HyperlaunchPath + "\Media\Fade\_Default"
        If Not FileIO.FileSystem.DirectoryExists(fade_path) Then MsgBox("Fade path not found.") : Exit Sub

        b = New Bitmap(Fade_Base_Resolution_Width, Fade_Base_Resolution_Height)
        g = Graphics.FromImage(b)
        'g.CompositingQuality = Drawing2D.CompositingQuality.HighQuality
        'g.InterpolationMode = Drawing2D.InterpolationMode.
        'g.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
        'g.PixelOffsetMode = Drawing2D.PixelOffsetMode.HighQuality

        'Layer 1
        l1 = getImage(sys, rom, "", "layer 1")
        'Layer 2
        If Fade_Layer_2_Prefix <> "" Then l2 = getImage(sys, rom, "", Fade_Layer_2_Prefix)
        'Layer 3 static
        If Fade_Layer_3_Static_Prefix <> "" Then l3s = getImage(sys, rom, "", Fade_Layer_3_Static_Prefix)
        'Layer 3 animated - progress
        l3a = getImage(sys, rom, "", "Layer 3")

        ':ayer 4 - animation
        Dim arr4 = {"Layer 4*.png", "Layer 4*.jpg", "Layer 4*.gif", "Layer 4*.bmp", "Layer 4*.tif"}
        Dim layer4 = FileIO.FileSystem.GetFiles(fade_path, FileIO.SearchOption.SearchTopLevelOnly, arr4).ToArray
        If layer4.Count > 0 Then
            Array.Sort(layer4)
            ReDim l4(layer4.Count - 1)
            For n As Integer = 0 To layer4.GetUpperBound(0)
                l4(n) = New Bitmap(layer4(n))
            Next

            l4_timer.Interval = Fade_Layer_4_FPS
            l4_timer.Enabled = True
        End If

        Me.BackgroundImage = b
        'Me.BackgroundImageLayout = ImageLayout.Stretch
        Me.BackgroundImageLayout = ImageLayout.None
        drawAllLayers()
    End Sub

    Sub drawAllLayers()
        Dim currentAspect = Me.Width / Me.Height

        'TEST NEW
        g.Clear(layer_params(1).c)
        If l1 IsNot Nothing Then g.DrawImage(l1, 0, 0, Me.Width, Me.Height)
        drawLayer(l2, layer_params(2).align, New Rectangle(layer_params(2).bounds.X, layer_params(2).bounds.Y, layer_params(2).bounds.Width, layer_params(2).bounds.Height))
        drawLayer(l3s, layer_params(3).align, New Rectangle(layer_params(3).bounds.X, layer_params(3).bounds.Y, layer_params(3).bounds.Width, layer_params(3).bounds.Height))
        drawLayer(l3a, layer_params(4).align, New Rectangle(layer_params(4).bounds.X, layer_params(4).bounds.Y, layer_params(4).bounds.Width, layer_params(4).bounds.Height))
        If l4 IsNot Nothing Then drawLayer(l4(l4_current_frame), layer_params(5).align, New Rectangle(layer_params(5).bounds.X, layer_params(5).bounds.Y, layer_params(5).bounds.Width, layer_params(5).bounds.Height), True)

        'INFO Layers
        For n As Integer = 0 To text_params.Count - 1
            Dim w, h As Integer
            Dim newBounds As Rectangle = text_params(n).bounds
            Dim scale = New PointF(CSng(Me.Width / Fade_Base_Resolution_Width), CSng(Me.Height / Fade_Base_Resolution_Height))
            Dim scale_factor As Double = 1

            If text_params(n).type = "TEXT" Then
                w = CInt(lt(n).Width * scale.X)
                h = CInt(lt(n).Height * scale.Y)
            ElseIf text_params(n).bounds.Width <> 0 And text_params(n).bounds.Height <> 0 Then
                w = CInt(text_params(n).bounds.Width * scale.X)
                h = CInt(text_params(n).bounds.Height * scale.Y)
            ElseIf text_params(n).bounds.Width <> 0 Then
                scale_factor = text_params(n).bounds.Width / lt(n).Width * scale.X
                w = CInt(lt(n).Width * scale_factor)
                h = CInt(lt(n).Height * scale_factor)
            ElseIf text_params(n).bounds.Height <> 0 Then
                scale_factor = text_params(n).bounds.Height / lt(n).Height * scale.Y
                w = CInt(lt(n).Width * scale_factor)
                h = CInt(lt(n).Height * scale_factor)
            Else
                w = CInt(lt(n).Width * scale.X)
                h = CInt(lt(n).Height * scale.Y)
            End If

            'Draw image
            newBounds = rescaleBounds(newBounds, Scale)
            g.DrawImage(lt(n), newBounds.X, newBounds.Y, w, h)
        Next

        Me.Refresh()
        Exit Sub
        'END TEST

        g.Clear(layer_params(1).c)
        If l1 IsNot Nothing Then g.DrawImage(l1, 0, 0, b.Width, b.Height)
        drawLayer(l2, layer_params(2).align, New Rectangle(layer_params(2).bounds.X, layer_params(2).bounds.Y, layer_params(2).bounds.Width, layer_params(2).bounds.Height))
        drawLayer(l3s, layer_params(3).align, New Rectangle(layer_params(3).bounds.X, layer_params(3).bounds.Y, layer_params(3).bounds.Width, layer_params(3).bounds.Height))
        drawLayer(l3a, layer_params(4).align, New Rectangle(layer_params(4).bounds.X, layer_params(4).bounds.Y, layer_params(4).bounds.Width, layer_params(4).bounds.Height))
        If l4 IsNot Nothing Then drawLayer(l4(l4_current_frame), layer_params(5).align, New Rectangle(layer_params(5).bounds.X, layer_params(5).bounds.Y, layer_params(5).bounds.Width, layer_params(5).bounds.Height), True)

        'Info layers
        For n As Integer = 0 To text_params.Count - 1
            'ASPECT CORRECTION
            Dim new_h = CInt(lt(n).Height / base_fade_aspect * currentAspect)

            'Draw image
            g.DrawImage(lt(n), text_params(n).bounds.X, text_params(n).bounds.Y, lt(n).Width, new_h)
        Next

        Me.Refresh()
    End Sub
    Function rescaleBounds(r As Rectangle, scale As PointF) As Rectangle
        Return New Rectangle(CInt(Math.Round(r.X * scale.X)), CInt(Math.Round(r.Y * scale.Y)), CInt(Math.Round(r.Width * scale.X)), CInt(Math.Round(r.Height * scale.Y)))
    End Function
    Sub drawLayer(l As Bitmap, align As String, bounds As Rectangle, Optional dont_rescale As Boolean = False)
        If l Is Nothing Then Exit Sub

        'TEST NEW ROUTINE
        Dim scale = New PointF(CSng(Me.Width / Fade_Base_Resolution_Width), CSng(Me.Height / Fade_Base_Resolution_Height))
        bounds = rescaleBounds(bounds, scale)

        'Dim center_X = CInt(b.Width / 2)
        'Dim center_Y = CInt(b.Height / 2)
        Dim center_X = CInt(Me.Width / 2)
        Dim center_Y = CInt(Me.Height / 2)

        Select Case align.ToUpper
            Case "Stretch and Lose Aspect".ToUpper
                g.DrawImage(l, 0, 0, b.Width, b.Height)
            Case "Stretch and Keep Aspect".ToUpper
                Dim scale_factor_to_fill_width = b.Width / l.Width
                Dim scale_factor_to_fill_height = b.Height / l.Height
                Dim scale_factor As Double = 0
                If l.Height * scale_factor_to_fill_height > b.Height Then
                    scale_factor = scale_factor_to_fill_width
                Else
                    scale_factor = scale_factor_to_fill_height
                End If
                g.DrawImage(l, 0, 0, CInt(b.Width * scale_factor), CInt(b.Height * scale_factor))
            Case "Center".ToUpper
                Dim w = CInt(l.Width * scale.X)
                Dim h = CInt(l.Height * scale.Y)
                'g.DrawImage(l, center_X - CInt(l.Width / 2), center_Y - CInt(l.Height / 2), l.Width, l.Height)
                g.DrawImage(l, center_X - CInt(w / 2), center_Y - CInt(h / 2), w, h)
            Case "Top Left".ToUpper
                g.DrawImage(l, 0, 0, l.Width, l.Height)
            Case "Top Right".ToUpper
                g.DrawImage(l, b.Width - l.Width, 0, l.Width, l.Height)
            Case "Bottom Left".ToUpper
                g.DrawImage(l, 0, b.Height - l.Height, l.Width, l.Height)
            Case "Bottom Right".ToUpper
                g.DrawImage(l, b.Width - l.Width, b.Height - l.Height, l.Width, l.Height)
            Case "Top Center".ToUpper
                g.DrawImage(l, center_X - CInt(l.Width / 2), 0, l.Width, l.Height)
            Case "Bottom Center".ToUpper
                g.DrawImage(l, center_X - CInt(l.Width / 2), b.Height - l.Height, l.Width, l.Height)
            Case "Left Center".ToUpper
                g.DrawImage(l, 0, center_Y - CInt(l.Height / 2), l.Width, l.Height)
            Case "Right Center".ToUpper
                g.DrawImage(l, b.Width - l.Width, center_Y - CInt(l.Height / 2), l.Width, l.Height)
            Case Else
                Dim w = bounds.Width
                Dim h = bounds.Height
                If w = 0 Then w = l.Width
                If h = 0 Then h = l.Height

                If Not dont_rescale Then
                    w = CInt(w * scale.X)
                    h = CInt(h * scale.Y)
                End If

                g.DrawImage(l, bounds.X, bounds.Y, w, h)
        End Select
    End Sub

    Sub l4_timer_tick() Handles l4_timer.Tick
        l4_current_frame = l4_current_frame + 1
        If l4_current_frame > l4.GetUpperBound(0) Then l4_current_frame = 0
        drawAllLayers()
    End Sub

    Function getImage(sys As String, rom As String, subfolder As String, prefix As String) As Bitmap
        Dim fade_path As New List(Of String)
        fade_path.Add(Class1.HyperlaunchPath + "\Media\" + subfolder + "\_Default")
        fade_path.Add(Class1.HyperlaunchPath + "\Media\Fade\" + sys + "\" + rom + "\" + subfolder)
        fade_path.Add(Class1.HyperlaunchPath + "\Media\Fade\" + sys + "\_Default\" + subfolder)
        fade_path.Add(Class1.HyperlaunchPath + "\Media\Fade\_Default\" + subfolder)

        For Each path In fade_path
            If FileIO.FileSystem.DirectoryExists(path) Then
                Dim arr = FileIO.FileSystem.GetFiles(path, FileIO.SearchOption.SearchTopLevelOnly, {prefix + "*.png", prefix + "*.jpg", prefix + ".gif", prefix + "*.bmp"})
                If arr.Count > 0 Then
                    Dim r As New Random()
                    Return New Bitmap(arr(r.Next(arr.Count)))
                End If
            End If
        Next
        Return Nothing
    End Function

    Overloads Function getFromIni(section As String, key As String, Def As String) As String
        Dim t = ini_sys.IniReadValue(section, key).Trim
        If t = "" Or t.ToLower = "use_global" Then t = ini_glb.IniReadValue(section, key).Trim
        If t = "" Then Return Def Else Return t
    End Function
    Overloads Function getFromIni(section As String, key As String, Def As Integer) As Integer
        Dim t = getFromIni(section, key, "DEFAULT")
        If t = "DEFAULT" Then Return Def
        If t.Contains("|") Then t = t.Split({"|"c}, StringSplitOptions.None)(0)
        If Not IsNumeric(t) Then Return Def
        Return CInt(t)
    End Function
    Overloads Function getFromIni(section As String, key As String, Def As Double) As Double
        Dim t = getFromIni(section, key, "DEFAULT")
        If t = "DEFAULT" Then Return Def
        If Not IsNumeric(t) Then Return Def
        Return Double.Parse(t)
    End Function

    'Resize form - keep aspect
    Private Sub FormI_RLFadePreview_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        If Me.FormBorderStyle = FormBorderStyle.None Then Exit Sub
        Height = CInt(Width * aspect) + aspect_adjust_height
        'Height = CInt(Width * 0.5625) 'keep 16/9 aspect, 0.5625 = 1080/1920
    End Sub
    Private Sub FormI_RLFadePreview_ResizeBegin(sender As Object, e As EventArgs) Handles Me.ResizeBegin
        SuspendLayout()
    End Sub
    Private Sub FormI_RLFadePreview_ResizeEnd(sender As Object, e As EventArgs) Handles Me.ResizeEnd
        ResumeLayout()
    End Sub

    'Fullscreen
    Dim storedBounds As Rectangle
    Private Sub FormI_RLFadePreview_DoubleClick(sender As Object, e As EventArgs) Handles Me.DoubleClick
        If Me.FormBorderStyle = FormBorderStyle.None Then
            Me.FormBorderStyle = FormBorderStyle.Sizable
            Me.Location = New Point(storedBounds.X, storedBounds.Y)
            Me.Size = New Size(storedBounds.Width, storedBounds.Height)
        Else
            storedBounds = Me.RestoreBounds
            Me.FormBorderStyle = FormBorderStyle.None
            Me.Location = New Point(0, 0)
            Me.Size = New Size(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height)
        End If

    End Sub
End Class