Imports System.IO
Imports System.Net
Imports System.Text.RegularExpressions
Public Class Form5_videoDownloader
    Dim a As Boolean = False
    Dim w As New WebClient
    Dim startTime, endTime As Double
    Dim urlsList As New List(Of String)
    Dim fileList As New List(Of String)
    Dim stateList As New List(Of Integer)
    Dim currentlyDownloading As Integer
    Dim qualityListList As New List(Of List(Of YouTubeVideoQuality))
    Dim WithEvents ListView1 As New myListView

    Private Sub Form5_videoDownloader_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Controls.Add(ListView1)
        ListView1.Location = New Point(15, 88)
        ListView1.Size = New Size(459, 200)
        ListView1.View = View.Details
        ListView1.FullRowSelect = True
        ListView1.HideSelection = False
        Dim columnheader As ColumnHeader
        columnheader = New ColumnHeader() With {.Text = "Title", .Width = 160} : ListView1.Columns.Add(columnheader)
        columnheader = New ColumnHeader() With {.Text = "State", .Width = 70} : ListView1.Columns.Add(columnheader)
        columnheader = New ColumnHeader() With {.Text = "Format", .Width = 70} : ListView1.Columns.Add(columnheader)
        columnheader = New ColumnHeader() With {.Text = "Length", .Width = 70} : ListView1.Columns.Add(columnheader)
        columnheader = New ColumnHeader() With {.Text = "Size", .Width = 70} : ListView1.Columns.Add(columnheader)
        ListView1.Controls.Add(ComboBox3)
        AddHandler ListView1.MouseUp, AddressOf ListView1_MouseUp
        AddHandler ListView1.SelectedIndexChanged, AddressOf ListView1_selectedchange
        'AddHandler AxVLCPlugin21.MediaPlayerTimeChanged, AddressOf b
    End Sub

    Private Sub Form5_videoDownloader_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        urlsList.Clear()
        ComboBox1.SelectedIndex = 1
        ComboBox3.Sorted = True
        'AxVLCPlugin21.Toolbar = False
    End Sub

    Private Sub b()
        'Label6.Text = AxVLCPlugin21.input.Position.ToString
    End Sub

    Private Sub Form5_videoDownloader_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'w.CancelAsync()
        'AxVLCPlugin21.playlist.stop()

        ''For i As Integer = 0 To fileList.Count - 1
        ''If fileList(i).EndsWith(".flv.avi", StringComparison.InvariantCultureIgnoreCase) Then
        ''fileList(i) = fileList(i).Replace(".flv.avi", ".flv")
        ''FileSystem.Rename(".\Downloaded Video\" + fileList(i) + ".avi", ".\Downloaded Video\" + fileList(i))
        ''End If
        ''Next
    End Sub

    Private Sub AxWindowsMediaPlayer1_PlayStateChange(ByVal sender As Object, ByVal e As AxWMPLib._WMPOCXEvents_PlayStateChangeEvent)
        If e.newState = 3 Then
            startTime = 0
            'endTime = AxWindowsMediaPlayer1.currentMedia.duration
            Label2.Text = milliseconds_to_string(0)
            'Label6.Text = milliseconds_to_string(AxWindowsMediaPlayer1.currentMedia.duration)
        End If
    End Sub

    Private Sub AxWindowsMediaPlayer1_PositionChange(ByVal sender As Object, ByVal e As AxWMPLib._WMPOCXEvents_PositionChangeEvent)
        'Label9.Text = milliseconds_to_string(AxWindowsMediaPlayer1.Ctlcontrols.currentPosition)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Clipboard.ContainsText Then
            TextBox1.Text = Clipboard.GetText
            Button2_Click(sender, e)
        End If
    End Sub 'get text from clipboard and press 'OK'

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        fileList.Add("")
        stateList.Add(0)
        urlsList.Add(TextBox1.Text)
        qualityListList.Add(New List(Of YouTubeVideoQuality))
        Dim lviewitem As ListViewItem
        lviewitem = New ListViewItem("scaning...")
        lviewitem.SubItems.Add("scaning...")
        lviewitem.SubItems.Add("") : lviewitem.SubItems.Add("") : lviewitem.SubItems.Add("")
        ListView1.Items.Add(lviewitem)
        TextBox1.Text = ""
        If Not BackgroundWorker1.IsBusy Then BackgroundWorker1.RunWorkerAsync()
    End Sub 'GO Download

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        ComboBox2.Items.Clear()
        If ComboBox1.SelectedIndex = 0 Then
            ComboBox2.Items.Add("320")
            ComboBox2.Items.Add("480")
            ComboBox2.Items.Add("640")
            ComboBox2.Items.Add("854")
            ComboBox2.SelectedIndex = 2
        Else
            ComboBox2.Items.Add("640")
            ComboBox2.Items.Add("1280")
            ComboBox2.Items.Add("1920")
            ComboBox2.Items.Add("2048")
            ComboBox2.SelectedIndex = 0
        End If
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim c As Integer = 0
        Do While c <> stateList.Count
            c = stateList.Count
            For i = 0 To c - 1
                If stateList(i) = 0 Then
                    Dim yd As New YouTubeDownloader
                    qualityListList(i) = yd.GetYouTubeVideoUrls(urlsList(i))
                    stateList(i) = 1 'Scanned
                    If Not Me.IsDisposed Then Me.Invoke(New Deleg(AddressOf changeItem), i, qualityListList(i))
                End If
            Next
        Loop
    End Sub

    Delegate Sub Deleg(ByVal i As Integer, ByVal l As List(Of YouTubeVideoQuality))
    Public Sub changeItem(ByVal i As Integer, ByVal l As List(Of YouTubeVideoQuality))
        If l.Count = 0 Then
            stateList(i) = -1 'Error
            ListView1.Items(i).Text = "Invalid url"
            ListView1.Items(i).SubItems(1).Text = "Invalid data"
        Else
            ListView1.Items(i).Text = WebUtility.HtmlDecode(l(0).VideoTitle)
            Dim t As String
            Dim videoLength As TimeSpan = TimeSpan.FromSeconds(l(0).Length)
            If (videoLength.Hours > 0) Then
                t = String.Format("{0}:{1}:{2}", videoLength.Hours, videoLength.Minutes, videoLength.Seconds)
            Else
                t = String.Format("{0}:{1}", videoLength.Minutes, videoLength.Seconds)
            End If
            ListView1.Items(i).SubItems(3).Text = t

            Dim yvq As YouTubeVideoQuality = Nothing
            For i2 As Integer = 0 To l.Count - 1
                'For Each y As YouTubeVideoQuality In l
                If l(i2).Extention.ToUpper = ComboBox1.SelectedItem.ToString.ToUpper And l(i2).Dimension.Width = CInt(ComboBox2.SelectedItem) Then
                    yvq = l(i2)
                    l(0).currentlySelected = i2
                    Exit For
                End If
            Next
            If yvq IsNot Nothing Then
                'l(0).currentlySelected 
                ListView1.Items(i).SubItems(1).Text = "queued"
                ListView1.Items(i).SubItems(2).Text = yvq.Extention.ToUpper + " " + yvq.Dimension.Width.ToString + "x" + yvq.Dimension.Height.ToString
                ListView1.Items(i).SubItems(4).Text = yvq.VideoSize.ToString
                ListView1.Items(i).ForeColor = Color.DarkBlue
                stateList(i) = 2 'Prepared for download
                If Not BackgroundWorker2.IsBusy And Not w.IsBusy Then BackgroundWorker2.RunWorkerAsync()
            Else
                l(0).currentlySelected = -1
                ListView1.Items(i).SubItems(1).Text = "Please, select format."
                stateList(i) = 3 'Waiting for format select
            End If
        End If
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Dim cur As Integer = -1
        For i As Integer = 0 To stateList.Count - 1
            If stateList(i) = 2 Then cur = i : Exit For
        Next
        If cur <> -1 Then
            stateList(cur) = 3 'Downloading
            AddHandler w.DownloadProgressChanged, AddressOf w_Progress

            Dim cur_format As Integer = qualityListList(cur)(0).currentlySelected
            Dim filename As String = WebUtility.HtmlDecode(qualityListList(cur)(0).VideoTitle)
            filename = filename.Replace("\", "-").Replace("/", "-")
            filename = filename.Replace(":", "-").Replace("*", "").Replace("?", "")
            filename = filename.Replace("""", "'").Replace("<", "").Replace(">", "").Replace("|", "") + "." + qualityListList(cur)(cur_format).Extention
            fileList(cur) = filename
            currentlyDownloading = cur
            w.DownloadFileAsync(New Uri(qualityListList(cur)(cur_format).DownloadUrl), ".\Downloaded Video\" + filename)
        End If
    End Sub
    Private Sub w_Progress(ByVal sender As Object, ByVal e As System.Net.DownloadProgressChangedEventArgs)
        If e.BytesReceived = e.TotalBytesToReceive Then
            If Not Me.IsDisposed Then Me.Invoke(New Deleg2(AddressOf changeLabel), 1, "Done")
        Else
            If Not Me.IsDisposed Then Me.Invoke(New Deleg2(AddressOf changeLabel), 0, e.BytesReceived.ToString + " / " + e.TotalBytesToReceive.ToString)
        End If
    End Sub
    Delegate Sub Deleg2(ByVal i As Integer, ByVal l As String)
    Public Sub changeLabel(ByVal i As Integer, ByVal s As String)
        If i = 0 Then
            ListView1.Items(currentlyDownloading).ForeColor = Color.DarkGoldenrod
            ListView1.Items(currentlyDownloading).SubItems(1).Text = "Downloading..."
            ListView1.Items(currentlyDownloading).SubItems(4).Text = s
        ElseIf i = 1 Then
            stateList(currentlyDownloading) = 4
            Dim cur_format As Integer = qualityListList(currentlyDownloading)(0).currentlySelected
            ListView1.Items(currentlyDownloading).ForeColor = Color.DarkGreen
            ListView1.Items(currentlyDownloading).SubItems(1).Text = "Done"
            ListView1.Items(currentlyDownloading).SubItems(4).Text = qualityListList(currentlyDownloading)(cur_format).VideoSize.ToString
            Do While BackgroundWorker2.IsBusy Or w.IsBusy : Application.DoEvents() : Loop
            BackgroundWorker2.RunWorkerAsync()
        End If
    End Sub

    Private lvItem As ListViewItem
    Private Sub ListView1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        lvItem = ListView1.GetItemAt(e.X, e.Y)
        If Not (lvItem Is Nothing) Then
            ComboBox3.Items.Clear()
            For Each y As YouTubeVideoQuality In qualityListList(lvItem.Index)
                ComboBox3.Items.Add(y.Extention.ToUpper + " " + y.Dimension.Width.ToString + "x" + y.Dimension.Height.ToString)
            Next

            Dim ClickedItem As Rectangle = lvItem.Bounds
            Dim col02Left = ClickedItem.Left + ListView1.Columns(0).Width + ListView1.Columns(1).Width
            If ((col02Left + ListView1.Columns(2).Width) < 0) Then     'Verify if column is completely scrolled off to the left.
                Return
            ElseIf (col02Left < 0) Then                                'Verify if column is partially scrolled off to the left.
                'Determine if column extends beyond right side of ListView.
                If ((col02Left + ListView1.Columns(2).Width) > ListView1.Width) Then
                    'Right side of cell out of view.
                    'Set width of column to match width of ListView.
                    'ClickedItem.Width = ListView1.Width
                    'ClickedItem.X = 0
                    ClickedItem.X = 0
                    ClickedItem.Width = ListView1.Width - 2
                Else
                    'Right side of cell is in view.
                    'ClickedItem.Width = ListView1.Columns(2).Width + ClickedItem.Left
                    'ClickedItem.X = 2
                    ClickedItem.Width = ListView1.Columns(2).Width + col02Left
                    ClickedItem.X = 0
                End If
            ElseIf col02Left + ListView1.Columns(2).Width > ListView1.Width Then 'Verify if right bound is visible
                ClickedItem.X = col02Left
                ClickedItem.Width = ListView1.Width - col02Left
            ElseIf (ListView1.Columns(2).Width > ListView1.Width) Then 'Verify if column width grater then listview width BUT left is in view
                ClickedItem.X = col02Left
                ClickedItem.Width = ListView1.Width - col02Left
            Else                                                       'Entire column is in view
                'ClickedItem.Width = ListView1.Columns(2).Width
                'ClickedItem.X = 2
                ClickedItem.X = col02Left
                ClickedItem.Width = ListView1.Columns(2).Width
            End If

            ' Adjust the top to account for the location of the ListView.
            'ClickedItem.Y += ListView1.Top
            'ClickedItem.X += ListView1.Left + ListView1.Columns(0).Width + ListView1.Columns(1).Width
            ' Assign calculated bounds to the ComboBox.
            ComboBox3.Bounds = ClickedItem
            ' Set default text for ComboBox to match the item that is clicked.
            ComboBox3.Text = lvItem.SubItems(2).Text
            ' Display the ComboBox, and make sure that it is on top with focus.
            ComboBox3.Visible = True
            ComboBox3.BringToFront()
            ComboBox3.Focus()
        Else
            ComboBox3.Visible = False
        End If
    End Sub
    Private Sub ComboBox3_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.Leave
        ComboBox3.Visible = False
    End Sub
    Private Sub ListView1_selectedchange(ByVal sender As Object, ByVal e As System.EventArgs)
        'For i As Integer = 0 To fileList.Count - 1
        'If fileList(i).EndsWith(".flv.avi", StringComparison.InvariantCultureIgnoreCase) Then
        'fileList(i) = fileList(i).Replace(".flv.avi", ".flv")
        'FileSystem.Rename(".\Downloaded Video\" + fileList(i) + ".avi", ".\Downloaded Video\" + fileList(i))
        'End If
        'Next

        If ListView1.SelectedIndices.Count = 0 Then Exit Sub
        Dim sel As Integer = ListView1.SelectedIndices(0)
        If stateList(sel) = 4 Then
            'If fileList(sel).EndsWith("FLV", StringComparison.InvariantCultureIgnoreCase) Then
            'FileSystem.Rename(".\Downloaded Video\" + fileList(sel), ".\Downloaded Video\" + fileList(sel) + ".avi")
            'fileList(sel) = fileList(sel) + ".avi"
            'End If

            'AxVLCPlugin21.playlist.stop()
            Dim t As String = New Uri(Application.StartupPath + "\Downloaded Video\" + fileList(sel)).AbsoluteUri
            'Dim id As Integer = AxVLCPlugin21.playlist.add(t) : AxVLCPlugin21.playlist.playItem(id)
        End If
    End Sub

    Private Function milliseconds_to_string(ByVal d As Double) As String
        Dim m As Integer = CInt(CLng(d) \ 60)
        Dim h As Integer = m \ 60
        If h > 0 Then d = d - h * 60 * 60 : m = m - h * 60
        If m > 0 Then d = d - m * 60
        d = Math.Round(d, 3)
        Return Format(h, "00") + ":" + Format(m, "00") + ":" + Format(d, "00.000").Replace(",", ".")
    End Function
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        'startTime = AxWindowsMediaPlayer1.Ctlcontrols.currentPosition
        'Label2.Text = milliseconds_to_string(AxWindowsMediaPlayer1.Ctlcontrols.currentPosition)
        If endTime < startTime Then
            Button5.Enabled = False : Label6.ForeColor = Color.Red
        Else
            Button5.Enabled = True : Label6.ForeColor = SystemColors.ControlText
        End If
    End Sub
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        'endTime = AxWindowsMediaPlayer1.Ctlcontrols.currentPosition
        'Label6.Text = milliseconds_to_string(AxWindowsMediaPlayer1.Ctlcontrols.currentPosition)
        If endTime < startTime Then
            Button5.Enabled = False : Label6.ForeColor = Color.Red
        Else
            Button5.Enabled = True : Label6.ForeColor = SystemColors.ControlText
        End If
    End Sub
    Private Sub TrackBar1_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar1.Scroll
        'AxVLCPlugin21.input.Position = TrackBar1.Value / 1000
    End Sub
End Class

Public Class YouTubeDownloader
    Public Function GetYouTubeVideoUrls(ByVal VideoUrl As String) As List(Of YouTubeVideoQuality)
        Dim urls As New List(Of YouTubeVideoQuality)
        'For Each VideoUrl As String In VideoUrls

        Dim html As String
        Try
            Dim w As New WebClient
            w.Encoding = System.Text.Encoding.UTF8
            html = w.DownloadString(VideoUrl)
        Catch ex As Exception
            Return urls
        End Try

        'Dim html As String = YouTube_Helper.DownloadWebPage(VideoUrl)
        'If html = "-1" Then Return urls
        Dim title As String = GetTitle(html)
        For Each videoLink In ExtractUrls(html)
            Dim q As YouTubeVideoQuality = New YouTubeVideoQuality
            q.VideoUrl = VideoUrl
            q.VideoTitle = title
            q.DownloadUrl = videoLink + "&title=" + title
            If Not getSize(q) Then Continue For
            q.Length = Long.Parse(Regex.Match(html, """length_seconds"":(.+?)[,}]", RegexOptions.Singleline).Groups(1).ToString())
            Dim IsWide As Boolean = IsWideScreen(html)
            If getQuality(q, IsWide) Then urls.Add(q)
        Next
        'Next
        If urls.Count = 0 Then 'RETRY
            Try
                Dim w As New WebClient
                w.Encoding = System.Text.Encoding.UTF8
                html = w.DownloadString(VideoUrl)
            Catch ex As Exception
                Return urls
            End Try
            For Each videoLink In ExtractUrls(html)
                Dim q As YouTubeVideoQuality = New YouTubeVideoQuality
                q.VideoUrl = VideoUrl
                q.VideoTitle = title
                q.DownloadUrl = videoLink + "&title=" + title
                If Not getSize(q) Then Continue For
                q.Length = Long.Parse(Regex.Match(html, """length_seconds"":(.+?)[,}]", RegexOptions.Singleline).Groups(1).ToString())
                Dim IsWide As Boolean = IsWideScreen(html)
                If getQuality(q, IsWide) Then urls.Add(q)
            Next
        End If
        Return urls
    End Function

    Public Function GetTitle(ByVal RssDoc As String) As String
        Dim str14 As String = YouTube_Helper.GetTxtBtwn(RssDoc, "'VIDEO_TITLE': '", "'", 0)
        If str14 = "" Then str14 = YouTube_Helper.GetTxtBtwn(RssDoc, """title"" content=""", """", 0)
        If str14 = "" Then str14 = YouTube_Helper.GetTxtBtwn(RssDoc, "&title=", "&", 0)
        str14 = str14.Replace("\", "").Replace("'", "&#39;").Replace("""", "&quot;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("+", " ")
        Return str14
    End Function

    Public Function IsWideScreen(ByVal html As String) As Boolean
        Dim res As Boolean = False
        Dim match As String = Regex.Match(html, "'IS_WIDESCREEN':\s+(.+?)\s+", RegexOptions.Singleline).Groups(1).ToString().ToLower().Trim()
        res = ((match = "true") Or (match = "true,"))
        Return res
    End Function

    Private Function getQuality(ByVal q As YouTubeVideoQuality, ByVal _Wide As Boolean) As Boolean
        Dim iTagValue As Integer
        Dim itag As String = Regex.Match(q.DownloadUrl, "itag=([1-9]?[0-9]?[0-9])", RegexOptions.Singleline).Groups(1).ToString()
        If itag <> "" Then
            If (Integer.TryParse(itag, iTagValue) = False) Then iTagValue = 0
            Select Case iTagValue
                Case 5 : q.SetQuality("flv", New Size(320, DirectCast(IIf(_Wide, 180, 240), Integer))) : Exit Select
                Case 6 : q.SetQuality("flv", New Size(480, DirectCast(IIf(_Wide, 270, 360), Integer))) : Exit Select
                Case 17 : q.SetQuality("3gp", New Size(176, DirectCast(IIf(_Wide, 99, 144), Integer))) : Exit Select
                Case 18 : q.SetQuality("mp4", New Size(640, DirectCast(IIf(_Wide, 360, 480), Integer))) : Exit Select
                Case 22 : q.SetQuality("mp4", New Size(1280, DirectCast(IIf(_Wide, 720, 960), Integer))) : Exit Select
                Case 34 : q.SetQuality("flv", New Size(640, DirectCast(IIf(_Wide, 360, 480), Integer))) : Exit Select
                Case 35 : q.SetQuality("flv", New Size(854, DirectCast(IIf(_Wide, 480, 640), Integer))) : Exit Select
                Case 36 : q.SetQuality("3gp", New Size(320, DirectCast(IIf(_Wide, 180, 240), Integer))) : Exit Select
                Case 37 : q.SetQuality("mp4", New Size(1920, DirectCast(IIf(_Wide, 1080, 1440), Integer))) : Exit Select
                Case 38 : q.SetQuality("mp4", New Size(2048, DirectCast(IIf(_Wide, 1152, 1536), Integer))) : Exit Select
                Case 43 : q.SetQuality("webm", New Size(640, DirectCast(IIf(_Wide, 360, 480), Integer))) : Exit Select
                Case 44 : q.SetQuality("webm", New Size(854, DirectCast(IIf(_Wide, 480, 640), Integer))) : Exit Select
                Case 45 : q.SetQuality("webm", New Size(1280, DirectCast(IIf(_Wide, 720, 960), Integer))) : Exit Select
                Case 46 : q.SetQuality("webm", New Size(1920, DirectCast(IIf(_Wide, 1080, 1440), Integer))) : Exit Select
                Case 82 : q.SetQuality("3D.mp4", New Size(480, DirectCast(IIf(_Wide, 270, 360), Integer))) : Exit Select '3D
                Case 83 : q.SetQuality("3D.mp4", New Size(640, DirectCast(IIf(_Wide, 360, 480), Integer))) : Exit Select       '3D
                Case 84 : q.SetQuality("3D.mp4", New Size(1280, DirectCast(IIf(_Wide, 720, 960), Integer))) : Exit Select   '3D
                Case 85 : q.SetQuality("3D.mp4", New Size(1920, DirectCast(IIf(_Wide, 1080, 1440), Integer))) : Exit Select  '3D
                Case 100 : q.SetQuality("3D.webm", New Size(640, DirectCast(IIf(_Wide, 360, 480), Integer))) : Exit Select '3D
                Case 101 : q.SetQuality("3D.webm", New Size(640, DirectCast(IIf(_Wide, 360, 480), Integer))) : Exit Select    '3D
                Case 102 : q.SetQuality("3D.webm", New Size(1280, DirectCast(IIf(_Wide, 720, 960), Integer))) : Exit Select '3D
                Case 120 : q.SetQuality("live.flv", New Size(1280, DirectCast(IIf(_Wide, 720, 960), Integer))) : Exit Select 'Live-streaming - should be ignored?
                Case Else : q.SetQuality("itag-" + itag, New Size(0, 0)) : Exit Select       'unknown or parse error
            End Select
            Return True
        End If
        Return False
    End Function

    Private Function ExtractUrls(ByVal html As String) As List(Of String)
        Dim urls As New List(Of String)
        Dim DataBlockStart As String = """url_encoded_fmt_stream_map"":\s+""(.+?)&" 'Marks start of Javascript Data Block
        html = Uri.UnescapeDataString(Regex.Match(html, DataBlockStart, RegexOptions.Singleline).Groups(1).ToString())

        Dim firstPatren As String = html.Substring(0, html.IndexOf("=") + 1)
        Dim matchs = Regex.Split(html, firstPatren)
        For i As Integer = 0 To matchs.Length - 1
            matchs(i) = firstPatren + matchs(i)
        Next
        For Each match In matchs
            If Not match.Contains("url=") Then Continue For

            Dim url As String = YouTube_Helper.GetTxtBtwn(match, "url=", "\u0026", 0)
            If url = "" Then url = YouTube_Helper.GetTxtBtwn(match, "url=", ",url", 0)
            If url = "" Then url = YouTube_Helper.GetTxtBtwn(match, "url=", """,", 0)

            Dim sig As String = YouTube_Helper.GetTxtBtwn(match, "sig=", "\u0026", 0)
            If sig = "" Then url = YouTube_Helper.GetTxtBtwn(match, "sig=", ",sig", 0)
            If sig = "" Then url = YouTube_Helper.GetTxtBtwn(match, "sig=", """,", 0)

            While (url.EndsWith(",")) Or (url.EndsWith(".")) Or (url.EndsWith(""""))
                url = url.Remove(url.Length - 1, 1)
            End While

            While (sig.EndsWith(",")) Or (sig.EndsWith(".")) Or (sig.EndsWith(""""))
                sig = sig.Remove(sig.Length - 1, 1)
            End While

            If String.IsNullOrEmpty(url) Then Continue For
            If Not String.IsNullOrEmpty(sig) Then url += "&signature=" + sig
            urls.Add(url)
        Next
        Return urls
    End Function

    Private Function getSize(ByVal q As YouTubeVideoQuality) As Boolean
        Try
            'Dim fileInfoRequest As HttpWebRequest = HttpWebRequest.Create(q.DownloadUrl)
            Dim fileInfoRequest As System.Net.WebRequest = HttpWebRequest.Create(q.DownloadUrl)
            'Dim fileInfoResponse As HttpWebResponse = fileInfoRequest.GetResponse()
            Dim fileInfoResponse As System.Net.WebResponse = fileInfoRequest.GetResponse()
            Dim bytesLength As Long = fileInfoResponse.ContentLength
            fileInfoRequest.Abort()
            fileInfoResponse.Close()
            If bytesLength <> -1 Then q.SetSize(bytesLength) : Return True Else Return False
        Catch ex As Exception
            Return False
        End Try
    End Function
End Class

Public Class YouTube_Helper
    Public Function UrlDecode(ByVal str As String) As String
        Return System.Web.HttpUtility.UrlDecode(str)
    End Function

    Public Function isValidUrl(ByVal url As String) As Boolean
        Dim pattern As String = "^(http|ftp|https):\/\/[\w\-_]+(\.[\w\-_]+)+([\w\-\.,@?^=%&amp;:/~\+#]*[\w\-\@?^=%&amp;/~\+#])?$"
        Dim regex As Regex = New Regex(pattern, RegexOptions.Compiled Or RegexOptions.IgnoreCase)
        Return regex.IsMatch(url)
    End Function

    Public Overloads Shared Function GetTxtBtwn(ByVal input As String, ByVal start As String, ByVal _end As String, ByVal startIndex As Integer) As String
        Return GetTxtBtwn(input, start, _end, startIndex, False)
    End Function

    Public Overloads Shared Function GetLastTxtBtwn(ByVal input As String, ByVal start As String, ByVal _end As String, ByVal startIndex As Integer) As String
        Return GetTxtBtwn(input, start, _end, startIndex, True)
    End Function

    Private Overloads Shared Function GetTxtBtwn(ByVal input As String, ByVal start As String, ByVal _end As String, ByVal startIndex As Integer, ByVal UseLastIndexOf As Boolean) As String
        Dim index1 As Integer = CInt(IIf(UseLastIndexOf, input.LastIndexOf(start, startIndex), input.IndexOf(start, startIndex)))
        If index1 = -1 Then Return ""
        index1 += start.Length
        Dim index2 As Integer = input.IndexOf(_end, index1)
        If index2 = -1 Then Return input.Substring(index1)
        Return input.Substring(index1, index2 - index1)
    End Function

    Public Function Split(ByVal input As String, ByVal pattren As String) As String()
        Return Regex.Split(input, pattren)
    End Function

    'Public Overloads Shared Function DownloadWebPage(ByVal Url As String) As String
    'Return DownloadWebPage(Url, Nothing)
    'End Function

    'Public Overloads Shared Function DownloadWebPage(ByVal Url As String, ByVal stopLine As String) As String
    'Try
    'Dim WebRequestObject As HttpWebRequest = HttpWebRequest.Create(Url)
    ''WebRequestObject.Proxy = InitialProxy()
    'WebRequestObject.Proxy = Nothing
    'WebRequestObject.UserAgent = ".NET Framework/2.0"
    'Dim Response As System.Net.WebResponse = WebRequestObject.GetResponse()
    'Dim WebStream As Stream = Response.GetResponseStream()
    'Dim Reader As StreamReader = New StreamReader(WebStream)
    'Dim line As String, PageContent As String = ""
    'If stopLine Is Nothing Then
    'PageContent = Reader.ReadToEnd()
    'Else
    'While Not Reader.EndOfStream
    'line = Reader.ReadLine()
    'PageContent += line + Environment.NewLine
    'If line.Contains(stopLine) Then Exit While
    'End While
    'End If
    'Reader.Close()
    'WebStream.Close()
    'Response.Close()
    'Return PageContent
    'Catch ex As Exception
    'Return "-1"
    'End Try
    'End Function

    Public Function GetVideoIDFromUrl(ByVal Url As String) As String
        Url = Url.Substring(Url.IndexOf("?") + 1)
        Dim props() As String = Url.Split("&"c)
        Dim videoid As String = ""
        For Each prop As String In props
            If prop.StartsWith("v=") Then videoid = prop.Substring(prop.IndexOf("v=") + 2)
        Next
        Return videoid
    End Function

    'Public Shared Function InitialProxy() As IWebProxy
    'Dim address As String = getIEProxy()
    'If Not String.IsNullOrEmpty(address) Then
    'Dim proxy As WebProxy = New WebProxy(address)
    'proxy.Credentials = CredentialCache.DefaultNetworkCredentials
    'Return proxy
    'Else
    'Return Nothing
    'End If
    'End Function 'TUT VOZMOJNO BAG

    'Private Shared Function getIEProxy() As String
    'Dim p = WebRequest.DefaultWebProxy
    'If p Is Nothing Then Return Nothing
    'Dim _webProxy As WebProxy = Nothing
    'If p Is _webProxy Then
    '_webProxy = p
    'Else
    'Dim t As Type = p.GetType()
    'Dim s = t.GetProperty("WebProxy", &HFFF).GetValue(p, Nothing)
    '        _webProxy = s
    '    End If
    '    If (_webProxy Is Nothing Or _webProxy.Address Is Nothing Or String.IsNullOrEmpty(_webProxy.Address.AbsolutePath)) Then Return Nothing
    '    Return _webProxy.Address.Host
    'End Function 'TUT VOZMOJNO BAG
End Class

Public Class FormatLeftTime
    Private TimeUnitsNames() As String = {"Milli", "Sec", "Min", "Hour", "Day", "Month", "Year", "Decade", "Century"}
    Private TimeUnitsValue() As Integer = {1000, 60, 60, 24, 30, 12, 10, 10} 'refrernce unit is milli
    Public Function Format(ByVal millis As Long) As String
        Dim _format As String = ""
        For i As Integer = 0 To TimeUnitsValue.Length - 1
            Dim y As Long = millis Mod TimeUnitsValue(i)
            millis = CLng(millis / TimeUnitsValue(i))
            If y = 0 Then Continue For
            _format = y.ToString + " " + TimeUnitsNames(i) + " , " + _format
        Next

        Dim trimChars() As Char = {","c, " "c}
        _format = _format.Trim(trimChars)
        If _format = "" Then Return "0 Sec" Else Return _format
    End Function
End Class

Public Class YouTubeVideoQuality
    Public VideoTitle As String
    Public Extention As String
    Public DownloadUrl As String
    Public VideoUrl As String
    Public VideoSize As Long
    Public Dimension As Size
    Public Length As Long
    Public currentlySelected As Integer

    Public Overrides Function ToString() As String
        'Return MyBase.ToString()
        Return Extention + " File " + Dimension.Width.ToString + "x" + Dimension.Height.ToString
    End Function

    Public Sub SetQuality(ByVal Extention As String, ByVal Dimension As Size)
        Me.Extention = Extention
        Me.Dimension = Dimension
    End Sub

    Public Sub SetSize(ByVal size As Long)
        Me.VideoSize = size
    End Sub
End Class

Public Class myListView
    Inherits ListView
    Private Const WM_HSCROLL As Integer = &H114
    Private Const WM_VSCROLL As Integer = &H115
    Protected Overrides Sub WndProc(ByRef msg As Message)
        ' Look for the WM_VSCROLL or the WM_HSCROLL messages.
        If ((msg.Msg = WM_VSCROLL) Or (msg.Msg = WM_HSCROLL)) Then
            ' Move focus to the ListView to cause ComboBox to lose focus.
            Me.Focus()
        End If

        ' Pass message to default handler.
        MyBase.WndProc(msg)
    End Sub
End Class
