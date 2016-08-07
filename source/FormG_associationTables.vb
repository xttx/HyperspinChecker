Imports System.ComponentModel
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic.FileIO.FileSystem

Public Class FormG_associationTables
    Const separator As String = "\\\|"
    Dim refr As Boolean = False
    Dim cur_table As New List(Of List(Of String))
    Dim cur_table_headers As New List(Of String)
    Dim c_add As Integer = 0
    Dim c_updt As Integer = 0
    Dim c_total As Integer = 0
    Dim oldHeight As Integer = Me.Height
    Dim oldWidth As Integer = Me.Width
    Dim WithEvents bg_from_dat As New BackgroundWorker() With {.WorkerReportsProgress = True}
    Dim WithEvents bg_from_folder As New BackgroundWorker() With {.WorkerReportsProgress = True}

    'Load form - fill system list
    Private Sub FormG_associationTables_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Language.localize(Me)
        For Each sys As String In Form1.ComboBox1.Items
            ComboBox1.Items.Add(sys)
        Next
    End Sub

    'Select System
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ListBox1.Items.Clear()
        cur_table.Clear()
        cur_table_headers.Clear()
        If FileExists(".\A-Tables\" + ComboBox1.SelectedItem.ToString.Trim + ".atbl") Then
            FileOpen(1, ".\A-Tables\" + ComboBox1.SelectedItem.ToString.Trim + ".atbl", OpenMode.Input)
            Dim s As String = LineInput(1)
            FileClose(1)

            For Each i As String In s.Split({separator}, StringSplitOptions.None)
                If i.ToUpper = "CRC" Or i.ToUpper = "MD5" Or i.ToUpper = "SHA1" Then Continue For
                ListBox1.Items.Add(i)
            Next
        Else
            ListBox1.Items.Add("This table does not exist.")
        End If
    End Sub

    'Load table
    Private Sub loadTable()
        cur_table.Clear()
        cur_table_headers.Clear()
        Dim sys As String = ""
        If ComboBox1.InvokeRequired Then
            ComboBox1.Invoke(Sub() sys = ComboBox1.SelectedItem.ToString.Trim())
        Else
            sys = ComboBox1.SelectedItem.ToString.Trim
        End If
        If FileExists(".\A-Tables\" + sys + ".atbl") Then
            FileOpen(1, ".\A-Tables\" + sys + ".atbl", OpenMode.Input)
            'headers
            Dim s As String = LineInput(1)
            For Each i As String In s.Split({separator}, StringSplitOptions.None)
                cur_table.Add(New List(Of String))
                cur_table_headers.Add(i)
            Next

            'Main table
            Do While Not EOF(1)
                s = LineInput(1)
                Dim arr() As String = s.Split({separator}, StringSplitOptions.None)
                For i As Integer = 0 To arr.Count - 1
                    cur_table(i).Add(arr(i))
                Next
            Loop
            FileClose(1)
        End If
    End Sub

    'Append from .dat file
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ComboBox1.SelectedIndex < 0 Then MsgBox("You have to select a system.") : Exit Sub
        If TextBox1.Text.Trim = "" Then MsgBox("You have to enter new set name.") : Exit Sub
        If ListBox1.Items.Contains(TextBox1.Text.Trim) Then MsgBox("This set name already exist in the table/") : Exit Sub

        Dim fb As New OpenFileDialog
        fb.Title = "Open clrMamePro dat or xml"
        fb.Filter = "All accepted files|*.dat;*.xml|ClrMamePro/Redump/No-Intro/Tosec dat|*.dat|Hyperspin/ClrMamePro/Redump/No-Intro/Tosec xml|*.xml|All files|*.*"
        fb.ShowDialog()
        If fb.FileName = "" Then MsgBox("No file selected.") : Exit Sub
        If Not FileIO.FileSystem.FileExists(fb.FileName) Then MsgBox("File """ + fb.FileName + """ does not exist.") : Exit Sub
        bg_from_dat.RunWorkerAsync(fb.FileName)
    End Sub
    Private Sub Button1_Click_bg(o As Object, e As DoWorkEventArgs) Handles bg_from_dat.DoWork
        Dim fb As String = DirectCast(e.Argument, String)

        Dim isXmlFormat As Boolean = True
        Dim isXmlFormatHS As Boolean = False
        Dim msg_noCRC_Shown As Boolean = False
        Dim msg_multipleRoms_Shown As Boolean = False

        Label4.Invoke(Sub() Label4.Visible = True)
        Label4.Invoke(Sub() Label4.Text = "Loading Table...")
        loadTable()
        If cur_table.Count = 0 Then
            cur_table.Add(New List(Of String))
            cur_table.Add(New List(Of String))
            cur_table.Add(New List(Of String))
            cur_table_headers.Add("CRC")
            cur_table_headers.Add("MD5")
            cur_table_headers.Add("SHA1")
        End If

        Dim l As New List(Of String)
        For i As Integer = 0 To cur_table(0).Count - 1
            l.Add("")
        Next
        cur_table.Add(l)
        cur_table_headers.Add(TextBox1.Text.Trim)
        c_add = 0
        c_updt = 0
        c_total = cur_table(0).Count

        Label4.Invoke(Sub() Label4.Text = "Loading XML...")
        Dim xdat As New Xml.XmlDocument
        Try
            xdat.Load(fb)
        Catch ex As Exception
            isXmlFormat = False
        End Try
        Label4.Invoke(Sub() Label4.Visible = False)

        Dim name As String = ""
        Dim description As String = ""
        Dim crc As String = ""
        Dim md5 As String = ""
        Dim sha As String = ""
        ProgressBar1.Invoke(Sub() ProgressBar1.Value = 0)
        If isXmlFormat Then
            'new format (xml)
            Dim nodes = xdat.SelectNodes("/datafile/game")
            If nodes.Count = 0 Then
                'try HS xml
                nodes = xdat.SelectNodes("/menu/game")
                isXmlFormatHS = True
            End If
            If nodes.Count = 0 Then MsgBox("XML contains 0 games. Exiting.") : Exit Sub
            ProgressBar1.Invoke(Sub() ProgressBar1.Maximum = nodes.Count)
            For Each node As Xml.XmlNode In nodes
                name = node.Attributes.GetNamedItem("name").Value
                If node.SelectNodes("description").Count > 0 Then
                    description = node.SelectSingleNode("description").InnerText
                Else
                    description = ""
                End If
                If node.SelectNodes("rom").Count = 1 Then
                    crc = node.SelectSingleNode("rom").Attributes.GetNamedItem("crc").Value
                    md5 = node.SelectSingleNode("rom").Attributes.GetNamedItem("md5").Value
                    sha = node.SelectSingleNode("rom").Attributes.GetNamedItem("sha1").Value
                ElseIf node.SelectNodes("crc").Count = 1 Then
                    'try hs format
                    crc = node.SelectSingleNode("crc").InnerText.Trim.ToUpper
                    md5 = ""
                    sha = ""
                Else
                    'Multiple roms
                    crc = ""
                    md5 = ""
                    sha = ""
                End If
                If Not msg_noCRC_Shown AndAlso node.SelectNodes("rom").Count = 0 Then
                    MsgBox("At least one of this dat file entry, have no crc. There is no much sense here, and this will be skipped.")
                    msg_noCRC_Shown = True
                End If
                If Not msg_multipleRoms_Shown AndAlso node.SelectNodes("rom").Count > 1 Then
                    MsgBox("At least one of this dat file entry, have multiple roms, or not at all. This is not handled yet.")
                    msg_multipleRoms_Shown = True
                End If
                If Not crc.Trim = "" Then append_entry_sub(name, crc, md5, sha)
                ProgressBar1.BeginInvoke(Sub() ProgressBar1.Value += 1)
            Next
        Else
            'old format (romcenter)
            Dim dat_file As String = IO.File.ReadAllText(fb)
            Dim games() As String = Regex.Split(dat_file, "game", RegexOptions.IgnoreCase)
            ProgressBar1.Invoke(Sub() ProgressBar1.Maximum = games.Count)
            For Each game As String In games
                If game.ToLower.Contains("clrmamepro") Then Continue For
                Try
                    name = Regex.Match(game, "name\s*""(.*)""", RegexOptions.IgnoreCase).Groups(1).Value
                    description = Regex.Match(game, "description\s*""(.*)""", RegexOptions.IgnoreCase).Groups(1).Value
                    crc = Regex.Match(game, "rom\s*" + Regex.Escape("(") + ".*crc\s*(\S*)", RegexOptions.IgnoreCase).Groups(1).Value
                Catch ex As Exception
                    MsgBox(ex.Message)
                    Exit Sub
                End Try
                If Not crc.Trim = "" Then append_entry_sub(name, crc)
                ProgressBar1.BeginInvoke(Sub() ProgressBar1.Value += 1)
            Next
        End If

        Label4.Invoke(Sub() Label4.Text = "Saving asociation table...")
        Label4.Invoke(Sub() Label4.Visible = True)
        saveCurTable()
        Label4.Invoke(Sub() Label4.Visible = False)
        ProgressBar1.Invoke(Sub() ProgressBar1.Value = 0)
    End Sub


    'Append from rom folder
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If ComboBox1.SelectedIndex < 0 Then MsgBox("You have to select a system.") : Exit Sub
        If TextBox1.Text.Trim = "" Then MsgBox("You have to enter new set name.") : Exit Sub
        If ListBox1.Items.Contains(TextBox1.Text.Trim) Then MsgBox("This set name already exist in the table/") : Exit Sub

        Dim fb As New FolderBrowserDialog
        fb.Description = "Open rom folder"
        fb.ShowDialog()
        If fb.SelectedPath.Trim = "" Then MsgBox("No file selected.") : Exit Sub
        If Not FileIO.FileSystem.DirectoryExists(fb.SelectedPath) Then MsgBox("Directory """ + fb.SelectedPath + """ does not exist.") : Exit Sub
        bg_from_folder.RunWorkerAsync(fb.SelectedPath)
    End Sub
    Private Sub Button2_Click_bg(o As Object, e As DoWorkEventArgs) Handles bg_from_folder.DoWork
        Dim fb As String = DirectCast(e.Argument, String)

        Dim msg_notSingleFileInArch_Shown As Boolean = False

        Label4.Invoke(Sub() Label4.Visible = True)
        Label4.Invoke(Sub() Label4.Text = "Loading Table...")
        loadTable()
        If cur_table.Count = 0 Then
            cur_table.Add(New List(Of String))
            cur_table.Add(New List(Of String))
            cur_table.Add(New List(Of String))
            cur_table_headers.Add("CRC")
            cur_table_headers.Add("MD5")
            cur_table_headers.Add("SHA1")
        End If
        Label4.Invoke(Sub() Label4.Visible = False)

        Dim l As New List(Of String)
        For i As Integer = 0 To cur_table(0).Count - 1
            l.Add("")
        Next
        cur_table.Add(l)
        cur_table_headers.Add(TextBox1.Text.Trim)
        c_add = 0
        c_updt = 0
        c_total = cur_table(0).Count

        ProgressBar1.Invoke(Sub() ProgressBar1.Value = 0)
        Dim arch As New Class7_archives
        Dim hash As New Class6_hash()
        Dim files = GetFiles(fb)
        ProgressBar1.Invoke(Sub() ProgressBar1.Maximum = files.Count)
        For Each f In files
            Dim crc As String = ""
            Dim gameName As String = ""
            If arch.isArchive(f) Then
                arch.setFile(f)
                If arch.ArchiveFileData.Count = 1 Then
                    gameName = arch.ArchiveFileData(0).FileName
                    If gameName.Contains("\") Then gameName = gameName.Substring(gameName.LastIndexOf("\") + 1)
                    If gameName.Contains(".") Then gameName = gameName.Substring(0, gameName.LastIndexOf("."))

                    crc = Hex(arch.ArchiveFileData(0).Crc).Trim.ToUpper
                ElseIf Not msg_notSingleFileInArch_Shown Then
                    msg_notSingleFileInArch_Shown = True
                    MsgBox("At least one archive contains more or less than one file. This not handled yet. Skipping.")
                End If
            Else
                gameName = GetFileInfo(f).Name
                If gameName.Contains(".") Then gameName = gameName.Substring(0, gameName.LastIndexOf("."))

                crc = Class6_hash.GetCRC32(f).Trim.ToUpper
            End If
            append_entry_sub(gameName, crc)
            ProgressBar1.BeginInvoke(Sub() ProgressBar1.Value += 1)
        Next

        Label4.Invoke(Sub() Label4.Text = "Saving asociation table...")
        Label4.Invoke(Sub() Label4.Visible = True)
        saveCurTable()
        Label4.Invoke(Sub() Label4.Visible = False)
        ProgressBar1.Invoke(Sub() ProgressBar1.Value = 0)
    End Sub

    Private Sub saveCurTable()
        Dim tmp As String = ""
        If Not DirectoryExists(".\A-Tables") Then CreateDirectory(".\A-Tables")
        Dim sys As String = ""
        If ComboBox1.InvokeRequired Then
            ComboBox1.Invoke(Sub() sys = ComboBox1.SelectedItem.ToString.Trim())
        Else
            sys = ComboBox1.SelectedItem.ToString.Trim
        End If
        FileOpen(1, ".\A-Tables\" + sys + ".atbl", OpenMode.Output)
        For Each h As String In cur_table_headers
            tmp = tmp + h + separator
        Next
        PrintLine(1, tmp.Substring(0, tmp.Length - separator.Length))

        Dim rowCount As Integer = cur_table(0).Count
        For i As Integer = 0 To rowCount - 1
            tmp = ""
            For Each l In cur_table
                tmp = tmp + l(i) + separator
            Next
            PrintLine(1, tmp.Substring(0, tmp.Length - separator.Length))
        Next
        FileClose(1)
        Me.Invoke(Sub() ComboBox1_SelectedIndexChanged(ComboBox1, New EventArgs))
        MsgBox(c_updt.ToString + " games was updated (of " + c_total.ToString + "), and " + c_add.ToString + " games was added")
    End Sub

    Private Sub append_entry_sub(name As String, crc As String, Optional md5 As String = "", Optional sha As String = "")
        crc = crc.TrimStart("0"c)
        Dim ind As Integer = cur_table(0).IndexOf(crc.ToUpper)
        If ind >= 0 Then
            'check and update entry
            If md5.Trim <> "" Then
                Dim md5_orig As String = cur_table(1)(ind).Trim.ToUpper
                If md5_orig = "" Then cur_table(1)(ind) = md5.Trim.ToUpper
                If Not md5_orig = "" And Not md5.Trim.ToUpper = md5_orig Then MsgBox(name + " crc match but md5 does not! Skipping.") : Exit Sub
            End If
            If sha.Trim <> "" Then
                Dim sha_orig As String = cur_table(2)(ind).Trim.ToUpper
                If sha_orig = "" Then cur_table(2)(ind) = sha.Trim.ToUpper
                If Not sha_orig = "" And Not sha.Trim.ToUpper = sha_orig Then MsgBox(name + " crc match but sha1 does not! Skipping.") : Exit Sub
            End If
            cur_table(cur_table.Count - 1)(ind) = name
            c_updt += 1
        Else
            'add entry
            cur_table(0).Add(crc.Trim.ToUpper)
            cur_table(1).Add(md5.Trim.ToUpper)
            cur_table(2).Add(sha.Trim.ToUpper)
            cur_table(cur_table.Count - 1).Add(name)
            If cur_table.Count > 4 Then
                For i As Integer = 3 To cur_table.Count - 2
                    cur_table(i).Add("")
                Next
            End If
            c_add += 1
        End If
    End Sub

    'Show table
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If refr Then Exit Sub
        If CheckBox1.Checked Then
            If ComboBox1.SelectedIndex < 0 Then MsgBox("You have to select a system.") : refr = True : CheckBox1.Checked = False : refr = False : Exit Sub
            If Not FileExists(".\A-Tables\" + ComboBox1.SelectedItem.ToString.Trim + ".atbl") Then MsgBox("This table does not exist.") : refr = True : CheckBox1.Checked = False : refr = False : Exit Sub
            oldHeight = Me.Height
            oldWidth = Me.Width
            Me.MaximizeBox = True
            Me.MinimizeBox = True
            Me.FormBorderStyle = FormBorderStyle.Sizable

            loadTable()
            DataGridView1.Columns.Clear()
            DataGridView1.Rows.Clear()
            For Each c In cur_table_headers
                DataGridView1.Columns.Add(c, c)
            Next
            DataGridView1.Rows.Add(cur_table(0).Count)
            For c As Integer = 0 To cur_table.Count - 1
                Dim counter As Integer = 0
                For r As Integer = 0 To cur_table(c).Count - 1
                    DataGridView1.Rows(r).Cells(c).Value = cur_table(c)(r)
                    If cur_table(c)(r) <> "" Then counter += 1
                Next
                DataGridView1.Columns(c).HeaderText += "(" + counter.ToString + ")"
            Next
            DataGridView1.Top = 38
            DataGridView1.Height = Me.Height - 80
            DataGridView1.Width = Me.Width - 40
            DataGridView1.Visible = True
        Else
            Me.WindowState = FormWindowState.Normal
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            DataGridView1.Visible = False
            Me.FormBorderStyle = FormBorderStyle.Fixed3D
            Me.Height = oldHeight
            Me.Width = oldWidth
        End If
    End Sub
End Class