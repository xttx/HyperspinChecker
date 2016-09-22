Imports System.ComponentModel
Imports Microsoft.VisualBasic.FileIO.FileSystem

'TODO - delete files from list, when renaming

Public Class Form2_checkMissingInOtherFolders
    Structure renamer_param
        Dim dstPath As String
    End Structure
    Dim refr As Boolean = True
    Dim override_copy_mode(5) As Boolean
    Dim restore_textbox_focus As Boolean = False
    Dim crc_cache As New Dictionary(Of String, Dictionary(Of String, String))
    Dim corresponding_games As New Dictionary(Of String, String)
    Dim allowRename As MsgBoxResult = MsgBoxResult.Retry
    Dim allowRenameInArchive As MsgBoxResult = MsgBoxResult.Retry
    Dim allowDeleteAllButGameInArchive As MsgBoxResult = MsgBoxResult.Retry
    Dim selectedMediaType As Integer = -1
    Dim WithEvents bg_fillist As New BackgroundWorker() With {.WorkerReportsProgress = True}
    Dim WithEvents bg_rename As New BackgroundWorker() With {.WorkerReportsProgress = True}

    Private Sub Form2_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Language.localize(Me)
        ProgressBar1.Value = 0
        ComboBox1.SelectedIndex = Class1.i : refr = False
        Select Case ComboBox1.SelectedIndex
            Case 0
                TextBox1.Text = Class1.romPath
            Case 1
                TextBox1.Text = Class1.videoPath
            Case 2
                TextBox1.Text = Class1.HyperspinPath + "Media\" + Form1.ComboBox1.SelectedItem.ToString + "\Themes\"
            Case 3
                TextBox1.Text = Class1.HyperspinPath + "Media\" + Form1.ComboBox1.SelectedItem.ToString + "\Images\Wheel\"
            Case 4
                TextBox1.Text = Class1.HyperspinPath + "Media\" + Form1.ComboBox1.SelectedItem.ToString + "\Images\Artwork1\"
            Case 5
                TextBox1.Text = Class1.HyperspinPath + "Media\" + Form1.ComboBox1.SelectedItem.ToString + "\Images\Artwork2\"
            Case 6
                TextBox1.Text = Class1.HyperspinPath + "Media\" + Form1.ComboBox1.SelectedItem.ToString + "\Images\Artwork3\"
            Case 7
                TextBox1.Text = Class1.HyperspinPath + "Media\" + Form1.ComboBox1.SelectedItem.ToString + "\Images\Artwork4\"
        End Select

        If ComboBox1.SelectedIndex = 0 Then
            CheckBox1.Enabled = True : CheckBox1.Tag = ""
            CheckBox2.Enabled = True : CheckBox2.Tag = ""
        Else
            CheckBox1.Enabled = False : CheckBox1.Tag = "DISABLED"
            CheckBox2.Enabled = False : CheckBox2.Tag = "DISABLED"
        End If
    End Sub

    'Change combobox - Media Type
    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If refr Then Exit Sub
        fillList()
        If ComboBox1.SelectedIndex = 0 Then
            CheckBox1.Enabled = True : CheckBox1.Tag = ""
            CheckBox2.Enabled = True : CheckBox2.Tag = ""
        Else
            CheckBox1.Enabled = False : CheckBox1.Tag = "DISABLED"
            CheckBox2.Enabled = False : CheckBox2.Tag = "DISABLED"
        End If
    End Sub

    'Change textbox - path
    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged
        fillList()
    End Sub

    'Change radiobutton - switch missing/all
    Private Sub RadioButton1_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles RadioButton1.CheckedChanged
        fillList()
    End Sub

    'Fill filelist
    Private Sub fillList()
        ListBox1.Items.Clear()
        corresponding_games.Clear()
        If Not DirectoryExists(TextBox1.Text) Then Exit Sub
        If TextBox1.Focused Then restore_textbox_focus = True Else restore_textbox_focus = False
        For Each c As Control In Me.Controls
            If Not c.Name.ToUpper.StartsWith("LISTBOX") Then c.Enabled = False
        Next
        ProgressBar1.Value = 0
        selectedMediaType = ComboBox1.SelectedIndex
        bg_fillist.RunWorkerAsync(TextBox1.Text)
    End Sub
    Private Sub fillList_BG(o As Object, e As DoWorkEventArgs) Handles bg_fillist.DoWork
        Dim index As Integer
        Dim progress As Integer = 0
        Dim fWoExt As String = ""
        Dim allowedExt As String = ""
        Dim useCRC As Boolean = CheckBox2.Checked
        Dim path As String = DirectCast(e.Argument, String)
        If Microsoft.VisualBasic.FileIO.FileSystem.DirectoryExists(path) Then
            Dim realdir As String = Microsoft.VisualBasic.FileIO.FileSystem.GetDirectoryInfo(path).FullName.ToUpper
            If Not realdir.EndsWith("\") Then realdir = realdir + "\"
            If useCRC AndAlso Not crc_cache.ContainsKey(realdir) Then crc_cache.Add(realdir, New Dictionary(Of String, String))

            Dim files As Collections.ObjectModel.ReadOnlyCollection(Of String)
            files = Microsoft.VisualBasic.FileIO.FileSystem.GetFiles(path)
            ProgressBar1.Invoke(Sub() ProgressBar1.Maximum = files.Count)
            For Each f As String In files
                If Not useCRC Then
                    'Use names
                    f = f.Substring(f.LastIndexOf("\") + 1)
                    fWoExt = f.Substring(0, f.LastIndexOf("."))
                    index = Class1.romlist.IndexOf(fWoExt.ToLower)
                    fillList_addToList(index, f, False)
                Else
                    'use crc
                    Dim crc As String = ""
                    Dim z As New Class7_archives
                    If z.isArchive(f) Then
                        'if it is archive
                        z.setFile(f)
                        For Each fileInfo In z.ArchiveFileData
                            If crc_cache(realdir).ContainsKey(f.ToUpper) Then crc = crc_cache(realdir)(f.ToUpper) Else crc = Hex(fileInfo.Crc).ToUpper : crc_cache(realdir).Add(f.ToUpper, crc)
                            index = Class1.data_crc.IndexOf(crc)
                            fillList_addToList(index, f)
                        Next
                    ElseIf CheckBox1.Checked Then
                        'if it is NOT an archive
                        If crc_cache(realdir).ContainsKey(f.ToUpper) Then crc = crc_cache(realdir)(f.ToUpper) Else crc = Class6_hash.GetCRC32(f).ToUpper : crc_cache(realdir).Add(f.ToUpper, crc)
                        index = Class1.data_crc.IndexOf(crc)
                        fillList_addToList(index, f)
                    End If
                End If
                'ProgressBar1.Value += 1
                progress += 1 : bg_fillist.ReportProgress(progress)
            Next
        End If
    End Sub
    Private Sub fillList_BG_progress(o As Object, p As ProgressChangedEventArgs) Handles bg_fillist.ProgressChanged
        ProgressBar1.Value = p.ProgressPercentage
    End Sub
    Private Sub fillList_BG_complete() Handles bg_fillist.RunWorkerCompleted
        For Each c As Control In Me.Controls
            If c.Tag Is Nothing OrElse c.Tag.ToString <> "DISABLED" Then c.Enabled = True
        Next
        If restore_textbox_focus Then TextBox1.Focus()
        ProgressBar1.Value = 0
        Label3.Text = "Total found: " + ListBox1.Items.Count.ToString
    End Sub
    'Fill filelist - Sub
    Private Sub fillList_addToList(index As Integer, f As String, Optional filenameContainPath As Boolean = True)
        Dim fname As String = f
        If index >= 0 Then
            If corresponding_games.Keys.Contains(Class1.romlist(index).ToString) Then Exit Sub
            If filenameContainPath Then fname = fname.Substring(fname.LastIndexOf("\") + 1)
            If RadioButton1.Checked Then
                If Class1.data(index).Substring(selectedMediaType, 1) = "N" Then
                    corresponding_games.Add(Class1.romlist(index).ToString, f) : ListBox1.BeginInvoke(Sub() ListBox1.Items.Add(fname))
                End If
            Else
                corresponding_games.Add(Class1.romlist(index).ToString, f) : ListBox1.BeginInvoke(Sub() ListBox1.Items.Add(fname))
            End If
        End If
    End Sub

    'CRC mode toggle
    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        fillList()
    End Sub
    'CRC - scan all/missing toggle
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox2.Checked Then fillList()
    End Sub

    'GO
    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        If ListBox1.Items.Count = 0 Then MsgBox("Nothing to copy.") : Exit Sub

        Dim dstPath As String = ""
        ProgressBar1.Value = 0
        Select Case ComboBox1.SelectedIndex
            Case 0
                dstPath = Class1.romPath
            Case 1
                dstPath = Class1.videoPath
            Case 2
                dstPath = Class1.HyperspinPath + "Media\" + Form1.ComboBox1.SelectedItem.ToString + "\Themes\"
            Case 3
                dstPath = Class1.HyperspinPath + "Media\" + Form1.ComboBox1.SelectedItem.ToString + "\Images\Wheel\"
            Case 4
                dstPath = Class1.HyperspinPath + "Media\" + Form1.ComboBox1.SelectedItem.ToString + "\Images\Artwork1\"
            Case 5
                dstPath = Class1.HyperspinPath + "Media\" + Form1.ComboBox1.SelectedItem.ToString + "\Images\Artwork2\"
            Case 6
                dstPath = Class1.HyperspinPath + "Media\" + Form1.ComboBox1.SelectedItem.ToString + "\Images\Artwork3\"
            Case 7
                dstPath = Class1.HyperspinPath + "Media\" + Form1.ComboBox1.SelectedItem.ToString + "\Images\Artwork4\"
        End Select

        For Each c As Control In Me.Controls
            If Not c.Name.ToUpper.StartsWith("LISTBOX") Then c.Enabled = False
        Next

        Class7_archives.lastRespons_rename = MsgBoxResult.Retry
        Class7_archives.lastRespons_keepOne = MsgBoxResult.Retry
        Class7_archives.rename = Class7_archives.answer.ask_once
        Class7_archives.keep_only_one = Class7_archives.answer.ask_once

        'Ovverwride matcher copy mode
        Dim frm As Form1 = DirectCast(Application.OpenForms("Form1"), Form1)
        override_copy_mode(0) = frm.AssocOption_fileInDiffFolder_copy.Checked
        override_copy_mode(1) = frm.AssocOption_fileInDiffFolder_copyToHS.Checked
        override_copy_mode(2) = frm.AssocOption_fileInDiffFolder_move.Checked
        override_copy_mode(3) = frm.AssocOption_fileInDiffFolder_moveToHS.Checked
        override_copy_mode(4) = frm.AssocOption_fileInHsFolder_copy.Checked
        override_copy_mode(5) = frm.AssocOption_fileInHsFolder_move.Checked
        frm.AssocOption_fileInDiffFolder_copy.Checked = False
        frm.AssocOption_fileInDiffFolder_copyToHS.Checked = True
        frm.AssocOption_fileInDiffFolder_move.Checked = False
        frm.AssocOption_fileInDiffFolder_moveToHS.Checked = False
        frm.AssocOption_fileInHsFolder_copy.Checked = False
        frm.AssocOption_fileInHsFolder_move.Checked = True

        Dim param As New renamer_param With {.dstPath = dstPath}
        bg_rename.RunWorkerAsync(param)
    End Sub
    Private Sub Button2_Click_BG(o As Object, e As DoWorkEventArgs) Handles bg_rename.DoWork
        Dim progress As Integer = 0
        Dim param As renamer_param = DirectCast(e.Argument, renamer_param)
        Dim frm As Form1 = DirectCast(Application.OpenForms("Form1"), Form1)

        If Not CheckBox2.Checked Then
            'name mode
            ProgressBar1.Invoke(Sub() ProgressBar1.Maximum = ListBox1.Items.Count)
            For Each filename As String In ListBox1.Items
                Dim fNameWoExt = filename
                If fNameWoExt.Contains(".") Then fNameWoExt = fNameWoExt.Substring(0, fNameWoExt.LastIndexOf("."))
                'CopyFile(TextBox1.Text + "\" + filename, param.dstPath + "\" + filename)
                frm.matcher_class.associate_copyMove(TextBox1.Text, param.dstPath, fNameWoExt, filename)
                progress += 1
                If progress Mod 3 = 0 Then bg_rename.ReportProgress(progress)
            Next
        Else
            'crc mode
            ProgressBar1.Invoke(Sub() ProgressBar1.Maximum = corresponding_games.Count)

            'Dim z As New Class7_archives

            Dim realdir As String = Microsoft.VisualBasic.FileIO.FileSystem.GetDirectoryInfo(TextBox1.Text).FullName.ToUpper
            If Not realdir.EndsWith("\") Then realdir = realdir + "\"
            If Not crc_cache.ContainsKey(realdir) Then MsgBox("CRC cache is empty.") : Exit Sub
            For Each item In corresponding_games
                Dim f = GetFileInfo(item.Value).Name
                Dim fWoExt As String = f.Substring(0, f.LastIndexOf("."))

                If fWoExt.ToUpper <> item.Key.ToUpper Then
                    If allowRename = MsgBoxResult.Retry Then allowRename = MsgBox("At least one of files have incorrect name, do you want to rename files while copying?", MsgBoxStyle.YesNo)
                End If
                If allowRename = MsgBoxResult.Yes Then
                    frm.matcher_class.associate_copyMove(TextBox1.Text, param.dstPath, item.Key, f)
                Else
                    frm.matcher_class.associate_copyMove(TextBox1.Text, param.dstPath, fWoExt, f)
                End If

                'If Not z.isArchive(item.Value) Then
                '    'file mode
                '    If fWoExt.ToUpper <> item.Key.ToUpper Then
                '        If allowRename = MsgBoxResult.Retry Then allowRename = MsgBox("At least one of files have incorrect name, do you want to rename files while copying?", MsgBoxStyle.YesNo)
                '    End If
                '    If fWoExt.ToUpper <> item.Key.ToUpper Then
                '        'Button2_Click_sub(item.Value, param.dstPath, item.Key)
                '    End If
                'Else
                'archive mode
                '    z.setFile(item.Value)
                '    'Dim rename As Boolean = False
                '    If fWoExt.ToUpper <> item.Key.ToUpper Then
                '        If allowRename = MsgBoxResult.Retry Then allowRename = MsgBox("At least one of files have incorrect name, do you want to rename files while copying?", MsgBoxStyle.YesNo)
                '    End If

                '    'check if filename inside have wrong name, and ask if it should recompress to rename file inside
                '    Dim arcF As String = ""
                '    For Each a In z.ArchiveFileData
                '        If Hex(a.Crc).ToUpper = crc_cache(realdir)(item.Value.ToUpper) Then arcF = a.FileName : Exit For
                '    Next
                '    If arcF = "" Then MsgBox("Needed file not found in archive. Something wrong. Aborting.") : Exit Sub
                '    If arcF.Contains("\") And arcF.IndexOf("\") <> arcF.Length - 1 Then arcF = arcF.Substring(arcF.LastIndexOf("\") + 1)
                '    Dim arcFWoExt As String = arcF : If arcF.Contains(".") Then arcFWoExt = arcF.Substring(0, arcF.LastIndexOf("."))
                '    If arcFWoExt.ToUpper <> item.Key.ToUpper Then
                '        If allowRenameInArchive = MsgBoxResult.Retry Then allowRenameInArchive = MsgBox("At least one of files, INSIDE ARCHIVE have incorrect name, do you want to rename files while copying?", MsgBoxStyle.YesNo)
                '        If allowRenameInArchive = MsgBoxResult.Yes Then
                '            'rename file in archive
                '            'check if it have multiple files and ask if it should just keep one, when rearchiving
                '            If z.ArchiveFileData.Count > 1 Then
                '                If allowDeleteAllButGameInArchive = MsgBoxResult.Retry Then allowDeleteAllButGameInArchive = MsgBox("At least one of archive contains more than one files. Do you want to remove other files and only keep needed file?", MsgBoxStyle.YesNo)
                '            End If
                '            Button2_Click_sub(item.Value, param.dstPath, item.Key, True, crc_cache(realdir)(item.Value.ToUpper))
                '        Else
                '            'just copy archive
                '            Button2_Click_sub(item.Value, param.dstPath, item.Key)
                '        End If
                '    End If
                'End If
                progress += 1
                If progress Mod 3 = 0 Then bg_rename.ReportProgress(progress)
            Next
        End If
    End Sub
    Private Sub Button2_Click_BG_Progress(o As Object, e As ProgressChangedEventArgs) Handles bg_rename.ProgressChanged
        ProgressBar1.Value = e.ProgressPercentage
    End Sub
    Private Sub Button2_Click_BG_Complete() Handles bg_rename.RunWorkerCompleted
        For Each c As Control In Me.Controls
            c.Enabled = True
        Next

        Dim frm As Form1 = DirectCast(Application.OpenForms("Form1"), Form1)
        frm.AssocOption_fileInDiffFolder_copy.Checked = override_copy_mode(0)
        frm.AssocOption_fileInDiffFolder_copyToHS.Checked = override_copy_mode(1)
        frm.AssocOption_fileInDiffFolder_move.Checked = override_copy_mode(2)
        frm.AssocOption_fileInDiffFolder_moveToHS.Checked = override_copy_mode(3)
        frm.AssocOption_fileInHsFolder_copy.Checked = override_copy_mode(4)
        frm.AssocOption_fileInHsFolder_move.Checked = override_copy_mode(5)
        MsgBox("Done.")
        Me.Close()
    End Sub
    'GO Sub
    Private Sub Button2_Click_sub(src As String, dstPath As String, gamename As String, Optional archiveMode As Boolean = False, Optional crc As String = "")
        Dim f = GetFileInfo(src).Name
        Dim ext As String = GetFileInfo(src).Extension
        Dim fWoExt As String = f.Substring(0, f.LastIndexOf("."))
        Dim i1 As String = GetFileInfo(src).FullName
        Dim i2 As String = ""
        If Not archiveMode Then
            If fWoExt.ToUpper <> gamename.ToUpper And allowRename = MsgBoxResult.Yes Then
                i2 = GetFileInfo(dstPath + "\" + gamename + ext).FullName
            Else
                i2 = GetFileInfo(dstPath + "\" + f).FullName
            End If
            If Not i1.ToUpper = i2.ToUpper Then CopyFile(i1, i2)
        Else
            Dim z As New Class7_archives
            z.setFile(src)
            Dim only_keep_crc As Boolean = False
            If z.ArchiveFileData.Count > 1 And allowDeleteAllButGameInArchive = MsgBoxResult.Yes Then only_keep_crc = True
            If allowRename = MsgBoxResult.Yes Then
                z.renameInArchive(dstPath + "\" + gamename, gamename, crc, only_keep_crc)
            Else
                z.renameInArchive(dstPath + "\" + fWoExt, gamename, crc, only_keep_crc)
            End If
        End If
    End Sub

    'button ...
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim fb As New FolderBrowserDialog
        If fb.ShowDialog = DialogResult.OK Then
            TextBox1.Text = fb.SelectedPath
        End If
    End Sub
End Class