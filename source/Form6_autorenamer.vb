Imports System.ComponentModel
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic.FileIO.FileSystem
'TODO - scan in background
'TODO - need to handle archives when renaming
'TODO - what is this error, when renaming Atari ST?
Public Class Form6_autorenamer
    'Dim local_list1 As New List(Of String)
    'Dim local_list2 As New List(Of String)
    Dim WithEvents bg_scaner As New BackgroundWorker()
    Dim WithEvents bg_scaner_crc As New BackgroundWorker()
    Dim WithEvents bg_renamer As New BackgroundWorker()
    Dim clean_game_names, clean_rom_names As List(Of String)
    Dim fileCrc As New Dictionary(Of String, String)
    Dim allowRenameInsideArchive As MsgBoxResult = MsgBoxResult.Retry
    Dim allowKeepOnlyNeeded As MsgBoxResult = MsgBoxResult.Retry
    Structure listboxes_local_lists
        Dim list1 As ListBox.ObjectCollection
        Dim list2 As ListBox.ObjectCollection
        Dim path As String
    End Structure

    'Button Collect Info
    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        If Not CheckBox8.Checked And Not IsNumeric(TextBox3.Text) Then MsgBox("Please, check your min ratio % param.") : Exit Sub

        For Each c As Control In Me.Controls
            If Not c.Name.ToUpper.StartsWith("DATAGRID") Then c.Enabled = False
        Next
        'Button2.Enabled = False
        'Button3.Enabled = False
        DataGridView1.Rows.Clear()
        ProgressBar1.Minimum = 0
        ProgressBar1.Value = 0
        Dim lists As New listboxes_local_lists With {.list1 = Form1.ListBox1.Items, .list2 = Form1.ListBox2.Items}

        If CheckBox8.Checked Then
            'crc mode
            lists.path = Form1.TextBox4.Text + "\"
            bg_scaner_crc.WorkerReportsProgress = True
            bg_scaner_crc.RunWorkerAsync(lists)
        Else
            'names mode
            bg_scaner.WorkerReportsProgress = True
            bg_scaner.RunWorkerAsync(lists)
        End If
    End Sub
    Private Sub Button2_Click_BGWork(sender As Object, e As DoWorkEventArgs) Handles bg_scaner.DoWork
        'Preprocess
        fileCrc.Clear()
        Dim progress As Integer = 0
        Dim min_ratio As Integer = CInt(TextBox3.Text)
        Dim curGameName, curRomName As String
        clean_game_names = New List(Of String)
        clean_rom_names = New List(Of String)
        Dim list_of_added_roms = New List(Of String)
        Dim list_of_added_games = New List(Of String)
        Dim lists As listboxes_local_lists = DirectCast(e.Argument, listboxes_local_lists)
        'For Each item As String In Form1.ListBox1.Items
        For Each item As String In lists.list1
            curGameName = item.ToUpper.Trim
            If CheckBox1.Checked Then curGameName = Remove_paranteses(curGameName)
            If CheckBox2.Checked Then curGameName = Remove_brackets(curGameName)
            If CheckBox3.Checked Then curGameName = Remove_Special(curGameName)
            If CheckBox4.Checked Then curGameName = Remove_words(curGameName)
            If CheckBox5.Checked Then curGameName = Convert_roman(curGameName)
            clean_game_names.Add(Remove_spaces(curGameName))
        Next
        'For Each item As DataRowView In Form1.ListBox2.Items
        For Each item As DataRowView In lists.list2
            curRomName = item.Item(0).ToString.ToUpper.Trim
            If CheckBox1.Checked Then curRomName = Remove_paranteses(curRomName)
            If CheckBox2.Checked Then curRomName = Remove_brackets(curRomName)
            If CheckBox3.Checked Then curRomName = Remove_Special(curRomName)
            If CheckBox4.Checked Then curRomName = Remove_words(curRomName)
            If CheckBox5.Checked Then curRomName = Convert_roman(curRomName)
            clean_rom_names.Add(Remove_spaces(curRomName))
        Next

        'Comparing
        Me.Invoke(New MethodInvoker(Sub() ProgressBar1.Maximum = clean_rom_names.Count))
        Dim counter1 As Integer = 0, counter2 As Integer = 0
        Dim numbersInRomName As String = "", numbersInGameName As String = ""
        Dim stringSplit1(), stringSplit2() As String
        Dim commonList As Collections.Generic.IEnumerable(Of String)
        Dim similarity As Double
        Dim t1, t2 As String
        For Each curRomName In clean_rom_names
            counter1 = 0
            stringSplit1 = curRomName.Split(" "c)
            If CheckBox7.Checked Then numbersInRomName = Regex.Replace(curRomName, "[^\d]", "")
            For Each curGameName In clean_game_names
                If CheckBox7.Checked Then numbersInGameName = Regex.Replace(curGameName, "[^\d]", "")
                If numbersInRomName = numbersInGameName Then
                    stringSplit2 = curGameName.Split(" "c)
                    commonList = stringSplit1.Intersect(stringSplit2)
                    similarity = 100 * (commonList.Count * 2) / (stringSplit1.Length + stringSplit2.Length)
                    If similarity >= min_ratio Then
                        Dim X As String = "X"
                        't1 = Form1.ListBox1.Items(counter1).ToString
                        t1 = lists.list1(counter1).ToString
                        't2 = DirectCast(Form1.ListBox2.Items(counter2), DataRowView).Item(0).ToString
                        t2 = DirectCast(lists.list2(counter2), DataRowView).Item(0).ToString
                        If list_of_added_roms.Contains(t2) Then X = "" 'Else list_of_added_roms.Add(t2)
                        If list_of_added_games.Contains(t1) Then X = "" 'Else list_of_added_games.Add(t1)
                        'DataGridView1.Rows.Add({t1, t2, Math.Round(similarity, 2, MidpointRounding.AwayFromZero).ToString, X})
                        Me.Invoke(New MethodInvoker(Sub() DataGridView1.Rows.Add({t1, t2, Math.Round(similarity, 2, MidpointRounding.AwayFromZero).ToString, X})))
                        If X = "X" AndAlso Not list_of_added_roms.Contains(t2) Then list_of_added_roms.Add(t2)
                        If X = "X" AndAlso Not list_of_added_games.Contains(t1) Then list_of_added_games.Add(t1)
                    End If
                End If
                counter1 += 1
            Next
            counter2 += 1
            progress += 1
            If progress Mod 10 = 0 Then bg_scaner.ReportProgress(progress)
            'ProgressBar1.Value += 1
            'If ProgressBar1.Value Mod 10 = 0 Then ProgressBar1.Refresh()
        Next
    End Sub
    Private Sub Button2_Click_BGWork_complete() Handles bg_scaner.RunWorkerCompleted
        'Button2.Enabled = True
        For Each c As Control In Me.Controls
            If Not c.Name.ToUpper.StartsWith("DATAGRID") Then c.Enabled = True
        Next
        ProgressBar1.Value = 0
        DataGridView1.AutoResizeColumns() : DataGridView1.AutoResizeColumnHeadersHeight()
        If DataGridView1.Rows.Count > 0 Then Button3.Enabled = True Else Button3.Enabled = False
    End Sub
    Private Sub Button2_Click_BGWork_crc_mode(sender As Object, e As DoWorkEventArgs) Handles bg_scaner_crc.DoWork
        fileCrc.Clear()
        Dim z As New Class7_archives
        Dim lists As listboxes_local_lists = DirectCast(e.Argument, listboxes_local_lists)
        Dim path As String = lists.path
        Dim progress As Integer = 0

        Dim crc As String = ""
        Dim index As Integer = -1
        Me.Invoke(New MethodInvoker(Sub() ProgressBar1.Maximum = lists.list2.Count))
        Dim list_of_added_roms = New List(Of String)
        Dim list_of_added_games = New List(Of String)
        For Each item As DataRowView In lists.list2
            Dim file As String = item(0).ToString
            If z.isArchive(item(0).ToString) Then
                z.setFile(path + item(0).ToString)
                For Each a_data In z.ArchiveFileData
                    crc = Hex(a_data.Crc).ToUpper
                    index = Class1.data_crc.IndexOf(crc)
                    If index >= 0 AndAlso Class1.data(index).Substring(0, 1) = "N" Then
                        Dim x As String = "X"
                        Dim game As String = Class1.romlist(index).ToUpper
                        Dim RealFilePath As String = GetFileInfo(path + item(0).ToString).FullName.ToUpper
                        fileCrc.Add(RealFilePath, crc)
                        If list_of_added_roms.Contains(file.ToUpper) Then x = ""
                        If list_of_added_games.Contains(game.ToUpper) Then x = ""
                        'DataGridView1.Rows.Add({Class1.romlist(index), item(0).ToString, "100", x})
                        Me.Invoke(New MethodInvoker(Sub() DataGridView1.Rows.Add({Class1.romlist(index), item(0).ToString, "100", x})))
                        If x = "X" AndAlso Not list_of_added_roms.Contains(file.ToUpper) Then list_of_added_roms.Add(file.ToUpper)
                        If x = "X" AndAlso Not list_of_added_games.Contains(game.ToUpper) Then list_of_added_games.Add(game.ToUpper)
                    End If
                Next
            Else
                crc = Class6_hash.GetCRC32(path + item(0).ToString).ToUpper
                index = Class1.data_crc.IndexOf(crc)
                If index >= 0 AndAlso Class1.data(index).Substring(0, 1) = "N" Then
                    Dim x As String = "X"
                    Dim game As String = Class1.romlist(index).ToUpper
                    If list_of_added_roms.Contains(file.ToUpper) Then x = ""
                    If list_of_added_games.Contains(game.ToUpper) Then x = ""
                    'DataGridView1.Rows.Add({Class1.romlist(index), item(0).ToString, "100", x})
                    Me.Invoke(New MethodInvoker(Sub() DataGridView1.Rows.Add({Class1.romlist(index), item(0).ToString, "100", x})))
                    If x = "X" AndAlso Not list_of_added_roms.Contains(file.ToUpper) Then list_of_added_roms.Add(file.ToUpper)
                    If x = "X" AndAlso Not list_of_added_games.Contains(game.ToUpper) Then list_of_added_games.Add(game.ToUpper)
                End If
            End If
            progress += 1
            If progress Mod 10 = 0 Then bg_scaner_crc.ReportProgress(progress)
            'ProgressBar1.Value += 1
            'If ProgressBar1.Value Mod 10 = 0 Then ProgressBar1.Refresh()
        Next
    End Sub
    Private Sub Button2_Click_BGWork_crc_mode_complete() Handles bg_scaner_crc.RunWorkerCompleted
        ProgressBar1.Value = 0
        'Button2.Enabled = True
        For Each c As Control In Me.Controls
            If Not c.Name.ToUpper.StartsWith("DATAGRID") Then c.Enabled = True
        Next
        DataGridView1.AutoResizeColumns() : DataGridView1.AutoResizeColumnHeadersHeight()
        If DataGridView1.Rows.Count > 0 Then Button3.Enabled = True Else Button3.Enabled = False
    End Sub

    'Button RENAME
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Button2.Enabled = False : Button3.Enabled = False

        ProgressBar1.Minimum = 0
        ProgressBar1.Maximum = DataGridView1.Rows.Count
        ProgressBar1.Value = 0
        bg_renamer.WorkerReportsProgress = True
        bg_renamer.RunWorkerAsync(Form1.TextBox4.Text)
    End Sub
    Private Sub Button3_Click_BGWork(sender As Object, e As DoWorkEventArgs) Handles bg_renamer.DoWork
        allowRenameInsideArchive = MsgBoxResult.Retry
        allowKeepOnlyNeeded = MsgBoxResult.Retry
        Dim fname As String = ""
        Dim ext As String = ""
        Dim path As String = e.Argument.ToString
        Dim progress As Integer = 0
        Dim z As New Class7_archives
        If Not path.EndsWith("\") Then path = path + "\"
        For Each r As DataGridViewRow In DataGridView1.Rows
            If r.Cells(3).Value.ToString = "X" Then
                Try
                    Dim f = GetFiles(path, FileIO.SearchOption.SearchTopLevelOnly, {r.Cells(1).Value.ToString + ".*"})
                    If f.Count > 0 Then
                        fname = f(0)
                        ext = fname.Substring(fname.LastIndexOf("."))
                        Rename(fname, path + r.Cells(0).Value.ToString + ext)
                        If z.isArchive(path + r.Cells(0).Value.ToString + ext) Then
                            'Handle archives
                            Dim realFilePath As String = GetFileInfo(path + r.Cells(0).Value.ToString + ext).FullName.ToUpper
                            z.setFile(realFilePath)
                            If fileCrc.ContainsKey(GetFileInfo(f(0)).FullName.ToUpper) Then
                                For Each a In z.ArchiveFileData
                                    If Hex(a.Crc).ToUpper = fileCrc(GetFileInfo(f(0)).FullName.ToUpper) Then
                                        Dim a_fName As String = a.FileName
                                        If a_fName.Contains("\") Then a_fName = a_fName.Substring(a_fName.LastIndexOf("\") + 1)
                                        Dim a_fNameWoExt As String = a_fName
                                        If a_fNameWoExt.Contains(".") Then a_fNameWoExt = a_fNameWoExt.Substring(0, a_fNameWoExt.LastIndexOf("."))
                                        If a_fNameWoExt.ToUpper <> r.Cells(0).Value.ToString.ToUpper Then
                                            If allowRenameInsideArchive = MsgBoxResult.Retry Then
                                                allowRenameInsideArchive = MsgBox("At least one of files, INSIDE ARCHIVE have incorrect name, do you want to rename files while copying?", MsgBoxStyle.YesNo)
                                            End If
                                            If allowRenameInsideArchive = MsgBoxResult.Yes Then
                                                Dim realfNameWoExt As String = realFilePath
                                                If realfNameWoExt.Contains(".") Then realfNameWoExt = realfNameWoExt.Substring(0, realfNameWoExt.LastIndexOf("."))
                                                If z.ArchiveFileData.Count > 1 Then
                                                    If allowKeepOnlyNeeded = MsgBoxResult.Retry Then
                                                        allowKeepOnlyNeeded = MsgBox("At least one of archive contains more than one files. Do you want to remove other files and only keep needed file?", MsgBoxStyle.YesNo)
                                                    End If
                                                    If allowKeepOnlyNeeded = MsgBoxResult.Yes Then
                                                        'rename inside arch and keep only crc
                                                        z.renameInArchive(realfNameWoExt, r.Cells(0).Value.ToString, fileCrc(GetFileInfo(f(0)).FullName.ToUpper), True)
                                                    Else
                                                        'jest rename inside arch
                                                        z.renameInArchive(realfNameWoExt, r.Cells(0).Value.ToString, fileCrc(GetFileInfo(f(0)).FullName.ToUpper))
                                                    End If
                                                Else
                                                    z.renameInArchive(realfNameWoExt, r.Cells(0).Value.ToString, fileCrc(GetFileInfo(f(0)).FullName.ToUpper))
                                                End If
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End If
            progress += 1
            If progress Mod 10 = 0 Then bg_renamer.ReportProgress(progress)
        Next
    End Sub
    Private Sub Button3_Click_BGWork_complete() Handles bg_renamer.RunWorkerCompleted
        ProgressBar1.Value = 0
        Button2.Enabled = True
        MsgBox("Done!")
        Me.Close()
    End Sub

    Private Function Remove_paranteses(s As String) As String
        'If s.IndexOf("(") >= 0 Then s = s.Substring(0, s.IndexOf("("))
        'Return s.Trim

        Dim lastIndex As Integer = 0
        Dim ind1 As Integer = 0, ind2 As Integer = 0
        Dim cnt As String = ""
        Do While s.IndexOf("(", lastIndex) > 0
            ind1 = s.IndexOf("(", lastIndex)
            ind2 = s.IndexOf(")", ind1)
            If ind2 > 0 And Not ind2 = s.Length - 1 Then
                s = s.Substring(0, ind1) + s.Substring(ind2 + 1)
            Else
                'Not pair or ')' is a last symbol in string
                s = s.Substring(0, ind1).Trim
                Exit Do
            End If
            s = s.Replace("  ", " ").Replace("  ", " ").Replace("  ", " ").Trim
        Loop
        Return s.Trim
    End Function
    Private Function Remove_brackets(s As String) As String
        'If s.IndexOf("[") >= 0 Then s = s.Substring(0, s.IndexOf("["))
        'Return s.Trim

        Dim lastIndex As Integer = 0
        Dim ind1 As Integer = 0, ind2 As Integer = 0
        Do While s.IndexOf("[", lastIndex) > 0
            ind1 = s.IndexOf("[", lastIndex)
            ind2 = s.IndexOf("]", ind1)
            If ind2 > 0 And Not ind2 = s.Length - 1 Then
                s = s.Substring(0, ind1) + s.Substring(ind2 + 1)
            Else
                'Not pair or ']' is a last symbol in string
                s = s.Substring(0, ind1).Trim
                Exit Do
            End If
            s = s.Replace("  ", " ").Replace("  ", " ").Replace("  ", " ").Trim
        Loop
        Return s.Trim
    End Function
    Private Function Remove_Special(s As String) As String
        If TextBox1.Text.Trim = "" Then Return s
        For Each c As Char In TextBox1.Text.Trim
            If c <> "" Then s = s.Replace(c, " ")
        Next
        Return s
    End Function
    Private Function Remove_words(s As String) As String
        If TextBox2.Text.Trim = "" Then Return s
        For Each w As String In TextBox2.Text.Trim.Split({","}, StringSplitOptions.RemoveEmptyEntries)
            w = w.Trim.ToUpper
            If w <> "" Then s = s.Replace(" " + w, " ").Replace(w + " ", " ")
        Next
        Return s
    End Function
    Private Function Convert_roman(s As String) As String
        Dim t As String = s
        For Each word As String In t.Split({" "}, StringSplitOptions.RemoveEmptyEntries)
            If word.Replace("I", "").Replace("V", "").Replace("X", "").Trim = "" Then s = s.Replace(word, Convert_roman2(word))
        Next
        Return s
    End Function
    Private Function Convert_roman2(s As String) As String
        Dim ch As String
        Dim new_value, old_value As Integer
        Dim result As Integer = 0

        old_value = 1000
        For i As Integer = 0 To s.Length - 1
            ch = s.Substring(i, 1)
            Select Case ch
                Case "I"
                    new_value = 1
                Case "V"
                    new_value = 5
                Case "X"
                    new_value = 10
                Case "L"
                    new_value = 50
                Case "C"
                    new_value = 100
                Case "D"
                    new_value = 500
                Case "M"
                    new_value = 1000
            End Select
            If new_value > old_value Then
                result = result + new_value - 2 * old_value
            Else
                result = result + new_value
            End If
            old_value = new_value
        Next
        Return result.ToString
    End Function
    Private Function Remove_spaces(s As String) As String
        Return s.Replace("      ", " ").Replace("     ", " ").Replace("    ", " ").Replace("   ", " ").Replace("  ", " ")
    End Function

    Private Sub BGWork_progress(sender As Object, p As ProgressChangedEventArgs) Handles bg_renamer.ProgressChanged, bg_scaner.ProgressChanged, bg_scaner_crc.ProgressChanged
        'ProgressBar1.Value += 1
        'If ProgressBar1.Value Mod 10 = 0 Then ProgressBar1.Refresh()
        ProgressBar1.Value = p.ProgressPercentage
        ProgressBar1.Refresh()
    End Sub

    'Cell chenge "X" - dblclick and keypress
    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If DataGridView1.Item(3, e.RowIndex).Value.ToString = "X" Then DataGridView1.Item(3, e.RowIndex).Value = "" Else DataGridView1.Item(3, e.RowIndex).Value = "X"
    End Sub
    Private Sub DataGridView1_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles DataGridView1.KeyDown
        If e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Space Then
            Dim r As Integer = DataGridView1.SelectedCells(0).RowIndex
            If DataGridView1.Item(3, r).Value.ToString = "X" Then DataGridView1.Item(3, r).Value = "" Else DataGridView1.Item(3, r).Value = "X"
        End If
    End Sub

    'Form show
    Private Sub Form6_autorenamer_Shown(sender As Object, e As System.EventArgs) Handles Me.Shown
        Dim enable As Boolean = False
        Form1.Button1.Enabled = enable
        Form1.Button2_moveUnneeded.Enabled = enable
        Form1.Button5_Associate.Enabled = enable
        Form1.GroupBox1.Enabled = enable
        Form1.GroupBox3.Enabled = enable

        Button2.Enabled = True
        DataGridView1.Rows.Clear()
    End Sub
    'Form Close
    Private Sub Form6_autorenamer_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim enable As Boolean = True
        Form1.Button1.Enabled = enable
        Form1.Button2_moveUnneeded.Enabled = enable
        Form1.Button5_Associate.Enabled = enable
        Form1.GroupBox1.Enabled = enable
        Form1.GroupBox3.Enabled = enable
    End Sub
End Class