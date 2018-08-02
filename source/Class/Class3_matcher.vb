Imports Microsoft.VisualBasic.FileIO
Imports System.Text.RegularExpressions

Public Class Class3_matcher
#Region "Declarations"
    Dim dt_games As New DataTable
    Dim dt_files As New DataTable
    Public crc_for_archs As String = ""
    Dim frm As Form1 = DirectCast(Application.OpenForms("Form1"), Form1)
    Dim countAll As Integer = -1, countFound As Integer = -1, countNotFound As Integer = -1
    Dim countMatchesFiles As Integer = -1, countNotMatchesFiles As Integer = -1, countAllFiles As Integer = -1
    Private WithEvents Button5_Associate As Button = frm.Button5_Associate
    Private WithEvents ButtonStrip1 As Button = frm.ButtonStrip1
    Private WithEvents Button7 As Button = frm.Button7
    Private WithEvents Button20_markAsFound As Button = frm.Button20_markAsFound
    Private WithEvents ComboBox3 As ComboBox = frm.ComboBox3
    Private WithEvents ComboBox7 As ComboBox = frm.ComboBox7
    Private WithEvents TextBox4 As TextBox = frm.TextBox4
    Private WithEvents TextStrip1 As TextBox = frm.TextStrip1
    Private WithEvents CheckBox3 As CheckBox = frm.CheckBox3
    Private WithEvents CheckBox10 As CheckBox = frm.CheckBox10
    Private WithEvents CheckBox11 As CheckBox = frm.CheckBox11
    Private WithEvents CheckBox12 As CheckBox = frm.CheckBox12
    Private WithEvents CheckBox13 As CheckBox = frm.CheckBox13
    Private WithEvents TextBox26 As TextBox = frm.TextBox26
    Private WithEvents TextBox27 As TextBox = frm.TextBox27

    Private WithEvents RadioStrip1 As RadioButton = frm.RadioStrip1
    Private WithEvents RadioStrip2 As RadioButton = frm.RadioStrip2
    Private WithEvents RadioButton1 As RadioButton = frm.RadioButton1
    Private WithEvents RadioButton2 As RadioButton = frm.RadioButton2
    Private WithEvents RadioButton3 As RadioButton = frm.RadioButton3
    Private WithEvents RadioButton4 As RadioButton = frm.RadioButton4
    Private WithEvents RadioButton5 As RadioButton = frm.RadioButton5
    Private WithEvents RadioButton6 As RadioButton = frm.RadioButton6

    Private WithEvents checkBox27 As CheckBox = frm.CheckBox27
    Private WithEvents listbox1 As ListBox = frm.ListBox1
    Private WithEvents listbox2 As ListBox = frm.ListBox2
    'Friend WithEvents myContextMenu7 As New ToolStripDropDownMenu 'autorenamer
    Public Shared autofilter_regex As String = "%[A-Za-z]{4}[A-Za-z]*"
    Public Shared autofilter_regex_options() As Boolean = {False, False}

    Dim listbox_searchAsYouTypeStr As String = ""
    Dim WithEvents listbox_searchAsYouTypeTimer As New Timer With {.Interval = 1000, .Enabled = False}
#End Region

    'Constructor
    Sub New()
        dt_games.Columns.Add("name")
        dt_files.Columns.Add("name")
        dt_files.Columns.Add("nameWithoutHyphen")
    End Sub

    'matcher - fill database
    Public Sub matcher_remplirDatabaseEntryList()
        countAll = 0
        countFound = 0
        countNotFound = 0
        With frm
            .ListBox1.DataSource = Nothing
            .ListBox1.BeginUpdate()
            dt_games.Clear()

            If .ComboBox3.SelectedIndex < 0 Then Exit Sub
            If .ComboBox1.SelectedIndex < 0 Then Exit Sub
            If .DataGridView1.Rows.Count = 0 Then
                dt_games.Rows.Add({"...Please, make ""check"" in summary page,"})
                dt_games.Rows.Add({"before using this future"})
                .ListBox1.DataSource = dt_games
                .ListBox1.DisplayMember = "name" : .ListBox1.ValueMember = "name"
                .ListBox1.EndUpdate() : Exit Sub
            End If
            .Label20.Text = "Refreshing Database Entry" : .Label20.BackColor = Color.Red : .Label20.Refresh()
            Dim cellindex As Integer
            For Each row As DataGridViewRow In .DataGridView1.Rows
                If .ComboBox3.SelectedIndex = 0 Then
                    cellindex = 2
                ElseIf .ComboBox3.SelectedIndex = 1 Then
                    cellindex = 3
                ElseIf .ComboBox3.SelectedIndex = 2 Then
                    cellindex = 5
                ElseIf .ComboBox3.SelectedIndex = 3 Then
                    cellindex = 6
                ElseIf .ComboBox3.SelectedIndex = 4 Then
                    cellindex = 7
                ElseIf .ComboBox3.SelectedIndex = 5 Then
                    cellindex = 8
                ElseIf .ComboBox3.SelectedIndex = 6 Then
                    cellindex = 9
                ElseIf .ComboBox3.SelectedIndex = 7 Then
                    cellindex = 4
                ElseIf .ComboBox3.SelectedIndex = 8 Then
                    cellindex = 10
                ElseIf .ComboBox3.SelectedIndex = 9 Then
                    cellindex = 11
                ElseIf .ComboBox3.SelectedIndex = 10 Then
                    cellindex = 12
                ElseIf .ComboBox3.SelectedIndex = 11 Then
                    cellindex = 13
                ElseIf .ComboBox3.SelectedIndex = 12 Then
                    cellindex = 14
                ElseIf .ComboBox3.SelectedIndex = 13 Then
                    cellindex = 15
                ElseIf .ComboBox3.SelectedIndex = 14 Then
                    cellindex = 16
                ElseIf .ComboBox3.SelectedIndex = 15 Then
                    cellindex = 17
                ElseIf .ComboBox3.SelectedIndex = 16 Then
                    cellindex = 18

                End If
                Dim retVal As Boolean = matcher_remplirDatabaseEntryList_addRow(row, cellindex)
                If retVal Then countFound += 1 Else countNotFound += 1
                countAll += 1
            Next

            .ListBox1.DataSource = dt_games
            .ListBox1.DisplayMember = "name"
            .ListBox1.ValueMember = "name"
            .ListBox1.EndUpdate()

            matcher_update_total_labels()
            'If .ListBox1.Items.Count = 0 And .RadioButton1.Checked Then .ListBox1.Items.Add("No matched " + .ComboBox3.SelectedItem.ToString)
            'If .ListBox1.Items.Count = 0 And .RadioButton2.Checked Then .ListBox1.Items.Add("No missing " + .ComboBox3.SelectedItem.ToString)
            'If .ListBox1.Items.Count = 0 And .RadioButton3.Checked Then .ListBox1.Items.Add("No " + .ComboBox3.SelectedItem.ToString + " found in current system database")
            .Label20.Text = "Ready" : .Label20.BackColor = Color.LightGreen : .Label20.Refresh()
        End With
    End Sub

    'matcher - fill database SUB - add db entry to listbox if it's needed in current show mode
    Private Function matcher_remplirDatabaseEntryList_addRow(ByVal row As DataGridViewRow, ByVal cellIndex As Integer) As Boolean
        Dim retVal As Boolean = False
        With frm
            If row.Cells(cellIndex).Value.ToString = "YES" Then retVal = True
            If .RadioButton1.Checked Then If row.Cells(cellIndex).Value.ToString = "YES" Then dt_games.Rows.Add({row.Cells(1).Value})
            If .RadioButton2.Checked Then If row.Cells(cellIndex).Value.ToString = "NO" Then dt_games.Rows.Add({row.Cells(1).Value})
            If .RadioButton3.Checked Then dt_games.Rows.Add({row.Cells(1).Value})
        End With
        Return retVal
    End Function

    'Update TOTAL labels
    Private Sub matcher_update_total_labels()
        With frm
            If countFound = -1 And countNotFound = -1 And countAll = -1 Then
                .Label7.Text = "Total: "
                .Label9.Text = "Total: "
                .Label42.Visible = False : .Label43.Visible = False : .Label44.Visible = False
                .Label45.Visible = False : .Label46.Visible = False : .Label47.Visible = False
            Else
                If Not .CheckBox27.Checked Then
                    .Label7.Text = "Total: " + .ListBox1.Items.Count.ToString
                    .Label9.Text = "Total: " + .ListBox2.Items.Count.ToString
                    .Label42.Visible = False : .Label43.Visible = False : .Label44.Visible = False
                    .Label45.Visible = False : .Label46.Visible = False : .Label47.Visible = False
                    Exit Sub
                End If

                .Label42.Font = New Font(.Label42.Font, FontStyle.Regular) : .Label42.ForeColor = SystemColors.ControlText
                .Label43.Font = New Font(.Label43.Font, FontStyle.Regular) : .Label43.ForeColor = SystemColors.ControlText
                .Label44.Font = New Font(.Label44.Font, FontStyle.Regular) : .Label44.ForeColor = SystemColors.ControlText
                If RadioButton1.Checked Then .Label42.Font = New Font(.Label42.Font, FontStyle.Bold) : .Label42.ForeColor = Color.DarkRed
                If RadioButton2.Checked Then .Label43.Font = New Font(.Label43.Font, FontStyle.Bold) : .Label43.ForeColor = Color.DarkRed
                If RadioButton3.Checked Then .Label44.Font = New Font(.Label44.Font, FontStyle.Bold) : .Label44.ForeColor = Color.DarkRed
                .Label7.Text = "Total - "
                .Label42.Text = "With Media: " + countFound.ToString + ", " : If countFound = 0 Then .Label42.Text = "With Media: NO, "
                .Label43.Text = "Without Media: " + countNotFound.ToString + ", " : If countNotFound = 0 Then .Label43.Text = "Without Media: NO, "
                .Label44.Text = "Both: " + countAll.ToString : If countAll = 0 Then .Label44.Text = "Both: NO"
                .Label42.Left = .Label7.Right : .Label43.Left = .Label42.Right : .Label44.Left = .Label43.Right
                .Label42.Visible = True : .Label43.Visible = True : .Label44.Visible = True

                .Label45.Font = New Font(.Label45.Font, FontStyle.Regular) : .Label45.ForeColor = SystemColors.ControlText
                .Label46.Font = New Font(.Label46.Font, FontStyle.Regular) : .Label46.ForeColor = SystemColors.ControlText
                .Label47.Font = New Font(.Label47.Font, FontStyle.Regular) : .Label47.ForeColor = SystemColors.ControlText
                If RadioButton4.Checked Then .Label45.Font = New Font(.Label45.Font, FontStyle.Bold) : .Label45.ForeColor = Color.DarkRed
                If RadioButton5.Checked Then .Label46.Font = New Font(.Label46.Font, FontStyle.Bold) : .Label46.ForeColor = Color.DarkRed
                If RadioButton6.Checked Then .Label47.Font = New Font(.Label47.Font, FontStyle.Bold) : .Label47.ForeColor = Color.DarkRed
                .Label9.Text = "Total - "
                .Label45.Text = "Matched: " + countMatchesFiles.ToString + ", " : If countMatchesFiles = 0 Then .Label45.Text = "Matched: NO, "
                .Label46.Text = "UnMatched: " + countNotMatchesFiles.ToString + ", " : If countNotMatchesFiles = 0 Then .Label46.Text = "UnMatched: NO, "
                .Label47.Text = "Both: " + countAllFiles.ToString : If countAllFiles = 0 Then .Label47.Text = "Both: NO"
                .Label45.Left = .Label9.Right : .Label46.Left = .Label45.Right : .Label47.Left = .Label46.Right
                .Label45.Visible = True : .Label46.Visible = True : .Label47.Visible = True
            End If
        End With
    End Sub

    'Associate CLICK
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5_Associate.Click
        With frm
            If .ListBox1.SelectedIndex < 0 Then MsgBox("Select a database entry to associate a file to.") : Exit Sub
            If .ListBox2.SelectedIndex < 0 Then MsgBox("Select a file to associate to selected database entry.") : Exit Sub
            crc_for_archs = ""
            Class7_archives.rename = Class7_archives.answer.useDefaultForm1Settings
            Class7_archives.keep_only_one = Class7_archives.answer.useDefaultForm1Settings
            Dim ext As String = ""
            Dim l1_selected_game As String = DirectCast(.ListBox1.SelectedItem, DataRowView).Item(0).ToString
            Dim l2_selected_file As String = DirectCast(.ListBox2.SelectedItem, DataRowView).Item(0).ToString

            Dim src_dir As String = FileSystem.GetDirectoryInfo(.TextBox4.Text).FullName.ToLower
            Dim dst_dir As String = FileSystem.GetDirectoryInfo(get_HS_Path_of_selected_media()).FullName.ToLower
            If Not src_dir.EndsWith("\") Then src_dir = src_dir + "\" : If Not dst_dir.EndsWith("\") Then dst_dir = dst_dir + "\"

            If Not .CheckBox3.Checked Then
                'If not subfoldered mode
                If l2_selected_file.Contains(".") Then ext = l2_selected_file.Substring(l2_selected_file.LastIndexOf("."))
            End If

            '''''''''''Actual Copy/Rename
            .Label20.Text = "Renaming..." : .Label20.BackColor = Color.Red : .Label20.Refresh()
            Dim res() As String = associate_copyMove(src_dir, dst_dir, l1_selected_game, l2_selected_file, .CheckBox3.Checked, .CheckBox10.Checked)
            If res(0) <> "" Then
                MsgBox(res(0)) : .Label20.Text = "Ready" : .Label20.BackColor = Color.LightGreen : .Label20.Refresh() : Exit Sub
            End If
            '''''''''''END Actual Copy/Rename

            'If we have renamed a good named file, we have to remove it from found
            'TODO TEST LOGIC WITH "subFOLDERED MODE"
            Dim l2_selected_file_woExt As String = l2_selected_file.ToLower
            If Not .CheckBox3.Checked Then l2_selected_file_woExt = l2_selected_file.Substring(0, l2_selected_file.LastIndexOf("."))
            If src_dir = dst_dir And Class1.romlist.Contains(l2_selected_file_woExt) Then
                Dim exist As Boolean = False
                For Each s As String In getCurMediaExtensionWildcards()
                    s = s.Replace("*", l2_selected_file_woExt)
                    If FileSystem.FileExists(src_dir + "\" + s) Then exist = True : Exit For
                Next
                If Not exist Then Button20_markAsFound_sub(l2_selected_file_woExt, Button20_markAsFound_comboToCol(), True)
            End If
            'add to found if needed
            Dim copyToHsFolder As Boolean = .AssocOption_fileInDiffFolder_copyToHS.Checked Or .AssocOption_fileInDiffFolder_moveToHS.Checked
            If src_dir = dst_dir Or copyToHsFolder Then Button20_markAsFound_sub(l1_selected_game, Button20_markAsFound_comboToCol())

            'Moving selected index on box2 (filelist)
            Dim currentTopIndex As Integer = .ListBox2.TopIndex
            Dim currentSelectedIndex As Integer = .ListBox2.SelectedIndex
            .ListBox2.BeginUpdate()
            If (src_dir = dst_dir Or Not copyToHsFolder) Then
                'V svoey direktorii ILI (NE v svoey i NE kopiruem) - i.e. src filename is changed
                If .RadioButton5.Checked Then
                    'show unmatched files (need to remove file from list)
                    dt_files.Rows.Remove(DirectCast(.ListBox2.SelectedItem, DataRowView).Row)
                Else
                    'show matched or both files (need to change filename and select next file)
                    DirectCast(.ListBox2.SelectedItem, DataRowView).Item(0) = l1_selected_game + ext
                    currentTopIndex += 1
                    currentSelectedIndex += 1
                End If
            Else
                'V drugoy directorii ili v svoey no copiruem - i.e. src filename is not changed
                currentTopIndex += 1
                currentSelectedIndex = .ListBox2.SelectedIndex + 1
            End If
            Dim b As New BindingContext : b(dt_files).EndCurrentEdit()
            If currentTopIndex >= 0 Then .ListBox2.TopIndex = currentTopIndex
            If .ListBox2.Items.Count > currentSelectedIndex Then .ListBox2.SelectedIndex = currentSelectedIndex Else .ListBox2.SelectedIndex = currentSelectedIndex - 1
            .ListBox2.EndUpdate()

            'Moving selected index on box1 (DB entry list)
            If .RadioButton2.Checked And (src_dir = dst_dir Or copyToHsFolder) Then
                currentSelectedIndex = .ListBox1.SelectedIndex
                dt_games.Rows.Remove(DirectCast(.ListBox1.SelectedItem, DataRowView).Row)
            Else
                currentSelectedIndex = .ListBox1.SelectedIndex + 1
            End If
            If .ListBox1.Items.Count > currentSelectedIndex Then .ListBox1.SelectedIndex = currentSelectedIndex Else .ListBox1.SelectedIndex = currentSelectedIndex - 1

            countFound += 1 : countNotFound -= 1
            countMatchesFiles += 1 : countNotMatchesFiles -= 1 : matcher_update_total_labels()
            .Label20.Text = "Ready" : .Label20.BackColor = Color.LightGreen : .Label20.Refresh()
        End With
    End Sub

    'Associate SUB - creating file operation array
    Public Function associate_copyMove(src_dir As String, dst_dir As String, gameName As String, fName As String, Optional subfoldered_mode As Boolean = False, Optional create_sub_folder As Boolean = False) As String()
        'In RL mode we just copy (or rename) directory, and don't bother with what is inside
        Dim RL_Mode = subfoldered_mode And frm.ComboBox3.SelectedIndex >= 9

        Dim ext As String = ""
        Dim op As New List(Of String())
        Try
            src_dir = FileSystem.GetDirectoryInfo(src_dir).FullName.ToLower
            dst_dir = FileSystem.GetDirectoryInfo(dst_dir).FullName.ToLower
            If Not subfoldered_mode Then
                'SINGLEFILE MODE
                ext = fName.Substring(fName.LastIndexOf("."))
                If src_dir = dst_dir Then
                    'src is in HS folder
                    If create_sub_folder Then dst_dir = dst_dir + "\" + gameName

                    If frm.AssocOption_fileInHsFolder_copy.Checked Then
                        'Copy (duplicate)
                        op.Add({"FILECOPY", src_dir + "\" + fName, dst_dir + "\" + gameName + ext})
                    ElseIf frm.AssocOption_fileInHsFolder_move.Checked Then
                        'Move (rename)
                        op.Add({"FILERENAME", src_dir + "\" + fName, dst_dir + "\" + gameName + ext})
                    End If
                Else
                    'src is in different folder
                    If frm.AssocOption_fileInDiffFolder_copy.Checked Then
                        'Copy in place
                        op.Add({"FILECOPY", src_dir + "\" + fName, src_dir + "\" + gameName + ext})
                    ElseIf frm.AssocOption_fileInDiffFolder_move.Checked Then
                        'Move in place
                        op.Add({"FILERENAME", src_dir + "\" + fName, src_dir + "\" + gameName + ext})
                    ElseIf frm.AssocOption_fileInDiffFolder_copyToHS.Checked Then
                        'Copy to HS folder
                        op.Add({"FILECOPY", src_dir + "\" + fName, dst_dir + "\" + gameName + ext})
                    ElseIf frm.AssocOption_fileInDiffFolder_moveToHS.Checked Then
                        'Move to HS folder
                        op.Add({"FILERENAME", src_dir + "\" + fName, dst_dir + "\" + gameName + ext})
                    End If
                End If
            Else
                'SUBFOLDERED MODE
                Dim list As New List(Of String)
                Dim wildcards() As String = getCurMediaExtensionWildcards()
                list = FileSystem.GetFiles(src_dir + "\" + fName, FileIO.SearchOption.SearchTopLevelOnly, wildcards).ToList
                If list.Count = 0 Then Return {"This folder does not contain any files that match rom's extensions.", ""}
                If list.Count > 1 And Not RL_Mode Then Return {"This folder contains multiple files that match rom's extensions. This situation is not handled yet, sorry." + vbCrLf + "The program simply doesn't know what file need to be renamed. Try use 'Override Extension' in options.", ""}

                ext = list.Item(0).Substring(list.Item(0).LastIndexOf("."))
                Dim filename As String = list.Item(0).Substring(list.Item(0).LastIndexOf("\") + 1)

                If src_dir = dst_dir Then
                    'src is in HS folder
                    If frm.AssocOption_fileInHsFolder_copy.Checked Then
                        'Copy (duplicate)
                        op.Add({"DIRCOPY", src_dir + "\" + fName, dst_dir + "\" + gameName + "\"})

                        If Not RL_Mode Then
                            Dim res0 As String = associate_fileOP(op)
                            If res0 <> "" Then Return {res0, ext}
                            op = New List(Of String())
                            op.Add({"FILERENAME", src_dir + "\" + gameName + "\" + filename, dst_dir + "\" + gameName + "\" + gameName + ext, "0"})
                        End If
                    ElseIf frm.AssocOption_fileInHsFolder_move.Checked Then
                        'Move (rename)
                        If Not RL_Mode Then op.Add({"FILERENAME", list.Item(0), dst_dir + "\" + fName + "\" + gameName + ext, "0"})
                        op.Add({"DIRRENAME", src_dir + "\" + fName, gameName, "0"})
                    End If
                Else
                    'src is in different folder
                    If frm.AssocOption_fileInDiffFolder_copy.Checked Then
                        'Copy in place
                        op.Add({"DIRCOPY", src_dir + "\" + fName, src_dir + "\" + gameName + "\"})

                        If Not RL_Mode Then
                            Dim res0 As String = associate_fileOP(op)
                            If res0 <> "" Then Return {res0, ext}
                            op = New List(Of String())
                            op.Add({"FILERENAME", src_dir + "\" + gameName + "\" + filename, src_dir + "\" + gameName + "\" + gameName + ext, "0"})
                        End If
                    ElseIf frm.AssocOption_fileInDiffFolder_move.Checked Then
                        'Move in place
                        If Not RL_Mode Then op.Add({"FILERENAME", list.Item(0), src_dir + "\" + fName + "\" + gameName + ext, "0"})
                        op.Add({"DIRRENAME", src_dir + "\" + fName, gameName, "0"})
                    ElseIf frm.AssocOption_fileInDiffFolder_copyToHS.Checked Then
                        'Copy to HS folder
                        op.Add({"DIRCOPY", src_dir + "\" + fName, dst_dir + "\" + gameName + "\"})

                        If Not RL_Mode Then
                            Dim res0 As String = associate_fileOP(op)
                            If res0 <> "" Then Return {res0, ext}
                            op = New List(Of String())
                            op.Add({"FILERENAME", dst_dir + "\" + gameName + "\" + filename, dst_dir + "\" + gameName + "\" + gameName + ext, "0"})
                        End If
                    ElseIf frm.AssocOption_fileInDiffFolder_moveToHS.Checked Then
                        'Move to HS folder
                        'TODO
                        MsgBox("Moving directory is not yet implimented.")
                    End If
                End If
            End If

            Dim res As String = associate_fileOP(op)
            If res <> "" Then Return {res, ext}
        Catch ex As Exception
            Return {ex.Message, ext}
        End Try
        Return {"", ext}
    End Function

    'Associate SUB - checks .cue and actual performing files operations
    Private Function associate_fileOP(ByVal op As List(Of String())) As String
        Dim msg As String = ""
        Dim restoreCue As Boolean = False
        Dim tmp As New List(Of String())
        For i As Integer = 0 To op.Count - 1
            Dim o() As String = op(i)
            If o(0) = "FILECOPY" Or o(0) = "FILERENAME" Then
                If Not FileSystem.FileExists(o(1)) Then msg = "Source file: """ + o(1) + """ does not exist." : Exit For
                '2nd part of existence check
                If FileSystem.FileExists(o(2)) Then
                    If o.Length < 4 Then msg = "Destination file: """ + o(2) + """ already exists." : Exit For Else op(i)(0) = "skip"
                End If

                Dim tmpExt = o(1).Substring(o(1).LastIndexOf(".") + 1)
                Dim tmppath = o(1).Substring(0, o(1).LastIndexOf("\") + 1)
                Dim tmpfileName = o(1).Substring(o(1).LastIndexOf("\") + 1)
                Dim tmpfileNameWOext = tmpfileName.Substring(0, tmpfileName.LastIndexOf("."))
                Dim newpath = o(2).Substring(0, o(2).LastIndexOf("\") + 1)
                Dim newfileName = o(2).Substring(o(2).LastIndexOf("\") + 1)
                Dim newfileNameWOext = newfileName.Substring(0, newfileName.LastIndexOf("."))

                'check .cue
                Dim listFilesInCue(1) As List(Of String)
                Dim listaudio As Boolean = DirectCast(IIf(o(0) = "FILECOPY", True, False), Boolean)
                If tmpExt.ToLower = "cue" And Not frm.CheckBox6.Checked Then
                    listFilesInCue = associate_listFilesFromCue(o(1), listaudio)
                    If listFilesInCue(0).Count = 0 Then msg = "Image file reference in """ + tmpfileName + """ not found. Check your .cue in notepad." : Exit For
                    If listFilesInCue(0)(0) = "" Then msg = "Image file reference in """ + tmpfileName + """ is empty. Check your .cue in notepad." : Exit For

                    'Rename image itself
                    Dim foundFileExt = listFilesInCue(0)(0).Substring(listFilesInCue(0)(0).LastIndexOf(".") + 1)
                    tmp.Add({o(0), tmppath + listFilesInCue(0)(0), newpath + newfileNameWOext + "." + foundFileExt})

                    'Rename imageName in .CUE
                    associate_rewriteCue(o(1), newfileNameWOext + "." + foundFileExt)

                    'if FILECOPY this will copy audio files
                    For Each f As String In listFilesInCue(1)
                        tmp.Add({o(0), tmppath + f, newpath + f})
                    Next

                    'just to put 'restore cue from backup' action in UNDO list
                    'tmp.Add({"CUERENAMED", o(1)})
                    restoreCue = True
                Else
                    'if file being renamed is iso or bin or one of the list in options, we have to search for .cue
                    For Each e As String In frm.TextBox16.Text.ToLower.Split(","c)
                        If e.Trim = tmpExt.ToLower Then
                            For Each f As String In FileSystem.GetFiles(tmppath, SearchOption.SearchTopLevelOnly, {"*.cue"})
                                'listFilesInCue = associate_listFilesFromCue(tmppath + tmpfileNameWOext + ".cue", listaudio)
                                listFilesInCue = associate_listFilesFromCue(f, listaudio)
                                If listFilesInCue(0)(0).ToLower = tmpfileName.ToLower Then
                                    'tmp.Add({o(0), f, newpath + f.Substring(f.LastIndexOf("\") + 1)})
                                    tmp.Add({o(0), f, newpath + newfileNameWOext + ".cue"})
                                    associate_rewriteCue(f, newfileNameWOext + "." + tmpExt)
                                    For Each fA As String In listFilesInCue(1)
                                        tmp.Add({o(0), tmppath + fA, newpath + fA})
                                    Next
                                End If
                            Next
                        End If
                    Next
                End If

                'check pairs (mdf/mds list)
                For Each l As String In frm.ListBox4.Items
                    If l.ToLower.Contains(tmpExt.ToLower) Then
                        For Each ext As String In l.Split(","c)
                            If ext.Trim.ToLower = tmpExt.ToLower Then Continue For
                            If FileSystem.FileExists(tmppath + tmpfileNameWOext + "." + ext) Then
                                Dim newfile = o(2).Substring(0, o(2).LastIndexOf(".") + 1) + ext
                                tmp.Add({o(0), tmppath + tmpfileNameWOext + "." + ext, newfile})
                            End If
                        Next
                    End If
                Next

                'MSU-1 - pcm files
                If tmpExt.ToLower = "sfc" Or tmpExt.ToLower = "smc" Then
                    Dim pcm_files = FileSystem.GetFiles(tmppath, SearchOption.SearchTopLevelOnly, {tmpfileNameWOext + "-*.pcm"})
                    For Each pcm In pcm_files
                        Dim newfile = o(2).Substring(0, o(2).LastIndexOf(".")) + pcm.Substring(pcm.LastIndexOf("-"))
                        tmp.Add({o(0), pcm, newfile})
                    Next
                End If
            End If
            If o(0) = "DIRCOPY" Then
                If Not FileSystem.DirectoryExists(o(1)) Then msg = "Source directory: """ + o(1) + """ does not exist." : Exit For
                If FileSystem.DirectoryExists(o(2)) Then
                    If o.Length < 4 Then msg = "Destination directory: """ + o(2) + """ already exists." : Exit For Else op(i)(0) = "skip"
                End If
            End If
            If o(0) = "DIRRENAME" Then
                If Not FileSystem.DirectoryExists(o(1)) Then msg = "Source directory: """ + o(1) + """ does not exist." : Exit For
                Dim path As String = o(1).Substring(0, o(1).LastIndexOf("\") + 1)
                If FileSystem.DirectoryExists(path + o(2)) Then
                    If o.Length < 4 Then msg = "Destination directory: """ + path + o(2) + """ already exists." : Exit For Else op(i)(0) = "skip"
                End If
            End If
            If o(0) = "DIRMOVE" Then
                If Not FileSystem.DirectoryExists(o(1)) Then msg = "Source directory: """ + o(1) + """ does not exist." : Exit For
                If FileSystem.DirectoryExists(o(2)) Then
                    If o.Length < 4 Then msg = "Destination directory: """ + o(2) + """ already exists." : Exit For Else op(i)(0) = "skip"
                End If
            End If
        Next
        If msg <> "" Then Return msg + vbCrLf + "File operation aborted. Nothing changed."

        'Actual performing file operations and undo array fill
        Dim z As New Class7_archives
        frm.undo.Add(New List(Of String))
        frm.undo_humanReadable.Add(New List(Of String))
        Dim undoIndex = frm.undo.Count - 1
        If tmp.Count > 0 Then op.InsertRange(1, tmp)
        Dim archive_handled As Boolean = False
        Try
            For Each o In op
                If o(0) = "FILECOPY" Then
                    If z.isArchive(o(1)) Then
                        z.setFile(o(1))
                        Dim archNameWoExt = o(2)
                        If archNameWoExt.Contains(".") Then archNameWoExt = archNameWoExt.Substring(0, archNameWoExt.LastIndexOf("."))
                        Dim gameName = FileSystem.GetFileInfo(o(2)).Name
                        If gameName.Contains(".") Then gameName = gameName.Substring(0, gameName.LastIndexOf("."))
                        archive_handled = z.renameInArchiveIfNeeded(archNameWoExt, crc_for_archs)
                        If archive_handled Then
                            frm.undo(undoIndex).Add("ARCHIVE?" + o(1) + "?" + o(2))
                            frm.undo_humanReadable(undoIndex).Add("Recompressed Archive " + o(1) + " to " + o(2))
                            Class1.Log("Recompressed Archive " + o(1) + " to " + o(2))
                        End If
                    End If

                    If Not archive_handled Then
                        FileSystem.CopyFile(o(1), o(2))
                        If o(1).Substring(o(1).LastIndexOf(".") + 1).ToLower = "cue" And restoreCue Then
                            frm.undo(undoIndex).Add("RESTORECUE?" + o(1))
                            frm.undo(undoIndex).Add("FILEREMOVE?" + o(2))
                            frm.undo_humanReadable(undoIndex).Add("Copy File")
                            frm.undo_humanReadable(undoIndex).Add("- rewrote .cue" + o(1))
                            frm.undo_humanReadable(undoIndex).Add("- copy " + o(1) + " to " + o(2))
                            Class1.Log("Copy File - Rewrote .CUE " + o(1) + ", copy " + o(1) + " to " + o(2))
                        Else
                            frm.undo(undoIndex).Add("FILEREMOVE?" + o(2))
                            frm.undo_humanReadable(undoIndex).Add("Copy File " + o(1) + " to " + o(2))
                            Class1.Log("Copy File " + o(1) + " to " + o(2))
                        End If
                    End If
                End If
                If o(0) = "FILERENAME" Then
                    If z.isArchive(o(1)) Then
                        z.setFile(o(1))
                        Dim archNameWoExt = o(2)
                        If archNameWoExt.Contains(".") Then archNameWoExt = archNameWoExt.Substring(0, archNameWoExt.LastIndexOf("."))
                        Dim gameName = FileSystem.GetFileInfo(o(2)).Name
                        If gameName.Contains(".") Then gameName = gameName.Substring(0, gameName.LastIndexOf("."))
                        archive_handled = z.renameInArchiveIfNeeded(archNameWoExt, crc_for_archs)
                        If archive_handled Then
                            FileSystem.DeleteFile(o(1))
                            frm.undo(undoIndex).Add("ARCHIVE?" + o(1) + "?" + o(2))
                            frm.undo_humanReadable(undoIndex).Add("Recompressed Archive " + o(1) + " to " + o(2))
                            frm.undo(undoIndex).Add("DELETE?" + o(1))
                            frm.undo_humanReadable(undoIndex).Add("Delete " + o(1))
                            Class1.Log("Recompressed Archive " + o(1) + " to " + o(2))
                            Class1.Log("Delete " + o(1))
                        End If
                    End If

                    If Not archive_handled Then
                        FileSystem.MoveFile(o(1), o(2))
                        If o(1).Substring(o(1).LastIndexOf(".") + 1).ToLower = "cue" And restoreCue Then
                            frm.undo(undoIndex).Add("RESTORECUE?" + o(1))
                            frm.undo(undoIndex).Add("FILEREMOVE?" + o(2))
                            frm.undo_humanReadable(undoIndex).Add("Move File")
                            frm.undo_humanReadable(undoIndex).Add("- rewrote .cue" + o(1))
                            frm.undo_humanReadable(undoIndex).Add("- move " + o(1) + " to " + o(2))
                            Class1.Log("Move File - Rewrote .CUE " + o(1) + ", move " + o(1) + " to " + o(2))
                        Else
                            frm.undo(undoIndex).Add("FILERENAME?" + o(2) + "?" + o(1))
                            frm.undo_humanReadable(undoIndex).Add("Move File " + o(1) + " to " + o(2))
                            Class1.Log("Move File " + o(1) + " to " + o(2))
                        End If
                    End If
                End If
                If o(0) = "DIRCOPY" Then
                    FileSystem.CopyDirectory(o(1), o(2))
                    frm.undo(undoIndex).Add("DIRREMOVE?" + o(2))
                    frm.undo_humanReadable(undoIndex).Add("Copy Directory " + o(1) + " to " + o(2))
                    Class1.Log("Copy Directory " + o(1) + " to " + o(2))
                End If
                If o(0) = "DIRMOVE" Then
                    FileSystem.MoveDirectory(o(1), o(2))
                    frm.undo(undoIndex).Add("DIRMOVE?" + o(2) + "?" + o(1))
                    frm.undo_humanReadable(undoIndex).Add("Move Directory " + o(1) + " to " + o(2))
                    Class1.Log("Move Directory " + o(1) + " to " + o(2))
                End If
                If o(0) = "DIRRENAME" Then
                    FileSystem.RenameDirectory(o(1), o(2))
                    Dim name As String = o(1).Substring(o(1).LastIndexOf("\") + 1)
                    Dim path As String = o(1).Substring(0, o(1).LastIndexOf("\") + 1)
                    frm.undo(undoIndex).Add("DIRRENAME?" + path + o(2) + "?" + name)
                    frm.undo_humanReadable(undoIndex).Add("Rename Directory " + o(1) + " to " + o(2))
                    Class1.Log("Rename Directory " + o(1) + " to " + o(2))
                End If
            Next
        Catch ex As Exception
            If frm.undo(undoIndex).Count = 0 Then frm.undo.RemoveAt(undoIndex) : frm.undo_humanReadable.RemoveAt(undoIndex)
            Return ex.Message
        End Try
        Return ""
    End Function

    'Associate SUB - list files listed in .cue
    Private Function associate_listFilesFromCue(ByVal filename As String, Optional ByVal listAudio As Boolean = True) As List(Of String)()
        Dim line As String
        Dim inFilename As String
        Dim list(1) As List(Of String) : list(0) = New List(Of String) : list(1) = New List(Of String)
        Dim binaryfound As Boolean = False
        FileOpen(1, filename, OpenMode.Input)
        Do While Not EOF(1)
            line = LineInput(1)
            If line.IndexOf("file", System.StringComparison.InvariantCultureIgnoreCase) >= 0 Then
                inFilename = line.Substring(line.IndexOf("""") + 1)
                inFilename = inFilename.Substring(0, inFilename.LastIndexOf(""""))
                If inFilename.Contains("\") Then inFilename = inFilename.Substring(inFilename.LastIndexOf("\") + 1)

                If line.IndexOf("binary", System.StringComparison.InvariantCultureIgnoreCase) >= 0 Then
                    binaryfound = True
                    list(0).Add(inFilename)
                    If Not listAudio Then Exit Do
                Else
                    list(1).Add(inFilename)
                End If
            End If
        Loop
        FileClose(1)
        Return list
    End Function

    'Associate SUB - rewrite .cue
    Private Sub associate_rewriteCue(ByVal filename As String, ByVal replaceImageBy As String)
        Dim line As String
        Dim firstBinaryHandled As Boolean = False
        Dim list As New List(Of String)
        FileOpen(1, filename, OpenMode.Input)
        Do While Not EOF(1)
            line = LineInput(1)
            If line.IndexOf("file", System.StringComparison.InvariantCultureIgnoreCase) >= 0 Then
                If line.IndexOf("binary", System.StringComparison.InvariantCultureIgnoreCase) >= 0 And Not firstBinaryHandled Then
                    Dim tmp As String = line.Substring(line.IndexOf("""") + 1)
                    tmp = tmp.Substring(0, tmp.LastIndexOf(""""))

                    'if filename contains path, and this path exist, than we have nothing to rename
                    If line.Contains("\") Then
                        If FileSystem.FileExists(line) Then Exit Sub
                    End If
                    line = line.Replace(tmp, replaceImageBy)
                    firstBinaryHandled = True
                End If
            End If
            list.Add(line)
        Loop
        FileClose(1)

        'do backup
        Dim backupN As Integer = 0
        If FileSystem.FileExists(filename + ".backup") Then
            Do While Microsoft.VisualBasic.FileIO.FileSystem.FileExists(frm.xmlPath + ".backup" + backupN.ToString)
                backupN += 1
            Loop
            FileSystem.CopyFile(filename, filename + ".backup" + backupN.ToString)
        Else
            FileSystem.CopyFile(filename, filename + ".backup")
        End If

        'write .cue file
        FileOpen(1, filename, OpenMode.Output)
        For Each s As String In list
            PrintLine(1, s)
        Next
        FileClose(1)
    End Sub

    'UNDO click
    Private Sub associate_undo() Handles ButtonStrip1.Click
        If frm.undo.Count = 0 Then MsgBox("Nothing to undo.") : Exit Sub
        Dim undoIndex As Integer = frm.undo.Count - 1
        Dim msg As String = "The following operations will be executed:" + vbCrLf
        For i As Integer = frm.undo(undoIndex).Count - 1 To 0 Step -1
            Dim m() As String = frm.undo(undoIndex)(i).Split("?"c)
            m(1) = m(1).Replace("\\", "\")
            If m(0) = "FILEREMOVE" Then msg = msg + "Remove file: " + m(1) + vbCrLf
            If m(0) = "FILERENAME" Then msg = msg + "Rename file: " + m(1) + " to " + m(2).Replace("\\", "\") + vbCrLf
            If m(0) = "DIRREMOVE" Then msg = msg + "Remove directory: " + m(1) + vbCrLf
            If m(0) = "DIRRENAME" Then msg = msg + "Rename directory: " + m(1) + " to " + m(2).Replace("\\", "\") + vbCrLf
            If m(0) = "DIRMOVE" Then msg = msg + "Move directory: " + m(1) + " to " + m(2).Replace("\\", "\") + vbCrLf
            If m(0) = "RESTORECUE" Then msg = msg + "Restore .cue file: " + m(1) + " from backup" + vbCrLf
        Next
        msg = msg + "Is it OK?"
        If MsgBox(msg, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
            If MsgBox("Remove this operation from undo list?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                frm.undo.RemoveAt(undoIndex)
                frm.undo_humanReadable.RemoveAt(undoIndex)
            End If
            Exit Sub
        End If

        For i As Integer = frm.undo(undoIndex).Count - 1 To 0 Step -1
            Try
                Dim o() As String = frm.undo(undoIndex)(i).Split("?"c)
                If o(0) = "FILEREMOVE" Then FileSystem.DeleteFile(o(1))
                If o(0) = "FILERENAME" Then FileSystem.MoveFile(o(1), o(2))
                If o(0) = "DIRREMOVE" Then FileSystem.DeleteDirectory(o(1), DeleteDirectoryOption.DeleteAllContents)
                If o(0) = "DIRRENAME" Then FileSystem.RenameDirectory(o(1), o(2))
                If o(0) = "DIRMOVE" Then FileSystem.MoveDirectory(o(1), o(2))
                If o(0) = "RESTORECUE" Then
                    Dim backupN As Integer = 0
                    If FileSystem.FileExists(o(1) + ".backup") Then
                        Do While Microsoft.VisualBasic.FileIO.FileSystem.FileExists(frm.xmlPath + ".backup" + backupN.ToString)
                            backupN += 1
                        Loop
                        If backupN = 0 Then FileSystem.MoveFile(o(1) + ".backup", o(1), True)
                        If backupN > 0 Then FileSystem.MoveFile(o(1) + ".backup" + (backupN - 1).ToString, o(1), True)
                    End If
                End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation)
            End Try
        Next
        frm.undo.RemoveAt(undoIndex)
        frm.undo_humanReadable.RemoveAt(undoIndex)
        MsgBox("You have to reCheck to reflect changes")
    End Sub

    'get path to media based on matcher mediaselect Combobox
    Private Function get_HS_Path_of_selected_media() As String
        With frm
            If .ComboBox3.SelectedIndex = 0 Then
                'TODO Better handle multiple rompaths
                If ComboBox7.Visible Then
                    Return ComboBox7.SelectedItem.ToString
                Else
                    Return Class1.romPath.Split({"|"c}, StringSplitOptions.RemoveEmptyEntries)(0)
                End If
            ElseIf .ComboBox3.SelectedIndex = 1 Then
                Return Class1.videoPath
            ElseIf .ComboBox3.SelectedIndex = 2 Then
                Return Class1.HyperspinPath + "Media\" + .ComboBox1.SelectedItem.ToString + "\Images\Wheel\"
            ElseIf .ComboBox3.SelectedIndex = 3 Then
                Return Class1.HyperspinPath + "Media\" + .ComboBox1.SelectedItem.ToString + "\Images\Artwork1\"
            ElseIf .ComboBox3.SelectedIndex = 4 Then
                Return Class1.HyperspinPath + "Media\" + .ComboBox1.SelectedItem.ToString + "\Images\Artwork2\"
            ElseIf .ComboBox3.SelectedIndex = 5 Then
                Return Class1.HyperspinPath + "Media\" + .ComboBox1.SelectedItem.ToString + "\Images\Artwork3\"
            ElseIf .ComboBox3.SelectedIndex = 6 Then
                Return Class1.HyperspinPath + "Media\" + .ComboBox1.SelectedItem.ToString + "\Images\Artwork4\"
            ElseIf .ComboBox3.SelectedIndex = 7 Then
                Return Class1.HyperspinPath + "Media\" + .ComboBox1.SelectedItem.ToString + "\Themes\"
            ElseIf .ComboBox3.SelectedIndex = 8 Then
                Return Class1.HyperspinPath + "Media\" + .ComboBox1.SelectedItem.ToString + "\Sound\Background Music\"
            ElseIf .ComboBox3.SelectedIndex = 9 Then
                Return Class1.HyperlaunchPath + "\Media\Artwork\" + .ComboBox1.SelectedItem.ToString + "\"
            ElseIf .ComboBox3.SelectedIndex = 10 Then
                Return Class1.HyperlaunchPath + "\Media\Backgrounds\" + .ComboBox1.SelectedItem.ToString + "\"
            ElseIf .ComboBox3.SelectedIndex = 11 Then
                Return Class1.HyperlaunchPath + "\Media\Bezels\" + .ComboBox1.SelectedItem.ToString + "\"
            ElseIf .ComboBox3.SelectedIndex = 12 Then
                Return Class1.HyperlaunchPath + "\Media\Fade\" + .ComboBox1.SelectedItem.ToString + "\"
            ElseIf .ComboBox3.SelectedIndex = 13 Then
                Return Class1.HyperlaunchPath + "\Media\Guides\" + .ComboBox1.SelectedItem.ToString + "\"
            ElseIf .ComboBox3.SelectedIndex = 14 Then
                Return Class1.HyperlaunchPath + "\Media\Manuals\" + .ComboBox1.SelectedItem.ToString + "\"
            ElseIf .ComboBox3.SelectedIndex = 15 Then
                Return Class1.HyperlaunchPath + "\Media\Music\" + .ComboBox1.SelectedItem.ToString + "\"
            ElseIf .ComboBox3.SelectedIndex = 16 Then
                Return Class1.HyperlaunchPath + "\Media\Videos\" + .ComboBox1.SelectedItem.ToString + "\"
            End If
            Return ""
        End With
    End Function

    'matcher options
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        frm.myContextMenu2.Show(Cursor.Position.X, Cursor.Position.Y)
    End Sub

    'markAsFound
    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20_markAsFound.Click
        With frm
            Dim count As Integer = 0
            Dim fWoExt As String = ""
            If .ComboBox3.SelectedIndex < 0 Then MsgBox("Please, select media type.") : Exit Sub
            If Microsoft.VisualBasic.FileIO.FileSystem.DirectoryExists(.TextBox4.Text) Then
                Dim col As Integer = Button20_markAsFound_comboToCol()
                Dim list As System.Collections.ObjectModel.ReadOnlyCollection(Of String)
                If .CheckBox3.Checked Then
                    list = FileSystem.GetDirectories(.TextBox4.Text)
                Else
                    list = FileSystem.GetFiles(.TextBox4.Text)
                End If

                For Each fi As String In list
                    Dim f As String = fi.Substring(fi.LastIndexOf("\") + 1)
                    If .CheckBox3.Checked Then
                        fWoExt = f
                        If FileSystem.GetFiles(fi, SearchOption.SearchTopLevelOnly, getCurMediaExtensionWildcards).Count = 0 Then Continue For
                    Else
                        fWoExt = f.Substring(0, f.LastIndexOf("."))
                    End If
                    If Button20_markAsFound_sub(fWoExt, col) Then count += 1
                Next
            End If
            MsgBox("Done. " + count.ToString + " entry affected.")
            matcher_remplirDatabaseEntryList()
        End With
    End Sub

    'matcher media combobox TO datagrid col
    Private Function Button20_markAsFound_comboToCol() As Integer
        If frm.ComboBox3.SelectedIndex = 0 Then Return 2
        If frm.ComboBox3.SelectedIndex = 1 Then Return 3
        If frm.ComboBox3.SelectedIndex = 2 Then Return 5
        If frm.ComboBox3.SelectedIndex = 3 Then Return 6
        If frm.ComboBox3.SelectedIndex = 4 Then Return 7
        If frm.ComboBox3.SelectedIndex = 5 Then Return 8
        If frm.ComboBox3.SelectedIndex = 6 Then Return 9
        If frm.ComboBox3.SelectedIndex = 7 Then Return 4
        If frm.ComboBox3.SelectedIndex = 8 Then Return 10
        If frm.ComboBox3.SelectedIndex = 9 Then Return 11
        If frm.ComboBox3.SelectedIndex = 10 Then Return 12
        If frm.ComboBox3.SelectedIndex = 11 Then Return 13
        If frm.ComboBox3.SelectedIndex = 12 Then Return 14
        If frm.ComboBox3.SelectedIndex = 13 Then Return 15
        If frm.ComboBox3.SelectedIndex = 14 Then Return 16
        If frm.ComboBox3.SelectedIndex = 15 Then Return 17
        If frm.ComboBox3.SelectedIndex = 16 Then Return 18
        Return 0
    End Function

    'markAsFoundSub
    Private Function Button20_markAsFound_sub(ByVal fWoExt As String, ByVal col As Integer, Optional ByVal remove As Boolean = False) As Boolean
        If Class1.romlist.Contains(fWoExt.ToLower) Then
            'mark in grid
            For r = 0 To frm.DataGridView1.Rows.Count - 1
                If frm.DataGridView1.Item(1, r).Value.ToString.ToLower = fWoExt.ToLower Then
                    If Not remove Then
                        frm.DataGridView1.Item(col, r).Value = "YES"
                        frm.DataGridView1.Item(col, r).Style.BackColor = Class1.colorYES
                    Else
                        frm.DataGridView1.Item(col, r).Value = "NO"
                        frm.DataGridView1.Item(col, r).Style.BackColor = Class1.colorNO
                    End If
                    Exit For
                End If
            Next

            'mark in class1.data
            Dim v As String
            If Not remove Then v = "Y" Else v = "N"
            Dim index As Integer = Class1.romlist.IndexOf(fWoExt.ToLower)
            Mid(Class1.data(index), col - 1, 1) = v

            'mark in romfound list
            'check 'col' to know, that we are dealing with ROM and not other media
            If col = 2 Then
                If Not remove Then
                    If Not Class1.romFoundlist.Contains(fWoExt.ToLower) Then
                        Class1.romFoundlist.Add(fWoExt.ToLower) : Return True
                    End If
                Else
                    If Class1.romFoundlist.Contains(fWoExt.ToLower) Then
                        Class1.romFoundlist.Remove(fWoExt.ToLower) : Return True
                    End If
                End If
            End If
        End If
        Return False
    End Function

    Private Function getCurMediaExtensionWildcards() As String()
        If frm.ComboBox3.SelectedIndex = 1 Then
            Dim t As String = ""
            If frm.CheckBox11.Checked Then t = "*.flv"
            If frm.CheckBox12.Checked Then If t = "" Then t = "*.mp4" Else t = t + ",*.mp4"
            If frm.CheckBox13.Checked Then If t = "" Then t = "*.png" Else t = t + ",*.png"
            Return t.Split(","c)
        End If
        If frm.ComboBox3.SelectedIndex = 7 Then Return {"*.zip"}
        If frm.ComboBox3.SelectedIndex = 8 Then Return {"*.mp3"}

        'RL Media
        If frm.ComboBox3.SelectedIndex = 9 Then Return {"*.png", "*.jpg"}
        If frm.ComboBox3.SelectedIndex = 10 Then Return {"*.png", "*.jpg"}
        If frm.ComboBox3.SelectedIndex = 11 Then Return {"Bezel*.ini"}
        If frm.ComboBox3.SelectedIndex = 12 Then Return {"Layer*.png", "Layer*.jpg"}
        If frm.ComboBox3.SelectedIndex = 13 Then Return {"*.*"}
        If frm.ComboBox3.SelectedIndex = 14 Then Return {"*.pdf", "*.txt", "page*.png", "page*.jpg"}
        If frm.ComboBox3.SelectedIndex = 15 Then Return {"*.mp3"}
        If frm.ComboBox3.SelectedIndex = 16 Then Return {"*.avi", "*.mp4"}

        If frm.ComboBox3.SelectedIndex <> 0 Then Return {"*.png"}

        Dim str As String
        If frm.TextStrip1.Text <> "" Then str = frm.TextStrip1.Text Else str = frm.TextBox3.Text

        Dim s() As String = {}
        If str.Trim = "" Then Return {"*.*"}
        For Each w In str.Split(","c)
            w = w.Trim
            w = "*." + w
            ReDim Preserve s(s.Length)
            s(s.Length - 1) = w
        Next
        Return s
    End Function

    'Change dir in matcher (fill files list)
    Private Sub TextBox4_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged
        If TextBox4.Text.Trim <> "" AndAlso IO.Path.GetInvalidPathChars.Intersect(TextBox4.Text).Count = 0 Then
            Dim tmp As String = FileSystem.GetDirectoryInfo(TextBox4.Text).FullName
            If tmp.ToUpper.Trim <> TextBox4.Text.ToUpper.Trim Then TextBox4.Text = tmp.Trim : Exit Sub
        End If

        With frm
            .Button20_markAsFound.Enabled = False
            Dim fWoExt As String = ""

            .ListBox2.DataSource = Nothing
            .ListBox2.BeginUpdate()
            dt_files.Clear()

            countMatchesFiles = 0 : countNotMatchesFiles = 0 : countAllFiles = 0
            If FileSystem.DirectoryExists(.TextBox4.Text) Then
                .Label20.Text = "Refreshing Files" : .Label20.BackColor = Color.Red : .Label20.Refresh()

                Dim dst As String = ""
                Dim src As String = FileSystem.GetDirectoryInfo(.TextBox4.Text).FullName.ToLower
                If .ComboBox3.SelectedIndex >= 0 Then dst = FileSystem.GetDirectoryInfo(get_HS_Path_of_selected_media()).FullName.ToLower

                If Not src.EndsWith("\") Then src = src + "\" : If Not dst.EndsWith("\") Then dst = dst + "\"

                Dim w() As String = getCurMediaExtensionWildcards()
                Dim list As System.Collections.ObjectModel.ReadOnlyCollection(Of String) = Nothing
                If w(0) <> "" Then
                    If .CheckBox3.Checked Then
                        list = FileSystem.GetDirectories(.TextBox4.Text)
                    Else
                        list = FileSystem.GetFiles(.TextBox4.Text, SearchOption.SearchTopLevelOnly, w)
                    End If

                    If list.Count > 0 And src <> dst Then .Button20_markAsFound.Enabled = True

                    Dim c As Boolean
                    For Each fi As String In list
                        c = True
                        Dim f As String = fi.Substring(fi.LastIndexOf("\") + 1)
                        If .CheckBox3.Checked Then
                            fWoExt = f

                            'In subFoldered mode, we need to check if file exist inside a folder 

                            'If not rl media, we have to alter wildecards to include romName to search in subfolder
                            If frm.ComboBox3.SelectedIndex <= 8 Then w = (From t In w Select t.Replace("*.", f + ".")).ToArray

                            If FileSystem.GetFiles(fi, SearchOption.SearchTopLevelOnly, w).Count = 0 Then c = False
                        Else
                            If f.Contains(".") Then
                                fWoExt = f.Substring(0, f.LastIndexOf("."))
                            Else
                                fWoExt = f
                            End If
                        End If

                        countAllFiles += 1
                        Dim found As Boolean = c And Class1.romlist.Contains(fWoExt.ToLower)
                        If found Then countMatchesFiles += 1 Else countNotMatchesFiles += 1

                        'If .RadioButton4.Checked Then If c And Class1.romlist.Contains(fWoExt.ToLower) Then .ListBox2.Items.Add(f) 
                        'If .RadioButton5.Checked Then If Not c Or Not Class1.romlist.Contains(fWoExt.ToLower) Then .ListBox2.Items.Add(f) 

                        'If .RadioButton4.Checked And found Then .ListBox2.Items.Add(f)
                        If .RadioButton4.Checked And found Then dt_files.Rows.Add({f, f.Replace("-", " ")})
                        'If .RadioButton5.Checked And Not found Then .ListBox2.Items.Add(f)
                        If .RadioButton5.Checked And Not found Then dt_files.Rows.Add({f, f.Replace("-", " ")})
                        'If .RadioButton6.Checked Then .ListBox2.Items.Add(f)
                        If .RadioButton6.Checked Then dt_files.Rows.Add({f, f.Replace("-", " ")})
                    Next
                End If
            End If

            'Sorting table, and reassign to itself
            Dim dv As DataView = dt_files.DefaultView
            'dv.Sort = dt_files.Columns(0).ColumnName + " asc"
            dv.Sort = dt_files.Columns(1).ColumnName + " asc"
            dt_files = dv.ToTable

            .ListBox2.DataSource = dt_files
            .ListBox2.DisplayMember = "name"
            .ListBox2.ValueMember = "name"
            .ListBox2.EndUpdate()

            matcher_update_total_labels()
            '.Label9.Text = "Total: " + .ListBox2.Items.Count.ToString
            .Label20.Text = "Ready" : .Label20.BackColor = Color.LightGreen
        End With
    End Sub
    'Matcher custom dir keyUP
    Private Sub TextBox4_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox4.KeyUp
        If e.KeyCode = Keys.Enter Then
            TextBox4_TextChanged(sender, New System.EventArgs)
        End If
    End Sub

    'Override extension textbox
    Private Sub TextStrip1_changed(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextStrip1.TextChanged
        TextBox4_TextChanged(sender, New System.EventArgs)
    End Sub

    'Subfoldered - mode switch
    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        frm.subfoldered = frm.CheckBox3.Checked
        frm.CheckBox10.Checked = frm.CheckBox3.Checked
        If frm.subfoldered Then frm.CheckBox10.Enabled = False Else frm.CheckBox10.Enabled = True
        TextBox4_TextChanged(sender, New System.EventArgs)
    End Sub
    'Subfoldered - create subfolder for rom
    Private Sub CheckBox10_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox10.CheckedChanged
        frm.subfoldered2 = frm.CheckBox10.Checked
    End Sub

    'Show database switch (with media/without media/both)
    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged, RadioButton2.CheckedChanged, RadioButton3.CheckedChanged
        If DirectCast(sender, RadioButton).Checked = True Then matcher_remplirDatabaseEntryList()
    End Sub
    'Show files switch (matched/unmatched/both)
    Private Sub RadioButton4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton4.CheckedChanged, RadioButton5.CheckedChanged, RadioButton6.CheckedChanged
        If DirectCast(sender, RadioButton).Checked = True Then TextBox4_TextChanged(sender, New System.EventArgs)
    End Sub

    'video ext changed
    Private Sub CheckBox11_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox11.CheckedChanged, CheckBox12.CheckedChanged, CheckBox13.CheckedChanged
        If frm.ComboBox3.SelectedIndex = 1 Then TextBox4_TextChanged(sender, New System.EventArgs)
    End Sub

    'Matcher, media type selection change
    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        With frm
            .TextBox4.Visible = True
            .ComboBox7.Visible = False

            .ListBox1.DataSource = Nothing
            .ListBox1.BeginUpdate() : dt_games.Clear() : .ListBox1.EndUpdate()

            CheckBox3.Enabled = False
            Dim tmpSubFoldered As Boolean = .subfoldered : CheckBox3.Checked = False : .subfoldered = tmpSubFoldered
            Dim tmpSubFoldered2 As Boolean = .subfoldered2 : CheckBox10.Checked = False : .subfoldered2 = tmpSubFoldered2 : CheckBox10.Enabled = False
            If .ComboBox3.SelectedIndex < 0 Then countFound = -1 : countNotFound = -1 : countAll = -1 : matcher_update_total_labels() : Exit Sub
            If .ComboBox1.SelectedIndex < 0 Then MsgBox("Please, select a system.") : .ComboBox3.SelectedIndex = -1 : Exit Sub
            If .ComboBox3.SelectedIndex = 0 Then CheckBox3.Enabled = True : CheckBox10.Enabled = True : CheckBox3.Checked = .subfoldered : CheckBox10.Checked = .subfoldered2
            If .ComboBox3.SelectedIndex >= 9 Then CheckBox3.Checked = True : CheckBox10.Checked = True 'RL Media always subfoldered

            If .RadioStrip1.Checked = True Then
                If .ComboBox3.SelectedIndex = 0 Then
                    If Class1.romPath.Contains("|") Then
                        .ComboBox7.Visible = True
                        .TextBox4.Visible = False
                        .ComboBox7.Items.Clear()
                        For Each t As String In Class1.romPath.Split({"|"}, StringSplitOptions.RemoveEmptyEntries)
                            If t.Trim <> "" AndAlso IO.Path.GetInvalidPathChars.Intersect(t).Count = 0 Then t = FileSystem.GetDirectoryInfo(t).FullName
                            .ComboBox7.Items.Add(t)
                        Next
                        .ComboBox7.Items.Add("Combined View")
                        .ComboBox7.SelectedIndex = 0
                    Else
                        TextBox4.Text = Class1.romPath
                    End If
                ElseIf .ComboBox3.SelectedIndex = 1 Then
                    TextBox4.Text = Class1.videoPath
                ElseIf .ComboBox3.SelectedIndex = 2 Then
                    TextBox4.Text = Class1.HyperspinPath + "Media\" + .ComboBox1.SelectedItem.ToString + "\Images\Wheel\"
                ElseIf .ComboBox3.SelectedIndex = 3 Then
                    TextBox4.Text = Class1.HyperspinPath + "Media\" + .ComboBox1.SelectedItem.ToString + "\Images\Artwork1\"
                ElseIf .ComboBox3.SelectedIndex = 4 Then
                    TextBox4.Text = Class1.HyperspinPath + "Media\" + .ComboBox1.SelectedItem.ToString + "\Images\Artwork2\"
                ElseIf .ComboBox3.SelectedIndex = 5 Then
                    TextBox4.Text = Class1.HyperspinPath + "Media\" + .ComboBox1.SelectedItem.ToString + "\Images\Artwork3\"
                ElseIf .ComboBox3.SelectedIndex = 6 Then
                    TextBox4.Text = Class1.HyperspinPath + "Media\" + .ComboBox1.SelectedItem.ToString + "\Images\Artwork4\"
                ElseIf .ComboBox3.SelectedIndex = 7 Then
                    TextBox4.Text = Class1.HyperspinPath + "Media\" + .ComboBox1.SelectedItem.ToString + "\Themes\"
                ElseIf .ComboBox3.SelectedIndex = 8 Then
                    TextBox4.Text = Class1.HyperspinPath + "Media\" + .ComboBox1.SelectedItem.ToString + "\Sound\Background Music\"
                ElseIf .ComboBox3.SelectedIndex = 9 Then
                    TextBox4.Text = Class1.HyperlaunchPath + "\Media\Artwork\" + .ComboBox1.SelectedItem.ToString + "\"
                ElseIf .ComboBox3.SelectedIndex = 10 Then
                    TextBox4.Text = Class1.HyperlaunchPath + "\Media\Backgrounds\" + .ComboBox1.SelectedItem.ToString + "\"
                ElseIf .ComboBox3.SelectedIndex = 11 Then
                    TextBox4.Text = Class1.HyperlaunchPath + "\Media\\Bezels\" + .ComboBox1.SelectedItem.ToString + "\"
                ElseIf .ComboBox3.SelectedIndex = 12 Then
                    TextBox4.Text = Class1.HyperlaunchPath + "\Media\Fade\" + .ComboBox1.SelectedItem.ToString + "\"
                ElseIf .ComboBox3.SelectedIndex = 13 Then
                    TextBox4.Text = Class1.HyperlaunchPath + "\Media\Guides\" + .ComboBox1.SelectedItem.ToString + "\"
                ElseIf .ComboBox3.SelectedIndex = 14 Then
                    TextBox4.Text = Class1.HyperlaunchPath + "\Media\Manuals\" + .ComboBox1.SelectedItem.ToString + "\"
                ElseIf .ComboBox3.SelectedIndex = 15 Then
                    TextBox4.Text = Class1.HyperlaunchPath + "\Media\Music\" + .ComboBox1.SelectedItem.ToString + "\"
                ElseIf .ComboBox3.SelectedIndex = 16 Then
                    TextBox4.Text = Class1.HyperlaunchPath + "\Media\Videos\" + .ComboBox1.SelectedItem.ToString + "\"
                End If
            End If
            matcher_remplirDatabaseEntryList()
        End With
    End Sub

    'Context Menu - Switch to custom path
    Private Sub contextMeny2radioChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioStrip1.CheckedChanged, RadioStrip2.CheckedChanged
        With frm
            If .RadioStrip1.Checked = True Then
                '.CheckStrip1.Enabled = False
                TextBox4.Enabled = False
                .Button6.Enabled = False
                TextBox4.Text = ""
                If .ComboBox3.SelectedIndex < 0 Then Exit Sub
                If .ComboBox1.SelectedIndex < 0 Then Exit Sub
                If .ComboBox3.SelectedIndex = 0 Then
                    If Class1.romPath.Contains("|") Then
                        .ComboBox7.Visible = True
                        .TextBox4.Visible = False
                        .ComboBox7.Items.Clear()
                        For Each t As String In Class1.romPath.Split({"|"}, StringSplitOptions.RemoveEmptyEntries)
                            If t.Trim <> "" AndAlso IO.Path.GetInvalidPathChars.Intersect(t).Count = 0 Then t = FileSystem.GetDirectoryInfo(t).FullName
                            .ComboBox7.Items.Add(t)
                        Next
                        .ComboBox7.Items.Add("Combined View")
                        .ComboBox7.SelectedIndex = 0
                    Else
                        TextBox4.Text = Class1.romPath
                    End If
                ElseIf ComboBox3.SelectedIndex = 1 Then
                    TextBox4.Text = Class1.videoPath
                ElseIf ComboBox3.SelectedIndex = 2 Then
                    TextBox4.Text = Class1.HyperspinPath + "Media\" + .ComboBox1.SelectedItem.ToString + "\Images\Wheel\"
                ElseIf ComboBox3.SelectedIndex = 3 Then
                    TextBox4.Text = Class1.HyperspinPath + "Media\" + .ComboBox1.SelectedItem.ToString + "\Images\Artwork1\"
                ElseIf ComboBox3.SelectedIndex = 4 Then
                    TextBox4.Text = Class1.HyperspinPath + "Media\" + .ComboBox1.SelectedItem.ToString + "\Images\Artwork2\"
                ElseIf ComboBox3.SelectedIndex = 5 Then
                    TextBox4.Text = Class1.HyperspinPath + "Media\" + .ComboBox1.SelectedItem.ToString + "\Images\Artwork3\"
                ElseIf ComboBox3.SelectedIndex = 6 Then
                    TextBox4.Text = Class1.HyperspinPath + "Media\" + .ComboBox1.SelectedItem.ToString + "\Images\Artwork4\"
                End If
            Else
                If .ComboBox1.SelectedIndex < 0 Then MsgBox("Please, select a system.") : .RadioStrip1.Checked = True : Exit Sub
                '.CheckStrip1.Enabled = True
                TextBox4.Enabled = True
                TextBox4.Visible = True
                .Button6.Enabled = True
                .ComboBox7.Visible = False
            End If
        End With
    End Sub

    'Settings Changed - Always show detailed total in matcher
    Private Sub CheckBox27_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles checkBox27.CheckedChanged
        matcher_update_total_labels()
    End Sub

    'Selecting path in multiple paths combobox
    Private Sub combobox7_selectedChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox7.SelectedIndexChanged
        If ComboBox7.SelectedIndex < ComboBox7.Items.Count - 1 Then
            'If selected <> "Combined View"
            TextBox4.Text = ComboBox7.SelectedItem.ToString
        End If
    End Sub

    'Filter change - DB filter
    Private Sub textbox26_textChange(sender As System.Object, e As System.EventArgs) Handles TextBox26.TextChanged
        Try
            dt_games.DefaultView.RowFilter = "[name] Like '" + TextBox26.Text.Replace("'", "''") + "%'"
        Catch ex As Exception

        End Try
    End Sub
    'Filter change - FileFilter
    Private Sub textbox27_textChange(sender As System.Object, e As System.EventArgs) Handles TextBox27.TextChanged
        Try
            dt_files.DefaultView.RowFilter = "[name] Like '" + TextBox27.Text.Replace("'", "''") + "%'"
        Catch ex As Exception

        End Try
    End Sub

    'Autofilter
    Private Sub listbox1_selection_change(sender As Object, e As System.EventArgs) Handles listbox1.SelectedIndexChanged
        If frm.AutofilterToolStripMenuItem.Checked Then
            If frm.ComboBox3.SelectedIndex = -1 Then Exit Sub
            If countAll <= 0 Then Exit Sub
            If listbox1.SelectedIndex < 0 Then Exit Sub

            Dim s = listbox1.SelectedItem.ToString

            Dim tmpl As String = ""
            Dim regex As String = autofilter_regex
            If regex.StartsWith("%") Then regex = regex.Substring(1) : tmpl = "%"
            Dim rgx As New System.Text.RegularExpressions.Regex(regex)
            Dim m As MatchCollection = rgx.Matches(s)
            If m.Count = 0 Then TextBox27.Text = "" : Exit Sub
            If m.Item(0).Groups.Count > 1 Then
                tmpl = tmpl + m.Item(0).Groups(1).Value
            Else
                tmpl = tmpl + m.Item(0).Groups(0).Value
            End If

            If autofilter_regex_options(0) = True Then
                If tmpl.IndexOf("[") > 0 Then tmpl = tmpl.Substring(0, tmpl.LastIndexOf("["))
            End If
            If autofilter_regex_options(1) = True Then
                If tmpl.IndexOf("(") > 0 Then tmpl = tmpl.Substring(0, tmpl.LastIndexOf("("))
            End If
            TextBox27.Text = tmpl.Trim

            'Old realisation
            'If s.IndexOf("(") >= 0 Then s = s.Substring(0, s.IndexOf("("))
            'If s.IndexOf("[") >= 0 Then s = s.Substring(0, s.IndexOf("["))

            'Dim rgx As New System.Text.RegularExpressions.Regex("[^a-zA-Z0-9 -]")
            's = rgx.Replace(s, "")

            'For Each word As String In s.Split({" "}, StringSplitOptions.None)
            'If word.Length > 3 Then TextBox27.Text = "%" + word : Exit Sub
            'Next
        End If
    End Sub

    'Listbox search as you type logic
    Private Sub ListBox_KeyDown(sender As Object, e As KeyEventArgs) Handles listbox1.KeyDown, listbox2.KeyDown
        Dim n As Integer
        Dim l = DirectCast(sender, ListBox)
        If l.Name.EndsWith("1") Then n = 1 Else n = 2
        If Not listbox_searchAsYouTypeStr.StartsWith(n.ToString) Then listbox_searchAsYouTypeStr = ""
        If listbox_searchAsYouTypeStr = "" Then listbox_searchAsYouTypeStr = n.ToString

        Dim ch = ChrW(e.KeyCode)
        If Char.IsLetterOrDigit(ch) Then
            listbox_searchAsYouTypeStr += ch.ToString.ToUpper

            Dim start = 0
            For i As Integer = start To DirectCast(sender, ListBox).Items.Count - 1
                Dim name = CType(l.Items(i), DataRowView).Row.Item("name").ToString.ToUpper
                If name.StartsWith(listbox_searchAsYouTypeStr.Substring(1)) Then
                    l.SelectedIndex = i : Exit For
                End If
            Next

            e.Handled = True
            e.SuppressKeyPress = True
            listbox_searchAsYouTypeTimer.Enabled = False
            listbox_searchAsYouTypeTimer.Enabled = True
        Else
            listbox_searchAsYouTypeStr = ""
            listbox_searchAsYouTypeTimer.Enabled = False
        End If
    End Sub
    Private Sub ListBox1_SearchAsYouTypeTimer() Handles listbox_searchAsYouTypeTimer.Tick
        listbox_searchAsYouTypeStr = ""
        listbox_searchAsYouTypeTimer.Enabled = False
    End Sub

    'Show autorenamer context menu
    'Private Sub Button_associate_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Button5_Associate.MouseDown
    'If e.Button = MouseButtons.Right Then
    'frm.myContextMenu7.Show(Cursor.Position.X, Cursor.Position.Y)
    'End If
    'End Sub
End Class

'Class ExtendedLabel
'Inherits Label

'Protected Overrides Sub OnPaint(e As PaintEventArgs)
'    MyBase.OnPaint(e)
'Dim myText = Text
''Text = ""
'Dim w As Integer
'Dim ary() As String = myText.Split({"|"c})
'    If ary.Count = 2 Then
'Dim fontBold As New Font(Font, FontStyle.Bold)
'Dim s1 As Size = TextRenderer.MeasureText(ary(0), Font)
'Dim s2 As Size = TextRenderer.MeasureText(ary(1), fontBold)
'Dim r1 As Rectangle = New Rectangle(New Point(0, 0), s1)
'Dim r2 As Rectangle = New Rectangle(r1.Right, r1.Top, r1.Width, r1.Height)
'        TextRenderer.DrawText(e.Graphics, ary(0), Font, r1, ForeColor)
'        TextRenderer.DrawText(e.Graphics, ary(1), fontBold, r2, ForeColor)
'        w = r1.Width + r2.Width
'    Else
'        TextRenderer.DrawText(e.Graphics, myText, Font, New Point(0, 0), ForeColor)
'        w = TextRenderer.MeasureText(myText, Font).Width
'    End If
'    Me.Width = w


''Text = myText
'End Sub
'End Class