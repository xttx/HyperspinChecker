Imports Microsoft.VisualBasic.FileIO
Imports System.Text.RegularExpressions

Public Class Class3_matcher
#Region "Declarations"
    Dim dt_files As New DataTable
    Dim countAll As Integer = -1, countFound As Integer = -1, countNotFound As Integer = -1
    Dim countMatchesFiles As Integer = -1, countNotMatchesFiles As Integer = -1, countAllFiles As Integer = -1
    Private WithEvents Button5_Associate As Button = Form1.Button5_Associate
    Private WithEvents ButtonStrip1 As Button = Form1.ButtonStrip1
    Private WithEvents Button7 As Button = Form1.Button7
    Private WithEvents Button20_markAsFound As Button = Form1.Button20_markAsFound
    Private WithEvents ComboBox3 As ComboBox = Form1.ComboBox3
    Private WithEvents ComboBox7 As ComboBox = Form1.ComboBox7
    Private WithEvents TextBox4 As TextBox = Form1.TextBox4
    Private WithEvents TextStrip1 As TextBox = Form1.TextStrip1
    Private WithEvents CheckBox3 As CheckBox = Form1.CheckBox3
    Private WithEvents CheckBox10 As CheckBox = Form1.CheckBox10
    Private WithEvents CheckBox11 As CheckBox = Form1.CheckBox11
    Private WithEvents CheckBox12 As CheckBox = Form1.CheckBox12
    Private WithEvents CheckBox13 As CheckBox = Form1.CheckBox13
    Private WithEvents TextBox27 As TextBox = Form1.TextBox27

    Private WithEvents RadioStrip1 As RadioButton = Form1.RadioStrip1
    Private WithEvents RadioStrip2 As RadioButton = Form1.RadioStrip2
    Private WithEvents RadioButton1 As RadioButton = Form1.RadioButton1
    Private WithEvents RadioButton2 As RadioButton = Form1.RadioButton2
    Private WithEvents RadioButton3 As RadioButton = Form1.RadioButton3
    Private WithEvents RadioButton4 As RadioButton = Form1.RadioButton4
    Private WithEvents RadioButton5 As RadioButton = Form1.RadioButton5
    Private WithEvents RadioButton6 As RadioButton = Form1.RadioButton6

    Private WithEvents checkBox27 As CheckBox = Form1.CheckBox27
    Private WithEvents listbox1 As ListBox = Form1.ListBox1
    Private WithEvents listbox2 As ListBox = Form1.ListBox2
    'Friend WithEvents myContextMenu7 As New ToolStripDropDownMenu 'autorenamer
    Public Shared autofilter_regex As String = "%[A-Za-z]{4}[A-Za-z]*"
    Public Shared autofilter_regex_options() As Boolean = {False, False}
#End Region

    'Constructor
    Sub New()
        dt_files.Columns.Add("name")
        dt_files.Columns.Add("nameWithoutHyphen")
    End Sub

    'matcher - remplir database
    Public Sub matcher_remplirDatabaseEntryList()
        countAll = 0
        countFound = 0
        countNotFound = 0
        With Form1
            .ListBox1.Items.Clear()
            If .ComboBox3.SelectedIndex < 0 Then Exit Sub
            If .ComboBox1.SelectedIndex < 0 Then Exit Sub
            If .DataGridView1.Rows.Count = 0 Then
                .ListBox1.Items.Add("...Please, make ""check"" in summary page,")
                .ListBox1.Items.Add("before using this future") : Exit Sub
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
                End If
                Dim retVal As Boolean = matcher_remplirDatabaseEntryList_addRow(row, cellindex)
                If retVal Then countFound += 1 Else countNotFound += 1
                countAll += 1
            Next

            matcher_update_total_labels()
            'If .ListBox1.Items.Count = 0 And .RadioButton1.Checked Then .ListBox1.Items.Add("No matched " + .ComboBox3.SelectedItem.ToString)
            'If .ListBox1.Items.Count = 0 And .RadioButton2.Checked Then .ListBox1.Items.Add("No missing " + .ComboBox3.SelectedItem.ToString)
            'If .ListBox1.Items.Count = 0 And .RadioButton3.Checked Then .ListBox1.Items.Add("No " + .ComboBox3.SelectedItem.ToString + " found in current system database")
            .Label20.Text = "Ready" : .Label20.BackColor = Color.LightGreen : .Label20.Refresh()
        End With
    End Sub

    'matcher - remplir database SUB
    Private Function matcher_remplirDatabaseEntryList_addRow(ByVal row As DataGridViewRow, ByVal cellIndex As Integer) As Boolean
        Dim retVal As Boolean = False
        With Form1
            If row.Cells(cellIndex).Value.ToString = "YES" Then retVal = True
            If .RadioButton1.Checked Then If row.Cells(cellIndex).Value.ToString = "YES" Then .ListBox1.Items.Add(row.Cells(1).Value)
            If .RadioButton2.Checked Then If row.Cells(cellIndex).Value.ToString = "NO" Then .ListBox1.Items.Add(row.Cells(1).Value)
            If .RadioButton3.Checked Then .ListBox1.Items.Add(row.Cells(1).Value)
        End With
        Return retVal
    End Function

    'Update TOTAL labels
    Private Sub matcher_update_total_labels()
        With Form1
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
        With Form1
            If .ListBox1.SelectedIndex < 0 Then MsgBox("Select a database entry to associate a file to.") : Exit Sub
            If .ListBox2.SelectedIndex < 0 Then MsgBox("Select a file to associate to selected database entry.") : Exit Sub

            Dim ext As String = ""
            Dim l1 As String = .ListBox1.SelectedItem.ToString
            'Dim l2 As String = .ListBox2.SelectedItem.ToString
            Dim l2 As String = DirectCast(.ListBox2.SelectedItem, DataRowView).Item(0).ToString

            Dim src As String = Microsoft.VisualBasic.FileIO.FileSystem.GetDirectoryInfo(.TextBox4.Text).FullName.ToLower
            Dim dst As String = Microsoft.VisualBasic.FileIO.FileSystem.GetDirectoryInfo(getPath()).FullName.ToLower
            If Not src.EndsWith("\") Then src = src + "\" : If Not dst.EndsWith("\") Then dst = dst + "\"

            '''''''''''HANDLE SUBFOLDERED MODE
            Dim list As System.Collections.ObjectModel.ReadOnlyCollection(Of String) = Nothing
            If .CheckBox3.Checked Then
                Dim wildcards() As String = strToWildcards()
                list = FileSystem.GetFiles(.TextBox4.Text + "\" + l2, FileIO.SearchOption.SearchTopLevelOnly, wildcards)
                If list.Count = 0 Then MsgBox("This folder does not contain any files that match rom's extensions.") : Exit Sub
                If list.Count > 1 Then MsgBox("This folder contains multiple files that match rom's extensions. This situation is not handled yet, sorry." + vbCrLf + "The program simply doesn't know what file need to be renamed. Try use 'Override Extension' in options.") : Exit Sub
            Else
                Dim tmpList = New List(Of String) : tmpList.Add(l2)
                list = New System.Collections.ObjectModel.ReadOnlyCollection(Of String)(tmpList)
                ext = l2.Substring(l2.LastIndexOf("."))
            End If
            '''''''''''END HANDLER OF SUBFOLDERED MODE

            '''''''''''Actual Copy/Rename
            Dim res() As String = associate_copyMove(src, dst, list)
            If res(0) <> "" Then
                MsgBox(res(0)) : .Label20.Text = "READY" : .Label20.BackColor = Color.LightGreen : .Label20.Refresh() : Exit Sub
            End If
            '''''''''''END Actual Copy/Rename

            'remove from found if needed
            'TODO TEST LOGIC WITH "subFOLDERED MODE"
            Dim l2woext As String = l2
            If Not .CheckBox3.Checked Then l2woext = l2.Substring(0, l2.LastIndexOf(".")).ToLower
            If src = dst And Class1.romlist.Contains(l2woext) Then
                Dim exist As Boolean = False
                For Each s As String In strToWildcards()
                    s = s.Replace("*", l2woext)
                    If FileSystem.FileExists(src + "\" + s) Then exist = True : Exit For
                Next
                If Not exist Then Button20_markAsFound_sub(l2woext, Button20_markAsFound_comboToCol(), True)
            End If
            'add to found if needed
            Dim copyToHsFolder As Boolean = .AssocOption_fileInDiffFolder_copyToHS.Checked Or .AssocOption_fileInDiffFolder_moveToHS.Checked
            If src = dst Or copyToHsFolder Then Button20_markAsFound_sub(l1, Button20_markAsFound_comboToCol())

            'Moving selected index
            'on box2 (filelist)
            Dim currentTopIndex As Integer = .ListBox2.TopIndex
            Dim currentSelectedIndex As Integer
            'V svoey direktorii ILI (NE v svoey i NE kopiruem)
            'If (src = dst Or Not .CheckStrip1.Checked) Then
            .ListBox2.BeginUpdate()
            If (src = dst Or Not copyToHsFolder) Then
                If .RadioButton5.Checked Then 'show unmatched files
                    currentSelectedIndex = .ListBox2.SelectedIndex
                    '.ListBox2.Items.RemoveAt(.ListBox2.SelectedIndex)
                    dt_files.Rows.Remove(DirectCast(.ListBox2.SelectedItem, DataRowView).Row)
                Else 'show matched or both files
                    '.ListBox2.Items(.ListBox2.SelectedIndex) = l1 + ext
                    DirectCast(.ListBox2.SelectedItem, DataRowView).Item(0) = l1 + ext
                    currentTopIndex += 1
                    currentSelectedIndex = .ListBox2.SelectedIndex + 1
                End If
            Else
                currentSelectedIndex = .ListBox2.SelectedIndex + 1
            End If
            Dim b As New BindingContext : b(dt_files).EndCurrentEdit()
            If currentTopIndex >= 0 Then .ListBox2.TopIndex = currentTopIndex
            If .ListBox2.Items.Count > currentSelectedIndex Then .ListBox2.SelectedIndex = currentSelectedIndex Else .ListBox2.SelectedIndex = currentSelectedIndex - 1
            .ListBox2.EndUpdate()

            'on box1 (DB entry list)
            If .RadioButton2.Checked And (src = dst Or copyToHsFolder) Then
                currentSelectedIndex = .ListBox1.SelectedIndex
                .ListBox1.Items.RemoveAt(.ListBox1.SelectedIndex)
            Else
                currentSelectedIndex = .ListBox1.SelectedIndex + 1
            End If
            If .ListBox1.Items.Count > currentSelectedIndex Then .ListBox1.SelectedIndex = currentSelectedIndex Else .ListBox1.SelectedIndex = currentSelectedIndex - 1

            countFound += 1 : countNotFound -= 1
            countMatchesFiles += 1 : countNotMatchesFiles -= 1 : matcher_update_total_labels()
            '.Label7.Text = "Total: " + .ListBox1.Items.Count.ToString
            '.Label9.Text = "Total: " + .ListBox2.Items.Count.ToString
            .Label20.Text = "READY" : .Label20.BackColor = Color.LightGreen : .Label20.Refresh()
        End With
    End Sub

    'Associate SUB - creating file operation array
    Private Function associate_copyMove(ByVal src As String, ByVal dst As String, Optional ByVal list As System.Collections.ObjectModel.ReadOnlyCollection(Of String) = Nothing) As String()
        With Form1
            Dim ext As String = ""
            Dim l1 As String = .ListBox1.SelectedItem.ToString
            'Dim l2 As String = .ListBox2.SelectedItem.ToString
            Dim l2 As String = DirectCast(.ListBox2.SelectedItem, DataRowView).Item(0).ToString
            Dim op As New List(Of String())
            Try
                If Not .CheckBox3.Checked Then
                    'SINGLEFILE MODE
                    ext = l2.Substring(l2.LastIndexOf("."))

                    If src = dst Then
                        If .CheckBox10.Checked Then dst = dst + "\" + l1

                        'src is in HS folder
                        If Form1.AssocOption_fileInHsFolder_copy.Checked Then
                            'Copy (duplicate)
                            op.Add({"FILECOPY", src + "\" + l2, dst + "\" + l1 + ext})
                        ElseIf Form1.AssocOption_fileInHsFolder_move.Checked Then
                            'Move (rename)
                            op.Add({"FILERENAME", src + "\" + l2, dst + "\" + l1 + ext})
                        End If
                    Else
                        'src is in different folder
                        If Form1.AssocOption_fileInDiffFolder_copy.Checked Then
                            'Copy in place
                            op.Add({"FILECOPY", src + "\" + l2, src + "\" + l1 + ext})
                        ElseIf Form1.AssocOption_fileInDiffFolder_move.Checked Then
                            'Move in place
                            op.Add({"FILERENAME", src + "\" + l2, src + "\" + l1 + ext})
                        ElseIf Form1.AssocOption_fileInDiffFolder_copyToHS.Checked Then
                            'Copy to HS folder
                            op.Add({"FILECOPY", src + "\" + l2, dst + "\" + l1 + ext})
                        ElseIf Form1.AssocOption_fileInDiffFolder_moveToHS.Checked Then
                            'Move to HS folder
                            op.Add({"FILERENAME", src + "\" + l2, dst + "\" + l1 + ext})
                        End If
                    End If
                Else
                    'SUBFOLDERED MODE
                    ext = list.Item(0).Substring(list.Item(0).LastIndexOf("."))
                    Dim filename As String = list.Item(0).Substring(list.Item(0).LastIndexOf("\") + 1)

                    If src = dst Then
                        'src is in HS folder
                        If Form1.AssocOption_fileInHsFolder_copy.Checked Then
                            'Copy (duplicate)
                            op.Add({"DIRCOPY", src + "\" + l2, dst + "\" + l1 + "\"})
                            Dim res0 As String = associate_fileOP(op)
                            If res0 <> "" Then Return {res0, ext}
                            op = New List(Of String())
                            op.Add({"FILERENAME", dst + "\" + l1 + "\" + filename, dst + "\" + l1 + "\" + l1 + ext, "0"})
                        ElseIf Form1.AssocOption_fileInHsFolder_move.Checked Then
                            'Move (rename)
                            op.Add({"FILERENAME", list.Item(0), dst + "\" + l2 + "\" + l1 + ext, "0"})
                            op.Add({"DIRRENAME", src + "\" + l2, l1, "0"})
                        End If
                    Else
                        'src is in different folder
                        If Form1.AssocOption_fileInDiffFolder_copy.Checked Then
                            'Copy in place
                            op.Add({"DIRCOPY", src + "\" + l2, src + "\" + l1 + "\"})
                            Dim res0 As String = associate_fileOP(op)
                            If res0 <> "" Then Return {res0, ext}
                            op = New List(Of String())
                            op.Add({"FILERENAME", src + "\" + l1 + "\" + filename, src + "\" + l1 + "\" + l1 + ext, "0"})
                        ElseIf Form1.AssocOption_fileInDiffFolder_move.Checked Then
                            'Move in place
                            op.Add({"FILERENAME", list.Item(0), src + "\" + l2 + "\" + l1 + ext, "0"})
                            op.Add({"DIRRENAME", src + "\" + l2, l1, "0"})
                        ElseIf Form1.AssocOption_fileInDiffFolder_copyToHS.Checked Then
                            'Copy to HS folder
                            op.Add({"DIRCOPY", src + "\" + l2, dst + "\" + l1 + "\"})
                            Dim res0 As String = associate_fileOP(op)
                            If res0 <> "" Then Return {res0, ext}
                            op = New List(Of String())
                            op.Add({"FILERENAME", dst + "\" + l1 + "\" + filename, dst + "\" + l1 + "\" + l1 + ext, "0"})
                        ElseIf Form1.AssocOption_fileInDiffFolder_moveToHS.Checked Then
                            'Move to HS folder

                        End If
                    End If
                End If

                Dim res As String = associate_fileOP(op)
                If res <> "" Then Return {res, ext}
            Catch ex As Exception
                Return {ex.Message, ext}
            End Try
            Return {"", ext}
        End With
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
                If tmpExt.ToLower = "cue" And Not Form1.CheckBox6.Checked Then
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
                    For Each e As String In Form1.TextBox16.Text.ToLower.Split(","c)
                        If e.Trim = tmpExt.ToLower Then
                            For Each f As String In FileSystem.GetFiles(tmppath, SearchOption.SearchTopLevelOnly, {"*.cue"})
                                listFilesInCue = associate_listFilesFromCue(tmppath + tmpfileNameWOext + ".cue", listaudio)
                                If listFilesInCue(0)(0).ToLower = tmpfileName.ToLower Then
                                    tmp.Add({o(0), f, newpath + f.Substring(f.LastIndexOf("\") + 1)})
                                    For Each fA As String In listFilesInCue(1)
                                        tmp.Add({o(0), tmppath + fA, newpath + fA})
                                    Next
                                End If
                            Next
                        End If
                    Next
                End If

                'check pairs (mdf/mds list)
                For Each l As String In Form1.ListBox4.Items
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
        Form1.undo.Add(New List(Of String))
        If tmp.Count > 0 Then op.InsertRange(1, tmp)
        Try
            For Each o In op
                If o(0) = "FILECOPY" Then
                    FileSystem.CopyFile(o(1), o(2))
                    If o(1).Substring(o(1).LastIndexOf(".") + 1).ToLower = "cue" And restoreCue Then
                        Form1.undo(Form1.undo.Count - 1).Add("RESTORECUE?" + o(1))
                        Form1.undo(Form1.undo.Count - 1).Add("FILEREMOVE?" + o(2))
                    Else
                        Form1.undo(Form1.undo.Count - 1).Add("FILEREMOVE?" + o(2))
                    End If
                End If
                If o(0) = "FILERENAME" Then
                    FileSystem.MoveFile(o(1), o(2))
                    If o(1).Substring(o(1).LastIndexOf(".") + 1).ToLower = "cue" And restoreCue Then
                        Form1.undo(Form1.undo.Count - 1).Add("RESTORECUE?" + o(1))
                        Form1.undo(Form1.undo.Count - 1).Add("FILEREMOVE?" + o(2))
                    Else
                        Form1.undo(Form1.undo.Count - 1).Add("FILERENAME?" + o(2) + "?" + o(1))
                    End If
                End If
                If o(0) = "DIRCOPY" Then FileSystem.CopyDirectory(o(1), o(2)) : Form1.undo(Form1.undo.Count - 1).Add("DIRREMOVE?" + o(2))
                If o(0) = "DIRMOVE" Then FileSystem.MoveDirectory(o(1), o(2)) : Form1.undo(Form1.undo.Count - 1).Add("DIRMOVE?" + o(2) + "?" + o(1))
                If o(0) = "DIRRENAME" Then
                    FileSystem.RenameDirectory(o(1), o(2))
                    Dim name As String = o(1).Substring(o(1).LastIndexOf("\") + 1)
                    Dim path As String = o(1).Substring(0, o(1).LastIndexOf("\") + 1)
                    Form1.undo(Form1.undo.Count - 1).Add("DIRRENAME?" + path + o(2) + "?" + name)
                End If
            Next
        Catch ex As Exception
            If Form1.undo(Form1.undo.Count - 1).Count = 0 Then Form1.undo.RemoveAt(Form1.undo.Count - 1)
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

    'Associate SUB
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
            Do While Microsoft.VisualBasic.FileIO.FileSystem.FileExists(Form1.xmlPath + ".backup" + backupN.ToString)
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
        If Form1.undo.Count = 0 Then MsgBox("Nothing to undo.") : Exit Sub
        Dim msg As String = "The following operations will be executed:" + vbCrLf
        For i As Integer = Form1.undo(Form1.undo.Count - 1).Count - 1 To 0 Step -1
            Dim m() As String = Form1.undo(Form1.undo.Count - 1)(i).Split("?"c)
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
            If MsgBox("Remove this operation from undo list?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then Form1.undo.RemoveAt(Form1.undo.Count - 1)
            Exit Sub
        End If

        For i As Integer = Form1.undo(Form1.undo.Count - 1).Count - 1 To 0 Step -1
            Try
                Dim o() As String = Form1.undo(Form1.undo.Count - 1)(i).Split("?"c)
                If o(0) = "FILEREMOVE" Then FileSystem.DeleteFile(o(1))
                If o(0) = "FILERENAME" Then FileSystem.MoveFile(o(1), o(2))
                If o(0) = "DIRREMOVE" Then FileSystem.DeleteDirectory(o(1), DeleteDirectoryOption.DeleteAllContents)
                If o(0) = "DIRRENAME" Then FileSystem.RenameDirectory(o(1), o(2))
                If o(0) = "DIRMOVE" Then FileSystem.MoveDirectory(o(1), o(2))
                If o(0) = "RESTORECUE" Then
                    Dim backupN As Integer = 0
                    If FileSystem.FileExists(o(1) + ".backup") Then
                        Do While Microsoft.VisualBasic.FileIO.FileSystem.FileExists(Form1.xmlPath + ".backup" + backupN.ToString)
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
        Form1.undo.RemoveAt(Form1.undo.Count - 1)
        MsgBox("You have to reCheck to reflect changes")
    End Sub

    'get path to media based on matcher mediaselect Combobox
    Private Function getPath() As String
        With Form1
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
            End If
            Return ""
        End With
    End Function

    'matcher options
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Form1.myContextMenu2.Show(Cursor.Position.X, Cursor.Position.Y)
    End Sub

    'markAsFound
    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20_markAsFound.Click
        With Form1
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
                        If FileSystem.GetFiles(fi, SearchOption.SearchTopLevelOnly, strToWildcards).Count = 0 Then Continue For
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
        If Form1.ComboBox3.SelectedIndex = 0 Then Return 2
        If Form1.ComboBox3.SelectedIndex = 1 Then Return 3
        If Form1.ComboBox3.SelectedIndex = 2 Then Return 5
        If Form1.ComboBox3.SelectedIndex = 3 Then Return 6
        If Form1.ComboBox3.SelectedIndex = 4 Then Return 7
        If Form1.ComboBox3.SelectedIndex = 5 Then Return 8
        If Form1.ComboBox3.SelectedIndex = 6 Then Return 9
        If Form1.ComboBox3.SelectedIndex = 7 Then Return 4
        If Form1.ComboBox3.SelectedIndex = 8 Then Return 10
        Return 0
    End Function

    'markAsFoundSub
    Private Function Button20_markAsFound_sub(ByVal fWoExt As String, ByVal col As Integer, Optional ByVal remove As Boolean = False) As Boolean
        If Class1.romlist.Contains(fWoExt.ToLower) Then
            'mark in grid
            For r = 0 To Form1.DataGridView1.Rows.Count - 1
                If Form1.DataGridView1.Item(1, r).Value.ToString.ToLower = fWoExt.ToLower Then
                    If Not remove Then
                        Form1.DataGridView1.Item(col, r).Value = "YES"
                        Form1.DataGridView1.Item(col, r).Style.BackColor = Form1.colorYES
                    Else
                        Form1.DataGridView1.Item(col, r).Value = "NO"
                        Form1.DataGridView1.Item(col, r).Style.BackColor = Form1.colorNO
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

    Private Function strToWildcards() As String()
        If Form1.ComboBox3.SelectedIndex = 1 Then
            Dim t As String = ""
            If Form1.CheckBox11.Checked Then t = "*.flv"
            If Form1.CheckBox12.Checked Then If t = "" Then t = "*.mp4" Else t = t + ",*.mp4"
            If Form1.CheckBox13.Checked Then If t = "" Then t = "*.png" Else t = t + ",*.png"
            Return t.Split(","c)
        End If
        If Form1.ComboBox3.SelectedIndex = 7 Then Return {"*.zip"}
        If Form1.ComboBox3.SelectedIndex = 8 Then Return {"*.mp3"}
        If Form1.ComboBox3.SelectedIndex <> 0 Then Return {"*.png"}

        Dim str As String
        If Form1.TextStrip1.Text <> "" Then str = Form1.TextStrip1.Text Else str = Form1.TextBox3.Text

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

        With Form1
            .Button20_markAsFound.Enabled = False
            Dim fWoExt As String = ""

            '.ListBox2.Items.Clear()
            .ListBox2.DataSource = Nothing
            .ListBox2.BeginUpdate()
            dt_files.Clear()

            countMatchesFiles = 0 : countNotMatchesFiles = 0 : countAllFiles = 0
            If FileSystem.DirectoryExists(.TextBox4.Text) Then
                .Label20.Text = "Refreshing Files" : .Label20.BackColor = Color.Red : .Label20.Refresh()

                Dim dst As String = ""
                Dim src As String = FileSystem.GetDirectoryInfo(.TextBox4.Text).FullName.ToLower
                If .ComboBox3.SelectedIndex >= 0 Then dst = FileSystem.GetDirectoryInfo(getPath()).FullName.ToLower

                If Not src.EndsWith("\") Then src = src + "\" : If Not dst.EndsWith("\") Then dst = dst + "\"

                Dim w() As String = strToWildcards()
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
                        'WE NEED TO CHECK FILES INSIDE FOLDERS WHEN subFoldered mode
                        If .CheckBox3.Checked Then
                            fWoExt = f
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

    'Subfoldered1
    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        Form1.subfoldered = Form1.CheckBox3.Checked
        Form1.CheckBox10.Checked = Form1.CheckBox3.Checked
        If Form1.subfoldered Then Form1.CheckBox10.Enabled = False Else Form1.CheckBox10.Enabled = True
        TextBox4_TextChanged(sender, New System.EventArgs)
    End Sub

    'Subfoldered2
    Private Sub CheckBox10_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox10.CheckedChanged
        Form1.subfoldered2 = Form1.CheckBox10.Checked
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged, RadioButton2.CheckedChanged, RadioButton3.CheckedChanged
        If DirectCast(sender, RadioButton).Checked = True Then matcher_remplirDatabaseEntryList()
    End Sub

    Private Sub RadioButton4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton4.CheckedChanged, RadioButton5.CheckedChanged, RadioButton6.CheckedChanged
        If DirectCast(sender, RadioButton).Checked = True Then TextBox4_TextChanged(sender, New System.EventArgs)
    End Sub

    'video ext changed
    Private Sub CheckBox11_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox11.CheckedChanged, CheckBox12.CheckedChanged, CheckBox13.CheckedChanged
        If Form1.ComboBox3.SelectedIndex = 1 Then TextBox4_TextChanged(sender, New System.EventArgs)
    End Sub

    'Matcher, media type selection change
    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        With Form1
            .TextBox4.Visible = True
            .ComboBox7.Visible = False
            .ListBox1.Items.Clear()
            CheckBox3.Enabled = False
            Dim tmpSubFoldered As Boolean = .subfoldered : CheckBox3.Checked = False : .subfoldered = tmpSubFoldered
            Dim tmpSubFoldered2 As Boolean = .subfoldered2 : CheckBox10.Checked = False : .subfoldered2 = tmpSubFoldered2 : CheckBox10.Enabled = False
            If .ComboBox3.SelectedIndex < 0 Then countFound = -1 : countNotFound = -1 : countAll = -1 : matcher_update_total_labels() : Exit Sub
            If .ComboBox1.SelectedIndex < 0 Then MsgBox("Please, select a system.") : .ComboBox3.SelectedIndex = -1 : Exit Sub
            If .ComboBox3.SelectedIndex = 0 Then CheckBox3.Enabled = True : CheckBox10.Enabled = True : CheckBox3.Checked = .subfoldered : CheckBox10.Checked = .subfoldered2

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
                End If
            End If
            matcher_remplirDatabaseEntryList()
        End With
    End Sub

    'Context Menu - Switch to custom path
    Private Sub contextMeny2radioChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioStrip1.CheckedChanged, RadioStrip2.CheckedChanged
        With Form1
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

    'Filter change - FileFilter
    Private Sub textbox27_textChange(sender As System.Object, e As System.EventArgs) Handles TextBox27.TextChanged
        Try
            dt_files.DefaultView.RowFilter = "[name] like '" + TextBox27.Text.Replace("'", "''") + "%'"
        Catch ex As Exception

        End Try
    End Sub

    'Autofilter
    Private Sub listbox1_selection_change(sender As Object, e As System.EventArgs) Handles listbox1.SelectedIndexChanged
        If Form1.AutofilterToolStripMenuItem.Checked Then
            If Form1.ComboBox3.SelectedIndex = -1 Then Exit Sub
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

    'Show autorenamer context menu
    'Private Sub Button_associate_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Button5_Associate.MouseDown
    'If e.Button = MouseButtons.Right Then
    'Form1.myContextMenu7.Show(Cursor.Position.X, Cursor.Position.Y)
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