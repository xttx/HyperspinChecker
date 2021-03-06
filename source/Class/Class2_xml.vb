﻿Imports Microsoft.VisualBasic.FileIO
Public Class Class2_xml
    Private WithEvents xml_crop As Button = Form1.Button18
    Private WithEvents xml_from_folder As ToolStripMenuItem = Form1.CreateDatabaseXMLFromRomFolderToolStripMenuItem
    Public Shared xml_from_folder_options_src As String = ""
    Public Shared xml_from_folder_options_dst As String = ""
    Public Shared xml_from_folder_options_adv As String = ""
    Public Shared xml_from_folder_options_fillCRC As Boolean = False
    Public Shared xml_from_folder_options_removeParanthesis As Boolean = False
    Public Shared xml_from_folder_options_removeParanthesis_exceptions As String = ""
    Public Shared xml_from_folder_options_mergeWithExisting As String = ""
    Public Shared xml_from_folder_options_includeFiles As Boolean = False
    Public Shared xml_from_folder_options_includeFolders As Boolean = False
    Public Shared xml_from_folder_options_includeFoldersCheck As Boolean = False
    Public Shared xml_from_folder_options_includeExtentions As String = ""

    'Private WithEvents xml_remove_clones As Button = Form1.Button28
    Private WithEvents xml_remove_clones As ToolStripMenuItem = Form1.RemoveClonesFromCurrentDBToolStripMenuItem
    Private WithEvents xml_compare As ToolStripMenuItem = Form1.CompareToolStripMenuItem
    Private WithEvents move_roms_from_txt_to_subfolder As ToolStripMenuItem = Form1.MoveRomsInProvidedListtxtToSubfolderToolStripMenuItem
    Private WithEvents update_delete_rom As Button = Form1.Button21
    Private WithEvents update_delete_rom_mm As ToolStripMenuItem = Form1.CommitDbEditionsToolStripMenuItem
    Private WithEvents main_table As DataGridView = Form1.DataGridView1

    'crop xml
    Private Sub xml_crop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles xml_crop.Click
        Dim xmlPath As String = Form1.xmlPath
        If Form1.DataGridView1.Rows.Count = 0 Then MsgBox("DataGrid must contain at least one entry") : Exit Sub
        Dim backupN As Integer = 0
        If Microsoft.VisualBasic.FileIO.FileSystem.FileExists(xmlPath + ".backup.xml") Then
            Do While Microsoft.VisualBasic.FileIO.FileSystem.FileExists(xmlPath + ".backup" + backupN.ToString + ".xml")
                backupN += 1
            Loop
            Microsoft.VisualBasic.FileIO.FileSystem.CopyFile(xmlPath, xmlPath + ".backup" + backupN.ToString + ".xml")
        Else
            Microsoft.VisualBasic.FileIO.FileSystem.CopyFile(xmlPath, xmlPath + ".backup.xml")
        End If

        Dim romName As String
        Dim x As New Xml.XmlDocument
        x.Load(xmlPath)
        For Each node As Xml.XmlNode In x.SelectNodes("/menu/game")
            romName = node.Attributes.GetNamedItem("name").Value
            If xml_crop.Text.ToLower.Contains("found") Then
                If Not Class1.romFoundlist.Contains(romName.ToLower) Then
                    node.ParentNode.RemoveChild(node)
                End If
            Else
                If Not Form3_mameFoldersFilter.filter.Contains(romName.ToLower) Then
                    node.ParentNode.RemoveChild(node)
                End If
            End If
        Next
        Dim w As Xml.XmlWriter = Xml.XmlWriter.Create(xmlPath, New Xml.XmlWriterSettings With {.Indent = True, .NewLineHandling = Xml.NewLineHandling.None})
        x.Save(w) : w.Close()
        MsgBox("Done!")
    End Sub

    'xmlFromFolder
    Private Sub xml_from_folder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles xml_from_folder.Click
        xml_from_folder_options_src = ""
        xml_from_folder_options_dst = ""
        Dim stat_files As Integer = 0
        Dim stat_folders As Integer = 0
        Dim stat_folders_skipped As Integer = 0
        Dim stat_merged_orig As Integer = 0
        Dim stat_merged_new As Integer = 0

        Dim f As New FormE_Create_database_XML_from_folder
        f.ShowDialog()

        xml_from_folder_options_src = xml_from_folder_options_src.Trim
        xml_from_folder_options_dst = xml_from_folder_options_dst.Trim
        xml_from_folder_options_includeExtentions = xml_from_folder_options_includeExtentions.Trim
        If xml_from_folder_options_includeExtentions = "" Then xml_from_folder_options_includeExtentions = "*"
        If xml_from_folder_options_src = "" Or xml_from_folder_options_dst = "" Then Exit Sub
        Dim x As New Xml.XmlDocument
        Dim menuElem As Xml.XmlElement = x.CreateElement("menu") : x.AppendChild(menuElem)

        If xml_from_folder_options_includeFiles Then
            Dim maskArr() As String = xml_from_folder_options_includeExtentions.Split(";"c)
            For i As Integer = 0 To maskArr.Count - 1
                maskArr(i) = "*." + maskArr(i)
            Next

            For Each file As String In FileSystem.GetFiles(xml_from_folder_options_src, SearchOption.SearchTopLevelOnly, maskArr)
                stat_files += 1
                xml_from_folder_Click_sub(x, menuElem, file)
            Next
        End If
        If xml_from_folder_options_includeFolders Then
            Dim exts() As String = xml_from_folder_options_includeExtentions.Split(";"c)
            For Each folder As String In FileSystem.GetDirectories(xml_from_folder_options_src, SearchOption.SearchTopLevelOnly)
                If xml_from_folder_options_includeFoldersCheck Then
                    Dim fld As String = folder.Substring(folder.LastIndexOf("\") + 1)
                    Dim fls(exts.Count - 1) As String
                    For i As Integer = 0 To exts.Count - 1
                        fls(i) = fld + "." + exts(i)
                    Next

                    Dim files() As String = FileSystem.GetFiles(folder, SearchOption.SearchTopLevelOnly, fls).ToArray
                    If files.Count > 0 Then stat_folders += 1 : xml_from_folder_Click_sub(x, menuElem, files(0)) Else stat_folders_skipped += 1
                Else
                    stat_folders += 1
                    Dim tmp As Boolean = xml_from_folder_options_fillCRC
                    xml_from_folder_options_fillCRC = False
                    xml_from_folder_Click_sub(x, menuElem, folder + ".tmp")
                    xml_from_folder_options_fillCRC = tmp
                End If
            Next
        End If

        If xml_from_folder_options_mergeWithExisting <> "" AndAlso FileSystem.FileExists(xml_from_folder_options_mergeWithExisting) Then
            Dim x_orig As New Xml.XmlDocument
            x_orig.Load(xml_from_folder_options_mergeWithExisting)
            Dim nodeOrigRoot As Xml.XmlNode = x_orig.SelectSingleNode("/menu")
            stat_merged_orig = x_orig.SelectNodes("/menu/game").Count

            For Each node As Xml.XmlNode In x.SelectNodes("/menu/game")
                Dim romName As String = node.Attributes.GetNamedItem("name").Value
                If x_orig.SelectSingleNode("/menu/game[@name=""" + romName + """]") Is Nothing Then nodeOrigRoot.AppendChild(x_orig.ImportNode(node, True))
            Next
            stat_merged_new = x_orig.SelectNodes("/menu/game").Count

            Dim w As Xml.XmlWriter = Xml.XmlWriter.Create(xml_from_folder_options_dst, New Xml.XmlWriterSettings With {.Indent = True, .NewLineHandling = Xml.NewLineHandling.None})
            x_orig.Save(w) : w.Close()

            Dim msg As String = "Database from selected rom folder created." + vbCrLf
            msg = msg + stat_files.ToString + " files and " + stat_folders.ToString + " folders added, " + stat_folders_skipped.ToString + " folders skipped" + vbCrLf
            msg = msg + stat_merged_orig.ToString + " entries was in original XML, and there are " + stat_merged_new.ToString + " entries now."
            MsgBox(msg)
        Else
            Dim w As Xml.XmlWriter = Xml.XmlWriter.Create(xml_from_folder_options_dst, New Xml.XmlWriterSettings With {.Indent = True, .NewLineHandling = Xml.NewLineHandling.None})
            x.Save(w) : w.Close()

            Dim msg As String = "Database from selected rom folder created." + vbCrLf
            msg = msg + stat_files.ToString + " files and " + stat_folders.ToString + " folders added, " + stat_folders_skipped.ToString + " folders skipped"
            MsgBox(msg)
        End If
    End Sub
    'xmlFromFolder SUB (add file)
    Private Sub xml_from_folder_Click_sub(x As Xml.XmlDocument, menuElem As Xml.XmlElement, file As String)
        Dim elemCrc As Xml.XmlElement = x.CreateElement("crc")
        Dim elemManufacturer As Xml.XmlElement = x.CreateElement("manufacturer")
        Dim elemYear As Xml.XmlElement = x.CreateElement("year")
        Dim elemGenre As Xml.XmlElement = x.CreateElement("genre")

        If xml_from_folder_options_fillCRC Then
            Dim crc As String = Class6_hash.GetCRC32(file)
            Dim lead As String = ""
            If crc.Length < 8 Then lead = New String("0"c, 8 - crc.Length)
            elemCrc.InnerText = lead + crc
        End If

        file = file.Substring(file.LastIndexOf("\") + 1)
        file = file.Substring(0, file.LastIndexOf("."))
        Dim gameElem As Xml.XmlElement = x.CreateElement("game")
        Dim name As String = file
        If xml_from_folder_options_removeParanthesis Then name = streap_brackets_accurate(name)
        gameElem.SetAttribute("name", name)
        Dim elemDescription As Xml.XmlElement = x.CreateElement("description")
        elemDescription.InnerText = name

        Dim val As String = ""
        val = getAttr(file, 0)
        If val <> "" Then elemYear.InnerText = val

        val = ""
        val = getAttr(file, 2)
        If val <> "" Then elemManufacturer.InnerText = val

        gameElem.AppendChild(elemDescription)
        gameElem.AppendChild(elemCrc)
        gameElem.AppendChild(elemManufacturer)
        gameElem.AppendChild(elemYear)
        gameElem.AppendChild(elemGenre)
        menuElem.AppendChild(gameElem)
    End Sub

    'Remove clones
    Private Sub xml_remove_clones_Click() Handles xml_remove_clones.Click
        If Form1.ComboBox1.SelectedIndex < 0 Then
            MsgBox("Please, select a system.")
            Exit Sub
        End If

        Dim xmlPath As String = Form1.xmlPath
        Dim backupN As Integer = 0
        If Microsoft.VisualBasic.FileIO.FileSystem.FileExists(xmlPath + ".backup.xml") Then
            Do While Microsoft.VisualBasic.FileIO.FileSystem.FileExists(xmlPath + ".backup" + backupN.ToString + ".xml")
                backupN += 1
            Loop
            Microsoft.VisualBasic.FileIO.FileSystem.CopyFile(xmlPath, xmlPath + ".backup" + backupN.ToString + ".xml")
        Else
            Microsoft.VisualBasic.FileIO.FileSystem.CopyFile(xmlPath, xmlPath + ".backup.xml")
        End If

        Dim tmpNode As Xml.XmlNode
        Dim x As New Xml.XmlDocument
        x.Load(xmlPath)
        For Each node As Xml.XmlNode In x.SelectNodes("/menu/game")
            tmpNode = node.SelectSingleNode("cloneof")
            If tmpNode.InnerText.Trim <> "" Then
                node.ParentNode.RemoveChild(node)
            End If
        Next
        Dim w As Xml.XmlWriter = Xml.XmlWriter.Create(xmlPath, New Xml.XmlWriterSettings With {.Indent = True, .NewLineHandling = Xml.NewLineHandling.None})
        x.Save(w) : w.Close()
        MsgBox("Done! Please, recheck To see changes.")
    End Sub

    'Move roms to subfolder
    Private Sub move_roms_from_txt_to_subfolder_Click() Handles move_roms_from_txt_to_subfolder.Click
        Dim f As New OpenFileDialog
        f.Title = "Select rom list"
        f.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        f.ShowDialog()
        If f.FileName = "" Then Exit Sub
        If Not FileSystem.FileExists(f.FileName) Then MsgBox("Can't open file") : Exit Sub
        Dim fWoExt As String = f.FileName.Substring(f.FileName.LastIndexOf("\") + 1)
        fWoExt = fWoExt.Substring(0, fWoExt.LastIndexOf("."))

        Dim fd As New FolderBrowserDialog
        fd.Description = "Select rom folder"
        fd.ShowDialog()
        If fd.SelectedPath = "" Then Exit Sub

        Dim s As String
        Dim p() As String
        Dim curFile As String
        FileOpen(1, f.FileName, OpenMode.Input)
        Do While Not EOF(1)
            s = LineInput(1)
            If s.Contains("=") Then s = s.Substring(0, s.IndexOf("="))
            p = {s + ".*"}
            For Each file As String In FileSystem.GetFiles(fd.SelectedPath, SearchOption.SearchTopLevelOnly, p)
                Try
                    curFile = file.Substring(file.LastIndexOf("\") + 1)
                    FileSystem.MoveFile(file, fd.SelectedPath + "\" + fWoExt + "\" + curFile)
                Catch ex As Exception
                    MsgBox("Can't move '" + file + "' to '" + fd.SelectedPath + "\" + fWoExt + "\" + vbCrLf + ex.Message)
                End Try
            Next
        Loop
        FileClose(1)
    End Sub

    'DIFF
    Private Sub xml_compare_Click() Handles xml_compare.Click
        If Form1.CompareAgainstCurrentDBToolStripMenuItem.Checked And Form1.ComboBox1.SelectedIndex < 0 Then MsgBox("In this mode you have to choose system first.") : Exit Sub

        Dim file1 As String = "", file2 As String = ""

        Dim f As New OpenFileDialog
        f.Title = "Select old xml"
        f.Filter = "Hyperspin Database file (*.xml)|*.xml|All Files (*.*)|*.*"
        f.ShowDialog()
        If f.FileName = "" Then Exit Sub
        If Not FileSystem.FileExists(f.FileName) Then MsgBox("Can't open file") : Exit Sub
        file1 = f.FileName

        If Form1.CompareAgainstCurrentDBToolStripMenuItem.Checked Then
            file2 = Form1.xmlPath
        Else
            Dim f2 As New OpenFileDialog
            f2.Title = "Select second xml to compare to"
            f2.Filter = "Hyperspin Database file (*.xml)|*.xml|All Files (*.*)|*.*"
            f2.ShowDialog()
            If f2.FileName = "" Then Exit Sub
            If Not FileSystem.FileExists(f.FileName) Then MsgBox("Can't open file") : Exit Sub
            file2 = f2.FileName
        End If

        Dim tmpNode As Xml.XmlNode
        Dim x1 As New Xml.XmlDocument, x2 As New Xml.XmlDocument
        Dim arr1 As New List(Of String), arr2 As New List(Of String)
        x1.Load(file1) : x2.Load(file2)
        For Each node As Xml.XmlNode In x1.SelectNodes("/menu/game")
            If Form1.CompareGameromNameToolStripMenuItem.Checked Then
                arr1.Add(streap_brackets(node.Attributes.GetNamedItem("name").Value))
            Else
                tmpNode = node.SelectSingleNode("description")
                arr1.Add(streap_brackets(tmpNode.InnerText))
            End If
        Next
        For Each node As Xml.XmlNode In x2.SelectNodes("/menu/game")
            If Form1.CompareGameromNameToolStripMenuItem.Checked Then
                arr2.Add(streap_brackets(node.Attributes.GetNamedItem("name").Value))
            Else
                tmpNode = node.SelectSingleNode("description")
                arr2.Add(streap_brackets(tmpNode.InnerText))
            End If
        Next

        Dim counter As Integer = 0
        Dim added As New List(Of String), deleted As New List(Of String)
        FileOpen(1, ".\diff.txt", OpenMode.Output)
        PrintLine(1, "Old xml: " + file1)
        PrintLine(1, "(" + arr1.Count.ToString + " games)")
        PrintLine(1, "New xml: " + file2)
        PrintLine(1, "(" + arr2.Count.ToString + " games)")
        PrintLine(1, "")
        PrintLine(1, "New in " + file2)
        PrintLine(1, "------------------------------------------------------------------------------")
        For Each v As String In arr2
            If Not arr1.Contains(v) Then PrintLine(1, v) : counter += 1 'added.Add(v)
        Next
        PrintLine(1, "------------(Total: " + counter.ToString + " games.)")

        counter = 0
        PrintLine(1, "")
        PrintLine(1, "Removed in " + file2)
        PrintLine(1, "------------------------------------------------------------------------------")
        For Each v As String In arr1
            If Not arr2.Contains(v) Then PrintLine(1, v) : counter += 1 'deleted.Add(v)
        Next
        PrintLine(1, "------------(Total: " + counter.ToString + " games.)")
        FileClose(1)
        System.Diagnostics.Process.Start("notepad.exe", ".\diff.txt")
    End Sub

    'xml repare
    Public Sub xml_repare(ByVal xmlPath As String)
        'Dim m_xmlr As Xml.XmlTextReader
        'm_xmlr = New Xml.XmlTextReader(xmlPath)
        'm_xmlr.WhitespaceHandling = Xml.WhitespaceHandling.None
        'While Not m_xmlr.EOF
        'm_xmlr.Read()
        'End While
        Dim s As String
        FileOpen(1, xmlPath, OpenMode.Input)
        FileOpen(2, xmlPath + ".tmp", OpenMode.Output)
        Do While Not EOF(1)
            s = LineInput(1)
            If s.Contains("&") Then
                ''''TUT VES' KOD
                Dim n = s.IndexOf("&", 0)
                Do While n >= 0
                    Dim sb As New System.Text.StringBuilder(s)
                    If s.Length > n + 5 Then
                        If s.Substring(n + 1, 5).Contains("&") Then s = sb.Insert(n + 1, "amp;").ToString : n = s.IndexOf("&", n + 1) : Continue Do
                        If s.Substring(n + 1, 5).Contains(";") Then n = s.IndexOf("&", n + 1) : Continue Do
                    End If
                    s = sb.Insert(n + 1, "amp;").ToString : n = s.IndexOf("&", n + 1)
                Loop
                '''''END TUT VES' KOD
            End If
            PrintLine(2, s)
        Loop
        FileClose(1)
        FileClose(2)

        Dim backupN As Integer = 0
        If FileSystem.FileExists(xmlPath + ".backup.xml") Then
            Do While Microsoft.VisualBasic.FileIO.FileSystem.FileExists(xmlPath + ".backup" + backupN.ToString + ".xml")
                backupN += 1
            Loop
            FileSystem.CopyFile(xmlPath, xmlPath + ".backup" + backupN.ToString + ".xml")
        Else
            FileSystem.CopyFile(xmlPath, xmlPath + ".backup.xml")
        End If
        FileSystem.MoveFile(xmlPath + ".tmp", xmlPath, True)
    End Sub

    'Delete selected roms from DB / Update DB
    Private Sub update_delete_rom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles update_delete_rom.Click
        If Form1.ComboBox1.SelectedIndex < 0 Then Exit Sub

        Dim romName As String = ""
        Dim xmlPath As String = Form1.xmlPath
        Dim x As New Xml.XmlDocument

        If main_table.RowCount = 0 Then
            Dim res As MsgBoxResult = MsgBox("The check-table is empty. Updating database in current state will just remove all games from xml. Are you sure you want to do it?", MsgBoxStyle.YesNo)
            If res = MsgBoxResult.No Then Exit Sub
            'TODO remove all entries from database
        End If

        'do backup
        Dim backupN As Integer = 0
        If FileSystem.FileExists(xmlPath + ".backupBeforeEdit.xml") Then
            Do While Microsoft.VisualBasic.FileIO.FileSystem.FileExists(xmlPath + ".backupBeforeEdit" + backupN.ToString + ".xml")
                backupN += 1
            Loop
            FileSystem.CopyFile(xmlPath, xmlPath + ".backupBeforeEdit" + backupN.ToString + ".xml")
        Else
            FileSystem.CopyFile(xmlPath, xmlPath + ".backupBeforeEdit.xml")
        End If

        x.Load(xmlPath)

        'Deleting
        Dim romToDel As String = ""
        For Each r As DataGridViewRow In Form1.editor_delete_command_list
            If Form1.editor_insert_command_list.Contains(r) Then Continue For

            If Form1.editor_update_command_list.ContainsKey(r) Then
                romToDel = Form1.editor_update_command_list(r).ToLower
            Else
                romToDel = r.Cells(1).Value.ToString.ToLower
            End If

            For Each node As Xml.XmlNode In x.SelectNodes("/menu/game")
                romName = node.Attributes.GetNamedItem("name").Value
                If romName.ToLower = romToDel Then
                    node.ParentNode.RemoveChild(node)
                End If
            Next
        Next

        'Updating romnames
        For Each r As DataGridViewRow In Form1.editor_update_command_list.Keys
            If Form1.editor_delete_command_list.Contains(r) Then Continue For
            If Form1.editor_insert_command_list.Contains(r) Then Continue For
            Dim oldRomName As String = Form1.editor_update_command_list(r)
            Dim newRomName As String = r.Cells(1).Value.ToString

            Dim node As Xml.XmlNode = x.SelectSingleNode("/menu/game[@name=""" + oldRomName + """]")
            node.Attributes.GetNamedItem("name").Value = newRomName
        Next

        'Insert
        Dim tmpnode As Xml.XmlNode
        For Each r As DataGridViewRow In Form1.editor_insert_command_list
            If Form1.editor_delete_command_list.Contains(r) Then Continue For

            Dim newNode As Xml.XmlNode = x.SelectSingleNode("menu").AppendChild(x.CreateNode(Xml.XmlNodeType.Element, "game", Nothing))
            newNode.Attributes.Append(x.CreateAttribute("name"))
            newNode.Attributes.GetNamedItem("name").Value = r.Cells(1).Value.ToString

            tmpnode = newNode.AppendChild(x.CreateNode(Xml.XmlNodeType.Element, "description", Nothing))
            If r.Cells(0).Value Is Nothing Then tmpnode.InnerText = "" Else tmpnode.InnerText = r.Cells(0).Value.ToString
            tmpnode = newNode.AppendChild(x.CreateNode(Xml.XmlNodeType.Element, "crc", Nothing))
            If r.Cells(11).Value Is Nothing Then tmpnode.InnerText = "" Else tmpnode.InnerText = r.Cells(11).Value.ToString
            tmpnode = newNode.AppendChild(x.CreateNode(Xml.XmlNodeType.Element, "manufacturer", Nothing))
            If r.Cells(12).Value Is Nothing Then tmpnode.InnerText = "" Else tmpnode.InnerText = r.Cells(12).Value.ToString
            tmpnode = newNode.AppendChild(x.CreateNode(Xml.XmlNodeType.Element, "year", Nothing))
            If r.Cells(13).Value Is Nothing Then tmpnode.InnerText = "" Else tmpnode.InnerText = r.Cells(13).Value.ToString
            tmpnode = newNode.AppendChild(x.CreateNode(Xml.XmlNodeType.Element, "genre", Nothing))
            If r.Cells(14).Value Is Nothing Then tmpnode.InnerText = "" Else tmpnode.InnerText = r.Cells(14).Value.ToString
        Next

        'Updating all other roms
        For Each r As Windows.Forms.DataGridViewRow In Form1.DataGridView1.Rows
            If Form1.editor_delete_command_list.Contains(r) Then Continue For
            If Form1.editor_insert_command_list.Contains(r) Then Continue For
            If Form1.editor_update_command_list.ContainsKey(r) Then Continue For

            Dim node As Xml.XmlNode = x.SelectSingleNode("/menu/game[@name=""" + r.Cells(1).Value.ToString + """]")
            tmpnode = node.SelectSingleNode("description")
            tmpnode.InnerText = r.Cells(0).Value.ToString
            tmpnode = node.SelectSingleNode("crc")
            If tmpnode Is Nothing Then tmpnode = node.AppendChild(x.CreateNode(Xml.XmlNodeType.Element, "crc", Nothing))
            tmpnode.InnerText = r.Cells(11).Value.ToString
            tmpnode = node.SelectSingleNode("manufacturer")
            If tmpnode Is Nothing Then tmpnode = node.AppendChild(x.CreateNode(Xml.XmlNodeType.Element, "manufacturer", Nothing))
            tmpnode.InnerText = r.Cells(12).Value.ToString
            tmpnode = node.SelectSingleNode("year")
            If tmpnode Is Nothing Then tmpnode = node.AppendChild(x.CreateNode(Xml.XmlNodeType.Element, "year", Nothing))
            tmpnode.InnerText = r.Cells(13).Value.ToString
            tmpnode = node.SelectSingleNode("genre")
            If tmpnode Is Nothing Then tmpnode = node.AppendChild(x.CreateNode(Xml.XmlNodeType.Element, "genre", Nothing))
            tmpnode.InnerText = r.Cells(14).Value.ToString
        Next

        Dim w As Xml.XmlWriter = Xml.XmlWriter.Create(xmlPath, New Xml.XmlWriterSettings With {.Indent = True, .NewLineHandling = Xml.NewLineHandling.None})

        If Form1.NoReorderinsertedLinesAddedToTheEndToolStripMenuItem.Checked Then
            'no reorder
            x.Save(w)
        ElseIf Form1.ReorderAlphabetycallyToolStripMenuItem.Checked Then
            'reorder alphabetically by romname
            Dim ordered = (From el In x.SelectNodes("/menu/game") Let name = DirectCast(el, Xml.XmlNode).Attributes.GetNamedItem("name").Value.ToString Order By name Ascending)
            Dim x_reordered As New Xml.XmlDocument
            Dim x_reordered_menuNode As Xml.XmlElement = x_reordered.CreateElement("menu")
            x_reordered.AppendChild(x_reordered_menuNode)

            For Each el In ordered
                x_reordered_menuNode.AppendChild(x_reordered.ImportNode(DirectCast(el.el, Xml.XmlElement), True))
            Next
            x_reordered.Save(w)
        ElseIf Form1.ReorderAlphabeticallyByDescriptionToolStripMenuItem.Checked Then
            'reorder alphabetically by description
            Dim ordered = (From el In x.SelectNodes("/menu/game") Let name = DirectCast(el, Xml.XmlNode).SelectSingleNode("description").InnerText Order By name Descending)
            Dim x_reordered As New Xml.XmlDocument
            Dim x_reordered_menuNode As Xml.XmlElement = x_reordered.CreateElement("menu")
            x_reordered.AppendChild(x_reordered_menuNode)

            For Each el In ordered
                x_reordered_menuNode.AppendChild(x_reordered.ImportNode(DirectCast(el.el, Xml.XmlElement), True))
            Next
            x_reordered.Save(w)
        ElseIf Form1.ReorderAsSeenInTheCheckTableToolStripMenuItem.Checked Then
            'reorder as seen in table
            Dim x_reordered As New Xml.XmlDocument
            Dim x_reordered_menuNode As Xml.XmlElement = x_reordered.CreateElement("menu")
            x_reordered.AppendChild(x_reordered_menuNode)
            For Each r As DataGridViewRow In Form1.DataGridView1.Rows
                Dim tmpNodes As Xml.XmlNodeList = x.SelectNodes("/menu/game[@name='" + r.Cells(1).Value.ToString + "']")
                If tmpNodes.Count = 0 Then
                    MsgBox("Error. Aborting update.") : Exit For
                ElseIf tmpNodes.Count = 1 Then
                    x_reordered_menuNode.AppendChild(x_reordered.ImportNode(tmpNodes(0), True))
                    x.SelectSingleNode("/menu").RemoveChild(tmpNodes(0))
                ElseIf tmpNodes.Count > 1 Then
                    For Each tmpNode2 As Xml.XmlNode In tmpNodes
                        If tmpNode2.SelectSingleNode("description").InnerText = r.Cells(0).Value.ToString Then
                            x_reordered_menuNode.AppendChild(x_reordered.ImportNode(tmpNode2, True))
                            x.SelectSingleNode("/menu").RemoveChild(tmpNode2)
                        End If
                    Next
                End If
            Next
            x_reordered.Save(w)
        End If
        w.Close()

        Form1.editor_delete_command_list.Clear()
        Form1.editor_update_command_list.Clear()
        Form1.editor_insert_command_list.Clear()
        MsgBox("Database Updated.")
    End Sub
    Private Sub update_delete_rom_mm_Click(sender As System.Object, e As System.EventArgs) Handles update_delete_rom_mm.Click
        update_delete_rom_Click(update_delete_rom, New System.EventArgs)
    End Sub

    'Delete selected roms from DB / Update DB !!!OLD CODE!!!
    Private Sub update_delete_rom_OLD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'IF DELETE ROM
        Dim xmlPath As String = Form1.xmlPath
        'If Form1.CheckBox17.Checked = False Then
        If Form1.DataGridView1.SelectedRows.Count = 0 Then MsgBox("You have to select a rom to delete") : Exit Sub

        If Form1.CheckBox16.Checked Then
            If MsgBox("Remove '" + Form1.DataGridView1.SelectedRows(0).Cells(1).Value.ToString + "' from hyperspin database?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then Exit Sub
        End If
        If Not Microsoft.VisualBasic.FileIO.FileSystem.FileExists(xmlPath + ".backupBeforeDelete.xml") Then
            Microsoft.VisualBasic.FileIO.FileSystem.CopyFile(xmlPath, xmlPath + ".backupBeforeDelete.xml")
        End If

        Dim romName As String
        Dim romToDel As String = Form1.DataGridView1.SelectedRows(0).Cells(1).Value.ToString.ToLower
        Dim x As New Xml.XmlDocument
        x.Load(xmlPath)

        For Each node As Xml.XmlNode In x.SelectNodes("/menu/game")
            romName = node.Attributes.GetNamedItem("name").Value
            If romName.ToLower = romToDel Then
                node.ParentNode.RemoveChild(node)
            End If
        Next
        Dim w As Xml.XmlWriter = Xml.XmlWriter.Create(xmlPath, New Xml.XmlWriterSettings With {.Indent = True, .NewLineHandling = Xml.NewLineHandling.None})
        x.Save(w) : w.Close()

        Form1.DataGridView1.Rows.Remove(Form1.DataGridView1.SelectedRows(0))
        If Form1.Label2.Text.Contains(" of ") Then
            Dim s As String = Form1.Label2.Text
            Dim c1, c2 As Integer
            c1 = CInt(s.Substring(9, s.IndexOf(" of ") - 9)) - 1
            c2 = CInt(s.Substring(s.IndexOf(" of ") + 4)) - 1
            Form1.Label2.Text = "Missing: " + c1.ToString + " of " + c2.ToString
        Else
            Form1.Label2.Text = "Total: " + Form1.DataGridView1.Rows.Count.ToString
        End If

        Dim i As Integer = Class1.romlist.IndexOf(romToDel)
        Class1.data.RemoveAt(i)
        Class1.data_crc.RemoveAt(i)
        Class1.romlist.RemoveAt(i)
        If Class1.romFoundlist.Contains(romToDel) Then
            Class1.romFoundlist.Remove(romToDel)
        End If
        'End If

        'IF UPDDATE DB
        'If Form1.CheckBox17.Checked = True Then
        If Form1.DataGridView1.Rows.Count = 0 Then MsgBox("The table contains no data, please, use 'CHECK' button.") : Exit Sub

        'do backup
        Dim backupN As Integer = 0
        If FileSystem.FileExists(xmlPath + ".backupBeforeEdit.xml") Then
            Do While Microsoft.VisualBasic.FileIO.FileSystem.FileExists(xmlPath + ".backupBeforeEdit" + backupN.ToString + ".xml")
                backupN += 1
            Loop
            FileSystem.CopyFile(xmlPath, xmlPath + ".backupBeforeEdit" + backupN.ToString + ".xml")
        Else
            FileSystem.CopyFile(xmlPath, xmlPath + ".backupBeforeEdit.xml")
        End If

        Dim tmpNode As Xml.XmlNode
        'Dim x As New Xml.XmlDocument
        x.Load(xmlPath)
        For Each r As Windows.Forms.DataGridViewRow In Form1.DataGridView1.Rows
            Dim node As Xml.XmlNode = x.SelectSingleNode("/menu/game[@name=""" + r.Cells(1).Value.ToString + """]")
            tmpNode = node.SelectSingleNode("description")
            tmpNode.InnerText = r.Cells(0).Value.ToString
            tmpNode = node.SelectSingleNode("crc")
            If tmpNode Is Nothing Then tmpNode = node.AppendChild(x.CreateNode(Xml.XmlNodeType.Element, "crc", Nothing))
            tmpNode.InnerText = r.Cells(11).Value.ToString
            tmpNode = node.SelectSingleNode("manufacturer")
            If tmpNode Is Nothing Then tmpNode = node.AppendChild(x.CreateNode(Xml.XmlNodeType.Element, "manufacturer", Nothing))
            tmpNode.InnerText = r.Cells(12).Value.ToString
            tmpNode = node.SelectSingleNode("year")
            If tmpNode Is Nothing Then tmpNode = node.AppendChild(x.CreateNode(Xml.XmlNodeType.Element, "year", Nothing))
            tmpNode.InnerText = r.Cells(13).Value.ToString
            tmpNode = node.SelectSingleNode("genre")
            If tmpNode Is Nothing Then tmpNode = node.AppendChild(x.CreateNode(Xml.XmlNodeType.Element, "genre", Nothing))
            tmpNode.InnerText = r.Cells(14).Value.ToString
        Next
        'Dim w As Xml.XmlWriter = Xml.XmlWriter.Create(xmlPath, New Xml.XmlWriterSettings With {.Indent = True, .NewLineHandling = Xml.NewLineHandling.None})
        x.Save(w) : w.Close()
        MsgBox("Database Updated.")
        'End If
    End Sub

    Private Function getAttr(ByVal s As String, ByVal n As Integer) As String
        Dim c(5) As Integer
        For i As Integer = 0 To 5
            c(i) = CInt(xml_from_folder_options_adv.Substring(i * 2, 2))
        Next

        Dim val As String = ""
        If c(1) = n Then
            If c(0) = 1 Then val = getInParenthesis(s, 1, False)
            If c(0) = 2 Then val = getInParenthesis(s, 1, True)
        ElseIf c(3) = n Then
            If c(2) = 1 Then val = getInParenthesis(s, 2, False)
            If c(2) = 2 Then val = getInParenthesis(s, 2, True)
        ElseIf c(5) = n Then
            If c(4) = 1 Then val = getInParenthesis(s, 3, False)
            If c(4) = 2 Then val = getInParenthesis(s, 3, True)
        End If
        Return val
    End Function

    Private Function getInParenthesis(ByVal s As String, ByVal n As Integer, Optional ByVal brackets As Boolean = False) As String
        Dim c1, c2 As Char
        If brackets Then c1 = "["c : c2 = "]"c Else c1 = "("c : c2 = ")"c

        Dim ret As String
        Dim i As Integer = s.IndexOf(c1) : If i < 0 Then Return ""
        Dim i1 As Integer = s.IndexOf(c2, i) : If i1 < 0 Then Return ""
        ret = s.Substring(i + 1, i1 - i - 1)
        If n = 1 Then Return ret

        i = s.IndexOf(c1, i1) : If i < 0 Then Return ""
        i1 = s.IndexOf(c2, i) : If i1 < 0 Then Return ""
        ret = s.Substring(i + 1, i1 - i - 1)
        If n = 2 Then Return ret

        i = s.IndexOf(c1, i1) : If i < 0 Then Return ""
        i1 = s.IndexOf(c2, i) : If i1 < 0 Then Return ""
        ret = s.Substring(i + 1, i1 - i - 1)
        Return ret
    End Function

    Public Shared Function streap_brackets(s As String) As String
        'If s.IndexOf("[") >= 0 Then
        's = s.Substring(0, s.IndexOf("["))
        'ElseIf s.IndexOf(".") >= 0 Then
        's = s.Substring(0, s.LastIndexOf("."))
        'End If

        'If s.IndexOf("(") >= 0 Then
        's = s.Substring(0, s.IndexOf("("))
        'ElseIf s.IndexOf(".") >= 0 Then
        's = s.Substring(0, s.LastIndexOf("."))
        'End If
        If s.IndexOf("[") >= 0 Or s.IndexOf("(") >= 0 Then
            If s.IndexOf("[") >= 0 Then s = s.Substring(0, s.IndexOf("["))
            If s.IndexOf("(") >= 0 Then s = s.Substring(0, s.IndexOf("("))
        End If
        Return s.ToUpper.Trim
    End Function

    Public Shared Function streap_brackets_accurate(s As String) As String
        'Dim r As New System.Text.RegularExpressions.Regex("\( (?> [^()]+ | \( (?<Depth>) | \) (?<-Depth>) )* (?(Depth)(?!)) \)", System.Text.RegularExpressions.RegexOptions.IgnorePatternWhitespace)
        's = r.Replace(s, "").Replace("  ", " ").Replace("  ", " ").Replace("  ", " ").Replace("  ", " ")
        'r = New System.Text.RegularExpressions.Regex("\[ (?> [^\[\]]+ | \[ (?<Depth>) | \] (?<-Depth>) )* (?(Depth)(?!)) \]", System.Text.RegularExpressions.RegexOptions.IgnorePatternWhitespace)
        's = r.Replace(s, "").Replace("  ", " ").Replace("  ", " ").Replace("  ", " ").Replace("  ", " ")

        '@"\(              # Match an opening parenthesis.
        '(?>             # Then either match (possessively):
        '[^()]+         #  any characters except parentheses
        '|               # or
        '\( (?<Depth>)  #  an opening paren (and increase the parens counter)
        '|               # or
        '\) (?<-Depth>) #  a closing paren (and decrease the parens counter).
        ')*              # Repeat as needed.
        '(?(Depth)(?!))   # Assert that the parens counter is at zero.
        '\)               # Then match a closing parenthesis.",

        Dim lastIndex As Integer = 0
        Dim ind1 As Integer = 0, ind2 As Integer = 0
        Dim cnt As String = ""
        Dim exc As Boolean = False
        Do While s.IndexOf("(", lastIndex) > 0
            exc = False
            ind1 = s.IndexOf("(", lastIndex)
            ind2 = s.IndexOf(")", ind1)
            If ind2 > 0 And Not ind2 = s.Length - 1 Then
                If xml_from_folder_options_removeParanthesis_exceptions.Trim <> "" Then
                    cnt = s.Substring(ind1 + 1, ind2 - ind1 - 1)
                    For Each ex As String In xml_from_folder_options_removeParanthesis_exceptions.Split({"|"}, StringSplitOptions.RemoveEmptyEntries)
                        If cnt.Trim.ToUpper.Contains(ex.Trim.ToUpper) Then exc = True : Exit For
                    Next
                End If

                If Not exc Then s = s.Substring(0, ind1) + s.Substring(ind2 + 1) Else lastIndex = ind1 + 1
            Else
                'Not pair or ')' is a last symbol in string
                s = s.Substring(0, ind1).Trim
                Exit Do
            End If
            s = s.Replace("  ", " ").Replace("  ", " ").Replace("  ", " ").Trim
        Loop

        lastIndex = 0
        Do While s.IndexOf("[", lastIndex) > 0
            exc = False
            ind1 = s.IndexOf("[", lastIndex)
            ind2 = s.IndexOf("]", ind1)
            If ind2 > 0 And Not ind2 = s.Length - 1 Then
                If xml_from_folder_options_removeParanthesis_exceptions.Trim <> "" Then
                    cnt = s.Substring(ind1 + 1, ind2 - ind1 - 1)
                    For Each ex As String In xml_from_folder_options_removeParanthesis_exceptions.Split({"|"}, StringSplitOptions.RemoveEmptyEntries)
                        If cnt.Trim.ToUpper.Contains(ex.Trim.ToUpper) Then exc = True : Exit For
                    Next
                End If

                If Not exc Then s = s.Substring(0, ind1) + s.Substring(ind2 + 1) Else lastIndex = ind1 + 1
            Else
                'Not pair or ']' is a last symbol in string
                s = s.Substring(0, ind1).Trim
                Exit Do
            End If
            s = s.Replace("  ", " ").Replace("  ", " ").Replace("  ", " ").Trim
        Loop

        Return s.Trim
    End Function

    'Press Insert or Delete in main table
    Public Sub main_table_key_down(sender As Object, e As KeyEventArgs) Handles main_table.KeyDown
        If Not Form1.AlowEditToolStripMenuItem.Checked Then Exit Sub

        If e.KeyCode = Keys.Delete Then
            If Not main_table.IsCurrentCellInEditMode Then
                If main_table.SelectedRows.Count > 0 Then
                    For Each r As DataGridViewRow In main_table.SelectedRows
                        Form1.editor_delete_command_list.Add(r)
                        main_table.Rows.Remove(r)
                    Next
                End If
            End If
        End If

        If e.KeyCode = Keys.Insert Then
            Dim i As Integer = 0
            If Not main_table.IsCurrentCellInEditMode Then
                If main_table.SelectedRows.Count = 0 Then
                    main_table.Rows.Add()
                    i = main_table.Rows.Count - 1
                    main_table.Rows(i).Selected = True
                    main_table.CurrentCell = main_table.Rows(i).Cells("col1")
                ElseIf main_table.SelectedRows.Count = 1 Then
                    i = main_table.SelectedRows(0).Index
                    main_table.Rows.Insert(i, 1)
                    main_table.Rows(i).Selected = True
                    main_table.CurrentCell = main_table.Rows(i).Cells("col1")
                End If
                Form1.editor_insert_command_list.Add(main_table.Rows(i))
            End If
        End If
    End Sub

    'Begin editing romname
    Public Sub main_table_begin_edit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles main_table.CellBeginEdit
        If e.ColumnIndex = 1 Then
            If Not Form1.editor_insert_command_list.Contains(main_table.Rows(e.RowIndex)) Then
                If Not Form1.editor_update_command_list.ContainsKey(main_table.Rows(e.RowIndex)) Then
                    Form1.editor_update_command_list.Add(main_table.Rows(e.RowIndex), main_table.Rows(e.RowIndex).Cells(1).Value.ToString)
                End If
            End If
        End If
    End Sub
End Class
