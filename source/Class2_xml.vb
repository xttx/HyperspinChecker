Imports Microsoft.VisualBasic.FileIO
Public Class Class2_xml
    Private WithEvents xml_crop As Button = Form1.Button18
    Private WithEvents xml_from_folder As Button = Form1.Button19
    Private WithEvents xml_remove_clones As Button = Form1.Button28
    Private WithEvents xml_compare As Button = Form1.Button31
    Private WithEvents move_roms_from_txt_to_subfolder As Button = Form1.Button29
    Private WithEvents update_delete_rom As Button = Form1.Button21
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
                If Not Form3.filter.Contains(romName.ToLower) Then
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
        Dim fd As New FolderBrowserDialog
        fd.Description = "Select folder to create XML from."
        fd.ShowDialog()
        If fd.SelectedPath = "" Then Exit Sub
        Dim x As New Xml.XmlDocument
        Dim menuElem As Xml.XmlElement = x.CreateElement("menu") : x.AppendChild(menuElem)

        For Each file As String In Microsoft.VisualBasic.FileIO.FileSystem.GetFiles(fd.SelectedPath)
            Dim elemCrc As Xml.XmlElement = x.CreateElement("crc")
            Dim elemManufacturer As Xml.XmlElement = x.CreateElement("manufacturer")
            Dim elemYear As Xml.XmlElement = x.CreateElement("year")
            Dim elemGenre As Xml.XmlElement = x.CreateElement("genre")

            If Form1.CheckBox21.Checked Then
                Dim crc As String = GetCRC32(file)
                Dim lead As String = ""
                If crc.Length < 8 Then lead = New String("0"c, 8 - crc.Length)
                elemCrc.InnerText = lead + crc
            End If

            file = file.Substring(file.LastIndexOf("\") + 1)
            file = file.Substring(0, file.LastIndexOf("."))
            Dim gameElem As Xml.XmlElement = x.CreateElement("game")
            gameElem.SetAttribute("name", file)
            Dim elemDescription As Xml.XmlElement = x.CreateElement("description")
            elemDescription.InnerText = file

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
        Next

        Dim fs As New SaveFileDialog
        fs.Title = "Select folder to put XML to."
        fs.Filter = "XML Files | *.xml"
        fs.RestoreDirectory = True
        fs.InitialDirectory = Class1.HyperspinPath + "\Databases"
        fs.ShowDialog()
        If fs.FileName = "" Then Exit Sub
        Dim w As Xml.XmlWriter = Xml.XmlWriter.Create(fs.FileName, New Xml.XmlWriterSettings With {.Indent = True, .NewLineHandling = Xml.NewLineHandling.None})
        x.Save(w) : w.Close()
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
        MsgBox("Done! Please, recheck to see changes.")
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

    Private Sub xml_compare_Click() Handles xml_compare.Click
        If Form1.RadioButton12.Checked And Form1.ComboBox1.SelectedIndex < 0 Then MsgBox("In this mode you have to choose system first.") : Exit Sub

        Dim file1 As String = "", file2 As String = ""

        Dim f As New OpenFileDialog
        f.Title = "Select old xml"
        f.Filter = "Hyperspin Database file (*.xml)|*.xml|All Files (*.*)|*.*"
        f.ShowDialog()
        If f.FileName = "" Then Exit Sub
        If Not FileSystem.FileExists(f.FileName) Then MsgBox("Can't open file") : Exit Sub
        file1 = f.FileName

        If Form1.RadioButton12.Checked Then
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
            If Form1.RadioButton14.Checked Then
                arr1.Add(streap_brackets(node.Attributes.GetNamedItem("name").Value))
            Else
                tmpNode = node.SelectSingleNode("description")
                arr1.Add(streap_brackets(tmpNode.InnerText))
            End If
        Next
        For Each node As Xml.XmlNode In x2.SelectNodes("/menu/game")
            If Form1.RadioButton14.Checked Then
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
        Class1.romlist.Remove(i)
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
        Dim val As String = ""
        If Form1.ComboBox13.SelectedIndex = n Then
            If Form1.ComboBox12.SelectedIndex = 1 Then val = getInParenthesis(s, 1, False)
            If Form1.ComboBox12.SelectedIndex = 2 Then val = getInParenthesis(s, 1, True)
        ElseIf Form1.ComboBox15.SelectedIndex = n Then
            If Form1.ComboBox14.SelectedIndex = 1 Then val = getInParenthesis(s, 2, False)
            If Form1.ComboBox14.SelectedIndex = 2 Then val = getInParenthesis(s, 2, True)
        ElseIf Form1.ComboBox17.SelectedIndex = n Then
            If Form1.ComboBox16.SelectedIndex = 1 Then val = getInParenthesis(s, 3, False)
            If Form1.ComboBox16.SelectedIndex = 2 Then val = getInParenthesis(s, 3, True)
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

    'Get CRC32
    Public Shared Function GetCRC32(ByVal sFileName As String) As String
        Try
            Dim FS As IO.FileStream = New IO.FileStream(sFileName, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read, 8192)
            Dim CRC32Result As Integer = &HFFFFFFFF
            Dim Buffer(4096) As Byte
            Dim ReadSize As Integer = 4096
            Dim Count As Integer = FS.Read(Buffer, 0, ReadSize)
            Dim CRC32Table(256) As Integer
            Dim DWPolynomial As Integer = &HEDB88320
            Dim DWCRC As Long
            Dim i As Integer, j As Integer, n As Integer

            'Create CRC32 Table
            For i = 0 To 255
                DWCRC = i
                For j = 8 To 1 Step -1
                    If CBool(DWCRC And 1) Then
                        DWCRC = ((DWCRC And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
                        DWCRC = DWCRC Xor DWPolynomial
                    Else
                        DWCRC = ((DWCRC And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
                    End If
                Next j
                CRC32Table(i) = CInt(DWCRC)
            Next i

            'Calcualting CRC32 Hash
            Do While (Count > 0)
                For i = 0 To Count - 1
                    n = (CRC32Result And &HFF) Xor Buffer(i)
                    CRC32Result = ((CRC32Result And &HFFFFFF00) \ &H100) And &HFFFFFF
                    CRC32Result = CRC32Result Xor CRC32Table(n)
                Next i
                Count = FS.Read(Buffer, 0, ReadSize)
            Loop
            FS.Close()
            Return Hex(Not (CRC32Result))
        Catch ex As Exception
            Return ""
        End Try
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
