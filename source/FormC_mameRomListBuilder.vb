Imports System.IO
Imports System.Threading
Imports Microsoft.VisualBasic.FileIO.FileSystem

Public Class FormC_mameRomListBuilder
    Public Const undef As String = "<undefined>"
    Public Const str_mechanical As String = "Mechanical"
    Public Const str_mechanical_not As String = "Non mechanical"
    Public Const str_parent As String = "Parent"
    Public Const str_clone As String = "Clone"

    Dim output As String = ""
    Dim machines As List(Of machine)
    Dim machines_choosen As New List(Of machine)
    Dim machines_choosen_folders As New List(Of machine)

    Dim props(10) As List(Of String)
    Dim props_obj(10) As List(Of item)
    Dim progress As Long = -1
    Dim progress_fLenght As Double = -1

    Dim n As Integer = 0
    Dim state As Integer = 0

    Dim ini As New IniFileApi()
    Dim lastopened_exe As String = ""
    Dim lastopened_xml As String = ""
    Dim lastopened_rom As String = ""
    Dim lastopened_dst As String = ""
    Dim lastopened_fld As String = ""

    Public Structure machine
        Dim name As String
        Dim description As String
        Dim cloneof As String
        Dim driver As String
        Dim manufacturer As String
        Dim year As String
        Dim bios As String
        Dim device As String
        Dim mechanical As Boolean
        Dim savetates As Boolean
        Dim display_type As String
        Dim resolution As String
        Dim rotate As Integer
        Dim cpu_clock As ULong
        Dim is_device As Boolean

        Dim status As status
        Dim emulation As status
        Dim emulation2 As String
        Dim graphic As status
        Dim color As status
        Dim sound As status

        Dim fileSize As Long
    End Structure
    Public Enum status
        preliminary
        incorrect
        good
        bad
        imperfect
    End Enum

    'Form load
    Private Sub FormC_mameRomListBuilder_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Language.localize(Me)
        ini.IniFile(Class1.confPath)
        TextBox1.Text = ini.IniReadValue("TOOL_mame_rom_builder", "MAME_exe_path")
        TextBox2.Text = ini.IniReadValue("TOOL_mame_rom_builder", "MAME_xml_path")
        TextBox3.Text = ini.IniReadValue("TOOL_mame_rom_builder", "MAME_romset_path")
        TextBox4.Text = ini.IniReadValue("TOOL_mame_rom_builder", "MAME_rom_dest_path")
        lastopened_exe = ini.IniReadValue("TOOL_mame_rom_builder", "MAME_last_opened_path_exe")
        lastopened_xml = ini.IniReadValue("TOOL_mame_rom_builder", "MAME_last_opened_path_xml")
        lastopened_rom = ini.IniReadValue("TOOL_mame_rom_builder", "MAME_last_opened_path_rom")
        lastopened_dst = ini.IniReadValue("TOOL_mame_rom_builder", "MAME_last_opened_path_dst")
        lastopened_fld = ini.IniReadValue("TOOL_mame_rom_builder", "MAME_last_opened_path_fld")
        If ini.IniReadValue("TOOL_mame_rom_builder", "generate_from_exe") = "0" Then RadioButton2.Checked = True
        If ini.IniReadValue("TOOL_mame_rom_builder", "split_romet") = "0" Then RadioButton4.Checked = True
    End Sub

    'Form closing - save config
    Private Sub FormC_mameRomListBuilder_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        ini.IniWriteValue("TOOL_mame_rom_builder", "MAME_exe_path", TextBox1.Text)
        ini.IniWriteValue("TOOL_mame_rom_builder", "MAME_xml_path", TextBox2.Text)
        ini.IniWriteValue("TOOL_mame_rom_builder", "MAME_romset_path", TextBox3.Text)
        ini.IniWriteValue("TOOL_mame_rom_builder", "MAME_rom_dest_path", TextBox4.Text)
        ini.IniWriteValue("TOOL_mame_rom_builder", "MAME_last_opened_path_exe", lastopened_exe)
        ini.IniWriteValue("TOOL_mame_rom_builder", "MAME_last_opened_path_xml", lastopened_xml)
        ini.IniWriteValue("TOOL_mame_rom_builder", "MAME_last_opened_path_rom", lastopened_rom)
        ini.IniWriteValue("TOOL_mame_rom_builder", "MAME_last_opened_path_dst", lastopened_dst)
        ini.IniWriteValue("TOOL_mame_rom_builder", "MAME_last_opened_path_fld", lastopened_fld)

        Dim tmp As Integer
        tmp = DirectCast(IIf(RadioButton1.Checked, 1, 0), Integer)
        ini.IniWriteValue("TOOL_mame_rom_builder", "generate_from_exe", tmp.ToString)
        tmp = DirectCast(IIf(RadioButton3.Checked, 1, 0), Integer)
        ini.IniWriteValue("TOOL_mame_rom_builder", "split_romet", tmp.ToString)
    End Sub

    'Form resize
    Private Sub FormC_mameRomListBuilder_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        Dim w = GroupBox2.Width - 18 - (6 * 4) - 13
        Dim lw = CInt(w / 5)
        ListBox1.Width = lw
        ListBox2.Width = lw
        Button4.Width = lw
        Button5.Width = lw

        ListBox3.Width = lw
        ListBox7.Width = lw
        ListBox3.Left = ListBox1.Left + ListBox1.Width + 6
        ListBox7.Left = ListBox1.Left + ListBox1.Width + 6
        Label4.Left = ListBox3.Left
        Label8.Left = ListBox3.Left
        Button6.Width = lw
        Button10.Width = lw
        Button6.Left = ListBox3.Left
        Button10.Left = ListBox3.Left

        ListBox4.Width = lw
        ListBox8.Width = lw
        ListBox4.Left = ListBox3.Left + ListBox3.Width + 6
        ListBox8.Left = ListBox3.Left + ListBox3.Width + 6
        Label5.Left = ListBox4.Left
        Label9.Left = ListBox4.Left
        Button7.Width = lw
        Button11.Width = lw
        Button7.Left = ListBox4.Left
        Button11.Left = ListBox4.Left

        ListBox5.Width = lw
        ListBox9.Width = lw
        ListBox5.Left = ListBox4.Left + ListBox4.Width + 6
        ListBox9.Left = ListBox4.Left + ListBox4.Width + 6
        Label6.Left = ListBox5.Left
        Label10.Left = ListBox5.Left
        Button8.Width = lw
        Button12.Width = lw
        Button8.Left = ListBox5.Left
        Button12.Left = ListBox5.Left

        ListBox6.Width = lw
        ListBox10.Width = lw
        ListBox6.Left = ListBox5.Left + ListBox5.Width + 6
        ListBox10.Left = ListBox5.Left + ListBox5.Width + 6
        Label7.Left = ListBox6.Left
        Label11.Left = ListBox6.Left
        Button9.Width = lw
        Button13.Width = lw
        Button9.Left = ListBox6.Left
        Button13.Left = ListBox6.Left
    End Sub

#Region "Main form"
    '... mame.exe path
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim fb As New OpenFileDialog
        fb.InitialDirectory = lastopened_exe
        fb.Filter = "MAME executable (*.exe)|*.exe"
        fb.ShowDialog()
        lastopened_exe = fb.FileName.Substring(0, fb.FileName.LastIndexOf("\") + 1)
        TextBox1.Text = fb.FileName
    End Sub

    '... listxml path
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim fb As New OpenFileDialog
        fb.InitialDirectory = lastopened_xml
        fb.Filter = "ListXML, generated with 'mame.exe -listxml' command (*.*)|*.*"
        fb.ShowDialog()
        lastopened_xml = fb.FileName.Substring(0, fb.FileName.LastIndexOf("\") + 1)
        TextBox2.Text = fb.FileName
    End Sub

    'Analyze press
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        state = 0
        progress = -1
        progress_fLenght = -1
        Timer1.Enabled = True
        If RadioButton1.Checked Then
            If Not FileExists(TextBox1.Text) Then MsgBox(TextBox1.Text + " does not exist.") : Exit Sub
            If Not TextBox1.Text.ToUpper.EndsWith("EXE") Then MsgBox(TextBox1.Text + " is not an executable.") : Exit Sub
            BackgroundWorker1.RunWorkerAsync()
            'Label1.Text = "Status: generate listXML"
        Else
            parse_listXML()
        End If
    End Sub

    'Generate listxml
    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim p As New ProcessStartInfo
        p.FileName = TextBox1.Text
        p.Arguments = "-listxml"
        p.WindowStyle = ProcessWindowStyle.Hidden
        p.CreateNoWindow = True
        p.UseShellExecute = False
        p.RedirectStandardOutput = True
        Dim pr As Process = Process.Start(p)

        output = pr.StandardOutput.ReadToEnd
        pr.WaitForExit()
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        parse_listXML()
    End Sub

    'Parse xml and fill listboxes
    Private Sub parse_listXML()
        state = 1
        BackgroundWorker2.RunWorkerAsync()
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Dim xml As New Xml.XmlDocument
        If RadioButton1.Checked Then
            Try
                xml.LoadXml(output)
            Catch ex As Exception
                MsgBox(ex.Message)
                Exit Sub
            End Try
        Else
            If Not FileExists(TextBox2.Text) Then MsgBox(TextBox2.Text + " does not exist.") : Exit Sub
            Try
                ''''''''''''''TEST
                'xml.Load(TextBox2.Text)
                progress_fLenght = Math.Round(FileLen(TextBox2.Text) / 1024 / 1024, 2)
                Dim stRead As StreamWrapper = New StreamWrapper(New FileStream(TextBox2.Text, FileMode.Open, FileAccess.Read))
                AddHandler stRead.ReadProgress, New StreamProgressHandler(AddressOf stRead_ReadProgress)
                xml.Load(New Xml.XmlTextReader(stRead))
                ''''''''''''''''''''''''
            Catch ex As Exception
                MsgBox(ex.Message)
                Exit Sub
            End Try
        End If
        progress = -1

        Dim nodes_machines As Xml.XmlNodeList = xml.SelectNodes("/mame/machine")

        'Old listXML format
        If nodes_machines.Count = 0 Then
            If xml.SelectNodes("/mame/game").Count > 1 Then
                MsgBox("Old mame xml format, before it was merged with software list. Not handled yet, sorry.") : Exit Sub
            End If
            MsgBox("No roms found. Possibly incorrect XML.") : Exit Sub
        End If

        'Retrieving machines
        state = 2
        machines = New List(Of machine)

        Dim tmp As String = ""

        For i As Integer = 0 To 10 : props(i) = New List(Of String) : Next
        For i As Integer = 0 To 10 : props_obj(i) = New List(Of item) : Next

        Dim cur_val As String
        Dim cur_item As item
        Dim ind As Integer = -1

        progress = 0
        progress_fLenght = nodes_machines.Count
        For Each x As Xml.XmlNode In nodes_machines
            'Parsing xml node and adding to rom list
            progress += 1
            If x.SelectSingleNode("rom") Is Nothing Then Continue For

            Dim cur_machine As New machine
            cur_machine.name = getNodeAttr(x, "name")
            cur_machine.description = getNodeValue(x, "description")

            If getNodeAttr(x, "isdevice").Trim.ToLower = "yes" Then
                cur_machine.is_device = True
                'machines.Add(cur_machine)
                'Continue For
            Else
                cur_machine.is_device = False
            End If

            cur_machine.year = getNodeValue(x, "year")
            cur_machine.manufacturer = getNodeValue(x, "manufacturer")
            cur_machine.driver = getNodeAttr(x, "sourcefile")
            cur_machine.cloneof = getNodeAttr(x, "cloneof")
            cur_machine.bios = getNodeAttr(x, "romof")
            cur_machine.device = getNodeAttr(x.SelectSingleNode("device"), "type")

            tmp = getNodeAttr(x.SelectSingleNode("chip[@tag='maincpu']"), "clock")
            If tmp <> "" Then
                cur_machine.cpu_clock = CULng(tmp)
            Else
                tmp = getNodeAttr(x.SelectSingleNode("chip[@tag='master']"), "clock")
                If tmp <> "" Then
                    cur_machine.cpu_clock = CULng(tmp)
                Else
                    cur_machine.cpu_clock = 0
                End If
            End If

            Dim display As Xml.XmlNode = x.SelectSingleNode("display")
            If display Is Nothing Then
                cur_machine.display_type = "no display"
                cur_machine.resolution = "no display"
                cur_machine.rotate = -1
            Else
                cur_machine.display_type = getNodeAttr(display, "type")
                cur_machine.resolution = getNodeAttr(display, "width") + "x" + getNodeAttr(display, "height")
                cur_machine.rotate = CInt(getNodeAttr(display, "rotate"))
            End If

            If getNodeAttr(x, "ismechanical").Trim.ToLower = "yes" Then
                cur_machine.mechanical = True
            Else
                cur_machine.mechanical = False
            End If

            Dim drv As Xml.XmlNode = x.SelectSingleNode("driver")
            If getNodeAttr(drv, "savestate").Trim.ToLower = "supported" Then
                cur_machine.savetates = True
            Else
                cur_machine.savetates = False
            End If
            cur_machine.emulation2 = getNodeAttr(drv, "emulation")

            machines.Add(cur_machine)

            'Fill listboxes
            For i As Integer = 1 To 10
                Select Case i
                    Case 1
                        cur_val = cur_machine.driver
                    Case 2
                        If cur_machine.cloneof = "" Then cur_val = str_parent Else cur_val = str_clone
                    Case 3
                        cur_val = cur_machine.manufacturer
                    Case 4
                        cur_val = cur_machine.year
                    Case 5
                        cur_val = cur_machine.bios
                    Case 6
                        cur_val = cur_machine.device
                    Case 7
                        If cur_machine.mechanical Then cur_val = str_mechanical Else cur_val = str_mechanical_not
                    Case 8
                        cur_val = cur_machine.cpu_clock.ToString
                    Case 9
                        cur_val = cur_machine.rotate.ToString
                    Case 10
                        cur_val = cur_machine.emulation2
                    Case Else
                        cur_val = ""
                End Select

                cur_item = New item With {.name = cur_val}
                ind = props(i).IndexOf(cur_val)
                If ind < 0 Then
                    props(i).Add(cur_val)
                    props_obj(i).Add(cur_item)
                Else
                    props_obj(i).Item(ind).count += 1
                End If
            Next
        Next
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Timer1.Enabled = False
        If props_obj(0) Is Nothing Then Label1.Text = "Status: Idle" : Exit Sub
        Label1.Text = "Filling lists..."
        Me.Refresh() : Label1.Refresh()

        Dim count As Integer = 0
        Static _comparerStd As System.Collections.Generic.IComparer(Of item) = New ComparerStd()
        Static _comparerNum As System.Collections.Generic.IComparer(Of item) = New ComparerNumber()
        Dim lists() As ListBox = {ListBox1, ListBox2, ListBox3, ListBox4, ListBox5, ListBox6, ListBox7, ListBox8, ListBox9, ListBox10}
        For i As Integer = 0 To 9
            Dim _list As New List(Of item)

            count = 0
            'lists(i).Items.Clear()
            For Each item As item In props_obj(i + 1)
                _list.Add(item)
                count += item.count
            Next
            _list.Add(New item With {.count = count, .name = "ALL"})
            If i = 7 Or i = 8 Then _list.Sort(_comparerNum) Else _list.Sort(_comparerStd)
            lists(i).DataSource = Nothing : lists(i).DataSource = _list
        Next

        Button14.Enabled = True
        GroupBox2.Enabled = True
        Label1.Text = "Done. Total games: " + machines.Count.ToString
    End Sub

    'GO button - Show file operation panel
    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Dim item As item

        'Creating lists of selected properties
        Dim list() As List(Of String) =
            {New List(Of String), New List(Of String), New List(Of String),
             New List(Of String), New List(Of String), New List(Of String),
             New List(Of String), New List(Of String), New List(Of String),
             New List(Of String)}

        Dim listboxes() As ListBox = {ListBox1, ListBox3, ListBox4,
                                      ListBox5, ListBox6, ListBox2,
                                      ListBox7, ListBox8, ListBox9,
                                      ListBox10}

        For i As Integer = 0 To listboxes.Count - 1
            For Each lItem In listboxes(i).SelectedItems
                item = DirectCast(lItem, item)
                If item.name.ToUpper = "ALL" Then list(i).Clear() : Exit For
                list(i).Add(item.name.ToUpper)
            Next
        Next

        Dim tmp As String = ""
        Dim skip As Boolean = False
        machines_choosen = New List(Of machine)
        For Each m As machine In machines
            'Driver
            If list(0).Count > 0 Then
                If Not list(0).Contains(m.driver.ToUpper) Then Continue For
            End If

            'Manufacturer
            tmp = m.manufacturer.ToUpper
            If tmp = "" Then tmp = undef
            If list(1).Count > 0 Then
                If Not list(1).Contains(tmp.ToUpper) Then Continue For
            End If

            'Year
            tmp = m.year.ToUpper
            If tmp = "" Then tmp = undef
            If list(2).Count > 0 Then
                If Not list(2).Contains(tmp.ToUpper) Then Continue For
            End If

            'Bios
            tmp = m.bios.ToUpper
            If tmp = "" Then tmp = undef
            If list(3).Count > 0 Then
                If Not list(3).Contains(tmp.ToUpper) Then Continue For
            End If

            'Device
            tmp = m.device.ToUpper
            If tmp = "" Then tmp = undef
            If list(4).Count > 0 Then
                If Not list(4).Contains(tmp.ToUpper) Then Continue For
            End If

            'Parent / clone (list 2)
            If m.cloneof = "" Then tmp = str_parent Else tmp = str_clone
            If list(5).Count > 0 Then
                If Not list(5).Contains(tmp.ToUpper) Then Continue For
            End If

            'Mechanical (list 7)
            If m.mechanical Then tmp = str_mechanical Else tmp = str_mechanical_not
            If list(6).Count > 0 Then
                If Not list(6).Contains(tmp.ToUpper) Then Continue For
            End If

            'CPU clock (list 8)
            tmp = m.cpu_clock.ToString
            If list(7).Count > 0 Then
                If Not list(7).Contains(tmp.ToUpper) Then Continue For
            End If

            'Rotate (list 9)
            tmp = m.rotate.ToString
            If tmp = "-1" Then tmp = undef
            If list(8).Count > 0 Then
                If Not list(8).Contains(tmp.ToUpper) Then Continue For
            End If

            'Emulation status (list 10)
            tmp = m.emulation2.ToUpper
            If tmp = "" Then tmp = undef
            If list(9).Count > 0 Then
                If Not list(9).Contains(tmp.ToUpper) Then Continue For
            End If

            machines_choosen.Add(m)
        Next

        GroupBox1.Enabled = False
        GroupBox2.Enabled = False
        GroupBox3.Parent = Me
        GroupBox3.Top = CInt(Me.Height / 2 - GroupBox3.Height / 2)
        GroupBox3.Left = CInt(Me.Width / 2 - GroupBox3.Width / 2)
        GroupBox3.Visible = True
        Label15.Text = "Games selected: " + machines_choosen.Count.ToString
    End Sub

    'Timer - Refresh status label
    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        n = n + 1
        If n > 6 Then n = 1
        If state = 0 Then Label1.Text = "Status: generate listXML"
        If state = 1 Then Label1.Text = "Status: Loading XML"
        If state = 2 Then Label1.Text = "Parsing XML"

        If progress >= 0 And state = 1 Then Label1.Text += " ( " + Math.Round(progress / 1024 / 1024, 2).ToString + "mb / " + progress_fLenght.ToString + "mb )"
        If progress >= 0 And state = 2 Then Label1.Text += " ( " + progress.ToString + " / " + progress_fLenght.ToString + " )"

        Label1.Text += " " + Strings.Space(n).Replace(" ", ".")
    End Sub

    'Stream progress change event handler
    Private Sub stRead_ReadProgress(ByVal position As Long)
        progress = position
        Application.DoEvents()
    End Sub

    'Helper function to get xmlNode value
    Public Function getNodeValue(ByVal x As Xml.XmlNode, ByVal childNode As String) As String
        If x.SelectSingleNode(childNode) Is Nothing Then
            Return ""
        Else
            Return x.SelectSingleNode(childNode).InnerText
        End If
    End Function

    'Helper function to get xmlNode attribute
    Public Function getNodeAttr(ByVal x As Xml.XmlNode, ByVal attr As String) As String
        If x Is Nothing Then Return ""
        If x.Attributes.GetNamedItem(attr) Is Nothing Then
            Return ""
        Else
            Return x.Attributes.GetNamedItem(attr).Value
        End If
    End Function

    'Mode change - generate from mame.ex, or set existing xml list
    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged, RadioButton2.CheckedChanged
        If RadioButton1.Checked Then
            Button1.Enabled = True : Button2.Enabled = False
            TextBox1.Enabled = True : TextBox2.Enabled = False
        Else
            Button1.Enabled = False : Button2.Enabled = True
            TextBox1.Enabled = False : TextBox2.Enabled = True
        End If
    End Sub

    'Sort button
    Private Sub sortList(sender As Object, e As EventArgs) Handles Button4.Click, Button5.Click, Button6.Click, Button7.Click, Button8.Click, Button9.Click, Button10.Click, Button11.Click, Button12.Click, Button13.Click
        Dim b As Button = DirectCast(sender, Button)
        Dim n As Integer = 0
        If IsNumeric(Microsoft.VisualBasic.Right(b.Name, 2)) Then n = CInt(Microsoft.VisualBasic.Right(b.Name, 2)) Else n = CInt(Microsoft.VisualBasic.Right(b.Name, 1))

        Dim cmp As IComparer(Of WindowsApplication1.item) = New ComparerStd
        Select Case b.Text
            Case "Sorted: Alphabetically"
                b.Text = "Sorted: By Count"
                cmp = New ComparerCount
            Case "Sorted: By Count"
                b.Text = "Sorted: By Size"
                cmp = New ComparerSize()
            Case "Sorted: By Size"
                b.Text = "Sorted: Alphabetically"
                If n = 11 Or n = 12 Then cmp = New ComparerNumber Else cmp = New ComparerStd
        End Select

        Dim listboxes() As ListBox = {ListBox1, ListBox3, ListBox4, ListBox5, ListBox6,
                                      ListBox2, ListBox7, ListBox8, ListBox9, ListBox10}

        Select Case n
            Case 4
                n = 0
            Case 6
                n = 1
            Case 7
                n = 2
            Case 8
                n = 3
            Case 9
                n = 4
            Case 5
                n = 5
            Case 10
                n = 6
            Case 11
                n = 7
            Case 12
                n = 8
            Case 13
                n = 9
        End Select

        Dim topind = listboxes(n).TopIndex
        Dim selected = New List(Of String)
        For Each sel As item In listboxes(n).SelectedItems
            selected.Add(sel.name)
        Next

        Dim _list = DirectCast(listboxes(n).DataSource, List(Of item))
        If _list IsNot Nothing Then
            listboxes(n).DataSource = Nothing
            _list.Sort(cmp)
            listboxes(n).DataSource = _list
        End If

        'restore selection
        listboxes(n).SelectedIndex = -1
        For i As Integer = 0 To listboxes(n).Items.Count - 1
            If selected.Contains(DirectCast(listboxes(n).Items(i), item).name) Then listboxes(n).SetSelected(i, True)
        Next
        listboxes(n).TopIndex = topind
    End Sub
#End Region

#Region "File Operations Panel"
    Dim list_unknownFiles As New List(Of String)
    Dim list_foundGames As New List(Of String)
    Dim list_foundGamesSelected As New List(Of String)
    Dim treehash As New Dictionary(Of String, nodeAndList)
    Structure nodeAndList
        Dim node As TreeNode
        Dim list As List(Of String)
    End Structure

    'File operations - ANALYZE
    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        If Not DirectoryExists(TextBox3.Text) Then MsgBox("Romset directory does not exist.") : Exit Sub

        Dim dirInfo As New DirectoryInfo(TextBox3.Text)
        Dim filesInfo() As FileInfo = dirInfo.GetFiles

        Dim m_all As New List(Of String)
        Dim m_chosen As New List(Of String)
        Dim m_clones As New Dictionary(Of String, List(Of String))
        For Each m As machine In machines
            m_all.Add(m.name.ToUpper)

            'We fill clones hashtable only if merged set selected, 
            'otherwise we skip it to save time
            If RadioButton4.Checked Then
                If m.cloneof <> "" Then
                    Dim tmp As New List(Of String)
                    If m_clones.ContainsKey(m.cloneof.ToUpper) Then
                        tmp = m_clones(m.cloneof.ToUpper)
                    Else
                        m_clones.Add(m.cloneof.ToUpper, tmp)
                    End If
                    tmp.Add(m.name)
                End If
            End If
        Next
        For Each m As machine In machines_choosen
            m_chosen.Add(m.name.ToUpper)
        Next

        'remove from local m_choosen list all games deselected from treeview
        If treehash.Count > 0 Then
            Dim ind As Integer = 0
            For Each n As nodeAndList In treehash.Values
                If n.node.Checked Then Continue For
                For Each s As String In n.list
                    ind = m_chosen.IndexOf(s.ToUpper)
                    If ind >= 0 Then m_chosen.RemoveAt(ind)
                Next
            Next
        End If

        Dim fNameWOExt As String = ""
        Dim n_known_files As Integer = 0
        Dim n_known_games As Integer = 0
        Dim n_unKnown_files As Integer = 0
        Dim n_space_total As Long = 0
        Dim n_space_known_files As Long = 0
        Dim n_space_selected_games As Long = 0
        Dim n_games_not_found As Integer = 0
        Dim n_selected_found As Integer = 0
        list_unknownFiles = New List(Of String)
        list_foundGames = New List(Of String)
        list_foundGamesSelected = New List(Of String)

        For Each f As FileInfo In filesInfo
            fNameWOExt = f.Name.ToUpper.Replace(f.Extension.ToUpper, "")

            Dim ind = m_all.IndexOf(fNameWOExt)
            If ind >= 0 Then
                n_known_files += 1
                n_known_games += 1
                list_foundGames.Add(f.Name)
                Dim length = f.Length
                n_space_known_files += length
                Dim m = machines(ind) : m.fileSize = length : machines(ind) = m

                'If merged romset add clones to foundlist
                If RadioButton4.Checked Then
                    If m_clones.ContainsKey(fNameWOExt) Then
                        For Each s As String In m_clones(fNameWOExt)
                            n_known_games += 1
                            list_foundGames.Add(s + ".clone")

                            If m_chosen.Contains(s.ToUpper) Then
                                n_selected_found += 1
                                list_foundGamesSelected.Add(s + ".clone")
                            End If
                        Next
                    End If
                End If
            Else
                n_unKnown_files += 1
                list_unknownFiles.Add(f.Name)
            End If

            If m_chosen.Contains(fNameWOExt) Then
                n_selected_found += 1
                n_space_selected_games += f.Length
                list_foundGamesSelected.Add(f.Name)
            End If
            n_space_total += f.Length
        Next

        fillListSizes()

        n_games_not_found = m_all.Count - n_known_games
        Label15.Text = "Games selected: " + m_chosen.Count.ToString
        Label14.Text = "Total files: " + filesInfo.Count.ToString
        Label16.Text = "Known files: " + n_known_files.ToString
        Label29.Text = "Known games: " + n_known_games.ToString
        Label17.Text = "Unknown files: " + n_unKnown_files.ToString
        Label18.Text = "Total space used: " + Math.Round(n_space_total / 1024 / 1024, 1).ToString + "mb"
        Label19.Text = "Total space used by known files: " + Math.Round(n_space_known_files / 1024 / 1024, 1).ToString + "mb"
        Label20.Text = "Total space used by selected games: " + Math.Round(n_space_selected_games / 1024 / 1024, 1).ToString + "mb"
        Label26.Text = "Total space used by not selected games: " + Math.Round((n_space_known_files - n_space_selected_games) / 1024 / 1024, 1).ToString + "mb"
        Label21.Text = "Number of found games: " + n_known_games.ToString
        Label22.Text = "Number of found selected games: " + n_selected_found.ToString
        Label27.Text = "Number of not found games: " + n_games_not_found.ToString
        Label28.Text = "Number of not found selected games: " + (m_chosen.Count - n_selected_found).ToString
    End Sub
    Private Sub fillListSizes()
        Static _comparerStd As System.Collections.Generic.IComparer(Of item) = New ComparerStd()
        Static _comparerNum As System.Collections.Generic.IComparer(Of item) = New ComparerNumber()

        Dim sizes(9) As Dictionary(Of String, Long)
        For i As Integer = 0 To 9
            sizes(i) = New Dictionary(Of String, Long)
            For Each item In props_obj(i + 1)
                sizes(i).Add(item.name, 0)
            Next
        Next

        For Each m As machine In machines
            sizes(0)(m.driver) += m.fileSize
            If m.cloneof = "" Then sizes(1)("Parent") += m.fileSize Else sizes(1)("Clone") += m.fileSize
            If m.manufacturer = "" Then sizes(2)("<undefined>") += m.fileSize Else sizes(2)(m.manufacturer) += m.fileSize
            If m.year = "" Then sizes(3)("<undefined>") += m.fileSize Else sizes(3)(m.year) += m.fileSize
            If m.bios = "" Then sizes(4)("<undefined>") += m.fileSize Else sizes(4)(m.bios) += m.fileSize
            If m.device = "" Then sizes(5)("<undefined>") += m.fileSize Else sizes(5)(m.device) += m.fileSize
            If m.mechanical Then sizes(6)("Mechanical") += m.fileSize Else sizes(6)("Non mechanical") += m.fileSize
            sizes(7)(m.cpu_clock.ToString) += m.fileSize
            If m.rotate = -1 Then sizes(8)("<undefined>") += m.fileSize Else sizes(8)(m.rotate.ToString) += m.fileSize
            If m.emulation2 = "" Then sizes(9)("<undefined>") += m.fileSize Else sizes(9)(m.emulation2) += m.fileSize
        Next

        Dim count As Integer = 0
        Dim topind As Integer = 0
        Dim lists() As ListBox = {ListBox1, ListBox2, ListBox3, ListBox4, ListBox5, ListBox6, ListBox7, ListBox8, ListBox9, ListBox10}
        Dim selected As New List(Of String)
        For n = 0 To 9
            topind = lists(n).TopIndex
            selected = New List(Of String)
            For Each sel As item In lists(n).SelectedItems
                selected.Add(sel.name)
            Next

            count = 0
            Dim _list As New List(Of item)
            For Each i As item In props_obj(n + 1)
                i.size = sizes(n)(i.name)
                _list.Add(i)
                count += i.count
            Next
            _list.Add(New item With {.count = count, .name = "ALL"})
            If n = 7 Or n = 8 Then _list.Sort(_comparerNum) Else _list.Sort(_comparerStd)
            lists(n).DataSource = Nothing : lists(n).DataSource = _list

            'restore selection
            lists(n).SelectedIndex = -1
            For i As Integer = 0 To lists(n).Items.Count - 1
                If selected.Contains(DirectCast(lists(n).Items(i), item).name) Then lists(n).SetSelected(i, True)
            Next
            lists(n).TopIndex = topind
        Next
    End Sub

    'File operations - Label - UnKnown files click
    Private Sub Label23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label23.Click
        FileOpen(1, ".\tmp.txt", OpenMode.Output)
        For Each s As String In list_unknownFiles
            PrintLine(1, s)
        Next
        PrintLine(1, "Total: " + list_unknownFiles.Count.ToString)
        FileClose(1)
        System.Diagnostics.Process.Start(".\tmp.txt")
    End Sub

    'File operations - Label - Not found games click
    Private Sub Label24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label24.Click
        Dim m_all As New List(Of String)
        For Each m As machine In machines
            m_all.Add(m.name.ToUpper)
        Next

        Dim ind As Integer = 0
        For Each i As String In list_foundGames
            If i.Contains(".") Then i = i.Substring(0, i.LastIndexOf("."))
            ind = m_all.IndexOf(i.ToUpper)
            If ind >= 0 Then
                m_all.RemoveAt(ind)
            Else
                MsgBox("Strange behaviour with game '" + i + "'" + vbCrLf + "It often means, you have both GAME.7z and GAME_CLONE.7z in merged set, in other words, your set is not proper.")
            End If
        Next

        FileOpen(1, ".\tmp.txt", OpenMode.Output)
        For Each s As String In m_all
            PrintLine(1, s)
        Next
        PrintLine(1, "Total: " + m_all.Count.ToString)
        FileClose(1)
        System.Diagnostics.Process.Start(".\tmp.txt")
    End Sub

    'File operations - Label - Not found selected games click
    Private Sub Label25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label25.Click
        Dim m_chosen As New List(Of String)
        For Each m As machine In machines_choosen
            m_chosen.Add(m.name.ToUpper)
        Next

        Dim ind As Integer = 0
        For Each i As String In list_foundGamesSelected
            If i.Contains(".") Then i = i.Substring(0, i.LastIndexOf("."))
            ind = m_chosen.IndexOf(i.ToUpper)
            If ind >= 0 Then
                m_chosen.RemoveAt(ind)
            Else
                MsgBox("Strange behaviour with game '" + i + "'")
            End If
        Next

        FileOpen(1, ".\tmp.txt", OpenMode.Output)
        For Each s As String In m_chosen
            PrintLine(1, s)
        Next
        PrintLine(1, "Total: " + m_chosen.Count.ToString)
        FileClose(1)
        System.Diagnostics.Process.Start(".\tmp.txt")
    End Sub

    'File operations - Treeview - Add mame folder
    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        Dim fb As New OpenFileDialog
        fb.InitialDirectory = lastopened_fld
        fb.Filter = "Mame folder files (*.ini)|*.ini"
        fb.ShowDialog()
        lastopened_fld = fb.FileName.Substring(0, fb.FileName.LastIndexOf("\") + 1)
        If Not FileExists(fb.FileName) Then Exit Sub

        treehash.Clear()
        TreeView1.Nodes.Clear()

        Dim ini As New IniFile()
        ini.Load(fb.FileName)
        Dim nodeList As New nodeAndList
        Dim iniSections As New List(Of String)
        For Each s As IniFile.IniSection In ini.Sections
            If s.Name.ToUpper = "FOLDER_SETTINGS" Then Continue For
            If s.Name.ToUpper = "ROOT_FOLDER" Then Continue For
            iniSections.Add(s.Name)
        Next
        iniSections.Sort()

        Dim t As TreeNode
        Dim ind As Integer
        Dim missing As New List(Of String)
        Dim tmpList As List(Of String)
        Dim count As String = ""
        For Each m As machine In machines
            missing.Add(m.name.ToUpper)
        Next
        For Each s As String In iniSections
            tmpList = New List(Of String)
            For Each k As IniFile.IniSection.IniKey In ini.GetSection(s).Keys
                tmpList.Add(k.Name)
                ind = missing.IndexOf(k.Name.ToUpper)
                If ind >= 0 Then missing.RemoveAt(ind)
            Next
            If tmpList.Count = 0 Then count = "" Else count = " (" + tmpList.Count.ToString + ")"

            If Not s.Contains("/") Then
                t = New TreeNode With {.Name = s, .Text = s + count, .Checked = True}
                nodeList = New nodeAndList With {.node = t, .list = tmpList}
                treehash.Add(s, nodeList)
                TreeView1.Nodes.Add(t)
            Else
                Dim part1 As String = s.Substring(0, s.IndexOf("/")).Trim
                Dim part2 As String = s.Substring(s.IndexOf("/") + 1).Trim

                If Not treehash.Keys.Contains(part1) Then
                    t = New TreeNode With {.Name = part1, .Text = part1 + count, .Checked = True}
                    nodeList = New nodeAndList With {.node = t, .list = New List(Of String)}
                    treehash.Add(part1, nodeList)
                    TreeView1.Nodes.Add(t)
                End If

                t = New TreeNode With {.Name = part2, .Text = part2 + count, .Checked = True}
                nodeList = New nodeAndList With {.node = t, .list = tmpList}
                treehash.Add(s, nodeList)
                treehash(part1).node.Nodes.Add(t)
            End If
        Next

        'Adding missing games
        t = New TreeNode With {.Name = "zMissing", .Text = "zMissing (" + missing.Count.ToString + ")", .Checked = True}
        nodeList = New nodeAndList With {.node = t, .list = missing}
        treehash.Add("zMissing", nodeList)
        TreeView1.Nodes.Add(t)

        TreeView1.Sort()
    End Sub

    'File operations - Treeview - Clear
    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        treehash.Clear()
        TreeView1.Nodes.Clear()
    End Sub

    'File Operations - Treeview - Check / Uncheck all nodes
    Private Sub Button_check_all_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_check_all.Click, Button_uncheck_all.Click
        If TreeView1.Nodes.Count = 0 Then Exit Sub
        Dim name As String = DirectCast(sender, Button).Name.ToUpper
        Dim check As Boolean = True
        If name.Contains("UNCHECK") Then check = False
        CheckAllTreeNodes(TreeView1.Nodes(0), check)
    End Sub

    'File Operations - Treeview - Check / Uncheck all nodes - SUB
    Private Sub CheckAllTreeNodes(ByVal node As TreeNode, ByVal check As Boolean)
        node.Checked = check
        If node.Nodes.Count > 0 Then CheckAllTreeNodes(node.Nodes(0), check)

        Do While node.NextNode IsNot Nothing
            node = node.NextNode
            node.Checked = check
            If node.Nodes.Count > 0 Then CheckAllTreeNodes(node.Nodes(0), check)
        Loop
    End Sub

    'File Operation - Hide Panel
    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        GroupBox1.Enabled = True
        GroupBox2.Enabled = True
        GroupBox3.Visible = False
    End Sub

    '... rompath
    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        Dim fb As New FolderBrowserDialog
        fb.SelectedPath = lastopened_rom
        fb.ShowDialog()
        lastopened_rom = fb.SelectedPath
        TextBox3.Text = fb.SelectedPath
    End Sub

    '... destination path
    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        Dim fb As New FolderBrowserDialog
        fb.SelectedPath = lastopened_dst
        fb.ShowDialog()
        lastopened_dst = fb.SelectedPath
        TextBox4.Text = fb.SelectedPath
    End Sub

    'Copy selected
    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        If Not DirectoryExists(TextBox3.Text) Then MsgBox("Source (romset) directory not exist.") : Exit Sub
        If Not DirectoryExists(TextBox4.Text) Then MsgBox("Destination directory not exist.") : Exit Sub

        Dim m_chosen As New List(Of String)
        For Each m As machine In machines_choosen
            m_chosen.Add(m.name.ToUpper)
        Next

        'remove from local m_choosen list all games deselected from treeview
        If treehash.Count > 0 Then
            Dim ind As Integer = 0
            For Each n As nodeAndList In treehash.Values
                If n.node.Checked Then Continue For
                For Each s As String In n.list
                    ind = m_chosen.IndexOf(s.ToUpper)
                    If ind >= 0 Then m_chosen.RemoveAt(ind)
                Next
            Next
        End If

        For Each m In m_chosen
            Dim files = GetFiles(TextBox3.Text, FileIO.SearchOption.SearchTopLevelOnly, {m + ".*"})
            For Each file In files
                FileCopy(file, TextBox4.Text + "\" + IO.Path.GetFileName(file))
            Next
        Next
        MsgBox("Done.")
    End Sub
    'Move selected
    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        If Not DirectoryExists(TextBox3.Text) Then MsgBox("Source (romset) directory not exist.") : Exit Sub
        If Not DirectoryExists(TextBox4.Text) Then MsgBox("Destination directory not exist.") : Exit Sub

        Dim m_chosen As New List(Of String)
        For Each m As machine In machines_choosen
            m_chosen.Add(m.name.ToUpper)
        Next

        'remove from local m_choosen list all games deselected from treeview
        If treehash.Count > 0 Then
            Dim ind As Integer = 0
            For Each n As nodeAndList In treehash.Values
                If n.node.Checked Then Continue For
                For Each s As String In n.list
                    ind = m_chosen.IndexOf(s.ToUpper)
                    If ind >= 0 Then m_chosen.RemoveAt(ind)
                Next
            Next
        End If

        For Each m In m_chosen
            Dim files = GetFiles(TextBox3.Text, FileIO.SearchOption.SearchTopLevelOnly, {m + ".*"})
            For Each file In files
                MoveFile(file, TextBox4.Text + "\" + IO.Path.GetFileName(file))
            Next
        Next
        MsgBox("Done.")
    End Sub
    'Generate bat to delete
    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        If Not DirectoryExists(TextBox3.Text) Then MsgBox("Source (romset) directory not exist.") : Exit Sub

        Dim m_chosen As New List(Of String)
        For Each m As machine In machines_choosen
            m_chosen.Add(m.name.ToUpper)
        Next

        'remove from local m_choosen list all games deselected from treeview
        If treehash.Count > 0 Then
            Dim ind As Integer = 0
            For Each n As nodeAndList In treehash.Values
                If n.node.Checked Then Continue For
                For Each s As String In n.list
                    ind = m_chosen.IndexOf(s.ToUpper)
                    If ind >= 0 Then m_chosen.RemoveAt(ind)
                Next
            Next
        End If

        Dim del As New List(Of String)
        For Each m In m_chosen
            Dim files = GetFiles(TextBox3.Text, FileIO.SearchOption.SearchTopLevelOnly, {m + ".*"})
            For Each file In files
                del.Add(file)
            Next
        Next
        del.Sort()

        FileOpen(1, TextBox3.Text + "\!del.bat", OpenMode.Output)
        For Each f As String In del
            PrintLine(1, "del """ + f + """")
        Next
        FileClose(1)

        MsgBox("Done. File '!del.bat' saved to romset dir.")
    End Sub
#End Region
End Class

Public Class item
    Public name As String = ""
    Public tag As String = ""
    Public count As Integer = 1
    Public size As Long = -1

    Public Overrides Function ToString() As String
        'Return MyBase.ToString()
        If name = "" Or name = "-1" Then name = FormC_mameRomListBuilder.undef
        If size = -1 Then
            Return name + " (" + count.ToString + ")"
        Else
            Return name + " (" + count.ToString + ")(" + Math.Round(size / 1024 / 1024, 1).ToString + "mb)"
        End If
    End Function
End Class

Public Delegate Sub StreamProgressHandler(ByVal position As Long)
Public Class StreamWrapper
    Inherits Stream

    Private m_baseStream As Stream
    Private m_readPosition As Long
    Private m_WaitHandle As WaitHandle = Nothing
    Public Event ReadProgress As StreamProgressHandler

    Public Sub New(ByVal baseStream As Stream)
        m_baseStream = baseStream
        m_readPosition = 0
    End Sub

    Public Overrides Function Read(ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer) As Integer
        Dim len As Integer = m_baseStream.Read(buffer, offset, count)
        m_readPosition += len
        RaiseEvent ReadProgress(m_readPosition)
        Return len
    End Function

    Protected Overrides Function CreateWaitHandle() As System.Threading.WaitHandle
        If m_WaitHandle Is Nothing Then m_WaitHandle = New Mutex(False)
        Return m_WaitHandle
    End Function

    Public Overrides Function EndRead(ByVal asyncResult As System.IAsyncResult) As Integer
        Dim len As Integer = m_baseStream.EndRead(asyncResult)
        m_readPosition += len
        RaiseEvent ReadProgress(m_readPosition)
        Return len
    End Function

    Public Overrides Sub Write(ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer)
        m_baseStream.Write(buffer, offset, count)
    End Sub

    Public Overrides Function BeginRead(ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer, ByVal callback As System.AsyncCallback, ByVal state As Object) As System.IAsyncResult
        Return MyBase.BeginRead(buffer, offset, count, callback, state)
    End Function

    Public Overrides Function BeginWrite(ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer, ByVal callback As System.AsyncCallback, ByVal state As Object) As System.IAsyncResult
        Return MyBase.BeginWrite(buffer, offset, count, callback, state)
    End Function

    Public Overrides ReadOnly Property CanRead As Boolean
        Get
            Return m_baseStream.CanRead
        End Get
    End Property

    Public Overrides ReadOnly Property CanSeek As Boolean
        Get
            Return m_baseStream.CanSeek
        End Get
    End Property

    Public Overrides ReadOnly Property CanWrite As Boolean
        Get
            Return m_baseStream.CanWrite
        End Get
    End Property

    Public Overrides Sub Close()
        MyBase.Close()
    End Sub

    Public Overrides Sub Flush()
        m_baseStream.Flush()
    End Sub

    Public Overrides ReadOnly Property Length As Long
        Get
            Return m_baseStream.Length
        End Get
    End Property

    Public Overrides Property Position As Long
        Get
            Return m_baseStream.Position
        End Get
        Set(ByVal value As Long)
            m_baseStream.Position = value
        End Set
    End Property

    Public Overrides Function Seek(ByVal offset As Long, ByVal origin As System.IO.SeekOrigin) As Long
        Return m_baseStream.Seek(offset, origin)
    End Function

    Public Overrides Sub SetLength(ByVal value As Long)
        m_baseStream.SetLength(value)
    End Sub
End Class

Class ComparerStd
    Implements IComparer(Of WindowsApplication1.item)

    Public Function Compare(ByVal x As item, ByVal y As item) As Integer Implements System.Collections.Generic.IComparer(Of item).Compare
        If x.name.ToUpper = "ALL" And y.name.ToUpper = "ALL" Then Return 0
        If x.name.ToUpper = "ALL" Then Return -1
        If y.name.ToUpper = "ALL" Then Return 1
        Return x.name.CompareTo(y.name)
    End Function
End Class
Class ComparerNumber
    Implements IComparer(Of WindowsApplication1.item)

    Public Function Compare(ByVal x As item, ByVal y As item) As Integer Implements System.Collections.Generic.IComparer(Of item).Compare
        If x.name.ToUpper = "ALL" And y.name.ToUpper = "ALL" Then Return 0
        If x.name.ToUpper = "ALL" Then Return -1
        If y.name.ToUpper = "ALL" Then Return 1

        If IsNumeric(x.name) And IsNumeric(y.name) Then
            Dim a As Long = Convert.ToInt64(x.name)
            Dim b As Long = Convert.ToInt64(y.name)
            Return a.CompareTo(b)
        End If
        Return x.name.CompareTo(y.name)
    End Function
End Class
Class ComparerSize
    Implements IComparer(Of WindowsApplication1.item)

    Public Function Compare(ByVal x As item, ByVal y As item) As Integer Implements System.Collections.Generic.IComparer(Of item).Compare
        If x.name.ToUpper = "ALL" And y.name.ToUpper = "ALL" Then Return 0
        If x.name.ToUpper = "ALL" Then Return -1
        If y.name.ToUpper = "ALL" Then Return 1
        Return x.size.CompareTo(y.size)
    End Function
End Class
Class ComparerCount
    Implements IComparer(Of WindowsApplication1.item)

    Public Function Compare(ByVal x As item, ByVal y As item) As Integer Implements System.Collections.Generic.IComparer(Of item).Compare
        If x.name.ToUpper = "ALL" And y.name.ToUpper = "ALL" Then Return 0
        If x.name.ToUpper = "ALL" Then Return -1
        If y.name.ToUpper = "ALL" Then Return 1
        Return x.count.CompareTo(y.count)
    End Function
End Class