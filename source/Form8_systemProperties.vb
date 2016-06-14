Imports Microsoft.VisualBasic.FileIO.FileSystem

Public Class Form8_systemProperties
    Dim controls_to_hide() As String = {"BUTTON6", "BUTTON7", "BUTTON8", "LISTVIEW1", "PANEL1"}

    Dim refr As Boolean = True
    Dim sys As String = ""
    Dim sysNormalCase As String = ""
    Dim _data As List(Of String)
    Dim _modules As List(Of String)
    Dim _emulators As Dictionary(Of String, List(Of String))
    Dim _HS_ini As List(Of String)
    Dim _mainMenu As List(Of String)
    Dim HL_Path As String = ""
    Dim current_module As String = ""
    Public Shared Event paths_updated(system As String)

    'Data property set/get to interchange data with main form
    Public Property data As List(Of String)
        Get
            Return _data
        End Get
        Set(value As List(Of String))
            refr = True
            If Panel1.Visible Then Button8_Click(Button8, New EventArgs)

            _data = value
            If _data.Count = 0 Then Me.Enabled = False : Exit Property
            Me.Enabled = True
            If _data(0) = "" Then CheckBox1.Checked = False Else CheckBox1.Checked = True
            If _data(1) = "" Then Label5.BackColor = Form1.colorNO Else Label5.BackColor = Form1.colorYES
            If _data(2) = "" Then Label1.BackColor = Form1.colorNO Else Label1.BackColor = Form1.colorYES
            If _data(3) = "" Then Label2.BackColor = Form1.colorNO Else Label2.BackColor = Form1.colorYES
            If _data(4) = "" Then Label6.BackColor = Form1.colorNO Else Label6.BackColor = Form1.colorYES
            If _data(9) = "" Then Label3.BackColor = Form1.colorNO Else Label3.BackColor = Form1.colorYES
            If _data(10) = "" Then Label4.BackColor = Form1.colorNO Else Label4.BackColor = Form1.colorYES
            If _data(5) = "" Then RadioButton2.Checked = True Else RadioButton1.Checked = True
            If _data(6) = "" Then
                TextBox1.Text = "" : TextBox2.Text = ""
                'TextBox1.Enabled = True : TextBox2.Enabled = True : ComboBox2.Enabled = True : CheckBox2.Enabled = True
            Else
                TextBox1.Text = "HL Path Not Set" : TextBox2.Text = "HL Path Not Set" : ComboBox1.SelectedIndex = -1
                RadioButton2.Checked = True : RadioButton1.Enabled = False
            End If
            If _data(7) = "" Then TextBox1.BackColor = Form1.colorNO : TextBox3.BackColor = Form1.colorNO Else TextBox1.BackColor = Form1.colorYES : TextBox3.BackColor = Form1.colorYES
            If _data(8) = "" Then TextBox2.BackColor = Form1.colorNO : TextBox4.BackColor = Form1.colorNO Else TextBox2.BackColor = Form1.colorYES : TextBox4.BackColor = Form1.colorYES

            sys = _data(11)
            sysNormalCase = _data(12)
            If sys <> "" AndAlso FileExists(Class1.HyperspinPath + "\Settings\" + sys + ".ini") Then
                Dim iniClass As New IniFile
                iniClass.Load(Class1.HyperspinPath + "\Settings\" + sys + ".ini")
                TextBox3.Text = iniClass.GetKeyValue("EXE INFO", "path").Trim
                TextBox4.Text = iniClass.GetKeyValue("EXE INFO", "rompath").Trim
            End If

            'Add main menu systems to listview
            ListView1.Items.Clear()
            Dim _mainMenuUpper As New List(Of String)
            For Each s As String In _mainMenu
                ListView1.Items.Add(New ListViewItem(s))
                _mainMenuUpper.Add(s.ToUpper)
            Next
            If Not _mainMenuUpper.Contains(sys) Then ListView1.Items.Insert(0, New ListViewItem(sys))

            If Not _mainMenuUpper.Contains(sys) Then
                ListView1.TopItem = ListView1.Items(0)
                ListView1.Items(0).Selected = True
            Else
                Dim ind As Integer = _mainMenuUpper.IndexOf(sys)
                If ind >= 3 Then ListView1.TopItem = ListView1.Items(ind - 3) Else ListView1.TopItem = ListView1.Items(0)
                ListView1.Items(ind).Selected = True
            End If

            refr = False
        End Set
    End Property
    'Modules property set/get to interchange modules data for current system with main form
    Public Property modules As List(Of String)
        Get
            Return _modules
        End Get
        Set(value As List(Of String))
            _modules = value

            'Fill modules list (only compatible here)
            ComboBox2.Items.Clear()
            For Each m As String In _modules
                ComboBox2.Items.Add(m)
            Next

            'Fill emulators list (only those, which module exist in modules list)
            ComboBox3.Items.Clear()
            For Each k As String In _emulators.Keys
                Dim m As String = _emulators(k)(1)
                For i As Integer = 0 To ComboBox2.Items.Count - 1
                    If ComboBox2.Items(i).ToString.ToUpper = m.ToUpper Then ComboBox3.Items.Add(k) : Exit For
                Next
            Next

            'Set hl rom path, and default emulator (if exist in emulator list)
            Dim emu As String = ""
            Dim ini As New IniFileApi
            If sys <> "" AndAlso FileIO.FileSystem.FileExists(HL_Path + "\Settings\" + sys + "\Emulators.ini") Then
                ini.path = HL_Path + "\Settings\" + sys + "\Emulators.ini"
                emu = ini.IniReadValue("ROMS", "Default_Emulator").Trim
                TextBox2.Text = ini.IniReadValue("ROMS", "Rom_Path").Trim
            Else
                TextBox2.Text = ""
            End If
            If emu <> "" And ComboBox3.Items.Contains(emu) Then ComboBox3.SelectedItem = emu Else ComboBox3.SelectedIndex = -1
        End Set
    End Property
    'One time init
    Public Sub oneTimeInit(HS_ini As List(Of String), mainMenu As List(Of String), emulators As Dictionary(Of String, List(Of String)))
        _mainMenu = mainMenu
        _HS_ini = HS_ini
        _emulators = emulators
        ComboBox1.Items.Clear()
        For Each s As String In _HS_ini
            ComboBox1.Items.Add(s)
        Next

        Dim ini As New IniFileApi
        ini.path = Class1.HyperspinPath + "\Settings\Settings.ini"
        HL_Path = ini.IniReadValue("Main", "Hyperlaunch_Path").Trim
        If HL_Path.ToUpper.EndsWith("EXE") Then HL_Path = HL_Path.Substring(0, HL_Path.LastIndexOf("\"))
    End Sub

    'Change HL emulator
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        current_module = ""
        If Not sys = "" And Not HL_Path = "" And ComboBox3.SelectedIndex >= 0 Then
            Dim emu As String = ComboBox3.SelectedItem.ToString
            If _emulators.Keys.Contains(emu) Then
                TextBox1.Text = _emulators(emu)(2)
                current_module = _emulators(emu)(1)
            Else
                TextBox1.Text = ""
            End If
            If current_module.Contains("\") Then current_module = current_module.Substring(current_module.LastIndexOf("\") + 1)

            If current_module <> "" Then
                For i As Integer = 0 To ComboBox2.Items.Count - 1
                    If ComboBox2.Items(i).ToString.ToUpper = current_module.ToUpper Then ComboBox2.SelectedIndex = i : Exit Sub
                Next
                ComboBox2.Items.Add(current_module)
                ComboBox2.SelectedIndex = ComboBox2.Items.Count - 1
                ComboBox2.ForeColor = Color.DarkRed
            End If
            Exit Sub
        End If
        ComboBox2.SelectedIndex = -1
        TextBox1.Text = ""
    End Sub

    'Use Hyperspin / hyperlaunch switch
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked Then
            ComboBox2.Enabled = True : ComboBox3.Enabled = True : CheckBox2.Enabled = True
            TextBox1.Enabled = True : TextBox2.Enabled = True
            Button1.Enabled = True : Button2.Enabled = True
            TextBox3.Enabled = False : TextBox4.Enabled = False
            Button4.Enabled = False : Button3.Enabled = False
        Else
            ComboBox2.Enabled = False : ComboBox3.Enabled = False : CheckBox2.Enabled = False
            TextBox1.Enabled = False : TextBox2.Enabled = False
            Button1.Enabled = False : Button2.Enabled = False
            TextBox3.Enabled = True : TextBox4.Enabled = True
            Button4.Enabled = True : Button3.Enabled = True
        End If
    End Sub

    ' ... browse for files/folders Buttons
    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click, Button2.Click, Button3.Click, Button4.Click
        Dim name As String = DirectCast(sender, Button).Name.ToLower
        If name = "button1" Then
            Dim fd As New OpenFileDialog
            fd.Filter = "Emulator exe (*.exe)|*.exe|All files (*.*)|*.*"
            fd.FileName = TextBox1.Text
            fd.ShowDialog()
            TextBox1.Text = fd.FileName
        ElseIf name = "button2" Then
            Dim fd As New FolderBrowserDialog
            fd.SelectedPath = TextBox2.Text
            fd.ShowDialog()
            TextBox2.Text = fd.SelectedPath
        ElseIf name = "button3" Then
            Dim fd As New OpenFileDialog
            fd.Filter = "Emulator exe (*.exe)|*.exe|All files (*.*)|*.*"
            fd.FileName = TextBox3.Text
            fd.ShowDialog()
            TextBox3.Text = fd.FileName
        ElseIf name = "button4" Then
            Dim fd As New FolderBrowserDialog
            fd.SelectedPath = TextBox4.Text
            fd.ShowDialog()
            TextBox4.Text = fd.SelectedPath
        End If
        'updateTextboxesColors()
    End Sub

    'Change path in textbox
    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged, TextBox2.TextChanged, TextBox3.TextChanged, TextBox4.TextChanged
        Dim name As String = DirectCast(sender, TextBox).Name.ToLower

        Dim tmpPath As String = ""
        If name = "textbox1" Then
            If TextBox1.Text.Trim = "" Or HL_Path = "" Then
                TextBox1.BackColor = Form1.colorNO
            Else
                tmpPath = TextBox1.Text
                If tmpPath.StartsWith(".") Then tmpPath = HL_Path + "\" + tmpPath
                If FileExists(tmpPath) Then
                    TextBox1.BackColor = Form1.colorYES

                    'Try to find module
                    If CheckBox2.Checked Then
                        ComboBox2.SelectedIndex = -1
                        Dim filenameWoExtension As String = tmpPath.Substring(tmpPath.LastIndexOf("\") + 1).ToUpper.Trim
                        If filenameWoExtension.Contains(".") Then filenameWoExtension = filenameWoExtension.Substring(0, filenameWoExtension.LastIndexOf("."))
                        For i As Integer = 0 To ComboBox2.Items.Count - 1
                            Dim moduleWoExtension As String = ComboBox2.Items(i).ToString.Trim.ToUpper
                            If moduleWoExtension.Contains(".") Then moduleWoExtension = moduleWoExtension.Substring(0, moduleWoExtension.LastIndexOf("."))
                            If moduleWoExtension.StartsWith(filenameWoExtension) Or filenameWoExtension.StartsWith(moduleWoExtension) Then ComboBox2.SelectedIndex = i : Exit For
                        Next
                    End If
                Else
                    TextBox1.BackColor = Form1.colorNO
                End If
            End If
        End If

        If name = "textbox2" Then
            If TextBox2.Text.Trim = "" Or HL_Path = "" Then
                TextBox2.BackColor = Form1.colorNO
            Else
                Dim somethingFound As Boolean = False
                Dim somethingNotFound As Boolean = False
                For Each tmp As String In TextBox2.Text.Split({"|"c}, StringSplitOptions.RemoveEmptyEntries)
                    If tmp.StartsWith(".") Then tmp = HL_Path + "\" + tmp
                    If DirectoryExists(tmp) Then somethingFound = True Else somethingNotFound = True
                Next

                If somethingFound And Not somethingNotFound Then TextBox2.BackColor = Form1.colorYES
                If somethingFound And somethingNotFound Then TextBox2.BackColor = Color.YellowGreen
                If Not somethingFound Then TextBox2.BackColor = Form1.colorNO
            End If
        End If

        If name = "textbox3" Then
            If TextBox3.Text.Trim = "" Then
                TextBox3.BackColor = Form1.colorNO
            Else
                tmpPath = TextBox3.Text
                If tmpPath.StartsWith(".") Then tmpPath = Class1.HyperspinPath + "\" + tmpPath
                If FileExists(tmpPath) Then TextBox3.BackColor = Form1.colorYES Else TextBox3.BackColor = Form1.colorNO
            End If
        End If

        If name = "textbox4" Then
            If TextBox4.Text.Trim = "" Then
                TextBox4.BackColor = Form1.colorNO
            Else
                tmpPath = TextBox4.Text
                If tmpPath.StartsWith(".") Then tmpPath = Class1.HyperspinPath + "\" + tmpPath
                If DirectoryExists(tmpPath) Then TextBox4.BackColor = Form1.colorYES Else TextBox4.BackColor = Form1.colorNO
            End If
        End If
    End Sub

    'UPDATE
    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        'If HL_Path 
        If RadioButton1.Checked And ComboBox3.SelectedIndex < 0 Then MsgBox("Emulator need to be set to use hyperlaunch.") : Exit Sub
        If RadioButton1.Checked And ComboBox2.SelectedIndex < 0 Then MsgBox("Module need to be set to use hyperlaunch.") : Exit Sub

        'Add/remove system in main menu
        If _data(0) = "" And CheckBox1.Checked And insert_new_system_at <> -1 Then
            'Add system
            Dim x As New Xml.XmlDocument
            x.Load(Class1.HyperspinPath + "\Databases\Main Menu\Main Menu.xml")
            Dim rootNode = x.SelectSingleNode("/menu")
            Dim gameNodes = x.SelectNodes("/menu/game")
            Dim newNode = x.CreateElement("game")
            newNode.SetAttribute("name", sysNormalCase)
            rootNode.InsertBefore(newNode, gameNodes(insert_new_system_at))
            x.Save(Class1.HyperspinPath + "\Databases\Main Menu\Main Menu.xml")
            _data(0) = "1"
        ElseIf _data(0) <> "" And Not CheckBox1.Checked And insert_new_system_at = -1 Then
            'Remove system
            Dim x As New Xml.XmlDocument
            x.Load(Class1.HyperspinPath + "\Databases\Main Menu\Main Menu.xml")
            Dim node_to_delete = x.SelectSingleNode("/menu/game[@name='" + sysNormalCase + "']")
            x.SelectSingleNode("/menu").RemoveChild(node_to_delete)
            x.Save(Class1.HyperspinPath + "\Databases\Main Menu\Main Menu.xml")
            _data(0) = ""
        End If

        'Emu & Rom Paths
        Dim found As Boolean = False
        Dim iniclass As New IniFile
        Dim iniclass2 As New IniFileApi

        'hyperlaunch
        Dim selectedEmuName As String = TextBox1.Text
        selectedEmuName = selectedEmuName.Substring(selectedEmuName.LastIndexOf("\") + 1).Trim()
        If selectedEmuName.Contains(".") Then selectedEmuName = selectedEmuName.Substring(0, selectedEmuName.LastIndexOf("."))

        If RadioButton1.Checked Then
            If ComboBox3.SelectedIndex < 0 Then MsgBox("You have to select an emulator, to use HyperLaunch")

            Dim emu As String = ComboBox3.SelectedItem.ToString
            Dim moduleWoExtension As String = ComboBox2.Items(ComboBox2.SelectedIndex).ToString.Trim()
            If moduleWoExtension.Contains(".") Then moduleWoExtension = moduleWoExtension.Substring(0, moduleWoExtension.LastIndexOf("."))

            If selectedEmuName.ToUpper <> emu.ToUpper And selectedEmuName.ToUpper <> moduleWoExtension.ToUpper Then
                Dim msg As String = "Your module name is: " + moduleWoExtension + vbCrLf + "Associated emulator is: " + emu + vbCrLf + "and your selected emulator is: " + selectedEmuName
                msg = msg + vbCrLf + "Your selected emulator doesn't match with module name nither with associated emulator name. Sometimes it happens. Are you sure this information is correct?"
                msg = msg + vbCrLf + "(emulator in [" + emu + "] section, in 'Global Emulators.ini' will be overwriten by your newly selected emulator)"
                Dim res As MsgBoxResult = MsgBox(msg, MsgBoxStyle.YesNo)
                If res = MsgBoxResult.No Then Exit Sub
            End If

            If FileExists(HL_Path + "\Settings\Global Emulators.ini") Then
                Dim ini As New IniFileApi
                ini.path = HL_Path + "\Settings\Global Emulators.ini"
                ini.IniWriteValue(emu, "Emu_Path", Absolute_Path_to_Relative((HL_Path + "\").Replace("\\", "\"), TextBox1.Text))

                If Not FileIO.FileSystem.FileExists(HL_Path + "\Settings\" + sys + "\Emulators.ini") Then
                    'MsgBox("System setting file does not exist: """ + HL_Path + "\Settings\" + sys + "\Emulators.ini""")
                    'Exit Sub
                    Dim msg As String = "System setting file does not exist: """ + HL_Path + "\Settings\" + sys + "\Emulators.ini"""
                    Dim frm As New FormF_createNewHL_system
                    frm.init(msg, HL_Path)
                    frm.ShowDialog(Me)
                    If frm.response = "" Then Exit Sub
                    CopyDirectory(HL_Path + "\Settings\" + frm.response, HL_Path + "\Settings\" + sys, True)
                End If

                ini.path = HL_Path + "\Settings\" + sys + "\Emulators.ini"
                ini.IniWriteValue("ROMS", "Default_Emulator", emu)
                ini.IniWriteValue("ROMS", "Rom_Path", Absolute_Path_to_Relative((HL_Path + "\").Replace("\\", "\"), TextBox2.Text))
            Else
                MsgBox("File does not exist: """ + HL_Path + "\Settings\Global Emulators.ini""") : Exit Sub
            End If
        End If

        RaiseEvent paths_updated(sys) : Me.Close() : Exit Sub
    End Sub

    'Absolute_Path_to_Relative
    Private Function Absolute_Path_to_Relative(path_from As String, path_to As String) As String
        If path_to.StartsWith(".") Then Return path_to
        If String.IsNullOrEmpty(path_from) Then Return ""
        If String.IsNullOrEmpty(path_to) Then Return ""

        If path_to.Substring(0, 1).ToUpper <> path_from.Substring(0, 1).ToUpper Then Return path_to
        Dim UriFrom = New Uri(path_from)
        Dim UriTo = New Uri(path_to)

        If UriFrom.Scheme <> UriTo.Scheme Then Return path_to

        Dim UriRelative As Uri = UriFrom.MakeRelativeUri(UriTo)
        Dim relativePath As String = Uri.UnescapeDataString(UriRelative.ToString)

        If UriTo.Scheme.ToUpperInvariant = "FILE" Then
            relativePath = relativePath.Replace(System.IO.Path.AltDirectorySeparatorChar, System.IO.Path.DirectorySeparatorChar)
        End If
        If Not relativePath.StartsWith(".") Then relativePath = ".\" + relativePath
        Return relativePath
    End Function

    'Always on top checkbox
    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        Me.TopMost = CheckBox3.Checked
    End Sub

    'Activate in mainMenu checkbox
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If refr Then Exit Sub
        If Not CheckBox1.Checked Then insert_new_system_at = -1 : Exit Sub

        If ListView1.SelectedItems.Count = 1 Then ListView1.SelectedItems(0).Text = sys
        For Each c As Control In Me.Controls
            If Not controls_to_hide.Contains(c.Name.ToUpper) Then c.Enabled = False
        Next

        Panel1.Visible = True
        ListView1.Focus()
    End Sub

    'Copy hs.ini
    Private Sub Button_copyFrom_Click(sender As Object, e As EventArgs) Handles Button_copyFrom.Click
        If ComboBox1.SelectedIndex < 0 Then Exit Sub
        Dim p As String = Class1.HyperspinPath + "\Settings\" + ComboBox1.SelectedItem.ToString
        Dim p_new As String = Class1.HyperspinPath + "\Settings\" + sysNormalCase + ".ini"
        If Not FileExists(p) Then MsgBox("File not exist." + vbCrLf + p) : Exit Sub
        If FileExists(p_new) Then MsgBox("File already exist." + vbCrLf + p_new) : Exit Sub
        FileCopy(p, p_new)
        Label6.BackColor = Form1.colorYES

        If FileExists(Class1.HyperspinPath + "\Settings\" + sys + ".ini") Then
            Dim iniClass As New IniFile
            iniClass.Load(Class1.HyperspinPath + "\Settings\" + sys + ".ini")
            TextBox3.Text = iniClass.GetKeyValue("EXE INFO", "path").Trim
            TextBox4.Text = iniClass.GetKeyValue("EXE INFO", "rompath").Trim

            iniClass.Load(Class1.HyperspinPath + "\Settings\Settings.ini")
            HL_Path = iniClass.GetKeyValue("Main", "Hyperlaunch_Path").Trim
            If HL_Path.ToUpper.EndsWith("EXE") Then HL_Path = HL_Path.Substring(0, HL_Path.LastIndexOf("\"))
            If FileIO.FileSystem.FileExists(HL_Path + "\Settings\" + sys + "\Emulators.ini") Then
                iniClass.Load(HL_Path + "\Settings\" + sys + "\Emulators.ini")
                Dim emu As String = iniClass.GetKeyValue("ROMS", "Default_Emulator").Trim
                TextBox2.Text = iniClass.GetKeyValue("ROMS", "Rom_Path").Trim
                Dim default_emulator = iniClass.GetKeyValue("ROMS", "Default_Emulator").Trim

                If emu <> "" Then
                    iniClass.Load(HL_Path + "\Settings\Global Emulators.ini")
                    TextBox1.Text = iniClass.GetKeyValue(emu, "Emu_Path").Trim
                    current_module = iniClass.GetKeyValue(emu, "Module").Trim
                End If
            End If
        End If

        RaiseEvent paths_updated(sys)
    End Sub

#Region "Region: Insert in mainMenu panel"
    'Main menu systems list selection change
    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView1.SelectedIndexChanged
        If refr Or ListView1.SelectedItems.Count <> 1 Then Exit Sub
        Dim c As Integer = 0
        Dim f As Font = ListView1.Font
        Dim f_bold As Font = New Font(f.FontFamily, f.SizeInPoints + 4, FontStyle.Bold)
        For Each l As ListViewItem In ListView1.Items
            If l.Selected Then
                l.Text = sys
                l.Font = f_bold
            Else
                l.Text = _mainMenu(c)
                l.Font = f
                c += 1
            End If
        Next
    End Sub

    'Set to top
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If ListView1.SelectedItems.Count <> 1 Then Exit Sub

        Dim c As Integer = 0
        Dim f As Font = ListView1.Font
        Dim f_bold As Font = New Font(f.FontFamily, f.SizeInPoints + 4, FontStyle.Bold)
        For i As Integer = 0 To ListView1.Items.Count - 1
            Dim l As ListViewItem = ListView1.Items(i)
            If i = 0 Then
                l.Text = sys
                l.Font = f_bold
            Else
                l.Text = _mainMenu(c)
                l.Font = f
                c += 1
            End If
        Next
        ListView1.TopItem = ListView1.Items(0)
        ListView1.Items(0).Selected = True
        ListView1.Focus()
    End Sub

    'Set to bottom
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If ListView1.SelectedItems.Count <> 1 Then Exit Sub

        Dim c As Integer = 0
        Dim f As Font = ListView1.Font
        Dim f_bold As Font = New Font(f.FontFamily, f.SizeInPoints + 4, FontStyle.Bold)
        For i As Integer = 0 To ListView1.Items.Count - 1
            Dim l As ListViewItem = ListView1.Items(i)
            If i = ListView1.Items.Count - 1 Then
                l.Text = sys
                l.Font = f_bold
            Else
                l.Text = _mainMenu(c)
                l.Font = f
                c += 1
            End If
        Next
        ListView1.TopItem = ListView1.Items(ListView1.Items.Count - 1)
        ListView1.Items(ListView1.Items.Count - 1).Selected = True
        ListView1.Focus()
    End Sub

    'OK
    Dim insert_new_system_at As Integer = -1
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Panel1.Visible = False
        For Each c As Control In Me.Controls
            If Not controls_to_hide.Contains(c.Name.ToUpper) Then c.Enabled = True
        Next
        insert_new_system_at = ListView1.SelectedIndices(0)
    End Sub

    'Cancel
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        refr = True : CheckBox1.Checked = False : refr = False
        Panel1.Visible = False
        For Each c As Control In Me.Controls
            If Not controls_to_hide.Contains(c.Name.ToUpper) Then c.Enabled = True
        Next
        insert_new_system_at = -1
    End Sub
#End Region
End Class