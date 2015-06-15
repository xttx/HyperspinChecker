Imports Microsoft.VisualBasic.FileIO.FileSystem

Public Class Form8_systemProperties
    Dim sys As String = ""
    Dim _data As List(Of String)
    Dim _modules As List(Of String)
    Dim HL_Path As String = ""
    Dim current_module As String = ""
    Public Shared Event paths_updated(system As String)

    'Data property set/get to interchange data with main form
    Public Property data As List(Of String)
        Get
            Return _data
        End Get
        Set(value As List(Of String))
            ComboBox2.Items.Clear()

            _data = value
            If _data.Count = 0 Then Me.Enabled = False : Exit Property
            Me.Enabled = True
            If _data(0) = "" Then CheckBox1.Checked = False Else CheckBox1.Checked = True
            If _data(1) = "" Then Label5.BackColor = Form1.colorNO Else Label5.BackColor = Form1.colorYES
            If _data(2) = "" Then Label1.BackColor = Form1.colorNO Else Label1.BackColor = Form1.colorYES
            If _data(3) = "" Then Label2.BackColor = Form1.colorNO Else Label2.BackColor = Form1.colorYES
            If _data(4) = "" Then Label6.BackColor = Form1.colorNO Else Label6.BackColor = Form1.colorYES
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

            sys = _data(9)
            If sys <> "" Then
                If FileExists(Class1.HyperspinPath + "\Settings\" + sys + ".ini") Then
                    Dim iniClass As New IniFile
                    iniClass.Load(Class1.HyperspinPath + "\Settings\" + sys + ".ini")
                    TextBox3.Text = iniClass.GetKeyValue("EXE INFO", "path").Trim
                    TextBox4.Text = iniClass.GetKeyValue("EXE INFO", "rompath").Trim

                    iniClass.Load(Class1.HyperspinPath + "\Settings\Settings.ini")
                    HL_Path = iniClass.GetKeyValue("Main", "Hyperlaunch_Path").Trim
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
            End If
        End Set
    End Property

    'Modules property set/get to interchange modules data for current system with main form
    Public Property modules As List(Of String)
        Get
            Return _modules
        End Get
        Set(value As List(Of String))
            _modules = value
            For Each m As String In _modules
                ComboBox2.Items.Add(m)
            Next

            If current_module <> "" Then
                For i As Integer = 0 To ComboBox2.Items.Count - 1
                    If ComboBox2.Items(i).ToString.ToUpper = current_module.ToUpper Then ComboBox2.SelectedIndex = i : Exit Property
                Next
                ComboBox2.Items.Add(current_module)
                ComboBox2.SelectedItem = ComboBox2.Items.Count - 1
                ComboBox2.ForeColor = Color.DarkRed
            End If
        End Set
    End Property

    'Use Hyperspin / hyperlaunch switch
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked Then
            TextBox1.Enabled = True : TextBox2.Enabled = True : ComboBox2.Enabled = True : CheckBox2.Enabled = True
            Button1.Enabled = True : Button2.Enabled = True
            TextBox3.Enabled = False : TextBox4.Enabled = False
            Button4.Enabled = False : Button3.Enabled = False
        Else
            TextBox1.Enabled = False : TextBox2.Enabled = False : ComboBox2.Enabled = False : CheckBox2.Enabled = False
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
            fd.ShowDialog()
            TextBox1.Text = fd.FileName
        ElseIf name = "button2" Then
            Dim fd As New FolderBrowserDialog
            fd.ShowDialog()
            TextBox2.Text = fd.SelectedPath
        ElseIf name = "button3" Then
            Dim fd As New OpenFileDialog
            fd.Filter = "Emulator exe (*.exe)|*.exe|All files (*.*)|*.*"
            fd.ShowDialog()
            TextBox3.Text = fd.FileName
        ElseIf name = "button4" Then
            Dim fd As New FolderBrowserDialog
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
                tmpPath = TextBox2.Text
                If tmpPath.StartsWith(".") Then tmpPath = HL_Path + "\" + tmpPath
                If DirectoryExists(tmpPath) Then TextBox2.BackColor = Form1.colorYES Else TextBox2.BackColor = Form1.colorNO
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
        If RadioButton1.Checked And ComboBox2.SelectedIndex < 0 Then MsgBox("Module need to be set to use hyperlaunch.") : Exit Sub

        Dim found As Boolean = False
        Dim iniclass As New IniFile
        Dim iniclass2 As New IniFileApi

        'hyperlaunch
        If RadioButton1.Checked Then
            If FileExists(HL_Path + "\Settings\Global Emulators.ini") Then
                iniclass.Load(HL_Path + "\Settings\Global Emulators.ini")
                For Each Section In iniclass.Sections
                    Dim emuSection = DirectCast(Section, IniFile.IniSection)
                    Dim emuName As String = emuSection.Name
                    Dim selectedEmuName As String = TextBox1.Text
                    selectedEmuName = selectedEmuName.Substring(selectedEmuName.LastIndexOf("\") + 1).Trim()
                    If selectedEmuName.Contains(".") Then selectedEmuName = selectedEmuName.Substring(0, selectedEmuName.LastIndexOf("."))
                    Dim moduleWoExtension As String = ComboBox2.Items(ComboBox2.SelectedIndex).ToString.Trim()
                    If moduleWoExtension.Contains(".") Then moduleWoExtension = moduleWoExtension.Substring(0, moduleWoExtension.LastIndexOf("."))

                    If iniclass.GetKeyValue(emuName, "Module") = ComboBox2.Items(ComboBox2.SelectedIndex).ToString Then
                        found = True
                        If selectedEmuName.ToUpper <> emuName.ToUpper And selectedEmuName.ToUpper <> moduleWoExtension.ToUpper Then
                            Dim msg As String = "Your module name is: " + moduleWoExtension + vbCrLf + "Associated emulator is: " + emuName + vbCrLf + "and your selected emulator is: " + selectedEmuName
                            msg = msg + vbCrLf + "Your selected emulator doesn't match with module name nither with associated emulator name. Sometimes it happens. Are you sure this information is correct?"
                            msg = msg + vbCrLf + "(emulator in [" + emuName + "] section, in 'Global Emulators.ini' will be overwriten by your newly selected emulator)"
                            Dim res As MsgBoxResult = MsgBox(msg, MsgBoxStyle.YesNo)
                            If res = MsgBoxResult.No Then Exit Sub
                        End If

                        iniclass2.path = HL_Path + "\Settings\Global Emulators.ini"
                        iniclass2.IniWriteValue(emuName, "Emu_Path", Absolute_Path_to_Relative((HL_Path + "\").Replace("\\", "\"), TextBox1.Text))

                        If FileIO.FileSystem.FileExists(HL_Path + "\Settings\" + sys + "\Emulators.ini") Then
                            iniclass2.path = HL_Path + "\Settings\" + sys + "\Emulators.ini"

                            iniclass2.IniWriteValue("ROMS", "Default_Emulator", emuName)
                            iniclass2.IniWriteValue("ROMS", "Rom_Path", Absolute_Path_to_Relative((HL_Path + "\").Replace("\\", "\"), TextBox2.Text))
                        Else
                            MsgBox("File does not exist: """ + HL_Path + "\Settings\" + sys + "\Emulators.ini""") : Exit Sub
                        End If
                        RaiseEvent paths_updated(sys) : Me.Close() : Exit Sub
                    End If
                Next
                If Not found Then MsgBox("There is no emulator associated with selected modul in 'Global Emulators.ini'. You have to either add new emulator associated with this module through HyperLaunchHQ, or select another module." + vbCrLf + "This feature will be implimented in HyperspinChecker soon, but not yet.") : Exit Sub
            Else
                MsgBox("File does not exist: """ + HL_Path + "\Settings\Global Emulators.ini""") : Exit Sub
            End If
        End If
    End Sub

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
End Class