Imports System.Text.RegularExpressions
Public Class FormJa2_ScriptConstructor

    'TODO - Registry isd
    'TODO - Fill default value from current ini/registry

    Public emuPath As String = ""
    Public modPath As String = ""
    Public returnScript As String = ""

    Dim emuPathsDict As New Dictionary(Of String, String)
    Dim modPathsDict As New Dictionary(Of String, String)
    Dim refreshing As Boolean = False
    Dim f As New Font("Courier New", 8)
    Dim ini_types As New Dictionary(Of String, String)
    Dim control_array As New Dictionary(Of String, List(Of control_elements))
    Dim available_variables As New List(Of String)
    Dim available_values_from_ini As New List(Of String)
    Dim available_variables_types() As String
    Dim available_defaults_types() As String
    'Dim available_variables_bind As New BindingSource

    Dim RegistryBrowserForm As New FormJa2b_RegistryBrowser

    Structure control_elements
        Dim l As Label
        Dim variable As ComboBox
        Dim variable_type As ComboBox
        Dim def As ComboBox
        Dim fulRow As CheckBox
        Dim buttonUp As Button
        Dim buttonDown As Button
        Dim buttonDel As Button
    End Structure

    Private Sub FormJa2_ScriptConstructor_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        available_defaults_types = {"{EmuPath}", "{RomPath}", "{RomPathRelative}", "{RLModulePath}", "{DesktopResolution}", "{DesktopResolutionW}", "{DesktopResolutionH}"}
        available_variables_types = {"Boolean", "Integer", "String", "Long", "FileName", "FilePath", "FolderPath", "Resolution", "ExeTarget", "ExeStartIn", "ExeParams", "ProcessName", "URI", "WinTitle", "xHotkey"}

        'Add emu inis and RL module inis to file chooser
        Dim c = 0
        Dim emuPathArr = emuPath.Split({"|"c}, StringSplitOptions.RemoveEmptyEntries)
        For Each path In emuPathArr
            emuPathsDict.Add("%EmuPath" + c.ToString + "%", path)
            If IO.Directory.Exists(path) Then
                Dim f_list = IO.Directory.GetFiles(path, "*.ini", IO.SearchOption.AllDirectories).ToList
                f_list.AddRange(IO.Directory.GetFiles(path, "*.xml", IO.SearchOption.AllDirectories))
                f_list.Sort()

                Dim emu_root = IO.Path.GetDirectoryName(path)
                Dim emu_dir = IO.Path.GetFileName(path)
                For Each f As String In f_list
                    ComboBox1.Items.Add("%EmuPath" + c.ToString + "%" + "\" + emu_dir + f.Substring(path.Length))
                Next
            End If
            c += 1
        Next

        c = 0
        Dim modPathArr = modPath.Split({"|"c}, StringSplitOptions.RemoveEmptyEntries)
        For Each path In modPathArr
            modPathsDict.Add("%RLModulePath" + c.ToString + "%", path)
            If IO.File.Exists(path) Then
                path = IO.Path.GetDirectoryName(path)
                Dim f_list = IO.Directory.GetFiles(path, "*.ini", IO.SearchOption.AllDirectories).ToList
                f_list.Sort()
                For Each f As String In f_list
                    ComboBox1.Items.Add("%RLModulePath" + c.ToString + "%" + f.Substring(path.Length))
                Next
            End If
            c += 1
        Next


        'available_variables_bind.DataSource = available_variables
        Button1.Width = Button1.Parent.Width - (Button1.Left * 2)
    End Sub

    'Form Closing - generate script
    Private Sub FormJa2_ScriptConstructor_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Dim ini_files = control_array.Keys.ToList
        ini_files.Sort()

        returnScript = "<?xml version=""1.0"" encoding=""UTF-8""?>" + vbCrLf
        returnScript += "<INISCHEMA>" + vbCrLf
        returnScript += "   <INIFILES>" + vbCrLf
        For Each ini In ini_files
            ini = ini.Trim

            'Parse sections
            Dim sections As New Dictionary(Of String, List(Of String()))
            For Each el In control_array(ini)
                Dim arr = el.variable.Text.Trim.Split("]"c)

                If arr.Count <> 2 OrElse Not arr(0).StartsWith("[") Then
                    Dim res = MsgBox("Problem in the ini [section]key format in " + ini + vbCrLf + el.variable.Text.Trim + vbCrLf + "Script will not be generated. Ok - close form, Cancel - continue edit.", MsgBoxStyle.OkCancel)
                    If res = MsgBoxResult.Cancel Then e.Cancel = True Else returnScript = ""
                    Exit Sub
                End If

                Dim cur_sect = arr(0).Substring(1)

                If Not sections.ContainsKey(cur_sect) Then sections.Add(cur_sect, New List(Of String()))
                sections(cur_sect).Add({arr(1), el.variable_type.Text.Trim, el.def.Text, el.fulRow.Checked.ToString.ToUpper})

                'TODO - ini types
            Next

            'TODO - Don't create empty sections, after deleting all variables
            returnScript += "       <INIFILE name=""" + ini + """ required=""false"" iniType=""" + ini_types(ini.ToUpper) + """>" + vbCrLf
            returnScript += "           <INITYPE>HS_Package</INITYPE>" + vbCrLf
            returnScript += "           <SECTIONS>" + vbCrLf
            For Each section In sections.Keys
                returnScript += "               <SECTION name=""" + section + """>" + vbCrLf
                returnScript += "                   <SECTIONTYPE>Global</SECTIONTYPE>" + vbCrLf
                returnScript += "                   <KEYS>" + vbCrLf
                For Each k In sections(section)
                    returnScript += "                       <KEY name=""" + k(0) + """ required=""false"" nullable=""false"">" + vbCrLf
                    returnScript += "                           <KEYTYPE>" + k(1) + "</KEYTYPE>" + vbCrLf
                    If k(3) = "TRUE" Then
                        returnScript += "                           <FULLROW>true</FULLROW>" + vbCrLf
                    End If
                    returnScript += "                           <DEFAULT>" + k(2) + "</DEFAULT>" + vbCrLf
                    returnScript += "                       </KEY>" + vbCrLf
                Next
                returnScript += "                   </KEYS>" + vbCrLf
                returnScript += "               </SECTION>" + vbCrLf
            Next
            returnScript += "           </SECTIONS>" + vbCrLf
            returnScript += "       </INIFILE>" + vbCrLf
        Next
        returnScript += "   </INIFILES>" + vbCrLf
        returnScript += "</INISCHEMA>" + vbCrLf
    End Sub

    'Choose a file
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.TextChanged
        'ComboBox2.SelectedItem = IO.Path.GetExtension(ComboBox1.Text).Substring(1)

        refreshing = True
        If Not ini_types.ContainsKey(ComboBox1.Text.Trim.ToUpper) Then
            If ComboBox1.Text.Trim.ToUpper.EndsWith(".INI") Then
                ComboBox2.SelectedIndex = 0
            ElseIf ComboBox1.Text.Trim.ToUpper.EndsWith(".XML") Then
                ComboBox2.SelectedIndex = 3
            ElseIf ComboBox1.Text.Trim.ToUpper.StartsWith("HK") Then
                ComboBox2.SelectedIndex = 4
            Else
                ComboBox2.SelectedIndex = -1
            End If
        Else
            ComboBox2.SelectedItem = ini_types(ComboBox1.Text.Trim.ToUpper)
        End If
        refreshing = False

        'update_variables()

        'To make sure, a ini_types entry is created for this file, even if combobox2 selection index (ini type) was not changed
        ComboBox2_SelectedIndexChanged(ComboBox2, New EventArgs)

        'available_variables_bind.ResetBindings(False)

        'Update FlowPanel, hide unneeded controls
        For Each kv In control_array
            Dim vis As Boolean = False
            If kv.Key = ComboBox1.Text.Trim.ToUpper Then vis = True

            For Each ctrl_el In control_array(kv.Key)
                ctrl_el.l.Visible = vis
                ctrl_el.variable.Visible = vis
                ctrl_el.variable_type.Visible = vis
                ctrl_el.def.Visible = vis
                ctrl_el.fulRow.Visible = vis
                ctrl_el.buttonDown.Visible = vis
                ctrl_el.buttonUp.Visible = vis
                ctrl_el.buttonDel.Visible = vis
            Next
        Next
    End Sub
    'Choose ini type
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.TextChanged
        If refreshing Then Exit Sub
        If ComboBox2.SelectedIndex < 0 Then available_variables.Clear() : available_values_from_ini.Clear() : Exit Sub
        If ComboBox1.Text.Trim.ToUpper = "" Then available_variables.Clear() : available_values_from_ini.Clear() : Exit Sub
        ini_types(ComboBox1.Text.Trim.ToUpper) = ComboBox2.Text.Trim
        update_variables()
    End Sub

    'Update available variables for current ini
    Private Sub update_variables()
        available_variables.Clear()
        available_values_from_ini.Clear()

        'Dim realPath = ComboBox1.Text.Trim.Replace("%EmuPath%", IO.Path.GetDirectoryName(emuPath)).Replace("%RLModulePath%", modPath)
        Dim realPath = ComboBox1.Text.Trim
        For Each kv In emuPathsDict
            If realPath.Contains(kv.Key) Then realPath = realPath.Replace(kv.Key, IO.Path.GetDirectoryName(kv.Value))
        Next
        For Each kv In modPathsDict
            If realPath.Contains(kv.Key) Then realPath = realPath.Replace(kv.Key, IO.Path.GetDirectoryName(kv.Value))
        Next

        If ComboBox2.SelectedIndex = 0 Then
            'INI Standard
            Dim ini As New IniFileApi With {.path = realPath}
            Dim sections = ini.IniListKey

            If sections.Count = 0 Then
                Dim msg = "No sections found in this ini. If it's non-standard ini, please, choose appropriate ini type:" + vbCrLf
                msg += "  -- MAME.ini - no sections, no '=' separator between key and value." + vbCrLf
                msg += "  -- UniEmu.ini - no sections, keys can contain dots, values can contains spaces."
                MsgBox(msg)
                Exit Sub
            End If

            For Each section In sections
                For Each key In ini.IniListKey(section)
                    available_variables.Add("[" + section + "]" + key)
                    available_values_from_ini.Add(ini.IniReadValue(section, key))
                Next
            Next
        ElseIf ComboBox2.SelectedIndex = 1 Then
            'MAME Ini
            Dim ini As New IniFile()
            ini.LoadMame(realPath, " "c)
            For Each k As IniFile.IniSection.IniKey In ini.GetSection("Main").Keys
                available_variables.Add("[Main]" + k.Name)
                available_values_from_ini.Add(k.Value)
            Next
        ElseIf ComboBox2.SelectedIndex = 2 Then
            'Emu 0.01 ini
            Dim ini As New IniFile()
            ini.LoadMame(realPath, "="c)
            For Each k As IniFile.IniSection.IniKey In ini.GetSection("Main").Keys
                available_variables.Add("[Main]" + k.Name)
                available_values_from_ini.Add(k.Value)
            Next
        ElseIf ComboBox2.SelectedIndex = 3 Then
            'XML
            Try
                Dim x As New Xml.XmlDocument
                x.Load(realPath)
                Dim arr As New List(Of String)
                getNodesRecurcively(x.DocumentElement, arr)
                For Each s In arr
                    available_variables.Add(s)
                Next
                'We need to fill available_values_from_ini, but I don't know yet how XMLs should actually work
            Catch ex As Exception
                MsgBox("There was an error to parse this xml. Looks like it's not a valid XML, this case is not supported yet, sorry.")
            End Try
        ElseIf ComboBox2.SelectedIndex = 4 Then
            'Registry
            'TODO
        Else
            ComboBox2.SelectedIndex = -1
        End If
    End Sub


    'Add variable
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ComboBox1.Text.Trim = "" Then MsgBox("Please, choose a file or registry key.") : Exit Sub

        Dim l As New Label With {.Font = f, .AutoSize = False, .TextAlign = ContentAlignment.MiddleLeft, .BorderStyle = BorderStyle.FixedSingle, .Size = New Size(30, 20), .Margin = New Padding(3)}
        Dim variable As New ComboBox With {.Width = 340}
        Dim variable_type As New ComboBox With {.Width = 90, .DropDownStyle = ComboBoxStyle.DropDownList}
        Dim def As New ComboBox With {.Width = 260}
        Dim fullRow As New CheckBox With {.Width = 20, .Text = "FR"}
        Dim up_button As New Button With {.Width = 20, .Height = 20, .Text = "↑", .TextAlign = ContentAlignment.MiddleCenter}
        Dim down_button As New Button With {.Width = 20, .Height = 20, .Text = "↓", .TextAlign = ContentAlignment.MiddleCenter}
        Dim del_button As New Button With {.Width = 20, .Height = 20, .Text = "x", .TextAlign = ContentAlignment.MiddleCenter}
        AddHandler variable.SelectedIndexChanged, AddressOf Combobox_ChangeVar
        AddHandler up_button.Click, AddressOf Button_UpDownVar_Click
        AddHandler down_button.Click, AddressOf Button_UpDownVar_Click
        AddHandler del_button.Click, AddressOf Button_DeleteVar_Click

        If Not control_array.ContainsKey(ComboBox1.Text.Trim.ToUpper) Then control_array.Add(ComboBox1.Text.Trim.ToUpper, New List(Of control_elements))

        def.Items.AddRange(available_defaults_types)
        variable_type.Items.AddRange(available_variables_types) : variable_type.SelectedIndex = 0
        variable.DataSource = available_variables.ToList 'We add .ToList() here to create new copy of a list, otherwise all combos will be bind to a single list, and change values at the same time
        'variable.DataSource = available_variables_bind.List

        Dim b As Button = DirectCast(sender, Button)
        FlowLayoutPanel1.Controls.Remove(b)
        FlowLayoutPanel1.Controls.AddRange({l, variable, variable_type, def, fullRow, up_button, down_button, del_button, b})

        Dim c As New control_elements
        c.l = l : c.variable = variable : c.variable_type = variable_type : c.def = def : c.fulRow = fullRow
        c.buttonUp = up_button : c.buttonDown = down_button : c.buttonDel = del_button
        control_array(ComboBox1.Text.Trim.ToUpper).Add(c)
        l.Text = Format(control_array(ComboBox1.Text.Trim.ToUpper).Count, "000")
        variable.Tag = c : up_button.Tag = c : down_button.Tag = c : del_button.Tag = c
        Combobox_ChangeVar(variable, New EventArgs, True) 'to update initial value

        b.Focus()
        FlowLayoutPanel1.ScrollControlIntoView(b)
        If Not ComboBox1.Items.Contains(ComboBox1.Text) Then ComboBox1.Items.Add(ComboBox1.Text)
    End Sub
    'Change variable
    Private Sub Combobox_ChangeVar(sender As Object, e As EventArgs, Optional force0 As Boolean = False)
        Dim c = DirectCast(sender, ComboBox)

        Dim ind = c.SelectedIndex
        If force0 Then ind = 0
        If ind < 0 Then Exit Sub

        Dim ctrl = DirectCast(c.Tag, control_elements)

        If available_values_from_ini.Count > ind Then
            Dim v = available_values_from_ini(ind)
            ctrl.def.Text = v

            If v.ToUpper = "TRUE" Or v.ToUpper = "FALSE" Then
                ctrl.variable_type.Text = "Boolean"
            ElseIf IsNumeric(v) Then
                ctrl.variable_type.Text = "Integer"
            ElseIf v.Length > 3 AndAlso (v.Substring(1, 2) = ":\" Or v.StartsWith(".\") Or v.StartsWith("..\")) Then
                If v.EndsWith("\") Then
                    ctrl.variable_type.Text = "FolderPath"
                Else
                    Dim ext = IO.Path.GetExtension(v)
                    If ext <> "" And ext.Length <= 6 Then
                        ctrl.variable_type.Text = "FilePath"
                    Else
                        ctrl.variable_type.Text = "FolderPath"
                    End If
                End If
            Else
                ctrl.variable_type.Text = "String"
            End If

        End If
    End Sub
    'Delete variable
    Private Sub Button_DeleteVar_Click(sender As Object, e As EventArgs)
        Dim b = DirectCast(sender, Button)
        Dim c = DirectCast(b.Tag, control_elements)

        control_array(ComboBox1.Text.Trim.ToUpper).Remove(c)
        If control_array(ComboBox1.Text.Trim.ToUpper).Count = 0 Then control_array.Remove(ComboBox1.Text.Trim.ToUpper)

        FlowLayoutPanel1.Controls.Remove(c.l)
        FlowLayoutPanel1.Controls.Remove(c.variable)
        FlowLayoutPanel1.Controls.Remove(c.variable_type)
        FlowLayoutPanel1.Controls.Remove(c.def)
        FlowLayoutPanel1.Controls.Remove(c.fulRow)
        FlowLayoutPanel1.Controls.Remove(c.buttonUp)
        FlowLayoutPanel1.Controls.Remove(c.buttonDown)
        FlowLayoutPanel1.Controls.Remove(c.buttonDel)
        FlowLayoutPanel1.Controls.Remove(b)
    End Sub
    'Up/down variable
    Private Sub Button_UpDownVar_Click(sender As Object, e As EventArgs)
        Dim b = DirectCast(sender, Button)
        Dim c = DirectCast(b.Tag, control_elements)

        Dim ind As Integer = Nothing
        Dim ind_swap As Integer = Nothing
        If b.Text = "↑" Then
            ind = control_array(ComboBox1.Text.Trim.ToUpper).IndexOf(c)
            If ind = 0 Then Exit Sub
            ind_swap = ind - 1
        ElseIf b.Text = "↓" Then
            ind = control_array(ComboBox1.Text.Trim.ToUpper).IndexOf(c)
            If ind = control_array(ComboBox1.Text.Trim.ToUpper).Count - 1 Then Exit Sub
            ind_swap = ind + 1
        End If

        Dim tmp_str As String = ""
        Dim tmp_cur As control_elements = control_array(ComboBox1.Text.Trim.ToUpper)(ind)
        Dim tmp_swap As control_elements = control_array(ComboBox1.Text.Trim.ToUpper)(ind_swap)

        tmp_str = tmp_cur.variable.Text : tmp_cur.variable.Text = tmp_swap.variable.Text : tmp_swap.variable.Text = tmp_str
        tmp_str = tmp_cur.variable_type.Text : tmp_cur.variable_type.Text = tmp_swap.variable_type.Text : tmp_swap.variable_type.Text = tmp_str
        tmp_str = tmp_cur.def.Text : tmp_cur.def.Text = tmp_swap.def.Text : tmp_swap.def.Text = tmp_str
        Dim f = tmp_cur.fulRow.Checked : tmp_cur.fulRow.Checked = tmp_swap.fulRow.Checked : tmp_swap.fulRow.Checked = f

        'tmp_cur.buttonUp.Tag = tmp_swap : tmp_cur.buttonDown.Tag = tmp_swap : tmp_cur.buttonDel.Tag = tmp_swap
        'tmp_swap.buttonUp.Tag = tmp_cur : tmp_swap.buttonDown.Tag = tmp_cur : tmp_swap.buttonDel.Tag = tmp_cur

        'control_array(ComboBox1.Text.Trim.ToUpper)(ind) = tmp_swap
        'control_array(ComboBox1.Text.Trim.ToUpper)(ind_swap) = tmp_cur
    End Sub

    'Utils
    Sub getNodesRecurcively(node As Xml.XmlNode, arr As List(Of String))
        For Each n As Xml.XmlNode In node.ChildNodes
            If n.NodeType <> Xml.XmlNodeType.Element Then Continue For

            Dim path As String = n.Name
            Dim cur = n
            Do While cur.ParentNode IsNot Nothing
                cur = cur.ParentNode
                If Not cur.Name.StartsWith("#") Then path = cur.Name + "/" + path Else path = "/" + path
            Loop
            If Not arr.Contains(path) Then arr.Add(path)

            getNodesRecurcively(n, arr)
        Next
    End Sub

    '...
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If ComboBox2.SelectedIndex < 0 Then Exit Sub

        If ComboBox2.SelectedItem.ToString.ToLower = "registry" Then
            RegistryBrowserForm.ShowDialog(Me)
            If RegistryBrowserForm.return_result.Trim <> "" Then ComboBox1.Text = RegistryBrowserForm.return_result.Trim
        End If
    End Sub
End Class