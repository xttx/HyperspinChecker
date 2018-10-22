Imports System.ComponentModel
Imports SevenZip

'TODO - Rocketlauncher Profiles
'TODO - RocketLauncher\Media\Keymapper\Controller Images
'TODO - Genres - only genre.xml is packed, need to pack additional xmls
'TODO - Strip paths from inis before packing
'TODO - handle rom in emu folder, emu in rom folder(!)
'TODO - Script...
'TODO - Disable controls when packing (including system changing and version/date/author textboxes)
'TODO - Option to check size

'TODO - romExtensions need to be saved per emulator
'TODO - forceRomInEmuFolder need to be saved per rom folder AND have to know, what emulator folder it belongs to (if multiple emulators)
'TODO - Emu folder will not be correctly detected, if emulator exe is in subfolder (i.e. Emulator\CoolEmu\Bin\x64\emu.exe)
'TODO - msgbox if not all media, required to play is set
'TODO - IsFullSystem param in generated script is always true yet (actually, this key is here just to give a better layout in Unpacking settings form. It can be replaced with something more useful)

'----------------------------------------------------------------------------------------------
'DONE - pack progress by media type (change icon to packed on already packed media)
'DONE - Get RL Emulator name and Module name
'DONE - Show "Force roms in emulator folder" option (if detected)
'DONE - pack per system emulators and per game emulators

Public Class FormJa_PackSystem
    Dim hs_path As String = Class1.HyperspinPath + "\"
    Dim rl_path As String = Class1.HyperlaunchPath + "\"

    Dim bg_arr As New Dictionary(Of CheckBox, BackgroundWorker)
    Dim bg_completed_counter As Integer = 0
    Dim pictureboxes As New Dictionary(Of CheckBox, PictureBox)
    Public bg_param As New Dictionary(Of CheckBox, checkerArguments)

    Dim zip_progress_media As New Dictionary(Of CheckBox, Integer)
    Dim zip_progress_media_files As New Dictionary(Of String, CheckBox)
    Dim zip_progress_media_lastFile As String = ""

    Dim emu_exe_by_name_dictionary As New Dictionary(Of String, String)
    Dim mod_ahk_by_name_dictionary As New Dictionary(Of String, String)
    Dim rom_extentions_by_emu_dictionary As New Dictionary(Of String, String)
    Dim force_roms_in_emu_path As New List(Of String)

    'Filters
    Dim gameListCache As New Dictionary(Of String, List(Of String))
    Dim game_filters_supporting_media As New List(Of CheckBox)

    Structure checkerArguments
        Dim path As String
        Dim path_packed As String
        Dim type As checkerArgumentsTypes
        Dim extensions() As String
        Dim checkbox As CheckBox
    End Structure
    Enum checkerArgumentsTypes
        file
        folder
    End Enum

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Dim img_ext = {"png", "jpg"}
        bg_param.Add(CheckBox1, New checkerArguments With {.path = hs_path + "Databases\%SYS%\%SYS%.xml", .type = checkerArgumentsTypes.file, .path_packed = "Database"})
        bg_param.Add(CheckBox2, New checkerArguments With {.path = hs_path + "Databases\%SYS%\Genre.xml", .type = checkerArgumentsTypes.file, .path_packed = "Database"})
        bg_param.Add(CheckBox3, New checkerArguments With {.path = hs_path + "Media\%SYS%\Images\Artwork1", .type = checkerArgumentsTypes.folder, .extensions = img_ext, .path_packed = "HS_Media"})
        bg_param.Add(CheckBox4, New checkerArguments With {.path = hs_path + "Media\%SYS%\Images\Artwork2", .type = checkerArgumentsTypes.folder, .extensions = img_ext, .path_packed = "HS_Media"})
        bg_param.Add(CheckBox5, New checkerArguments With {.path = hs_path + "Media\%SYS%\Images\Artwork3", .type = checkerArgumentsTypes.folder, .extensions = img_ext, .path_packed = "HS_Media"})
        bg_param.Add(CheckBox6, New checkerArguments With {.path = hs_path + "Media\%SYS%\Images\Artwork4", .type = checkerArgumentsTypes.folder, .extensions = img_ext, .path_packed = "HS_Media"})
        bg_param.Add(CheckBox8, New checkerArguments With {.path = hs_path + "Media\%SYS%\Images\Backgrounds", .type = checkerArgumentsTypes.folder, .extensions = img_ext, .path_packed = "HS_Media"})
        bg_param.Add(CheckBox9, New checkerArguments With {.path = hs_path + "Media\%SYS%\Images\Genre\Wheel", .type = checkerArgumentsTypes.folder, .extensions = img_ext, .path_packed = "HS_Media\Genre"})
        bg_param.Add(CheckBox40, New checkerArguments With {.path = hs_path + "Media\%SYS%\Images\Genre\Backgrounds", .type = checkerArgumentsTypes.folder, .extensions = img_ext, .path_packed = "HS_Media\Genre"})
        bg_param.Add(CheckBox10, New checkerArguments With {.path = hs_path + "Media\%SYS%\Images\Letters", .type = checkerArgumentsTypes.folder, .extensions = {"png"}, .path_packed = "HS_Media"})
        bg_param.Add(CheckBox11, New checkerArguments With {.path = hs_path + "Media\%SYS%\Images\Other\Pointer.png", .type = checkerArgumentsTypes.file, .path_packed = "HS_Media\Pointer"})
        bg_param.Add(CheckBox12, New checkerArguments With {.path = hs_path + "Media\%SYS%\Images\Particle", .type = checkerArgumentsTypes.folder, .extensions = {"*"}, .path_packed = "HS_Media"})
        bg_param.Add(CheckBox13, New checkerArguments With {.path = hs_path + "Media\%SYS%\Images\Special", .type = checkerArgumentsTypes.folder, .extensions = {"swf"}, .path_packed = "HS_Media"})
        bg_param.Add(CheckBox15, New checkerArguments With {.path = hs_path + "Media\%SYS%\Sound\Background Music", .type = checkerArgumentsTypes.folder, .extensions = {"mp3"}, .path_packed = "HS_Media\Sound"})
        bg_param.Add(CheckBox41, New checkerArguments With {.path = hs_path + "Media\%SYS%\Sound\System Start", .type = checkerArgumentsTypes.folder, .extensions = {"mp3"}, .path_packed = "HS_Media\Sound"})
        bg_param.Add(CheckBox43, New checkerArguments With {.path = hs_path + "Media\%SYS%\Sound\System Exit", .type = checkerArgumentsTypes.folder, .extensions = {"mp3"}, .path_packed = "HS_Media\Sound"})
        bg_param.Add(CheckBox42, New checkerArguments With {.path = hs_path + "Media\%SYS%\Sound\Wheel Sounds", .type = checkerArgumentsTypes.folder, .extensions = {"mp3"}, .path_packed = "HS_Media\Sound"})
        bg_param.Add(CheckBox44, New checkerArguments With {.path = hs_path + "Media\%SYS%\Sound\Wheel Click.mp3", .type = checkerArgumentsTypes.file, .path_packed = "HS_Media\Sound"})

        bg_param.Add(CheckBox14, New checkerArguments With {.path = hs_path + "Media\Main Menu\Images\Wheel\%SYS%.png", .type = checkerArgumentsTypes.file, .path_packed = "HS_Media\Main Menu\Wheel"})
        bg_param.Add(CheckBox16, New checkerArguments With {.path = hs_path + "Media\Main Menu\Themes\%SYS%.zip", .type = checkerArgumentsTypes.file, .path_packed = "HS_Media\Main Menu\Theme"})
        bg_param.Add(CheckBox18, New checkerArguments With {.path = hs_path + "Media\Main Menu\Video\%SYS%", .type = checkerArgumentsTypes.file, .extensions = {"flv", "mp4"}, .path_packed = "HS_Media\Main Menu\Video"})
        bg_param.Add(CheckBox17, New checkerArguments With {.path = hs_path + "Media\%SYS%\Images\Wheel", .type = checkerArgumentsTypes.folder, .path_packed = "HS_Media"})
        bg_param.Add(CheckBox19, New checkerArguments With {.path = hs_path + "Media\%SYS%\Themes", .type = checkerArgumentsTypes.folder, .extensions = {"zip"}, .path_packed = "HS_Media"})
        bg_param.Add(CheckBox20, New checkerArguments With {.path = hs_path + "Media\%SYS%\Video", .type = checkerArgumentsTypes.folder, .extensions = {"flv", "mp4"}, .path_packed = "HS_Media"})
        bg_param.Add(CheckBox39, New checkerArguments With {.path = hs_path + "Media\%SYS%\Video\Override Transitions", .type = checkerArgumentsTypes.folder, .extensions = {"flv", "mp4"}, .path_packed = "HS_Media\Video"})

        bg_param.Add(CheckBox23, New checkerArguments With {.path = rl_path + "Media\Artwork\%SYS%", .type = checkerArgumentsTypes.folder, .path_packed = "RL_Media\Artwork"})
        bg_param.Add(CheckBox25, New checkerArguments With {.path = rl_path + "Media\Backgrounds\%SYS%", .type = checkerArgumentsTypes.folder, .path_packed = "RL_Media\Backgrounds"})
        bg_param.Add(CheckBox26, New checkerArguments With {.path = rl_path + "Media\Bezels\%SYS%", .type = checkerArgumentsTypes.folder, .path_packed = "RL_Media\Bezels"})
        bg_param.Add(CheckBox27, New checkerArguments With {.path = rl_path + "Media\Controller\%SYS%", .type = checkerArgumentsTypes.folder, .path_packed = "RL_Media\Controller"})
        bg_param.Add(CheckBox28, New checkerArguments With {.path = rl_path + "Media\Fade\%SYS%", .type = checkerArgumentsTypes.folder, .path_packed = "RL_Media\Fade"})
        bg_param.Add(CheckBox29, New checkerArguments With {.path = rl_path + "Media\Guides\%SYS%", .type = checkerArgumentsTypes.folder, .path_packed = "RL_Media\Guides"})
        bg_param.Add(CheckBox30, New checkerArguments With {.path = rl_path + "Media\Icons\%SYS%.png", .type = checkerArgumentsTypes.file, .path_packed = "RL_Media\Icons"})
        bg_param.Add(CheckBox31, New checkerArguments With {.path = rl_path + "Media\Logos\%SYS%", .type = checkerArgumentsTypes.folder, .path_packed = "RL_Media\Logos"})
        bg_param.Add(CheckBox32, New checkerArguments With {.path = rl_path + "Media\Manuals\%SYS%", .type = checkerArgumentsTypes.folder, .path_packed = "RL_Media\Manuals"})
        bg_param.Add(CheckBox33, New checkerArguments With {.path = rl_path + "Media\MultiGame\%SYS%", .type = checkerArgumentsTypes.folder, .path_packed = "RL_Media\MultiGame"})
        bg_param.Add(CheckBox34, New checkerArguments With {.path = rl_path + "Media\Music\%SYS%", .type = checkerArgumentsTypes.folder, .path_packed = "RL_Media\Music"})
        bg_param.Add(CheckBox35, New checkerArguments With {.path = rl_path + "Media\Videos\%SYS%", .type = checkerArgumentsTypes.folder, .path_packed = "RL_Media\Videos"})
        bg_param.Add(CheckBox24, New checkerArguments With {.path = rl_path + "Data\Game Info\%SYS%.ini", .type = checkerArgumentsTypes.file, .path_packed = "Game Info\%SYS%.ini"})

        bg_param.Add(CheckBox38, New checkerArguments With {.path = rl_path + "Settings\%SYS%", .type = checkerArgumentsTypes.folder, .path_packed = "RL_Settings"})
        bg_param.Add(CheckBox22, New checkerArguments With {.path = hs_path + "Settings\%SYS%.ini", .type = checkerArgumentsTypes.file, .path_packed = "HS_Settings"})

        bg_param.Add(CheckBox7, New checkerArguments With {.path = "", .type = checkerArgumentsTypes.folder, .path_packed = "Emulator"})
        bg_param.Add(CheckBox21, New checkerArguments With {.path = "", .type = checkerArgumentsTypes.folder, .path_packed = "Roms", .extensions = {"*"}})
        bg_param.Add(CheckBox36, New checkerArguments With {.path = "", .type = checkerArgumentsTypes.file, .path_packed = "RL_Module"})

        game_filters_supporting_media.AddRange({CheckBox3, CheckBox4, CheckBox5, CheckBox6, CheckBox17, CheckBox19, CheckBox20, CheckBox23, CheckBox25, CheckBox26, CheckBox27, CheckBox28, CheckBox29, CheckBox31, CheckBox32, CheckBox33, CheckBox34, CheckBox35, CheckBox21})
    End Sub

    Private Sub FormJa_PackSystem_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For Each sys In Form1.ComboBox1.Items
            ComboBox1.Items.Add(sys)
        Next

        For Each chk In GroupBox1.Controls.OfType(Of CheckBox)
            Dim p As New PictureBox With {.SizeMode = PictureBoxSizeMode.StretchImage}
            p.Location = New Point(chk.Left + 130, chk.Top - 1)
            p.Size = New Size(19, 19)
            GroupBox1.Controls.Add(p)
            pictureboxes.Add(chk, p)

            Dim bg As New BackgroundWorker
            AddHandler bg.DoWork, AddressOf checkSystemBG
            bg_arr.Add(chk, bg)

            'Add this media type to filter enabler
            If game_filters_supporting_media.Contains(chk) Then CheckedListBox2.Items.Add(chk.Text, True)
        Next

        ComboBox1.SelectedIndex = Form1.ComboBox1.SelectedIndex
    End Sub

    Private Sub checkSystem() Handles ComboBox1.SelectedIndexChanged
        Button3.Enabled = False
        Button4.Enabled = False
        force_roms_in_emu_path.Clear()

        'Disable filters if we are selecting new system
        If CheckedListBox1.Tag Is Nothing OrElse Not CheckedListBox1.Tag.ToString = ComboBox1.SelectedItem.ToString.ToUpper.Trim Then
            RadioButton1.Checked = True
            CheckedListBox1.Tag = Nothing
        End If

        For Each c In pictureboxes.Keys
            c.Checked = False
            pictureboxes(c).Image = Nothing
        Next
        If ComboBox1.SelectedIndex < 0 Then Exit Sub

        ComboBox1.Enabled = False
        bg_completed_counter = 0

        'Autoset package path
        If TextBox1.Text.Contains("\") Then
            TextBox1.Text = TextBox1.Text.Substring(0, TextBox1.Text.LastIndexOf("\")) + "\" + ComboBox1.SelectedItem.ToString + ".HSPack.zip"
        Else
            TextBox1.Text = ".\SystemPackages\" + ComboBox1.SelectedItem.ToString + ".HSPack.zip"
        End If
        TextBox5.Text = ""

        Dim paths_arr = getEmuRomModulePaths() 'Returns array of lists {emu_names, emu_exes, emu_fldr, rom_exten, rom_paths, mod_paths}
        emu_exe_by_name_dictionary.Clear()
        mod_ahk_by_name_dictionary.Clear()
        rom_extentions_by_emu_dictionary.Clear()
        For i As Integer = 0 To paths_arr(0).Count - 1
            If paths_arr(0)(i) <> "" Then
                Dim emu_exe_relative_path = paths_arr(1)(i).Substring(IO.Path.GetDirectoryName(paths_arr(2)(i)).Length)
                If Not emu_exe_by_name_dictionary.ContainsKey(paths_arr(0)(i)) Then emu_exe_by_name_dictionary.Add(paths_arr(0)(i), emu_exe_relative_path)
                If Not mod_ahk_by_name_dictionary.ContainsKey(paths_arr(0)(i)) Then mod_ahk_by_name_dictionary.Add(paths_arr(0)(i), paths_arr(5)(i))
                If Not rom_extentions_by_emu_dictionary.ContainsKey(paths_arr(0)(i)) Then rom_extentions_by_emu_dictionary.Add(paths_arr(0)(i), paths_arr(3)(i))
            End If
        Next

        Dim param_emu = bg_param(CheckBox7)
        Dim param_rom = bg_param(CheckBox21)
        Dim param_mod = bg_param(CheckBox36)
        'If emu_path <> "" Then param_emu.path = IO.Path.GetDirectoryName(IO.Path.GetFullPath(emu_path)) Else param_emu.path = ""
        'param_rom.path = IO.Path.GetFullPath(rom_path)
        'param_mod.path = IO.Path.GetFullPath(mod_path)
        param_emu.path = String.Join("|", paths_arr(2))
        param_rom.path = String.Join("|", paths_arr(4))
        param_mod.path = String.Join("|", paths_arr(5))
        bg_param(CheckBox7) = param_emu
        bg_param(CheckBox21) = param_rom
        bg_param(CheckBox36) = param_mod

        For Each chk In GroupBox1.Controls.OfType(Of CheckBox)
            Do While bg_arr(chk).IsBusy
                Application.DoEvents()
            Loop

            bg_arr(chk).RunWorkerAsync(chk)
        Next
    End Sub
    Private Sub checkSystemBG(sender As Object, e As DoWorkEventArgs)
        Dim c = DirectCast(e.Argument, CheckBox)
        If Not bg_param.ContainsKey(c) Then Exit Sub

        Me.Invoke(Sub() pictureboxes(c).Image = My.Resources.saving)

        Dim param = bg_param(c)
        Dim sys = DirectCast(Me.Invoke(Function() ComboBox1.SelectedItem.ToString), String)

        If checkSystemBG_Check(param, sys) Then
            Me.Invoke(Sub() c.Enabled = True)
            Me.Invoke(Sub() c.Checked = True)
            Me.Invoke(Sub() pictureboxes(c).Image = My.Resources.check)
        Else
            Me.Invoke(Sub() c.Enabled = False)
            Me.Invoke(Sub() c.Checked = False)
            Me.Invoke(Sub() pictureboxes(c).Image = My.Resources.no)
        End If

        bg_completed_counter += 1
        If bg_completed_counter = bg_param.Count Then Me.Invoke(Sub() Button3.Enabled = True)
        If bg_completed_counter = bg_param.Count Then Me.Invoke(Sub() Button4.Enabled = True)
        If bg_completed_counter = bg_param.Count Then Me.Invoke(Sub() ComboBox1.Enabled = True)
        If bg_completed_counter = bg_param.Count Then Me.Invoke(Sub() ComboBox1.Focus())
    End Sub
    Public Function checkSystemBG_Check(param As checkerArguments, sys As String) As Boolean
        Dim path = param.path.Replace("%SYS%", sys)

        Dim res As Boolean = False
        For Each p In path.Split({"|"c}, StringSplitOptions.RemoveEmptyEntries)
            If param.type = checkerArgumentsTypes.file Then
                If param.extensions Is Nothing OrElse param.extensions.Count = 0 Then
                    If IO.File.Exists(p) Then res = True
                Else
                    For Each ext In param.extensions
                        If IO.File.Exists(p + "." + ext) Then res = True : Exit For
                    Next
                End If
            ElseIf param.type = checkerArgumentsTypes.folder Then
                If IO.Directory.Exists(p) Then
                    If param.extensions Is Nothing OrElse param.extensions.Count = 0 Then
                        res = True
                    Else
                        For Each ext In param.extensions
                            If IO.Directory.GetFiles(p, "*." + ext).Count > 0 Then res = True : Exit For
                        Next
                    End If
                End If
            End If
            If res Then Exit For
        Next

        Return res
    End Function
    Public Function getEmuRomModulePaths() As List(Of String)()
        Dim curEmuName As String = ""
        Dim ini_glb As New IniFileApi With {.path = rl_path + "Settings\Global Emulators.ini"}
        Dim ini_lcl As New IniFileApi With {.path = rl_path + "Settings\" + ComboBox1.SelectedItem.ToString + "\Emulators.ini"}
        Dim ini_gam As New IniFileApi With {.path = rl_path + "Settings\" + ComboBox1.SelectedItem.ToString + "\Games.ini"}

        Dim emu_names As New List(Of String)
        Dim emu_exes As New List(Of String)
        Dim emu_fldr As New List(Of String)
        Dim mod_paths As New List(Of String)
        Dim rom_exten As New List(Of String)
        Dim rom_paths = ini_lcl.IniReadValue("ROMS", "Rom_Path").Split({"|"c}, StringSplitOptions.RemoveEmptyEntries).ToList

        'Add default emulator
        curEmuName = ini_lcl.IniReadValue("ROMS", "Default_Emulator")
        If curEmuName <> "" Then
            emu_names.Add(ini_lcl.IniReadValue("ROMS", "Default_Emulator"))
            emu_exes.Add(ini_glb.IniReadValue(emu_names(0), "Emu_Path"))
            mod_paths.Add(ini_glb.IniReadValue(emu_names(0), "Module"))
            rom_exten.Add(ini_glb.IniReadValue(emu_names(0), "Rom_Extension"))
        End If

        'Add per system emulators
        For Each section In ini_lcl.IniListKey()
            If section.ToUpper = "ROMS" Then Continue For
            emu_names.Add(section)
            emu_exes.Add(ini_lcl.IniReadValue(section, "Emu_Path"))
            mod_paths.Add(ini_lcl.IniReadValue(section, "Module"))
            rom_exten.Add(ini_lcl.IniReadValue(section, "Rom_Extension"))
        Next

        'Add per game emulators
        For Each section In ini_gam.IniListKey()
            curEmuName = ini_gam.IniReadValue(section, "Emulator")
            If curEmuName <> "" Then
                emu_names.Add(curEmuName)
                emu_exes.Add(ini_glb.IniReadValue(curEmuName, "Emu_Path"))
                mod_paths.Add(ini_glb.IniReadValue(curEmuName, "Module"))
                rom_exten.Add(ini_glb.IniReadValue(curEmuName, "Rom_Extension"))
            End If
        Next

        'Normalize and cleanup all paths
        For i As Integer = 0 To emu_names.Count - 1
            If emu_exes(i).StartsWith(".") Then emu_exes(i) = rl_path + emu_exes(i)
            If emu_exes(i).EndsWith("\") Then emu_exes(i) = emu_exes(i).Substring(0, emu_exes(i).Length - 1)
            If emu_exes(i) <> "" Then emu_exes(i) = IO.Path.GetFullPath(emu_exes(i))
            If emu_exes(i) <> "" Then emu_fldr.Add(IO.Path.GetDirectoryName(emu_exes(i))) Else emu_fldr.Add("")

            If mod_paths(i).StartsWith(".") Then
                mod_paths(i) = rl_path + "Modules" + mod_paths(i).Substring(mod_paths(i).IndexOf("\"))
            ElseIf mod_paths(i) <> "" And Not IO.Path.IsPathRooted(mod_paths(i)) Then
                mod_paths(i) = rl_path + "Modules\" + IO.Path.GetFileNameWithoutExtension(mod_paths(i)) + "\" + mod_paths(i)
            End If

            If mod_paths(i) <> "" Then mod_paths(i) = IO.Path.GetFullPath(mod_paths(i))
        Next
        For i As Integer = 0 To rom_paths.Count - 1
            If rom_paths(i).StartsWith(".") Then rom_paths(i) = rl_path + rom_paths(i)
            If rom_paths(i).EndsWith("\") Then rom_paths(i) = rom_paths(i).Substring(0, rom_paths(i).Length - 1)
            If rom_paths(i) <> "" Then rom_paths(i) = IO.Path.GetFullPath(rom_paths(i))
        Next

        Return {emu_names, emu_exes, emu_fldr, rom_exten, rom_paths, mod_paths}
    End Function

    'Pack
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Button2.Enabled = False

        For Each c In pictureboxes.Keys
            If c.Checked Then
                pictureboxes(c).Image = My.Resources.saving
            End If
        Next

        Label7.Text = ""
        Label7.Visible = True
        ProgressBar1.Value = 0
        ProgressBar1.Visible = True

        'Dim emu_path = bg_param(CheckBox7).path.ToUpper     'TODO - multiple emu
        'Dim rom_path = bg_param(CheckBox21).path.ToUpper    'TODO - multiple roms
        'If rom_path.StartsWith(emu_path) Then
        '    Dim msg = "Roms are detected in emulator path. Should this be forced during unpack, or should user be allowed to change rom directory?" + vbCrLf
        '    msg += "Some emulators only support roms in their folder. 'Yes' - force unpacking roms to emu folder, 'No' - allow user change rom directory during install of this package."
        '    If MsgBox(msg, MsgBoxStyle.YesNo, "HS Packing system") = MsgBoxResult.Yes Then force_roms_in_emu_path = True
        'End If

        Dim emu_path = bg_param(CheckBox7).path     'TODO - multiple emu
        Dim rom_path = bg_param(CheckBox21).path    'TODO - multiple roms
        For Each ep In emu_path.Split({"|"c}, StringSplitOptions.RemoveEmptyEntries)
            Dim res As MsgBoxResult = 0
            For Each rp In rom_path.Split({"|"c}, StringSplitOptions.RemoveEmptyEntries)
                If rp.Trim.ToUpper.StartsWith(ep.Trim.ToUpper) Then
                    Dim msg = "Roms are detected in emulator path." + vbCrLf
                    msg += rp + vbCrLf
                    msg += "Should this be forced during unpack, or should user be allowed to change rom directory?" + vbCrLf
                    msg += "Some emulators only support roms in their folder. 'Yes' - force unpacking roms to emu folder, 'No' - allow user change rom directory during install of this package."
                    If res = 0 Then res = MsgBox(msg, MsgBoxStyle.YesNo, "HS Packing system")
                    If res = MsgBoxResult.Yes Then force_roms_in_emu_path.Add(IO.Path.GetFileName(rp) + "->" + IO.Path.GetFileName(ep) + rp.Substring(ep.Length))
                End If
            Next
        Next

        'Generate default script if needed
        Try
            GenerateDefaultScript()
        Catch ex As Exception
            MsgBox("The script is wrong formatted or incorrect save path provided. Aborting." + vbCrLf + vbCrLf + ex.Message)
            Exit Sub
        End Try

        Dim bg As New BackgroundWorker
        AddHandler bg.DoWork, AddressOf Button2_Click_Sub_BG
        AddHandler bg.RunWorkerCompleted, AddressOf Button2_Click_Sub_compete
        bg.RunWorkerAsync()
    End Sub
    Private Sub Button2_Click_Sub_BG(sender As Object, e As EventArgs)
        zip_progress_media.Clear()
        zip_progress_media_files.Clear()
        zip_progress_media_lastFile = ""

        Dim gameFiltersList = CheckedListBox1.Items.Cast(Of String).Select(Of String)(Function(s) s.ToUpper.Trim).ToList
        Dim gameFiltersEnabled = CheckedListBox2.CheckedItems.Cast(Of String).Select(Of String)(Function(s) s.ToUpper).ToList
        Dim sys = DirectCast(Me.Invoke(Function() ComboBox1.SelectedItem.ToString), String)
        Dim emu_path = bg_param(CheckBox7).path.ToUpper
        Dim rom_path = bg_param(CheckBox21).path.ToUpper
        Dim files As New Dictionary(Of String, String)
        For Each c In pictureboxes.Keys
            If c.Checked Then
                Dim param = bg_param(c)
                Dim path = param.path.Replace("%SYS%", sys)

                'Special case - module is a file to check, but a directory to pack
                If c Is CheckBox36 Then
                    param.type = checkerArgumentsTypes.folder
                    Dim newPath As New List(Of String)
                    For Each p In path.Split({"|"c}, StringSplitOptions.RemoveEmptyEntries)
                        newPath.Add(IO.Path.GetDirectoryName(p))
                    Next
                    path = String.Join("|", newPath)
                End If

                If param.type = checkerArgumentsTypes.file Then
                    If param.extensions Is Nothing OrElse param.extensions.Count = 0 Then
                        Dim f = path.Replace("\\", "\")
                        If IO.File.Exists(f) Then
                            Dim fName = IO.Path.GetFileName(f)
                            If c IsNot CheckBox36 And c IsNot CheckBox7 And c IsNot CheckBox21 Then
                                If IO.Path.GetFileNameWithoutExtension(fName).ToUpper = sys.ToUpper Then fName = "%SYS%" + IO.Path.GetExtension(fName)
                            End If
                            Dim zip_path = param.path_packed + "\" + fName
                            'Dim zip_path = param.path_packed
                            files.Add(zip_path, f)

                            zip_progress_media_files(zip_path.ToUpper) = c
                            If zip_progress_media.ContainsKey(c) Then zip_progress_media(c) += 1 Else zip_progress_media.Add(c, 1)
                        End If
                    Else
                        For Each ext In param.extensions
                            Dim f = (path + "." + ext).Replace("\\", "\")
                            If IO.File.Exists(f) Then
                                Dim fName = IO.Path.GetFileName(f)
                                If c IsNot CheckBox36 And c IsNot CheckBox7 And c IsNot CheckBox21 Then
                                    If IO.Path.GetFileNameWithoutExtension(fName).ToUpper = sys.ToUpper Then fName = "%SYS%" + IO.Path.GetExtension(fName)
                                End If
                                Dim zip_path = param.path_packed + "\" + fName
                                'Dim zip_path = param.path_packed + IO.Path.GetExtension(f)
                                files.Add(zip_path, f)

                                zip_progress_media_files(zip_path.ToUpper) = c
                                If zip_progress_media.ContainsKey(c) Then zip_progress_media(c) += 1 Else zip_progress_media.Add(c, 1)
                            End If
                        Next
                    End If
                ElseIf param.type = checkerArgumentsTypes.folder Then
                    Dim path_arr = path.Split({"|"c}, StringSplitOptions.RemoveEmptyEntries)
                    For Each p_el In path_arr
                        If IO.Directory.Exists(p_el) Then
                            Dim f_arr = IO.Directory.GetFiles(p_el, "*.*", IO.SearchOption.AllDirectories)
                            For Each f In f_arr
                                Dim skip = False

                                'Game Filters
                                If Not RadioButton1.Checked AndAlso game_filters_supporting_media.Contains(c) AndAlso gameFiltersEnabled.Contains(c.Text.ToUpper) Then
                                    Dim gameF = IO.Path.GetFileNameWithoutExtension(f)
                                    Dim gameD = IO.Path.GetFileName(IO.Path.GetDirectoryName(f))

                                    If gameD.ToUpper <> "_Default".ToUpper Then 'Skip RL _Default directory
                                        Dim ind = gameFiltersList.IndexOf(gameF.ToUpper.Trim)
                                        If ind = -1 Then ind = gameFiltersList.IndexOf(gameD.ToUpper.Trim)

                                        If ind <> -1 Then
                                            Dim checked = CheckedListBox1.GetItemCheckState(ind)
                                            If RadioButton2.Checked AndAlso checked <> CheckState.Checked Then skip = True 'Include mode
                                            If RadioButton3.Checked AndAlso checked = CheckState.Checked Then skip = True  'Exclude mode
                                        ElseIf RadioButton2.Checked Then
                                            'Don't include unrecognized items in include mode
                                            skip = True
                                        End If
                                    End If
                                End If

                                ''Check if roms are in emu folder
                                'If c Is CheckBox7 And rom_path.StartsWith(emu_path) Then 'TODO need to check EACH rom path, if multiple
                                '    'We are now packing the emulator
                                '    Dim cur_path = IO.Path.GetDirectoryName(f)
                                '    If cur_path.ToUpper.StartsWith(rom_path) Then Continue For
                                'End If
                                ''Check if emu is in roms folder
                                'If c Is CheckBox21 And emu_path.StartsWith(rom_path) Then 'TODO need to check EACH emu path, if multiple
                                '    'We are now packing roms
                                '    Dim cur_path = IO.Path.GetDirectoryName(f)
                                '    If cur_path.ToUpper.StartsWith(emu_path) Then Continue For
                                'End If

                                'I spent two f**** weeks to get the logic below work!
                                'It completely blowed my mind!!!

                                'Check if roms are in emu folder
                                If c Is CheckBox7 Then
                                    'We are now packing the emulator
                                    Dim cur_path = IO.Path.GetDirectoryName(f)
                                    For Each r In rom_path.Split({"|"c}, StringSplitOptions.RemoveEmptyEntries)
                                        'If this rom path is in current emulator folder
                                        If r.ToUpper.StartsWith(p_el.ToUpper) Then
                                            'If we are currently packing it - then skip
                                            If cur_path.ToUpper.StartsWith(r) Then skip = True : Exit For
                                        End If
                                    Next
                                End If
                                'Check if emu is in roms folder
                                If c Is CheckBox21 Then
                                    'We are now packing roms
                                    Dim cur_path = IO.Path.GetDirectoryName(f)
                                    For Each u In emu_path.Split({"|"c}, StringSplitOptions.RemoveEmptyEntries)
                                        'If this emulator is in current rom folder
                                        If u.ToUpper.StartsWith(p_el.ToUpper) Then
                                            'If we are currently packing it - then skip
                                            If cur_path.ToUpper.StartsWith(u) Then skip = True : Exit For
                                        End If
                                    Next
                                End If

                                'Case: roms are in emu folder
                                'Packing emul: \emulator
                                'Current file: \emulator\dataroms\rom.zip
                                '----------------------------------------

                                'Case: emulator is in rom folder
                                'Packing roms: \roms
                                'Current file: \roms\emulator\dice\dice.exe
                                'Current path: \roms\emulator\dice
                                'Conclusion: current path should not starts with emu_path
                                'I.E if "\roms\emulator\dice".StartsWith("\roms\emulator") Then - we are packing emulator, need to skip
                                '
                                'But if we have roms in emu folder
                                'Packing roms: \Emu\Data\Roms
                                'Current file: \Emu\Data\Roma\rom.zip
                                'Current path: \Emu\Data\Roms
                                'If (cur_path)"\Emu\Data\Roms".StartsWith((emu_path)"\Emu") - Condition meet, this will be skipped, But it must not!

                                If Not skip Then
                                    Dim zip_path = f.Substring(p_el.Length)
                                    If zip_path.StartsWith("\") Then zip_path = zip_path.Substring(1)

                                    Dim fName = IO.Path.GetFileName(p_el)
                                    If c IsNot CheckBox36 And c IsNot CheckBox7 And c IsNot CheckBox21 Then
                                        If IO.Path.GetFileNameWithoutExtension(fName).ToUpper = sys.ToUpper Then fName = "%SYS%" + IO.Path.GetExtension(fName)
                                        If IO.Path.GetFileNameWithoutExtension(zip_path).ToUpper = sys.ToUpper Then zip_path = IO.Path.GetDirectoryName(zip_path) + "\%SYS%" + IO.Path.GetExtension(zip_path)
                                    End If
                                    zip_path = param.path_packed + "\" + fName + "\" + zip_path
                                    files.Add(zip_path, f)

                                    zip_progress_media_files(zip_path.ToUpper) = c
                                    If zip_progress_media.ContainsKey(c) Then zip_progress_media(c) += 1 Else zip_progress_media.Add(c, 1)
                                End If
                            Next
                        End If
                    Next
                End If
            End If
        Next

        Dim dir = IO.Path.GetDirectoryName(TextBox1.Text)
        If Not IO.Directory.Exists(dir) Then
            Try
                IO.Directory.CreateDirectory(dir)
            Catch ex As Exception
                MsgBox(ex.Message) : Exit Sub
            End Try
        End If

        Dim isd = (dir + "\HS_Package.isd").Replace("\\", "\").Replace("\\", "\")
        If IO.File.Exists(isd) Then files.Add("HS_Package.isd", isd)

        If IO.File.Exists(dir + "\dummy") Then IO.File.Delete(dir + "\dummy")
        IO.File.Create(dir + "\dummy").Close()
        files.Add("~PackageParams\originalSystemName=" + sys, dir + "\dummy")
        If force_roms_in_emu_path.Count > 0 Then
            For Each f In force_roms_in_emu_path
                f = f.Replace("->", "==TO==")
                f = f.Replace("\", "]]]")
                files.Add("~PackageParams\forceRomInEmuFolder=" + f, dir + "\dummy")
            Next
        End If

        Dim zc As New SevenZipCompressor
        zc.CustomParameters.Add("mt", "on")
        zc.ArchiveFormat = OutArchiveFormat.SevenZip
        zc.CompressionMethod = CompressionMethod.Default
        zc.CompressionLevel = CompressionLevel.High
        AddHandler zc.Compressing, AddressOf Button2_Click_Sub_Progress
        AddHandler zc.FileCompressionStarted, AddressOf Button2_Click_Sub_Progress_File_Start
        AddHandler zc.FileCompressionFinished, AddressOf Button2_Click_Sub_Progress_File_End

        Try
            zc.CompressFileDictionary(files, TextBox1.Text)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        zc.ZipEncryptionMethod = Nothing
        zc = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()

        If IO.File.Exists(isd) Then IO.File.Delete(isd)
        If IO.File.Exists(dir + "\dummy") Then IO.File.Delete(dir + "\dummy")
    End Sub
    Private Sub Button2_Click_Sub_Progress(sender As Object, e As ProgressEventArgs)
        Me.Invoke(Sub() ProgressBar1.Value = CInt(e.PercentDone))
        Me.Invoke(Sub() Label7.Text = e.PercentDone.ToString + "%")
    End Sub
    Private Sub Button2_Click_Sub_Progress_File_Start(sender As Object, e As FileNameEventArgs)
        zip_progress_media_lastFile = e.FileName
    End Sub
    Private Sub Button2_Click_Sub_Progress_File_End(sender As Object, e As EventArgs)
        If zip_progress_media_files.ContainsKey(zip_progress_media_lastFile.ToUpper) Then
            Dim c = zip_progress_media_files(zip_progress_media_lastFile.ToUpper)
            zip_progress_media(c) -= 1
            If zip_progress_media(c) <= 0 Then pictureboxes(c).Image = My.Resources.smile
        End If
    End Sub
    Private Sub Button2_Click_Sub_compete()
        Label7.Visible = False
        ProgressBar1.Visible = False
        checkSystem()
        Button2.Enabled = True
    End Sub

    Private Sub GenerateDefaultScript()
        Dim dir = IO.Path.GetDirectoryName(TextBox1.Text)
        If Not IO.Directory.Exists(dir) Then IO.Directory.CreateDirectory(dir)
        Dim isd_path As String = dir + "\HS_Package.isd"

        Dim x As New Xml.XmlDocument
        TextBox5.Text = TextBox5.Text.Trim
        If TextBox5.Text = "" Then
            Dim script = ""
            script += "<?xml version=""1.0"" encoding=""UTF-8""?>" + vbCrLf
            script += "<INISCHEMA>" + vbCrLf
            script += "   <INIFILES>" + vbCrLf
            script += "   </INIFILES>" + vbCrLf
            script += "</INISCHEMA>" + vbCrLf
            x.LoadXml(script)
        Else
            x.LoadXml(TextBox5.Text)
        End If

        Dim node = x.SelectSingleNode("INISCHEMA/INIFILES")

        Dim n As Xml.XmlElement = x.CreateElement("INIFILE")
        n.SetAttribute("name", "PackageSettings") : n.SetAttribute("required", "false") : n.SetAttribute("iniType", "ini")
        'node.AppendChild(n)
        node.PrependChild(n) 'adds to the begining

        Dim n2 As Xml.XmlElement = x.CreateElement("INITYPE")
        n2.InnerText = "HS_Package"
        n.AppendChild(n2)

        Dim n3 As Xml.XmlElement = x.CreateElement("SECTIONS")
        n.AppendChild(n3)

        Dim n4 As Xml.XmlElement = x.CreateElement("SECTION")
        n4.SetAttribute("name", "Main")
        n3.AppendChild(n4)

        Dim n5 As Xml.XmlElement = x.CreateElement("SECTIONTYPE")
        n5.InnerText = "Global"
        n4.AppendChild(n5)

        Dim n6 As Xml.XmlElement = x.CreateElement("KEYS")
        n4.AppendChild(n6)

        'Author
        Dim k_auth As Xml.XmlElement = x.CreateElement("KEY")
        k_auth.SetAttribute("name", "Author")
        k_auth.SetAttribute("required", "false")
        k_auth.SetAttribute("nullable", "true")
        k_auth.SetAttribute("readonly", "true")
        n6.AppendChild(k_auth)
        Dim k_auth_type As Xml.XmlElement = x.CreateElement("KEYTYPE")
        k_auth_type.InnerText = "String"
        k_auth.AppendChild(k_auth_type)
        Dim k_auth_def As Xml.XmlElement = x.CreateElement("DEFAULT")
        k_auth_def.InnerText = TextBox2.Text.Trim
        k_auth.AppendChild(k_auth_def)

        'Comment
        Dim k_com As Xml.XmlElement = x.CreateElement("KEY")
        k_com.SetAttribute("name", "Comment")
        k_com.SetAttribute("required", "false")
        k_com.SetAttribute("nullable", "true")
        k_com.SetAttribute("readonly", "true")
        n6.AppendChild(k_com)
        Dim k_com_type As Xml.XmlElement = x.CreateElement("KEYTYPE")
        k_com_type.InnerText = "Text"
        k_com.AppendChild(k_com_type)
        Dim k_com_fullrow As Xml.XmlElement = x.CreateElement("FULLROW")
        k_com_fullrow.InnerText = "True"
        k_com.AppendChild(k_com_fullrow)
        Dim k_com_def As Xml.XmlElement = x.CreateElement("DEFAULT")
        k_com_def.InnerText = TextBox3.Text.Trim
        k_com.AppendChild(k_com_def)

        'Is Full System (all media required to play is in the package)
        Dim k_full As Xml.XmlElement = x.CreateElement("KEY")
        k_full.SetAttribute("name", "IsFullSystem")
        k_full.SetAttribute("required", "false")
        k_full.SetAttribute("nullable", "false")
        k_full.SetAttribute("readonly", "true")
        n6.AppendChild(k_full)
        Dim k_full_type As Xml.XmlElement = x.CreateElement("KEYTYPE")
        k_full_type.InnerText = "Boolean"
        k_full.AppendChild(k_full_type)
        Dim k_full_def As Xml.XmlElement = x.CreateElement("DEFAULT")
        k_full_def.InnerText = "True" 'TODO
        k_full.AppendChild(k_full_def)

        'Version/Date
        Dim k_ver As Xml.XmlElement = x.CreateElement("KEY")
        k_ver.SetAttribute("name", "Version")
        k_ver.SetAttribute("required", "false")
        k_ver.SetAttribute("nullable", "true")
        k_ver.SetAttribute("readonly", "true")
        n6.AppendChild(k_ver)
        Dim k_ver_type As Xml.XmlElement = x.CreateElement("KEYTYPE")
        k_ver_type.InnerText = "String"
        k_ver.AppendChild(k_ver_type)
        Dim k_ver_def As Xml.XmlElement = x.CreateElement("DEFAULT")
        k_ver_def.InnerText = TextBox4.Text.Trim
        k_ver.AppendChild(k_ver_def)

        'Emu params
        For Each kv In rom_extentions_by_emu_dictionary
            Dim k_ext As Xml.XmlElement = x.CreateElement("KEY")
            k_ext.SetAttribute("name", "EmuParams")
            k_ext.SetAttribute("required", "true")
            k_ext.SetAttribute("nullable", "false")
            k_ext.SetAttribute("hide", "true")
            k_ext.SetAttribute("emuName", kv.Key)
            n6.AppendChild(k_ext)
            Dim k_type As Xml.XmlElement = x.CreateElement("KEYTYPE")
            k_type.InnerText = "String"
            k_ext.AppendChild(k_type)
            Dim k_descr As Xml.XmlElement = x.CreateElement("DESCRIPTION")
            k_descr.InnerText = "Rom extensions for " + kv.Key + " emulator."
            k_ext.AppendChild(k_descr)
            Dim k_def As Xml.XmlElement = x.CreateElement("EXTENSIONS")
            k_def.InnerText = kv.Value.Trim
            k_ext.AppendChild(k_def)
            If emu_exe_by_name_dictionary.ContainsKey(kv.Key) Then
                Dim k_exe As Xml.XmlElement = x.CreateElement("EMUEXE")
                k_exe.InnerText = emu_exe_by_name_dictionary(kv.Key).Trim
                k_ext.AppendChild(k_exe)
            End If
            If mod_ahk_by_name_dictionary.ContainsKey(kv.Key) Then
                Dim module_rel_path = mod_ahk_by_name_dictionary(kv.Key).Trim
                module_rel_path = module_rel_path.Substring(rl_path.Length)
                If module_rel_path.StartsWith("\") Then module_rel_path = module_rel_path.Substring(1)
                module_rel_path = module_rel_path.Substring(module_rel_path.IndexOf("\"))

                Dim k_mod As Xml.XmlElement = x.CreateElement("MODULE")
                k_mod.InnerText = module_rel_path
                k_ext.AppendChild(k_mod)
            End If
        Next

        For Each forceParam In force_roms_in_emu_path
            Dim n7 As Xml.XmlElement = x.CreateElement("KEY")
            n7.SetAttribute("name", "ForceRomsInEmuFolder")
            n7.SetAttribute("required", "false")
            n7.SetAttribute("nullable", "false")
            n7.SetAttribute("hide", "true")
            n6.AppendChild(n7)

            Dim n8 As Xml.XmlElement = x.CreateElement("KEYTYPE")
            n8.InnerText = "String"
            n7.AppendChild(n8)

            Dim n9 As Xml.XmlElement = x.CreateElement("DEFAULT")
            'n9.InnerText = relative_rom_path
            n9.InnerText = forceParam
            n7.AppendChild(n9)
        Next

        Dim sb As New Text.StringBuilder
        Dim xs As New Xml.XmlWriterSettings With {.Indent = True, .IndentChars = "   ", .NewLineChars = vbCrLf, .NewLineHandling = Xml.NewLineHandling.Replace}
        Using xmlwriter = Xml.XmlWriter.Create(sb, xs)
            x.Save(xmlwriter)
        End Using

        'TextBox5.Text = x.OuterXml.Trim
        TextBox5.Text = sb.ToString
        x.Save(isd_path)
    End Sub

    'Customization script - construct button
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim f As New FormJa2_ScriptConstructor
        f.emuPath = bg_param(CheckBox7).path
        f.modPath = bg_param(CheckBox36).path
        f.ShowDialog(Me)
        TextBox5.Text = f.returnScript
    End Sub

    'FILTERS
    'Filters button
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If ComboBox1.SelectedIndex < 0 Then Exit Sub

        ComboBox1.Enabled = False
        If Not gameListCache.ContainsKey(ComboBox1.SelectedItem.ToString.Trim) Then
            Dim l = New List(Of String)
            gameListCache.Add(ComboBox1.SelectedItem.ToString.Trim, l)

            Dim x As New Xml.XmlDocument
            x.Load(hs_path + "Databases\" + ComboBox1.SelectedItem.ToString.Trim + "\" + ComboBox1.SelectedItem.ToString.Trim + ".xml")
            For Each gn As Xml.XmlNode In x.SelectNodes("/menu/game")
                Dim attr = gn.Attributes.GetNamedItem("name")
                If attr IsNot Nothing AndAlso attr.Value.Trim <> "" Then l.Add(attr.Value.Trim)
            Next
        End If

        If CheckedListBox1.Tag Is Nothing OrElse Not CheckedListBox1.Tag.ToString = ComboBox1.SelectedItem.ToString.ToUpper.Trim Then
            CheckedListBox1.Items.Clear()
            CheckedListBox1.Items.AddRange(gameListCache(ComboBox1.SelectedItem.ToString.Trim).ToArray)
            CheckedListBox1.Tag = ComboBox1.SelectedItem.ToString.ToUpper.Trim
        End If

        GroupBox2.Width = 500
        GroupBox2.Height = Me.Height - 70
        GroupBox2.Location = New Point(CInt((Me.Width / 2) - (GroupBox2.Width / 2)), 10)
        GroupBox2.Visible = True : GroupBox2.BringToFront()
    End Sub
    'Filters close button
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click, Button6.Click
        GroupBox2.Visible = False
        ComboBox1.Enabled = True
    End Sub
End Class