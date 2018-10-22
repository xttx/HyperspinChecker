Imports System.ComponentModel

Public Class FormJb_UnPackSystem
    'TODO - check multiple emulators / multiple roms folder / multiple modules case, with and without force rom to emu folder
    'TODO - check single and multiple rom folders with subfolders hyerarchy with and without ForceRomsInEmuFolder mode in "backup" and "rename" modes
    'TODO - check emu with no modules, modules with no emu, and roms with no emulator, emu with no roms cases

    'TODO - Expression to retrive a function from metakeyword: "replace\((".*?"),(".*?")\)"
    'TODO - Parser - ISD Parser does not parse XML and REGISTRY ini
    'TODO - Parser - Groups in isd are ignored
    'TODO - Parser - Only current %SystemName% ini file is added, need to check all ini/systems supported by module
    'TODO - Parser - Browse value buttons (for file/folder/process option type)
    'TODO - Parser - Don't add 'Rocket launcher modules' label in left button panel, if no RL ini are present
    'TODO - Parser - Add dummy label, on the place of browse button, if button is not present - otherwise layout can be messed

    '----------------------------------------------------------------------------------------------------------------
    'DONE - does unpacking a system with different system name work?
    'DONE - Don't create ROM folder if not needed (forced roms in emu folder)
    'DONE - Parser - need to handle multiple emu/rom/modules path - Need check

    '----------------------------------------------------------------------------------------------------------------
    'DONE - The whole RocketLauncher\Settings folder was backuped (probably when RL_Settings media was added)
    'DONE - If changing system name, some files (i.e. database) can not be detected in the zip anymore
    'DONE - Make romPath text box enabled again after changing package (textBox1.text)
    'DONE - Renaming multi paths emulators or roms code completely broken
    'DONE - ModulePath only kept for one module, if multiple modules are present
    'DONE - Default Emulator can be the wrong one, if multiple modules/emulator are present
    'DONE - RomExtensions from all emulators are aggregated, and all put to default emulator
    'DONE - Only last rom path is added to the system, if multiple rom folders are present
    'DONE - When renaming multi path emu/roms/module, don't need to rename those, which are not conflicting
    'DONE - After renaming EMULATOR folder, all roms forcedInEmuFolder should go to the NEW enulator folder
    'DONE - Only default emulator path and its module path is corrected. Need to change other emulators and their modules paths
    'DONE - Rom path is incorrectly set in HyperLaunch if emulator is renamed and roms are in emuFolder
    'DONE - Need separate handling for delete/backup/merge options for standart roms and forcedInEmuFolder roms
    'DONE - Add separate set of options for forcedRomsToEmFolder, because they should never be renamed
    'DONE - RomsInEmuFolder always shows as conflicting
    'DONE - After changing system name, combobox 'Set all conflicts to ..." no longer works
    'DONE - Show somewhere info about installed files and end paths
    'DONE - Better progress bar (current one stops at half)
    'DONE - Check system in background with progress bar
    'DONE - Ask to enable RL fade, if not active
    'DONE - After 'incorrect rom path' message, instalation progress bar is not removed
    'DONE - Allow user to change system order when adding a system to main menu
    'DONE - Disable controls while installing (including paths and system textboxes)
    'DONE - Change button background on the buttons on the flow panel, when this button is active
    'DONE - Roms end_path are expanded to the most deep subfolder. Need to keep them as they was in original setup
    'DONE - Check if emulator with this name is already set up in RocketLauncher GlobalEmulator.ini and add additional set of options to rename it, add as new or use existing, or overwrite (as it is now)
    'DONE - After unzipping a package, when selecting another (new) package without closing form, the previously parsed ini buttons (on the left side panel) does not resets
    'DONE - After unzipping a package, when selecting another (new) package, alert if not all options was set (i.e. emulator was not added in GlobalEmulator.ini)
    'DONE - After unzipping a package, when trying to install the same package second time, all ISD buttons are added another time as double
    'DONE - After unzipping a package, when changing emulator path, media_forceRomInEmuFolderParam variable can contains absolute path, when it MUST be relative, which makes it throw an error in function checkForcedRoms() (line 377) - IO.Path.GetFullPath((TextBox2.Text + "\" + media_forceRomInEmuFolderParam(z_folder.ToUpper)))
    'DONE - When trying to set incorrect emulator path (i.e. c:\aaa:\bbb) - throws error in the same place as above
    'DONE - Unbind tooltip from all descriptions labels, when ISD tab is removed
    'DONE - Parser - ISD Parser only display one module settings, if multiple modules are present
    'DONE - Parser - ISD Parser can't display by rom (%romname%) settings in RL Modules ISD
    'DONE - Parser - ISD Parser incorrectly writes MAME-type ini
    'DONE - Parser - When unpacking "Electronica BK Test" default RL module values are not set because it try set text= DESCT instead of VALUE
    'DONE - Parser - Textarea for package comment
    'DONE - Parser - Readonly fields
    'DONE - Parser - Tooltips for item descriptions

    Const PACKAGE_PARAM_PREFIX As String = "~PackageParams"

    Dim frm_pack As New FormJa_PackSystem
    Dim refr As Boolean = False
    Dim globalEmulatorsNeedsToBeSet As Boolean = False
    Dim toolTip As New CustomToolTip() With {.Owner = Me}
    Dim parser As New RL_ISD_Parser(toolTip)
    Dim flow As New FlowLayoutPanel
    Dim media_options As String() = {"Merge only new (no overwrite)", "Merge all (overwrite)", "Backup old and copy new", "Delete old and copy new"}
    Dim addSysSelector As New AddSystemToHSMainMenu

    Dim sys As String = ""
    Dim media_paths As New List(Of media_path)
    Dim media_paths_EmuRomModule_Refs(2) As Integer
    Dim media_paths_EmuRomModule_Paths() As List(Of String)
    Dim media_emuParams_FromSettingsISD As New Dictionary(Of String, String())
    Dim media_forceRomInEmuFolderParam As New Dictionary(Of String, String)
    Dim bg_check As New BackgroundWorker
    Dim WithEvents zip As SevenZip.SevenZipExtractor = Nothing

    Class media_path
        Public meda_name As String
        Public zip_path As String
        Public zip_files As List(Of String)
        Public zip_subfolder_names As List(Of String)
        Public unpack_path As String
        Public found As Boolean
        Public status As status
        Public ref As FormJa_PackSystem.checkerArguments
        Public associated_check As CheckBox
        Public associated_combo As ComboBox
        Public size As ULong
    End Class
    Enum status
        _Default
        NoConflict
        NeedMerge
    End Enum

    Private Sub FormJb_UnPackSystem_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        bg_check.WorkerSupportsCancellation = True
        AddHandler bg_check.DoWork, AddressOf checkMediaBG
        AddHandler bg_check.RunWorkerCompleted, AddressOf checkMediaBG_Complete

        ComboBox1.Items.AddRange(media_options)
        TextBox2.Text = (Class1.HyperspinPath + "\Emulators").Replace("\\", "\")
        TextBox3.Text = (Class1.HyperspinPath + "\Roms").Replace("\\", "\")

        flow.AutoScroll = True
        flow.WrapContents = False
        flow.Padding = New Padding(0)
        flow.FlowDirection = FlowDirection.TopDown

        flow.BorderStyle = BorderStyle.FixedSingle
        flow.Location = New Point(10, 10)
        flow.Size = New Size(200, GroupBox1.Height - 20)
        flow.Anchor = AnchorStyles.Left Or AnchorStyles.Top Or AnchorStyles.Bottom
        GroupBox1.Text = ""
        GroupBox1.Controls.Add(flow)

        Dim ll As New Label With {.Text = "Package Settings", .AutoSize = False, .Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleCenter}
        flow.Controls.Add(ll)

        'Create initial settings tabControl and button
        Dim kv_import = parser.createTabControl(GroupBox1, flow, "Import System Settings")
        kv_import.Key.TabPages.Add("Import System Settings")

        Dim tab = DirectCast(DirectCast(flow.Controls(1), Button).Tag, TabControl).TabPages(0)
        tab.Controls.Add(GroupBox_Paths)
        GroupBox_Paths.Location = New Point(10, 10)
        GroupBox_Paths.Width = tab.Width - 20
        GroupBox_Paths.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right

        tab.Controls.Add(GroupBox_Controls)
        GroupBox_Controls.Location = New Point(10, GroupBox_Paths.Top + GroupBox_Paths.Height + 10)
        GroupBox_Controls.Width = tab.Width - 20
        GroupBox_Controls.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right

        DirectCast(flow.Controls(1), Button).PerformClick()
    End Sub
    Private Sub FormJb_UnPackSystem_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If Cancel_if_global_Emulators_Needs_To_Be_Set() Then e.Cancel = True : Exit Sub
        parser.save_pending_ini()
    End Sub

    'Package selected
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs, Optional dontDetectSystem As Boolean = False) Handles TextBox1.TextChanged
        If refr Then Exit Sub
        If Cancel_if_global_Emulators_Needs_To_Be_Set() Then
            refr = True : TextBox1.Text = TextBox1.Tag.ToString : refr = False : Exit Sub
        End If

        sys = ""
        zip = Nothing
        'last_zip_file = ""
        media_paths.Clear()
        media_paths_EmuRomModule_Refs = {-1, -1, -1, -1}    '0 = Emu, 1 = rom, 2 = module, 3 = forced rom
        media_paths_EmuRomModule_Paths = {New List(Of String), New List(Of String), New List(Of String), New List(Of String)}
        media_forceRomInEmuFolderParam.Clear()
        Dim f_name = TextBox1.Text.Trim
        TextBox1.Tag = f_name

        'Remove all but two controls from flow panel, if some package was already installed
        RemoveISDButtons()

        'Remove conflict table (if it exists) from tabPage
        Dim tab = DirectCast(DirectCast(flow.Controls(1), Button).Tag, TabControl).TabPages(0)
        If tab.Controls.Count > 2 Then tab.Controls.RemoveAt(2) : GroupBox_Controls.Visible = False : toolTip.unBindAll()

        'Restore rom path textbox, if it was disabled by previous package
        If Not TextBox3.Enabled Then TextBox3.Enabled = True : TextBox3.Text = TextBox3.Tag.ToString : TextBox3.Tag = Nothing

        If Not IO.File.Exists(f_name) Then Exit Sub
        Try
            zip = New SevenZip.SevenZipExtractor(f_name)
            Dim tmp = zip.ArchiveFileData.Count 'This is just for init SevenZipExtractor to provoke it to throw error if something is wrong
        Catch ex As Exception
            zip = Nothing
            MsgBox(ex.Message) : Exit Sub
        End Try

        'Fill parameters - get original_system_name and media_forceRomInEmuFolderParam dictionary
        For Each info In zip.ArchiveFileData
            If info.IsDirectory Then Continue For
            CheckParam(info.FileName)
        Next

        'Detect system name from package filename
        If Not dontDetectSystem Then
            If sys <> "" Then
                TextBox4.Text = sys
            Else
                'We didn't found system param, so we are guessing system from filename
                TextBox4.Text = IO.Path.GetFileNameWithoutExtension(IO.Path.GetFileNameWithoutExtension(f_name))
            End If
        End If
        TextBox4.Text = TextBox4.Text.Trim

        checkMedia()
    End Sub
    Sub checkMedia()
        Button6.Visible = False
        PictureBox1.Visible = True

        If bg_check.IsBusy Then bg_check.CancelAsync()
        Do While bg_check.IsBusy
            'Threading.Thread.Sleep(10)
            Application.DoEvents()
        Loop
        bg_check.RunWorkerAsync()
    End Sub
    Sub checkMediaBG(sender As Object, e As DoWorkEventArgs)
        media_paths.Clear()
        media_paths_EmuRomModule_Refs = {-1, -1, -1, -1}    '0 = Emu, 1 = rom, 2 = module, 3 = forced rom
        media_paths_EmuRomModule_Paths = {New List(Of String), New List(Of String), New List(Of String), New List(Of String)}

        'Construct media_paths array from all available media in FormJa_PackSystem.bg_param
        For Each p In frm_pack.bg_param
            If e.Cancel Then Exit Sub
            Dim zip_path As String = p.Value.path_packed
            zip_path += "\" + IO.Path.GetFileName(p.Value.path)
            media_paths.Add(New media_path With {.meda_name = p.Key.Text, .zip_path = zip_path, .zip_files = New List(Of String), .zip_subfolder_names = New List(Of String), .found = False, .ref = p.Value})
        Next
        'Duplicate media_path entry for romsForcedInEmuFolder
        Dim romParam = frm_pack.bg_param(frm_pack.CheckBox21)
        media_paths.Insert(media_paths.Count - 1, New media_path With {.meda_name = "Roms (forced)", .zip_path = ":::", .zip_files = New List(Of String), .zip_subfolder_names = New List(Of String), .found = False, .ref = romParam})

        'Iterate through files in zip package and check conflicts (i.e. if some of media already exist in current HS path)
        If zip Is Nothing Then Exit Sub
        Dim nothing_found As Boolean = True
        Try
            For Each info In zip.ArchiveFileData
                If info.IsDirectory Then Continue For
                If info.FileName.ToUpper.StartsWith(PACKAGE_PARAM_PREFIX.ToUpper) Then Continue For

                For Each m In media_paths
                    If e.Cancel Then Exit Sub
                    If info.FileName.ToUpper.StartsWith(m.zip_path.ToUpper) Then

                        'Get Emu/Rom/Module path if it's an emu/rom/module file, and overwrite current media if forcedRom detected
                        Dim cur_m = Add_EmuRomModule_Path(m, info.FileName)

                        If Not cur_m.found Then
                            cur_m.found = True
                            nothing_found = False

                            'Set status NEED_MERGE in any case, but don't remove this status, if it was previously set by another conflicted emu/rom/module path in multiple paths
                            If frm_pack.checkSystemBG_Check(cur_m.ref, TextBox4.Text.Trim) Then
                                cur_m.status = status.NeedMerge
                            ElseIf cur_m.status = status._Default Then
                                cur_m.status = status.NoConflict
                            End If
                        End If
                        cur_m.zip_files.Add(info.FileName)
                        cur_m.size += info.Size
                    End If
                Next
            Next

            If nothing_found Then MsgBox("This is not valid HS System package.") : zip = Nothing : e.Result = "-1" : Exit Sub

            'Now, we have all emu paths filled, we need to recheck roms if forceRomInEmuFolderParam were detected
            checkForcedRoms()
            e.Result = "0"
        Catch ex As Exception
            MsgBox(ex.Message)
            'e.Result = ex.Message : Exit Sub
            e.Result = "-1" : Exit Sub
        End Try
    End Sub
    Sub checkMediaBG_Complete(sender As Object, e As RunWorkerCompletedEventArgs)
        If e.Result.ToString = "-1" Then Exit Sub
        If e.Cancelled Then Exit Sub

        refr = True
        If media_paths_EmuRomModule_Paths(1).Count = 0 Then
            If TextBox3.Enabled Then TextBox3.Tag = TextBox3.Text : TextBox3.Enabled = False

            If media_paths_EmuRomModule_Paths(3).Count = 0 Then
                TextBox3.Text = "This pack does not contain roms."
            Else
                TextBox3.Text = "All roms paths forced into Emulator folder in package settings."
            End If
        End If
        refr = False

        Dim tblWasCreated As Boolean = False
        Dim tbl As RL_ISD_Parser.DBLayoutPanel = Nothing
        Dim tab = DirectCast(DirectCast(flow.Controls(1), Button).Tag, TabControl).TabPages(0)
        GroupBox_Controls.Visible = True
        If tab.Controls.Count <= 2 Then
            Dim y = GroupBox_Controls.Top + GroupBox_Controls.Height + 10
            Dim grp As New GroupBox With {.Location = New Point(10, y), .Size = New Size(tab.Width - 20, tab.Height - y - 10)}
            grp.Anchor = AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top Or AnchorStyles.Bottom
            grp.Text = "The package contains following elements:"
            tab.Controls.Add(grp)

            tbl = New RL_ISD_Parser.DBLayoutPanel With {.AutoScroll = True, .ColumnCount = 3}
            'tbl.CellBorderStyle = TableLayoutPanelCellBorderStyle.Single 'Debug - To see layout
            tbl.ColumnStyles.Clear()
            tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25))
            tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 55))
            tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 10))
            tbl.Padding = New Padding(0, 0, SystemInformation.VerticalScrollBarWidth + 10, 0)
            grp.Controls.Add(tbl)
            tblWasCreated = True
        Else
            Dim grp = DirectCast(tab.Controls(2), GroupBox)
            tbl = DirectCast(grp.Controls(0), RL_ISD_Parser.DBLayoutPanel)
        End If

        'Iterate through all media, to add checkbox and options combobox to found media
        Dim r = 0
        tbl.SuspendLayout()
        For Each m In media_paths
            If Not tblWasCreated Then
                If m.found Then
                    m.associated_check = DirectCast(tbl.GetControlFromPosition(0, r), CheckBox)
                    Dim c = tbl.GetControlFromPosition(1, r)
                    If TypeOf c Is ComboBox Then m.associated_combo = DirectCast(c, ComboBox)
                    r = r + 1
                End If
            End If
            createConflictingMediaOptionsControls(m, tbl)
        Next

        If tblWasCreated Then
            'Add a dummy line at the end of the table, to keep layout when resizing table.
            tbl.Controls.Add(New Label() With {.Text = ""}) : tbl.Controls.Add(New Label() With {.Text = ""})
            tbl.Dock = DockStyle.Fill
        End If

        tbl.ResumeLayout()
        PictureBox1.Visible = False : Button6.Visible = True
    End Sub
    Function Add_EmuRomModule_Path(m As media_path, filename As String) As media_path
        Dim ind As Integer = -1
        Dim basePath As String = ""

        If m.ref.path_packed.ToUpper = "Emulator".ToUpper Then
            ind = 0 : basePath = TextBox2.Text
        ElseIf m.ref.path_packed.ToUpper = "Roms".ToUpper Then
            ind = 1 : basePath = TextBox3.Text : m.ref.extensions = {} 'Overwride extensions, or it will only detect files in the folder, and not the folder itself.
        ElseIf m.ref.path_packed.ToUpper = "RL_Module".ToUpper Then
            ind = 2 : basePath = Class1.HyperlaunchPath + "\Modules"
        End If

        If ind >= 0 Then
            Dim folder_name = filename.Substring(m.ref.path_packed.Length + 1)
            folder_name = folder_name.Substring(0, folder_name.IndexOf("\"))
            Dim path = (basePath + "\" + folder_name).Replace("\\", "\").Replace("\\", "\")

            'If it is Roms and actually forcedRoms - switch current media to appropriate entry 
            If ind = 1 AndAlso media_forceRomInEmuFolderParam.ContainsKey(folder_name.ToUpper) Then
                m = media_paths(media_paths.Count - 2) : ind = 3
            End If

            If media_paths_EmuRomModule_Refs(ind) < 0 Then media_paths_EmuRomModule_Refs(ind) = media_paths.IndexOf(m)

            If Not media_paths_EmuRomModule_Paths(ind).Contains(path, StringComparer.OrdinalIgnoreCase) Then
                media_paths_EmuRomModule_Paths(ind).Add(path)
                m.ref.path = path
                m.zip_subfolder_names.Add(folder_name)
                m.found = False 'We've just add another emu/rom or module path, we need to recheck it for conflicts
                m.ref.type = FormJa_PackSystem.checkerArgumentsTypes.folder 'Because module is a FILE type when packing
            End If
        End If

        Return m
    End Function
    Function CheckParam(zip_path As String) As Boolean
        If zip_path.ToUpper.StartsWith(PACKAGE_PARAM_PREFIX.ToUpper) Then
            Dim param = IO.Path.GetFileName(zip_path).Trim
            If param.ToUpper.StartsWith("forceRomInEmuFolder".ToUpper) Then
                param = param.Substring(param.IndexOf("=") + 1).Replace("]]]", "\")
                Dim param_arr = param.Split({"==TO=="}, StringSplitOptions.RemoveEmptyEntries)
                If param_arr.Count = 2 Then
                    media_forceRomInEmuFolderParam.Add(param_arr(0).ToUpper, param_arr(1))
                End If
            End If
            If param.ToUpper.StartsWith("originalSystemName".ToUpper) Then
                sys = param.Substring(param.IndexOf("=") + 1).Trim
            End If
            Return True
        End If
        Return False
    End Function
    Sub checkForcedRoms()
        If media_paths_EmuRomModule_Refs(3) = -1 Then Exit Sub
        If media_forceRomInEmuFolderParam.Count = 0 Then Exit Sub

        Dim forcedRomPath As New List(Of String)
        Dim romMedia = media_paths(media_paths_EmuRomModule_Refs(3))
        romMedia.ref.extensions = {} 'Overwride extensions, or it will only detect files in the folder, and not the folder itself.
        If romMedia.zip_subfolder_names.Count > 0 Then romMedia.status = status.NoConflict
        For i As Integer = 0 To media_paths_EmuRomModule_Paths(3).Count - 1
            Dim z_folder = romMedia.zip_subfolder_names(i)
            If media_forceRomInEmuFolderParam.ContainsKey(z_folder.ToUpper) Then
                'Dim targetRomFolder = IO.Path.GetFullPath(TextBox2.Text + "\" + media_forceRomInEmuFolderParam(z_folder.ToUpper))
                Dim targetRomFolder = (TextBox2.Text + "\" + media_forceRomInEmuFolderParam(z_folder.ToUpper)).Replace("\\", "\").Replace("\\", "\")
                For Each emu In media_paths_EmuRomModule_Paths(0)
                    If targetRomFolder.ToUpper.StartsWith(emu.ToUpper + "\") Then
                        romMedia.ref.path = targetRomFolder
                        forcedRomPath.Add(targetRomFolder)
                        Exit For
                    End If
                Next
            End If

            If frm_pack.checkSystemBG_Check(romMedia.ref, TextBox4.Text) Then
                romMedia.status = status.NeedMerge
            End If
        Next

        media_paths_EmuRomModule_Paths(3) = forcedRomPath
    End Sub

    'Change system name
    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        Dim tab = DirectCast(DirectCast(flow.Controls(1), Button).Tag, TabControl).TabPages(0)
        If tab.Controls.Count < 3 Then Exit Sub
        checkMedia()
    End Sub
    'Change emulator or rom path - recheck conflicts
    Private Sub TextBox_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged, TextBox3.TextChanged
        If refr Then Exit Sub
        Dim ind As Integer = -1
        Dim additional_combo_item As String = ""
        Dim t = DirectCast(sender, TextBox)

        If t Is TextBox2 Then
            'Emu path
            ind = 0 : additional_combo_item = "Rename Emulator Folder"
        ElseIf t Is TextBox3 Then
            'Rom path
            ind = 1 : additional_combo_item = "Rename Roms Folder"
        End If

        'If this media was not found - exiting.
        If media_paths.Count = 0 Then Exit Sub
        If ind < 0 OrElse media_paths_EmuRomModule_Refs(ind) < 0 Then Exit Sub
        If Not media_paths(media_paths_EmuRomModule_Refs(ind)).found Then Exit Sub
        If media_paths(media_paths_EmuRomModule_Refs(ind)).zip_files Is Nothing Then Exit Sub

        Dim f As New FormJa_PackSystem
        Dim m = media_paths(media_paths_EmuRomModule_Refs(ind))
        media_paths_EmuRomModule_Paths(ind).Clear()
        m.found = False
        m.status = status._Default
        For Each z In m.zip_files
            Add_EmuRomModule_Path(m, z) 'This sub will overwrite .found variable for each new path in multiple paths
            If Not m.found Then
                m.found = True

                'Set status NEED_MERGE in any case, but don't remove it if it is already set by another conflicted emu/rom/module path
                If f.checkSystemBG_Check(m.ref, TextBox4.Text) Then
                    m.status = status.NeedMerge
                ElseIf m.status = status._Default Then
                    m.status = status.NoConflict
                End If
            End If
        Next

        Dim arrToCheck As media_path() = {m}
        If ind = 0 AndAlso media_paths_EmuRomModule_Refs(3) > -1 Then checkForcedRoms() : arrToCheck = {m, media_paths(media_paths_EmuRomModule_Refs(3))}

        Dim tbl As RL_ISD_Parser.DBLayoutPanel = Nothing
        For Each mediaToCheck In arrToCheck
            If mediaToCheck.associated_check Is Nothing Then Continue For
            If tbl Is Nothing Then tbl = DirectCast(mediaToCheck.associated_check.Parent, RL_ISD_Parser.DBLayoutPanel) : tbl.SuspendLayout()
            createConflictingMediaOptionsControls(mediaToCheck, tbl)
        Next
        If tbl IsNot Nothing Then tbl.ResumeLayout()
    End Sub
    Sub createConflictingMediaOptionsControls(m As media_path, tbl As RL_ISD_Parser.DBLayoutPanel)
        If m.found Then
            Dim tip = ""
            Dim control2 As Control = Nothing

            'Construct toolTip and options items
            Dim comboboxItems As New List(Of String) : comboboxItems.AddRange(media_options)
            If m.meda_name.ToUpper = "Emulator".ToUpper Then
                comboboxItems.Add("Rename Emulator Folder")
                tip = String.Join(vbCrLf, media_paths_EmuRomModule_Paths(0))
            ElseIf m.meda_name.ToUpper = "Roms".ToUpper Then
                comboboxItems.Add("Rename Roms Folder")
                tip = String.Join(vbCrLf, media_paths_EmuRomModule_Paths(1))
            ElseIf m.meda_name.ToUpper = "Roms (forced)".ToUpper Then
                tip = String.Join(vbCrLf, media_paths_EmuRomModule_Paths(3))
            ElseIf m.meda_name.ToUpper = "RL Module".ToUpper Then
                comboboxItems.Add("Rename Module")
                tip = String.Join(vbCrLf, media_paths_EmuRomModule_Paths(2))
            Else
                tip = m.ref.path.Replace("%SYS%", TextBox4.Text.Trim)
            End If
            tip = "Path: " + vbCrLf + vbCrLf + tip.Replace("|", vbCrLf) + vbCrLf + "Size: " + (Math.Round(m.size / 1024 / 1024, 2)).ToString + "M"

            If m.status = status.NoConflict Then
                m.associated_combo = Nothing
                'control2 = New Label With {.Text = "No Conflict", .AutoSize = False, .Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleLeft}
                control2 = New Label With {.Text = "No Conflict", .AutoSize = False, .Anchor = AnchorStyles.Left Or AnchorStyles.Right, .TextAlign = ContentAlignment.MiddleLeft}
            ElseIf m.status = status.NeedMerge Then
                m.associated_combo = New ComboBox With {.DropDownStyle = ComboBoxStyle.DropDownList, .Anchor = AnchorStyles.Left Or AnchorStyles.Right}
                'm.associated_combo = New ComboBox With {.DropDownStyle = ComboBoxStyle.DropDownList, .Dock = DockStyle.Fill}
                m.associated_combo.Items.AddRange(comboboxItems.ToArray)
                If m.associated_combo.Items.Count > 4 Then m.associated_combo.SelectedIndex = 4 Else m.associated_combo.SelectedIndex = 0
                control2 = m.associated_combo
            End If

            'Check if controls already exist on the table
            If m.associated_check Is Nothing Then
                'Add options
                'm.associated_check = New CheckBox With {.Text = m.meda_name, .AutoSize = False, .Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleLeft, .CheckAlign = ContentAlignment.MiddleLeft, .Checked = True}
                m.associated_check = New CheckBox With {.Text = m.meda_name, .AutoSize = False, .Anchor = AnchorStyles.Left Or AnchorStyles.Right, .TextAlign = ContentAlignment.MiddleLeft, .CheckAlign = ContentAlignment.MiddleLeft, .Checked = True}
                tbl.Controls.Add(m.associated_check)
                tbl.Controls.Add(control2)

                'Add tooltip
                Dim p As New PictureBox With {.Width = 20, .Height = 20, .Dock = DockStyle.Right, .SizeMode = PictureBoxSizeMode.Zoom}
                p.Image = My.Resources.info
                tbl.Controls.Add(p)
                toolTip.Bind(p, tip, True)
            Else
                Dim pos = tbl.GetPositionFromControl(m.associated_check)

                'Change tooltip
                toolTip.Bind(tbl.GetControlFromPosition(2, pos.Row), tip, True)

                'Change options (combobox)
                If tbl.GetControlFromPosition(1, pos.Row).GetType Is control2.GetType Then
                    'No need to replace control if it is of the same type. 
                    'But we need to revert associated combo back, in this case
                    If m.associated_combo IsNot Nothing Then m.associated_combo = DirectCast(tbl.GetControlFromPosition(1, pos.Row), ComboBox)
                    Exit Sub
                End If

                tbl.Controls.Remove(tbl.GetControlFromPosition(1, pos.Row))
                tbl.Controls.Add(control2, 1, pos.Row)
            End If
        End If
    End Sub

    'Install package
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        'If Not IO.File.Exists(last_zip_file) Then Exit Sub
        If zip Is Nothing Then Exit Sub
        If Cancel_if_global_Emulators_Needs_To_Be_Set() Then Exit Sub

        flow.Enabled = False
        GroupBox_Paths.Enabled = False
        GroupBox_Controls.Enabled = False
        Dim tab = DirectCast(DirectCast(flow.Controls(1), Button).Tag, TabControl).TabPages(0)
        If tab.Controls.Count > 2 Then tab.Controls(2).Enabled = False

        'Remove all but two controls from flow panel, if this package was already installed
        RemoveISDButtons()

        Dim p As New RL_ISD_Parser.DBProgressBar
        p.Minimum = 0 : p.Maximum = 1001
        GroupBox_Controls.Controls.Add(p)
        p.Location = New Point(5, 15)
        p.Size = New Size(GroupBox_Controls.Width - 10, 30)
        p.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top
        p.BringToFront()

        Dim BG As New System.ComponentModel.BackgroundWorker
        AddHandler BG.DoWork, AddressOf Button6_Click_BG
        AddHandler BG.RunWorkerCompleted, AddressOf Button6_Click_BG_Complete
        BG.RunWorkerAsync(p)
    End Sub
    Private Sub Button6_Click_BG(sender As Object, e As System.ComponentModel.DoWorkEventArgs)
        Dim p = DirectCast(e.Argument, RL_ISD_Parser.DBProgressBar)
        e.Result = Nothing

        Dim emuPathRoot As String = DirectCast(Me.Invoke(Function() TextBox2.Text.Trim), String)
        Dim romPathRoot As String = DirectCast(Me.Invoke(Function() TextBox3.Text.Trim), String)
        Dim emuPathNeed = False
        Dim romPathNeed = False
        Dim settingFileProvided = False

        If media_paths_EmuRomModule_Refs(0) >= 0 Then
            If media_paths(media_paths_EmuRomModule_Refs(0)).associated_check.Checked Then
                If emuPathRoot = "" OrElse IO.Path.GetInvalidPathChars().Any(Function(c) emuPathRoot.Contains(c)) OrElse Not IO.Directory.Exists(IO.Path.GetPathRoot(emuPathRoot)) OrElse IO.File.Exists(emuPathRoot) Then
                    MsgBox("Invalid Emulator Path.") : Exit Sub
                Else
                    emuPathNeed = True
                End If
            End If
        End If

        If media_paths_EmuRomModule_Refs(1) >= 0 And TextBox3.Tag Is Nothing Then
            If media_paths(media_paths_EmuRomModule_Refs(1)).associated_check.Checked Then
                If romPathRoot = "" OrElse IO.Path.GetInvalidPathChars().Any(Function(c) romPathRoot.Contains(c)) OrElse Not IO.Directory.Exists(IO.Path.GetPathRoot(romPathRoot)) OrElse IO.File.Exists(romPathRoot) Then
                    MsgBox("Invalid Rom Path.") : Exit Sub
                Else
                    romPathNeed = True
                End If
            End If
        End If

        Try
            If emuPathNeed AndAlso Not IO.Directory.Exists(emuPathRoot) Then IO.Directory.CreateDirectory(emuPathRoot)
            If romPathNeed AndAlso Not IO.Directory.Exists(romPathRoot) Then IO.Directory.CreateDirectory(romPathRoot)
        Catch ex As Exception
            MsgBox(ex.Message) : Exit Sub
        End Try

        Dim emuExePath As New Dictionary(Of String, String)
        Dim romEndPath As New List(Of String)
        Dim romExtensionsArr As New List(Of String)
        Dim modules As New List(Of String())
        Dim pathToBackupOrDelete() As String = Nothing
        Dim sys = TextBox4.Text.Trim
        Dim zip_loc As New SevenZip.SevenZipExtractor(zip.FileName)
        Dim zipFilesCount = zip_loc.ArchiveFileNames.Count

        Dim tempDir = Class1.HyperspinPath + "\" + IO.Path.GetRandomFileName
        Do While IO.Directory.Exists(tempDir)
            tempDir = Class1.HyperspinPath + "\" + IO.Path.GetRandomFileName
        Loop
        IO.Directory.CreateDirectory(tempDir)
        AddHandler zip_loc.Extracting, AddressOf Button6_Click_BG_ExtractProgress
        zip_loc.ExtractArchive(IO.Path.GetFullPath(tempDir))

        p.Invoke(Sub() p.Value = 0)
        p.Invoke(Sub() p.Maximum = media_paths.Count * 10)

        Dim check_isd = Button6_Click_BG_CheckPackageISD(tempDir)
        Dim _forceRomInEmuFolderParam = check_isd(1).DeserializeToDictionary("???", "|")

        Dim RenameRootDir As New Dictionary(Of String, String)
        For Each m In media_paths
            If m.associated_check IsNot Nothing AndAlso m.associated_check.Checked Then
                Dim optionIndex As Integer = -1
                If m.associated_combo IsNot Nothing Then optionIndex = DirectCast(Me.Invoke(Function() m.associated_combo.SelectedIndex), Integer)

                'Special cases - per media
                If media_paths_EmuRomModule_Refs(0) >= 0 AndAlso m Is media_paths(media_paths_EmuRomModule_Refs(0)) Then
                    'Emulator
                    m.unpack_path = TextBox2.Text.Trim
                    pathToBackupOrDelete = media_paths_EmuRomModule_Paths(0).ToArray
                ElseIf media_paths_EmuRomModule_Refs(1) >= 0 AndAlso m Is media_paths(media_paths_EmuRomModule_Refs(1)) Then
                    'Roms
                    m.unpack_path = TextBox3.Text.Trim
                    romEndPath.AddRange(media_paths_EmuRomModule_Paths(1).ToArray)
                    pathToBackupOrDelete = media_paths_EmuRomModule_Paths(1).ToArray
                ElseIf media_paths_EmuRomModule_Refs(3) >= 0 AndAlso m Is media_paths(media_paths_EmuRomModule_Refs(3)) Then
                    'Roms forced in emu folder
                    m.unpack_path = ""
                    If check_isd(1) IsNot Nothing Then
                        Dim new_rom_paths As New List(Of String)
                        pathToBackupOrDelete = Nothing
                        For Each rom_path In media_paths_EmuRomModule_Paths(3)
                            Dim romSubFolder = IO.Path.GetFileName(rom_path)
                            'If media_forceRomInEmuFolderParam.ContainsKey(romSubFolder.ToUpper) Then
                            If _forceRomInEmuFolderParam.ContainsKey(romSubFolder.ToUpper) Then
                                'Dim targetRomFolder = media_forceRomInEmuFolderParam(romSubFolder.ToUpper)
                                Dim targetRomFolder = _forceRomInEmuFolderParam(romSubFolder.ToUpper)
                                For Each emu In media_paths_EmuRomModule_Paths(0)
                                    If targetRomFolder.ToUpper.StartsWith(emu.ToUpper) Then
                                        new_rom_paths.Add(targetRomFolder + "\" + romSubFolder)
                                        Exit For
                                    End If
                                Next
                            Else
                                new_rom_paths.Add(rom_path)
                            End If
                        Next
                        romEndPath.AddRange(new_rom_paths.ToArray)
                        pathToBackupOrDelete = new_rom_paths.ToArray
                    End If
                ElseIf media_paths_EmuRomModule_Refs(2) >= 0 AndAlso m Is media_paths(media_paths_EmuRomModule_Refs(2)) Then
                    'Module
                    m.unpack_path = m.ref.path
                    m.unpack_path = IO.Path.GetDirectoryName(m.unpack_path) 'Because it's a AHK file but we need a dir
                    pathToBackupOrDelete = media_paths_EmuRomModule_Paths(2).ToArray
                Else
                    'Any other media
                    m.unpack_path = m.ref.path
                    m.unpack_path = IO.Path.GetDirectoryName(m.unpack_path)
                    pathToBackupOrDelete = {m.ref.path}
                End If

                If optionIndex = 2 Then
                    'Backup
                    For Each path In pathToBackupOrDelete
                        path = path.Replace("%SYS%", sys).Replace("\\", "\").Replace("\\", "\")
                        Backup_FileOrFolder(path, False)
                    Next
                ElseIf optionIndex = 3 Then
                    'Delete all old
                    For Each path In pathToBackupOrDelete
                        path = path.Replace("%SYS%", sys).Replace("\\", "\").Replace("\\", "\")
                        If IO.File.Exists(path) Then IO.File.Delete(path)
                        If IO.Directory.Exists(path) Then IO.Directory.Delete(path, True)
                    Next
                ElseIf optionIndex = 4 Then
                    'Rename emu/roms/module
                    For Each path In pathToBackupOrDelete
                        If Not IO.Directory.Exists(path) Then Continue For

                        Dim c As Integer = 1
                        Do While IO.Directory.Exists(path + " (a" + c.ToString + ")")
                            c += 1
                        Loop
                        Dim newPath = path + " (a" + c.ToString + ")"
                        RenameRootDir.Add(path, newPath)
                    Next
                End If

                'Iterate through all files in media
                For Each zip_f In m.zip_files
                    If Not m.unpack_path.EndsWith("\") Then m.unpack_path += "\"
                    Dim unpack_file_path As String = Replace(zip_f, m.ref.path_packed, m.unpack_path, 1, 1, CompareMethod.Text)
                    unpack_file_path = unpack_file_path.Replace("%SYS%", sys).Replace("\\", "\").Replace("\\", "\")

                    Dim case_emu = media_paths_EmuRomModule_Refs(0) >= 0 AndAlso m Is media_paths(media_paths_EmuRomModule_Refs(0))
                    Dim case_rom = media_paths_EmuRomModule_Refs(1) >= 0 AndAlso m Is media_paths(media_paths_EmuRomModule_Refs(1))
                    Dim case_mod = media_paths_EmuRomModule_Refs(2) >= 0 AndAlso m Is media_paths(media_paths_EmuRomModule_Refs(2))
                    Dim case_romF = media_paths_EmuRomModule_Refs(3) >= 0 AndAlso m Is media_paths(media_paths_EmuRomModule_Refs(3))

                    If case_romF Then
                        'ForcedRom case - get new path from forcedRomToFolder param need to be set BEFORE renaming routine
                        'because forcedRomToFolder dictionary is based on the old, non-renamed emu path and need to be renamed after

                        'Check forceRomInEmuFolder
                        Dim romSubFolder = Replace(zip_f, m.ref.path_packed, "", 1, 1, CompareMethod.Text).TrimStart({"\"c})
                        romSubFolder = romSubFolder.Substring(0, romSubFolder.IndexOf("\"))
                        'If media_forceRomInEmuFolderParam.ContainsKey(romSubFolder.ToUpper) Then
                        If _forceRomInEmuFolderParam.ContainsKey(romSubFolder.ToUpper) Then
                            For Each emu In media_paths_EmuRomModule_Paths(0)
                                'Dim targetRomFolder = media_forceRomInEmuFolderParam(romSubFolder.ToUpper)
                                Dim targetRomFolder = _forceRomInEmuFolderParam(romSubFolder.ToUpper)
                                If targetRomFolder.ToUpper.StartsWith(emu.ToUpper) Then
                                    unpack_file_path = Replace(zip_f, m.ref.path_packed, targetRomFolder, 1, 1, CompareMethod.Text).Replace("\\", "\")
                                    Exit For
                                End If
                            Next
                        End If
                    End If

                    Dim original_unpack_file_before_renaming = unpack_file_path
                    For Each kv In RenameRootDir
                        If unpack_file_path.ToUpper.StartsWith(kv.Key.ToUpper + "\") Then
                            unpack_file_path = Replace(unpack_file_path, kv.Key.ToUpper + "\", kv.Value + "\", 1, 1, CompareMethod.Text).Replace("\\", "\")
                        End If
                    Next

                    'Special cases - per file
                    If case_emu AndAlso unpack_file_path.ToUpper.EndsWith(".EXE") Then
                        'Emulator EXE Case
                        For Each kv In media_emuParams_FromSettingsISD
                            If original_unpack_file_before_renaming.ToUpper.EndsWith(kv.Value(1).ToUpper) Then emuExePath.Add(kv.Key, unpack_file_path) : Exit For
                        Next
                    End If
                    If case_rom Or case_romF Then
                        'Roms Case
                        Dim romPath = IO.Path.GetDirectoryName(unpack_file_path)
                        'If Not romEndPath.Contains(romPath) Then romEndPath.Add(romPath) 'This adds rom path as deepest subfolder which is wrong

                        'Add current rom extension to list, which will be used if no rom extension provided in package settings
                        Dim ext = IO.Path.GetExtension(unpack_file_path).Trim
                        If ext.StartsWith(".") Then
                            ext = ext.Substring(1).ToLower
                            If Not romExtensionsArr.Contains(ext) Then romExtensionsArr.Add(ext)
                        End If
                    End If
                    If case_mod Then
                        'Module Case
                        If unpack_file_path.ToUpper.EndsWith(".AHK") Then
                            Dim module_dir_name = IO.Path.GetDirectoryName(unpack_file_path)
                            Dim module_file_name = IO.Path.GetFileNameWithoutExtension(unpack_file_path)
                            Dim module_rel_path = "..\" + IO.Path.GetFileName(IO.Path.GetDirectoryName(unpack_file_path)) + "\" + IO.Path.GetFileName(unpack_file_path)
                            modules.Add({module_dir_name, module_file_name, module_rel_path, unpack_file_path, original_unpack_file_before_renaming})

                            'modulePath = "..\" + IO.Path.GetFileName(IO.Path.GetDirectoryName(unpack_file_path)) + "\" + IO.Path.GetFileName(unpack_file_path)
                            'moduleFullPath = unpack_file_path
                        End If
                    End If
                    If m.ref.path_packed = "RL_Settings" And zip_f.ToUpper.EndsWith("Emulators.ini".ToUpper) Then settingFileProvided = True


                    Dim extract As Boolean = True
                    If optionIndex = 0 Then
                        'No overwrite
                        If IO.File.Exists(unpack_file_path) Then extract = False
                    Else
                        'Overwrite (1) or handled above (2 - backup, 3 - delete, 4 - emu/rom/module rename)
                        If IO.File.Exists(unpack_file_path) Then IO.File.Delete(unpack_file_path)
                    End If

                    If extract Then
                        Dim dir = IO.Path.GetDirectoryName(unpack_file_path)
                        If Not IO.Directory.Exists(dir) Then IO.Directory.CreateDirectory(dir)
                        IO.File.Move(tempDir + "\" + zip_f, unpack_file_path)
                    End If

                    'p.Invoke(Sub() p.Value += 1)
                Next

            ElseIf m.associated_check IsNot Nothing Then
                'p.Invoke(Sub() p.Value += m.zip_files.Count)
            End If

            p.Invoke(Sub() p.txt = "Step 2 of 2 - Moving media (" + CInt((p.Value + 1) / 10).ToString + " / " + media_paths.Count.ToString + ")")
            p.Invoke(Sub() p.Value += 10)
            p.Invoke(Sub() p.Value -= 1) 'Slow progressbar animation workaround
        Next

        'Send params for ISD parser
        Dim l As New List(Of String)
        l.Add(tempDir)
        l.AddRange(SetEmulators(emuPathNeed, emuPathRoot, emuExePath, romEndPath, romExtensionsArr, modules))
        e.Result = l
    End Sub
    Private Sub Button6_Click_BG_ExtractProgress(sender As Object, e As SevenZip.ProgressEventArgs)
        Dim p = GroupBox_Controls.Controls.OfType(Of RL_ISD_Parser.DBProgressBar).FirstOrDefault
        p.Invoke(Sub() p.txt = "Step 1 of 2 - Extracting archive (" + e.PercentDone.ToString + "%)")
        p.Invoke(Sub() p.Value = e.PercentDone * 10)
        p.Invoke(Sub() p.Value = p.Value - 1) 'Slow progressbar animation workaround
    End Sub
    Private Sub Button6_Click_BG_Complete(sender As Object, e As ComponentModel.RunWorkerCompletedEventArgs)
        'Remove progressbar
        Dim p = GroupBox_Controls.Controls.OfType(Of RL_ISD_Parser.DBProgressBar).FirstOrDefault
        If p IsNot Nothing Then GroupBox_Controls.Controls.Remove(p)

        'Reenable controls
        flow.Enabled = True
        GroupBox_Paths.Enabled = True
        GroupBox_Controls.Enabled = True
        Dim tab = DirectCast(DirectCast(flow.Controls(1), Button).Tag, TabControl).TabPages(0)
        If tab.Controls.Count > 2 Then tab.Controls(2).Enabled = True

        If e.Result Is Nothing Then Exit Sub
        Dim l = DirectCast(e.Result, List(Of String))

        parser.emuPath = l(1)
        parser.romPath = l(2)
        parser.romPathRel = l(3)
        If IO.File.Exists(l(0) + "\HS_Package.isd") Then
            parser.Parse(GroupBox1, flow, l(0) + "\HS_Package.isd")
        End If

        'Adding aditional options
        flow.Controls.OfType(Of Button)()(1).PerformClick() 'Click on packageSettings
        Dim tab_packageSettings = GroupBox1.Controls.OfType(Of TabControl)()(1) 'Get packageSettings tab
        Dim tabPage = tab_packageSettings.TabPages.Item(0)
        Dim tbl = tabPage.Controls.OfType(Of RL_ISD_Parser.DBLayoutPanel).FirstOrDefault
        Dim Cell0Right = tbl.GetControlFromPosition(0, 0).Right
        Dim Cell1Left = tbl.GetControlFromPosition(1, 0).Left
        Dim Cell1Width = tbl.GetControlFromPosition(1, 0).Width
        Dim Cell2Right = tbl.GetControlFromPosition(3, 0).Right
        Dim Cell3Left = tbl.GetControlFromPosition(4, 0).Left
        Dim rowCount = tbl.RowCount
        tbl.SuspendLayout()
        tbl.RowCount = tbl.RowCount + 5

        'Adding a dummy label to make space
        Dim tmpLbl = New Label() With {.Text = ""}
        tbl.Controls.Add(tmpLbl) : tbl.SetCellPosition(tmpLbl, New TableLayoutPanelCellPosition(0, rowCount)) : tbl.SetColumnSpan(tmpLbl, 6)

        'Adding "global emulator.ini" options
        Dim y = 30
        Dim g As New GroupBox With {.Text = "Adding emulators to RL ""Global Emulators.ini"":", .Anchor = AnchorStyles.Left Or AnchorStyles.Right}
        Dim Cell0Offset = g.Margin.Left + tbl.Padding.Left
        Dim ini_glb_emu = New IniFileApi With {.path = Class1.HyperlaunchPath + "\Settings\Global Emulators.ini"}
        Dim sections = ini_glb_emu.IniListKey().Select(Of String)(Function(s) s.ToUpper).ToList
        For Each emuSerialized In l(5).Split({"<*>"}, StringSplitOptions.RemoveEmptyEntries)
            globalEmulatorsNeedsToBeSet = True
            Dim emuParam = emuSerialized.Split({";"c})

            'Find new available name
            Dim c As Integer = 1
            Dim newAvailName = emuParam(0)
            Do While sections.Contains(newAvailName.ToUpper)
                newAvailName = emuParam(0) + " (" + c.ToString + ")" : c += 1
            Loop

            'Title
            Dim lbl_title As New Label With {.Text = emuParam(0), .AutoSize = False, .Height = 21, .TextAlign = ContentAlignment.MiddleRight, .Anchor = AnchorStyles.Left Or AnchorStyles.Top}
            g.Controls.Add(lbl_title) : lbl_title.Location = New Point(Cell0Right - lbl_title.Width - Cell0Offset, y)

            'Combobox or label (if no conflict)
            Dim conflict_control As Control = Nothing
            If Not sections.Contains(emuParam(0).ToUpper) Then
                conflict_control = New Label With {.Text = "No Conflict", .AutoSize = False, .Height = 21, .TextAlign = ContentAlignment.MiddleLeft, .Anchor = AnchorStyles.Left Or AnchorStyles.Top}
                g.Controls.Add(conflict_control) : conflict_control.Location = New Point(Cell1Left - Cell0Offset, y)
            Else
                Dim cmb_conflict As New ComboBox With {.DropDownStyle = ComboBoxStyle.DropDownList, .Width = 122, .Anchor = AnchorStyles.Left Or AnchorStyles.Top}
                cmb_conflict.Items.AddRange({"Overwrite", "Rename"})
                cmb_conflict.SelectedIndex = 1
                conflict_control = cmb_conflict
                g.Controls.Add(cmb_conflict) : cmb_conflict.Location = New Point(Cell1Left - Cell0Offset, y) : cmb_conflict.Width = Cell1Width
                AddHandler cmb_conflict.SelectedIndexChanged, Sub(o As Object, ea As EventArgs)
                                                                  Dim cmb = DirectCast(o, ComboBox)
                                                                  Dim arr = DirectCast(cmb.Tag, Object())
                                                                  Dim txt = DirectCast(arr(0), TextBox)
                                                                  txt.Text = DirectCast(arr(cmb.SelectedIndex + 1), String)
                                                              End Sub
            End If

            'Rename textbox
            Dim txt_rename As New TextBox With {.Text = newAvailName, .Anchor = AnchorStyles.Left Or AnchorStyles.Top}
            g.Controls.Add(txt_rename) : txt_rename.Location = New Point(Cell1Left - Cell0Offset + Cell1Width + 15, y)
            conflict_control.Tag = New Object() {txt_rename, emuParam(0), newAvailName}

            'Checkbox
            Dim chk As New CheckBox With {.Text = "", .Checked = True, .Anchor = AnchorStyles.Left Or AnchorStyles.Top}
            chk.Tag = New Object() {lbl_title, conflict_control, txt_rename, emuSerialized}
            g.Controls.Add(chk) : chk.Location = New Point(txt_rename.Location.X + txt_rename.Width + 15, y)

            y += 30
        Next

        'A button
        Dim b As New Button With {.Text = "Apply", .Anchor = AnchorStyles.Top Or AnchorStyles.Right}
        g.Controls.Add(b) : b.Location = New Point(g.Width - b.Width - 10, y)
        AddHandler b.Click, AddressOf SetEmulatorGlobalApplyButton

        g.Height = y + 35
        tbl.Controls.Add(g) : tbl.SetCellPosition(g, New TableLayoutPanelCellPosition(0, rowCount + 1)) : tbl.SetColumnSpan(g, 6)

        'RL Fade options
        Dim ini_lcl_rl = New IniFileApi With {.path = Class1.HyperlaunchPath + "\Settings\" + sys + "\RocketLauncher.ini"}
        Dim ini_glb_rl = New IniFileApi With {.path = Class1.HyperlaunchPath + "\Settings\Global RocketLauncher.ini"}
        g = New GroupBox With {.Text = "Rocket Launcher Fade:", .Height = 100, .Anchor = AnchorStyles.Left Or AnchorStyles.Right}

        'Fade IN
        Dim lbl_f1 As New Label With {.Text = "Fade in:", .AutoSize = True, .Anchor = AnchorStyles.Left Or AnchorStyles.Top}
        g.Controls.Add(lbl_f1) : lbl_f1.Location = New Point(Cell0Right - lbl_f1.Width - Cell0Offset, 30)
        Dim ci11 As New RL_ISD_Parser.control_info With {.associated_ini = ini_glb_rl, .associated_ini_section = "Fade", .key_name = "Fade_In"}
        Dim chk_f11 As New CheckBox With {.Text = "Global: ", .AutoSize = True, .Anchor = AnchorStyles.Left, .Tag = ci11, .CheckAlign = ContentAlignment.MiddleRight}
        If ini_glb_rl.IniReadValue("Fade", "Fade_In").Trim.ToLower = "true" Then chk_f11.Checked = True
        AddHandler chk_f11.CheckedChanged, AddressOf parser.Option_ValueChanged
        g.Controls.Add(chk_f11) : chk_f11.Location = New Point(Cell1Left - Cell0Offset, 29)
        Dim ci12 As New RL_ISD_Parser.control_info With {.associated_ini = ini_lcl_rl, .associated_ini_section = "Fade", .key_name = "Fade_In"}
        Dim lbl_f1a As New Label With {.Text = "Local:", .AutoSize = True, .Anchor = AnchorStyles.Left Or AnchorStyles.Top}
        g.Controls.Add(lbl_f1a) : lbl_f1a.Location = New Point(Cell2Right - lbl_f1a.Width - Cell0Offset, 30)
        Dim cmb1 As New ComboBox With {.DropDownStyle = ComboBoxStyle.DropDownList, .Anchor = AnchorStyles.Left Or AnchorStyles.Top, .Tag = ci12}
        cmb1.Items.AddRange({"true", "false", "use_global"}) : cmb1.Text = ini_lcl_rl.IniReadValue("Fade", "Fade_In")
        AddHandler cmb1.SelectedIndexChanged, AddressOf parser.Option_ValueChanged
        g.Controls.Add(cmb1) : cmb1.Location = New Point(Cell3Left + Cell0Offset, 27)

        'Fade Out
        Dim lbl_f2 As New Label With {.Text = "Fade out:", .AutoSize = True, .Anchor = AnchorStyles.Left Or AnchorStyles.Top}
        g.Controls.Add(lbl_f2) : lbl_f2.Location = New Point(Cell0Right - lbl_f2.Width - Cell0Offset, 60)
        Dim ci21 As New RL_ISD_Parser.control_info With {.associated_ini = ini_glb_rl, .associated_ini_section = "Fade", .key_name = "Fade_Out"}
        Dim chk_f21 As New CheckBox With {.Text = "Global: ", .AutoSize = True, .Anchor = AnchorStyles.Left, .Tag = ci21, .CheckAlign = ContentAlignment.MiddleRight}
        If ini_glb_rl.IniReadValue("Fade", "Fade_Out").Trim.ToLower = "true" Then chk_f21.Checked = True
        AddHandler chk_f21.CheckedChanged, AddressOf parser.Option_ValueChanged
        g.Controls.Add(chk_f21) : chk_f21.Location = New Point(Cell1Left - Cell0Offset, 59)
        Dim ci22 As New RL_ISD_Parser.control_info With {.associated_ini = ini_lcl_rl, .associated_ini_section = "Fade", .key_name = "Fade_Out"}
        Dim lbl_f2a As New Label With {.Text = "Local:", .AutoSize = True, .Anchor = AnchorStyles.Left Or AnchorStyles.Top}
        g.Controls.Add(lbl_f2a) : lbl_f2a.Location = New Point(Cell2Right - lbl_f2a.Width - Cell0Offset, 60)
        Dim cmb2 As New ComboBox With {.DropDownStyle = ComboBoxStyle.DropDownList, .Anchor = AnchorStyles.Left Or AnchorStyles.Top, .Tag = ci22}
        cmb2.Items.AddRange({"true", "false", "use_global"}) : cmb2.Text = ini_lcl_rl.IniReadValue("Fade", "Fade_Out")
        AddHandler cmb2.SelectedIndexChanged, AddressOf parser.Option_ValueChanged
        g.Controls.Add(cmb2) : cmb2.Location = New Point(Cell3Left + Cell0Offset, 57)


        'Dummy
        tmpLbl = New Label() With {.Text = ""}
        tbl.Controls.Add(tmpLbl) : tbl.SetCellPosition(tmpLbl, New TableLayoutPanelCellPosition(0, rowCount + 2)) : tbl.SetColumnSpan(tmpLbl, 6)
        tbl.Controls.Add(g) : tbl.SetCellPosition(g, New TableLayoutPanelCellPosition(0, rowCount + 3)) : tbl.SetColumnSpan(g, 6)
        'Final Dummy
        tmpLbl = New Label() With {.Text = ""}
        tbl.Controls.Add(tmpLbl) : tbl.SetCellPosition(tmpLbl, New TableLayoutPanelCellPosition(0, rowCount + 4)) : tbl.SetColumnSpan(tmpLbl, 6)

        tbl.ResumeLayout()
        'tbl.CellBorderStyle = TableLayoutPanelCellBorderStyle.Single 'DEBUG

        'Adding RL ISDs
        Dim ll = New Label With {.Text = "Rocket Launcher Module", .AutoSize = False, .Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleCenter}
        flow.Controls.Add(ll)

        For Each isd In l(4).Split({"|"c}, StringSplitOptions.RemoveEmptyEntries)
            If IO.File.Exists(isd) Then
                parser.modPath = IO.Path.GetDirectoryName(isd)
                parser.Parse(GroupBox1, flow, isd, TextBox4.Text)
            End If
        Next

        'Cleanup
        IO.Directory.Delete(l(0), True)

        'Show add-system-to-database prompt
        Dim x As New Xml.XmlDocument
        x.Load(Class1.HyperspinPath + "\Databases\Main Menu\Main Menu.xml")
        If x.SelectSingleNode("/menu/game[@name='" + TextBox4.Text.Trim + "']") Is Nothing Then
            addSysSelector.show(GroupBox1, TextBox4.Text.Trim)
        End If
    End Sub
    Private Function Button6_Click_BG_CheckPackageISD(tempDir As String) As String()
        media_emuParams_FromSettingsISD.Clear()
        'media_forceRomInEmuFolderParam.Clear()
        Dim res As String() = {Nothing, Nothing}

        If IO.File.Exists(tempDir + "\HS_Package.isd") Then
            Dim isd As New Xml.XmlDocument
            isd.Load(tempDir + "\HS_Package.isd")

            Dim node = isd.SelectSingleNode("INISCHEMA/INIFILES/INIFILE[@name='PackageSettings']")
            If node IsNot Nothing Then
                Dim node_MainSection = node.SelectSingleNode("SECTIONS/SECTION[@name='Main']")
                If node_MainSection IsNot Nothing Then
                    'Check emu params for rocketlauncher emulator setup
                    Dim nodeRomExtensionsString = node_MainSection.SelectNodes("KEYS/KEY[@name='EmuParams']")
                    If nodeRomExtensionsString IsNot Nothing AndAlso nodeRomExtensionsString.Count > 0 Then
                        Dim c = 0
                        For Each emuParam As Xml.XmlNode In nodeRomExtensionsString
                            Dim emuNameNode = emuParam.Attributes.GetNamedItem("emuName")
                            Dim emuExeNode = emuParam.SelectSingleNode("EMUEXE")
                            Dim emuExtNode = emuParam.SelectSingleNode("EXTENSIONS")
                            Dim emuModNode = emuParam.SelectSingleNode("MODULE")
                            If emuExeNode IsNot Nothing AndAlso emuNameNode IsNot Nothing AndAlso emuModNode IsNot Nothing Then
                                media_emuParams_FromSettingsISD.Add(emuNameNode.Value.ToUpper, {emuNameNode.Value, emuExeNode.InnerText.Trim, emuExtNode.InnerText.Trim, emuModNode.InnerText.Trim, c.ToString})
                            End If

                            If emuExtNode IsNot Nothing AndAlso emuExtNode.InnerText.Trim <> "" Then
                                If res(0) Is Nothing Then res(0) = ""
                                res(0) += emuNameNode.Value + ":" + emuExtNode.InnerText.Trim + ";"
                            End If
                            c += 1
                        Next
                    End If

                    'Check if rom path is forced emu path
                    Dim nodesForceRomInEmuPath = node_MainSection.SelectNodes("KEYS/KEY[@name='ForceRomsInEmuFolder']")
                    If nodesForceRomInEmuPath IsNot Nothing AndAlso nodesForceRomInEmuPath.Count > 0 Then
                        For Each force As Xml.XmlNode In nodesForceRomInEmuPath
                            Dim node_Value = force.SelectSingleNode("DEFAULT")
                            If node_Value IsNot Nothing AndAlso node_Value.InnerText.ToUpper.Trim <> "" Then
                                'Example value: roms-&gt;dice.0.9.win.x64\roms
                                Dim val_arr = node_Value.InnerText.Trim.Split({"->"}, StringSplitOptions.RemoveEmptyEntries)
                                If val_arr.Count = 2 Then
                                    Dim rom_relative_path = val_arr(1).Trim
                                    If Not rom_relative_path.StartsWith("\") Then rom_relative_path = "\" + rom_relative_path
                                    If rom_relative_path.EndsWith("\") Then rom_relative_path = rom_relative_path.Substring(rom_relative_path.Length - 1)

                                    If res(1) Is Nothing Then res(1) = ""
                                    res(1) += val_arr(0).ToUpper + "|" + IO.Path.GetDirectoryName(TextBox2.Text.Trim + rom_relative_path) + "???"
                                    'res(1) += val_arr(0) + "|" + IO.Path.GetDirectoryName(TextBox2.Text.Trim + rom_relative_path) + ";"
                                    'media_forceRomInEmuFolderParam.Add(val_arr(0).ToUpper, IO.Path.GetDirectoryName(TextBox2.Text.Trim + rom_relative_path))
                                End If
                            End If
                        Next
                    End If
                End If
            End If
        End If

        Return res
    End Function
    Private Function SetEmulators(emuPathNeed As Boolean, emuPathRoot As String, emuExePath As Dictionary(Of String, String), romEndPath As List(Of String), romExtensionsArr As List(Of String), modules As List(Of String())) As List(Of String)
        Dim ini = New IniFileApi With {.path = Class1.HyperlaunchPath + "\Settings\" + sys + "\Emulators.ini"}
        'Dim defaultEmulatorName = ""
        Dim localEmulatorsList As New List(Of String)
        'If settingFileProvided Then
        '    defaultEmulatorName = ini.IniReadValue("ROMS", "Default_Emulator")
        'End If

        'Set rom path
        Dim romEndPathRel = romEndPath.ToList
        If romEndPathRel.Count > 0 Then
            For i As Integer = 0 To romEndPathRel.Count - 1
                romEndPathRel(i) = Absolute_Path_to_Relative((Class1.HyperlaunchPath + "\").Replace("\\", "\"), romEndPathRel(i))
            Next
            ini.IniWriteValue("ROMS", "Rom_Path", String.Join("|", romEndPathRel)) 'Relative to HL folder, i.e. ..\Roms\Consoles\Sega SG-1000
        End If

        'Set default emulator
        'If defaultEmulatorName <> "" Then ini.IniWriteValue("ROMS", "Default_Emulator", defaultEmulatorName)
        'TODO - set default emulator, in case the package contains an emulator, but not RL Settings

        'Set local emulators paths/romExt/module
        If emuPathNeed And modules.Count > 0 Then
            For Each section In ini.IniListKey()
                If section.ToUpper = "ROMS" Then Continue For
                localEmulatorsList.Add(section.ToUpper)

                Dim param = SetEmulatorsConstructParam(section, emuExePath(section.ToUpper), media_emuParams_FromSettingsISD(section.ToUpper)(2), romExtensionsArr, modules, media_emuParams_FromSettingsISD(section.ToUpper)(3))
                SetEmulatorInIniFromConstructedString(ini, param)
                'SetEmulatorInIni(ini, section, emuExePath, media_emuParams_FromSettingsISD, romExtensionsArr, modules)
            Next
        End If

        ''media_emuParams: 0 - emuName, 1 - emuExeNode, 2 - emuExtNode, 3 - emuModNode, 4 - index
        ''Modules list: 0 - module_dir_name, 1 - module_file_name_no_EXT, 2 - module_rel_path, 3 - module_full_path, 4 - full_path_orig_before_renaming

        'Set global emulator paths/romExt/module
        ini = New IniFileApi With {.path = Class1.HyperlaunchPath + "\Settings\Global Emulators.ini"}
        Dim globalEmulators = ""
        For Each foundEmu In emuExePath.Keys
            If localEmulatorsList.Contains(foundEmu) Then Continue For

            Dim name = media_emuParams_FromSettingsISD(foundEmu)(0)
            globalEmulators += SetEmulatorsConstructParam(name, emuExePath(name.ToUpper), media_emuParams_FromSettingsISD(name.ToUpper)(2), romExtensionsArr, modules, media_emuParams_FromSettingsISD(name.ToUpper)(3)) + "<*>"
            'globalEmulators += SetEmulatorInIni(ini, media_emuParams_FromSettingsISD(foundEmu)(0), emuExePath, media_emuParams_FromSettingsISD, romExtensionsArr, modules) + "<*>"
        Next

        'Create params for ISD parser
        Dim emu_combined_paths As New List(Of String)
        Dim modules_combined_paths As New List(Of String)
        Dim media_emuParams_by_index As New Dictionary(Of Integer, String())
        For Each kv In media_emuParams_FromSettingsISD
            'Dim cur_emu_path As String = "NOT_DEFINED"
            Dim cur_module_path As String = "NOT_DEFINED"

            For Each mod_param In modules
                If mod_param(4).ToUpper.EndsWith(kv.Value(3).ToUpper) Then
                    cur_module_path = IO.Path.GetDirectoryName(mod_param(3)) + "\" + IO.Path.GetFileNameWithoutExtension(mod_param(3)) + ".ISD"
                    Exit For
                End If
            Next

            media_emuParams_by_index.Add(CInt(kv.Value(4)), {emuPathRoot, cur_module_path})
        Next
        For i As Integer = 0 To media_emuParams_by_index.Count - 1
            emu_combined_paths.Add(media_emuParams_by_index(i)(0))
            modules_combined_paths.Add(media_emuParams_by_index(i)(1))
        Next

        'Send params for ISD parser
        Dim l As New List(Of String)
        'l.Add(tempDir)
        l.Add(String.Join("|", emu_combined_paths))
        l.Add(String.Join("|", romEndPath))
        l.Add(String.Join("|", romEndPathRel))
        l.Add(String.Join("|", modules_combined_paths))
        l.Add(globalEmulators)
        Return l
    End Function
    Private Function SetEmulatorsConstructParam(name As String, emuExePath As String, romExtensionsOrig As String, romExtensionsNew As List(Of String), modules As List(Of String()), moduleFromISD As String) As String
        Dim str = name + ";"

        str += emuExePath.Trim + ";"

        If romExtensionsOrig.Trim <> "" Then
            str += romExtensionsOrig.Trim
        ElseIf romExtensionsNew.Count > 0 Then
            str += String.Join("|", romExtensionsNew)
        End If
        str += ";"

        'Set module
        'Modules list: 0 - module_dir_name, 1 - module_file_name_no_EXT, 2 - module_rel_path, 3 - module_full_path, 4 - full_path_orig_before_renaming
        For Each mod_param In modules
            If mod_param(4).Trim <> "" AndAlso moduleFromISD.Trim <> "" AndAlso mod_param(4).ToUpper.EndsWith(moduleFromISD.Trim.ToUpper) Then
                str += mod_param(2).Trim : Exit For
            End If
        Next

        Return str
    End Function
    Private Sub SetEmulatorInIniFromConstructedString(ini As IniFileApi, str As String)
        Dim emuParam = str.Split(";"c)
        If emuParam.Count <> 4 Then Exit Sub

        Dim name = emuParam(0).Trim
        Dim path = emuParam(1).Trim
        Dim ext = emuParam(2).Trim
        Dim modl = emuParam(3).Trim
        If name <> "" Then
            If path <> "" Then ini.IniWriteValue(name, "Emu_Path", Absolute_Path_to_Relative((Class1.HyperlaunchPath + "\").Replace("\\", "\"), path)) 'Relative to HL folder, i.e. ..\Emulators\Arcade\dice.0.9.win.x64\dice.exe
            If ext <> "" Then ini.IniWriteValue(name, "Rom_Extension", ext)
            If modl <> "" Then ini.IniWriteValue(name, "Module", modl) 'Relative to module folder itself, i.e. ..\Phoenix-Altmer\phoenix-altmer.ahk
        End If
    End Sub
    Private Sub SetEmulatorGlobalApplyButton(o As Object, e As EventArgs)
        Dim b As Button = DirectCast(o, Button)
        Dim ini = New IniFileApi With {.path = Class1.HyperlaunchPath + "\Settings\Global Emulators.ini"}
        For Each chk In b.Parent.Controls.OfType(Of CheckBox)
            If chk.Checked Then
                Dim arr = DirectCast(chk.Tag, Object())
                'INFO: arr = New Object() {lbl_title, conflict_control, txt_rename, emuSerialized}
                'conflict_control can be either combobox or label (if no conflicts)

                Dim emuParam = DirectCast(arr(3), String).Split(";"c)
                emuParam(0) = DirectCast(arr(2), TextBox).Text.Trim
                SetEmulatorInIniFromConstructedString(ini, String.Join(";", emuParam))
            End If
        Next

        b.Text = "Done"
        b.Parent.Enabled = False
        globalEmulatorsNeedsToBeSet = False
    End Sub
    'OBSOLETE
    Private Function SetEmulatorInIni(ini As IniFileApi, name As String, emuExePath As Dictionary(Of String, String), media_emuParams As Dictionary(Of String, String()), romExtensionsArr As List(Of String), modules As List(Of String())) As String
        Dim str = name + ";"

        'Set emulator exe
        If emuExePath(name.ToUpper).Trim <> "" Then
            str += emuExePath(name.ToUpper).Trim
            ini.IniWriteValue(name, "Emu_Path", Absolute_Path_to_Relative((Class1.HyperlaunchPath + "\").Replace("\\", "\"), emuExePath(name.ToUpper))) 'Relative to HL folder, i.e. ..\Emulators\Arcade\dice.0.9.win.x64\dice.exe
        End If
        str += ";"

        'Set rom extensions
        If media_emuParams(name.ToUpper)(2).Trim <> "" Then
            str += media_emuParams(name.ToUpper)(2)
            ini.IniWriteValue(name, "Rom_Extension", media_emuParams(name.ToUpper)(2))
        ElseIf romExtensionsArr.Count > 0 Then
            str += String.Join("|", romExtensionsArr)
            ini.IniWriteValue(name, "Rom_Extension", String.Join("|", romExtensionsArr))
        End If
        str += ";"

        'Set module
        'media_emuParams:  0 - emuName, 1 - emuExeNode, 2 - emuExtNode, 3 - emuModNode, 4 - index
        'Modules list:     0 - module_dir_name, 1 - module_file_name_no_EXT, 2 - module_rel_path, 3 - module_full_path, 4 - full_path_orig_before_renaming
        For Each mod_param In modules
            If mod_param(4).Trim <> "" And media_emuParams(name.ToUpper)(3).Trim <> "" Then
                If mod_param(4).ToUpper.EndsWith(media_emuParams(name.ToUpper)(3).ToUpper) Then
                    str += mod_param(2).Trim + "|"
                    If mod_param(2).Trim <> "" Then ini.IniWriteValue(name, "Module", mod_param(2).Trim) 'Relative to module folder itself, i.e. ..\Phoenix-Altmer\phoenix-altmer.ahk
                End If
            End If
        Next

        Return str
    End Function

    'RemoveISDButtons
    Sub RemoveISDButtons()
        'Remove all but two controls from flow panel, if some package was already installed
        For i As Integer = flow.Controls.Count - 1 To 2 Step -1
            Dim b = TryCast(flow.Controls(i), Button)
            If b IsNot Nothing Then
                Dim tab = DirectCast(b.Tag, TabControl)
                GroupBox1.Controls.Remove(tab)

                For Each page As TabPage In tab.TabPages
                    Dim tbl = page.Controls.OfType(Of TableLayoutPanel).FirstOrDefault
                    For Each lbl In tbl.Controls.OfType(Of Label)
                        toolTip.unBind(lbl)
                    Next
                Next
            End If

            flow.Controls.RemoveAt(i)
        Next
    End Sub
    'Alert if global emulators settings was net applied
    Function Cancel_if_global_Emulators_Needs_To_Be_Set() As Boolean
        If globalEmulatorsNeedsToBeSet Then
            Dim msg = "Warning, unapplied settings." + vbCrLf
            msg += "There are some emulators waiting to be added to RocketLauncher ""Global Emulator.ini""" + vbCrLf
            msg += "You have to review settings and click 'Apply' button in 'PackageSettings' tab to actually add them." + vbCrLf
            msg += "If you continue now, those emulators will not be added, and your installed system may not function." + vbCrLf
            msg += vbCrLf
            'msg += "Do you want to skip adding emulators and load a new package?"
            msg += "Do you want to skip adding emulators and continue?"
            If MsgBox(msg, MsgBoxStyle.YesNo) = MsgBoxResult.No Then Return True
        End If
        Return False
    End Function
    'Package conflicts - Check/UnCheck all
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click, Button5.Click
        Dim b = DirectCast(sender, Button)
        Dim c As Boolean = False
        If b.Text.ToUpper.StartsWith("CHECK") Then c = True

        For Each m In media_paths
            If m.associated_check IsNot Nothing Then m.associated_check.Checked = c
        Next
    End Sub
    'Package conflicts - Set All To
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedIndex < 0 Then Exit Sub
        For Each m In media_paths
            If m.associated_combo IsNot Nothing And m.status = status.NeedMerge Then
                If ComboBox1.SelectedIndex = 0 And m.associated_combo.Items.Count > media_options.Count Then
                    'If it's a special case with 5 items combobox (Emu/Rom/Module)
                    m.associated_combo.SelectedIndex = m.associated_combo.Items.Count - 1
                Else
                    m.associated_combo.SelectedIndex = ComboBox1.SelectedIndex
                End If
            End If
        Next

        ComboBox1.SelectedIndex = -1
    End Sub
    '... browse buttons
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, Button2.Click, Button3.Click
        Dim b = DirectCast(sender, Button)
        If b.Name.EndsWith("1") Then
            Dim fb As New OpenFileDialog
            fb.Filter = "HS Checker Packages|*.HSPack.zip|Archive files|*.zip;*.7z|All Files|*.*"
            fb.InitialDirectory = Application.StartupPath + ".\SystemPackages"
            fb.ShowDialog(Me)

            If fb.FileName.Trim = "" Then Exit Sub
            TextBox1.Text = fb.FileName.Trim
        ElseIf b.Name.EndsWith("2") Then

        ElseIf b.Name.EndsWith("3") Then

        End If
    End Sub
End Class

Public Class RL_ISD_Parser
    Public emuPath As String
    Public romPath As String
    Public romPathRel As String
    Public modPath As String

    Dim system As String = ""
    Dim refr As Boolean = False
    Dim pending_ini_save As New List(Of INI_File_Class_Interface)
    Dim info_image As Image = New Bitmap(18, 18)
    Dim toolTip As CustomToolTip

    Structure control_info
        Dim key_name As String
        Dim value_type As String
        Dim required As Boolean
        Dim file_extensions As String()
        Dim associated_control As Control
        Dim associated_ini As INI_File_Class_Interface
        Dim associated_ini_section As String
    End Structure

    Public Sub New(_toolTip As CustomToolTip)
        toolTip = _toolTip
        Dim g = Graphics.FromImage(info_image)
        g.DrawImage(My.Resources.info, New Rectangle(0, 0, info_image.Width, info_image.Height))
    End Sub

    Public Sub Parse(panel As Control, flow_for_buttons As FlowLayoutPanel, isd_path As String, Optional sys As String = "")
        Dim x As New Xml.XmlDocument()
        x.Load(isd_path)

        For Each INIFILE As Xml.XmlNode In x.SelectNodes("/INISCHEMA/INIFILES/INIFILE")
            Dim ini_name = INIFILE.Attributes.GetNamedItem("name").Value
            ini_name = evaluateKeywords(ini_name, isd_path, sys)

            'Skip all but current systems, in multisystem RL modules
            If sys <> "" Then
                Dim INITYPE = INIFILE.SelectSingleNode("INITYPE")
                If INITYPE IsNot Nothing Then
                    If INITYPE.InnerText.ToUpper = "System".ToUpper Then
                        If ini_name.Trim.ToUpper <> sys.Trim.ToUpper Then Continue For
                    End If
                End If

                system = sys
            End If

            Dim iniMetaType = "INI"
            If INIFILE.Attributes.GetNamedItem("iniType") IsNot Nothing Then iniMetaType = INIFILE.Attributes.GetNamedItem("iniType").Value.ToUpper

            Dim ini As INI_File_Class_Interface = Nothing
            Dim ini_file = ini_name.ToUpper
            If Not ini_file.Trim.ToUpper = "PackageSettings".ToUpper Then
                If Not ini_file.EndsWith(".INI") Then ini_file += ".ini" 'TODO: it can be xml or actually any extension
                If Not ini_file.StartsWith("\\") And Not ini_file.Substring(1, 2) = ":\" Then ini_file = IO.Path.GetDirectoryName(isd_path) + "\" + ini_file
                If Not IO.File.Exists(ini_file) Then Continue For

                If iniMetaType.Contains("MAME") Then
                    Dim ini_tmp = New IniFile() : ini_tmp.LoadMame(ini_file, " "c) : ini = ini_tmp
                    pending_ini_save.Add(ini)
                ElseIf iniMetaType.Contains("UNIEMU") Then
                    Dim ini_tmp = New IniFile() : ini_tmp.LoadMame(ini_file, "="c) : ini = ini_tmp
                Else
                    ini = New IniFileApi With {.path = ini_file}
                End If
            End If

            Dim tab = createTabControl(panel, flow_for_buttons, IO.Path.GetFileName(ini_name)).Key

            For Each SECTION As Xml.XmlNode In INIFILE.SelectNodes("SECTIONS/SECTION")
                Dim section_name = SECTION.Attributes.GetNamedItem("name").Value
                tab.TabPages.Add(section_name)

                Dim tbl As New DBLayoutPanel
                tbl.ColumnCount = 6
                tbl.AutoScroll = True
                tbl.Padding = New Padding(10)
                tbl.Dock = DockStyle.Fill
                tbl.BorderStyle = BorderStyle.FixedSingle
                'tbl.CellBorderStyle = TableLayoutPanelCellBorderStyle.Single
                tbl.ColumnStyles.Clear()
                tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 170))
                tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 30))
                tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 30))
                tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 170))
                tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 30))
                tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 30))
                AddHandler tbl.CellPaint, Sub(o As Object, e As TableLayoutCellPaintEventArgs)
                                              If e.Column = 0 Or e.Column = 3 Then
                                                  Dim ctrl = tbl.GetControlFromPosition(e.Column, e.Row)
                                                  If TypeOf ctrl Is Label AndAlso DirectCast(ctrl, Label).Visible = True AndAlso DirectCast(ctrl, Label).Text.Trim <> "" Then
                                                      Dim _x = e.CellBounds.Right - info_image.Width - 3
                                                      Dim _y = e.CellBounds.Top + CInt(e.CellBounds.Height / 2) - CInt(info_image.Height / 2)
                                                      e.Graphics.DrawImage(info_image, New Point(_x, _y))
                                                  End If
                                              End If
                                          End Sub
                'Dim w1 = tbl.GetColumnWidths(1)
                'Dim w2 = tbl.GetColumnWidths(2)
                tbl.SuspendLayout()

                'Create game list and alter table height, if it's perRom section
                Dim rom_list As ListBox = Nothing
                If section_name.ToUpper = "%ROMNAME%" Then rom_list = createRomList(tab, tbl, ini)

                tab.TabPages(tab.TabPages.Count - 1).Controls.Add(tbl)

                Dim curRow As Integer = 0
                Dim curCol As Integer = 0
                Dim freePlaceToTheRight As Integer = 0
                Dim keys As Xml.XmlNodeList = SECTION.SelectNodes("KEYS/KEY[not(@hide = 'true')]")
                Dim keys_left As Integer = keys.Count
                For Each KEY As Xml.XmlNode In keys
                    Dim values As Xml.XmlNodeList = Nothing
                    Dim values_node = KEY.SelectSingleNode("VALUES")
                    If values_node IsNot Nothing Then values = values_node.SelectNodes("VALUE")
                    Dim file_extensions As Xml.XmlNodeList = Nothing
                    Dim file_extensions_node = KEY.SelectSingleNode("FILEEXTENSIONS")
                    If file_extensions_node IsNot Nothing Then file_extensions = file_extensions_node.SelectNodes("FILEEXTENSION")
                    Dim read_only As String = "false"
                    Dim read_only_node = KEY.Attributes.GetNamedItem("readonly")
                    If read_only_node IsNot Nothing Then read_only = read_only_node.Value
                    Dim required As String = "false"
                    Dim required_node = KEY.Attributes.GetNamedItem("required")
                    If required_node IsNot Nothing Then required = required_node.Value
                    Dim nullable As String = "true"
                    Dim nullable_node = KEY.Attributes.GetNamedItem("nullable")
                    If nullable_node IsNot Nothing Then nullable = nullable_node.Value
                    Dim descr As String = ""
                    Dim descr_node = KEY.SelectSingleNode("DESCRIPTION")
                    If descr_node IsNot Nothing Then descr = descr_node.InnerText.Trim
                    Dim defaults As String = ""
                    Dim defaults_node = KEY.SelectSingleNode("DEFAULT")
                    If defaults_node IsNot Nothing Then defaults = defaults_node.InnerText.Trim
                    defaults = evaluateKeywords(defaults, isd_path, sys).Trim 'Evaluate meta %keywords% and {keywords}

                    Dim l As New Label
                    l.Text = KEY.Attributes.GetNamedItem("name").Value
                    l.Padding = New Padding(0, 0, 25, 0) : l.BackColor = Color.Transparent
                    l.AutoSize = False : l.Dock = DockStyle.Fill : l.TextAlign = ContentAlignment.MiddleRight
                    If descr <> "" Then toolTip.Bind(l, descr, New Rectangle(150, 0, 0, 0))


                    'Create control (combobox or textbox, depending of value type) and set it's default value
                    Dim ci As New control_info With {.key_name = l.Text, .value_type = KEY.SelectSingleNode("KEYTYPE").InnerText, .associated_ini = ini, .associated_ini_section = section_name}
                    Dim ctrl = getControlByType(ci, defaults, values, nullable, required, read_only, file_extensions)
                    ctrl(0).Anchor = AnchorStyles.Left Or AnchorStyles.Right
                    tbl.Controls.Add(l)
                    tbl.Controls.Add(ctrl(0))
                    'If defaults <> "" Then ctrl(0).Text = defaults Else If ini IsNot Nothing Then ctrl(0).Text = ini.IniReadValue(section_name, l.Text)
                    If defaults <> "" AndAlso ini IsNot Nothing Then Option_ValueChanged(ctrl(0), New EventArgs)

                    'A VERY tricky way to replicate RLUI module settings grid layout (control flow direction top-to-down, instead of left-to-right
                    If curCol = 0 AndAlso KEY.SelectSingleNode("FULLROW") IsNot Nothing AndAlso KEY.SelectSingleNode("FULLROW").InnerText.ToUpper = "TRUE" Then
                        'ctrl(0).Width = w1 * 2 + w2
                        tbl.SetColumnSpan(ctrl(0), 4)

                        tbl.SetCellPosition(l, New TableLayoutPanelCellPosition(0, curRow))
                        tbl.SetCellPosition(ctrl(0), New TableLayoutPanelCellPosition(1, curRow))
                        If ctrl.Count = 2 Then
                            ctrl(1).Size = New Size(ctrl(0).Height, ctrl(0).Height)
                            tbl.Controls.Add(ctrl(1))
                            tbl.SetCellPosition(ctrl(1), New TableLayoutPanelCellPosition(2, curRow))
                        End If

                        curRow += 1
                        keys_left -= 1
                    Else
                        'ctrl(0).Width = w1

                        If curCol > 0 Then
                            Do While tbl.GetColumnSpan(tbl.GetControlFromPosition(1, curRow)) > 1
                                curRow += 1
                            Loop
                        End If

                        tbl.SetCellPosition(l, New TableLayoutPanelCellPosition(curCol, curRow))
                        tbl.SetCellPosition(ctrl(0), New TableLayoutPanelCellPosition(curCol + 1, curRow))
                        If ctrl.Count = 2 Then
                            ctrl(1).Size = New Size(ctrl(0).Height, ctrl(0).Height)
                            tbl.Controls.Add(ctrl(1))
                            tbl.SetCellPosition(ctrl(1), New TableLayoutPanelCellPosition(curCol + 2, curRow))
                        End If

                        curRow += 1
                        keys_left -= 1
                        freePlaceToTheRight += 1
                        If curCol = 0 And freePlaceToTheRight >= keys_left Then curCol = 3 : tbl.RowCount = curRow + 1 : curRow = 0
                    End If
                Next
                tbl.ResumeLayout()
                If rom_list IsNot Nothing Then romList_SelectedIndexChanged(rom_list, New EventArgs) 'Disable controls for roms section
            Next
        Next
    End Sub
    'Evaluate meta keywords %keyword% and {KeyWord}
    Function evaluateKeywords(s As String, isd_path As String, sys As String) As String
        If s.Trim = "" Then Return ""

        Dim w = My.Computer.Screen.Bounds.Width.ToString
        Dim h = My.Computer.Screen.Bounds.Height.ToString

        Dim _modP = modPath : If _modP Is Nothing Then _modP = ""
        Dim _emuP = emuPath : If _emuP Is Nothing Then _emuP = ""
        Dim _romP = romPath : If _romP Is Nothing Then _romP = ""
        Dim _romPRel = romPathRel : If _romPRel Is Nothing Then _romPRel = ""


        Dim modPathsArr = _modP.Split({"|"c}, StringSplitOptions.RemoveEmptyEntries)
        Dim emuPathsArr = _emuP.Split({"|"c}, StringSplitOptions.RemoveEmptyEntries)
        Dim romPathsArr = _romP.Split({"|"c}, StringSplitOptions.RemoveEmptyEntries)
        Dim romPathsRelArr = _romPRel.Split({"|"c}, StringSplitOptions.RemoveEmptyEntries)

        'Use old vb replace, because it's case insensitive
        s = Strings.Replace(s, "%ModuleName%", IO.Path.GetFileNameWithoutExtension(isd_path), 1, -1, CompareMethod.Text)
        s = Strings.Replace(s, "%EmuPath%", _emuP, 1, -1, CompareMethod.Text)
        s = Strings.Replace(s, "%RomPath%", _romP, 1, -1, CompareMethod.Text)
        s = Strings.Replace(s, "%RomPathRelative%", _romPRel, 1, -1, CompareMethod.Text)
        s = Strings.Replace(s, "%RLModulePath%", _modP, 1, -1, CompareMethod.Text)

        s = Strings.Replace(s, "{ModuleName}", IO.Path.GetFileNameWithoutExtension(isd_path), 1, -1, CompareMethod.Text)
        s = Strings.Replace(s, "{EmuPath}", _emuP, 1, -1, CompareMethod.Text)
        s = Strings.Replace(s, "{RomPath}", _romP, 1, -1, CompareMethod.Text)
        s = Strings.Replace(s, "{RomPathRelative}", _romPRel, 1, -1, CompareMethod.Text)
        s = Strings.Replace(s, "{RLModulePath}", _modP, 1, -1, CompareMethod.Text)
        s = Strings.Replace(s, "{DesktopResolution}", w + "x" + h, 1, -1, CompareMethod.Text)
        s = Strings.Replace(s, "{DesktopResolutionW}", w, 1, -1, CompareMethod.Text)
        s = Strings.Replace(s, "{DesktopResolutionH}", h, 1, -1, CompareMethod.Text)

        For i As Integer = 0 To modPathsArr.Count - 1
            s = Strings.Replace(s, "%RLModulePath" + i.ToString + "%", modPathsArr(i), 1, -1, CompareMethod.Text)
            s = Strings.Replace(s, "{RLModulePath" + i.ToString + "}", modPathsArr(i), 1, -1, CompareMethod.Text)
        Next
        For i As Integer = 0 To emuPathsArr.Count - 1
            s = Strings.Replace(s, "%EmuPath" + i.ToString + "%", emuPathsArr(i), 1, -1, CompareMethod.Text)
            s = Strings.Replace(s, "{EmuPath" + i.ToString + "}", emuPathsArr(i), 1, -1, CompareMethod.Text)
        Next
        For i As Integer = 0 To romPathsArr.Count - 1
            s = Strings.Replace(s, "%RomPath" + i.ToString + "%", romPathsArr(i), 1, -1, CompareMethod.Text)
            s = Strings.Replace(s, "{RomPath" + i.ToString + "}", romPathsArr(i), 1, -1, CompareMethod.Text)
        Next
        For i As Integer = 0 To romPathsRelArr.Count - 1
            s = Strings.Replace(s, "%RomPathRelative" + i.ToString + "%", romPathsRelArr(i), 1, -1, CompareMethod.Text)
            s = Strings.Replace(s, "{RomPathRelative" + i.ToString + "}", romPathsRelArr(i), 1, -1, CompareMethod.Text)
        Next

        If sys <> "" Then s = Strings.Replace(s, "%SystemName%", sys.Trim, 1, -1, CompareMethod.Text)

        If s Is Nothing Then Return "" Else Return s
    End Function
    'Create a tab page and its button with a given name
    Public Function createTabControl(panel As Control, flow_for_buttons As FlowLayoutPanel, name As String) As KeyValuePair(Of TabControl, Button)
        Dim b As New Button
        flow_for_buttons.Controls.Add(b)
        b.Width = 170
        b.Height = 30
        b.Text = name
        AddHandler b.Click, AddressOf ChangeTab_ButtonClick

        Dim tab As New TabControl
        b.Tag = tab
        tab.Visible = False
        tab.Location = New Point(215, 10)
        tab.Size = New Size(panel.Width - 225, panel.Height - 20)
        tab.Anchor = AnchorStyles.Left Or AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Right
        panel.Controls.Add(tab)

        Return New KeyValuePair(Of TabControl, Button)(tab, b)
    End Function
    'Create a control and its browse button by type of the option
    Function getControlByType(ci As control_info, defaults As String, Optional values As Xml.XmlNodeList = Nothing, Optional nullable As String = "true", Optional required As String = "false", Optional read_only As String = "false", Optional f_ext As Xml.XmlNodeList = Nothing) As Control()
        Dim c As Control() = Nothing

        ci.associated_control = Nothing : ci.file_extensions = Nothing
        If defaults = "" AndAlso ci.associated_ini IsNot Nothing Then defaults = ci.associated_ini.IniReadValue(ci.associated_ini_section, ci.key_name)

        If f_ext IsNot Nothing Then
            Dim l As New List(Of String)
            For Each x As Xml.XmlNode In f_ext
                l.Add(x.InnerText)
            Next
            ci.file_extensions = l.ToArray
        End If
        nullable = nullable.Trim.ToLower
        required = required.Trim.ToLower
        read_only = read_only.Trim.ToLower

        Select Case ci.value_type.ToUpper
            Case "Boolean".ToUpper
                Dim cmb = New ComboBox With {.DropDownStyle = ComboBoxStyle.DropDownList, .Tag = ci}
                cmb.Items.Add(New SimpleListItem("true", "true"))
                cmb.Items.Add(New SimpleListItem("false", "false"))
                c = {cmb}
            Case "Binary".ToUpper
                Dim cmb = New ComboBox With {.DropDownStyle = ComboBoxStyle.DropDownList, .Tag = ci}
                cmb.Items.Add(New SimpleListItem("on", "on"))
                cmb.Items.Add(New SimpleListItem("off", "off"))
                c = {cmb}
            Case "Textarea".ToUpper, "text".ToUpper
                Dim txt = New TextBox With {.Text = defaults, .Tag = ci, .Multiline = True, .Height = 90}
                c = {txt}
            Case "Integer".ToUpper, "String".ToUpper, "Long".ToUpper, "ExeParams".ToUpper
                If values Is Nothing Then
                    Dim txt = New TextBox With {.Text = defaults, .Tag = ci}
                    c = {txt}
                Else
                    Dim cmb = New ComboBox With {.DropDownStyle = ComboBoxStyle.DropDownList, .Tag = ci}
                    For Each v As Xml.XmlNode In values
                        Dim vt = v.InnerText
                        Dim vd = v.Attributes.GetNamedItem("description")
                        If vd IsNot Nothing Then
                            cmb.Items.Add(New SimpleListItem(vt, vd.Value))
                        Else
                            cmb.Items.Add(New SimpleListItem(vt, vt))
                        End If
                        'If defaults <> "" And vt.ToUpper = defaults.ToUpper Then cmb.SelectedIndex = cmb.Items.Count - 1
                        If defaults <> "" And vt.ToUpper = defaults.ToUpper Then defaults = cmb.Items(cmb.Items.Count - 1).ToString
                    Next

                    'If nullable = "true" Then cmb.Items.Add(New SimpleListItem(" ", " "))
                    'cmb.Items.Add(New SimpleListItem("<Default>", "<Default>"))
                    'If cmb.SelectedIndex < 0 Then cmb.Text = "<Default>"
                    'AddHandler cmb.SelectedIndexChanged, AddressOf Option_ValueChanged
                    c = {cmb}
                End If
            Case "FilePath".ToUpper, "FolderPath".ToUpper, "FileName".ToUpper, "ExeTarget".ToUpper, "ExeStartIn".ToUpper, "URI".ToUpper
                Dim txt = New TextBox : txt.Text = defaults : ci.associated_control = txt : txt.Tag = ci
                Dim btn = New Button With {.Text = "...", .Tag = ci}
                c = {txt, btn}
            Case "ProcessName".ToUpper
                Dim txt = New TextBox : txt.Text = defaults : ci.associated_control = txt : txt.Tag = ci
                Dim btn = New Button With {.Text = "P", .Tag = ci}
                c = {txt, btn}
            Case "WinTitle".ToUpper
                Dim txt = New TextBox : txt.Text = defaults : ci.associated_control = txt : txt.Tag = ci
                Dim btn = New Button With {.Text = "F", .Tag = ci}
                c = {txt, btn}
            Case "Resolution".ToUpper
                Dim cmb = New ComboBox With {.DropDownStyle = ComboBoxStyle.DropDownList, .Tag = ci}
                cmb.Items.Add(New SimpleListItem("640x480", "640x480"))
                cmb.Items.Add(New SimpleListItem("800x600", "800x600"))
                cmb.Items.Add(New SimpleListItem("1024x768", "1024x768"))
                cmb.Items.Add(New SimpleListItem("1152x864", "1152x864"))
                cmb.Items.Add(New SimpleListItem("1280x720", "1280x720"))
                cmb.Items.Add(New SimpleListItem("1280x768", "1280x768"))
                cmb.Items.Add(New SimpleListItem("1280x800", "1280x800"))
                cmb.Items.Add(New SimpleListItem("1280x960", "1280x960"))
                cmb.Items.Add(New SimpleListItem("1280x1024", "1280x1024"))
                cmb.Items.Add(New SimpleListItem("1360x768", "1360x768"))
                cmb.Items.Add(New SimpleListItem("1366x768", "1366x768"))
                cmb.Items.Add(New SimpleListItem("1440x900", "1440x900"))
                cmb.Items.Add(New SimpleListItem("1600x900", "1600x900"))
                cmb.Items.Add(New SimpleListItem("1600x1024", "1600x1024"))
                cmb.Items.Add(New SimpleListItem("1600x1200", "1600x1200"))
                cmb.Items.Add(New SimpleListItem("1680x1050", "1680x1050"))
                cmb.Items.Add(New SimpleListItem("1920x1680", "1920x1080"))
                c = {cmb}
            Case "xHotkey".ToUpper
                Dim txt = New TextBox : txt.Text = defaults : ci.associated_control = txt : txt.Tag = ci
                Dim btn = New Button With {.Text = "X", .Tag = ci}
                c = {txt, btn}
            Case Else
                MsgBox("Congratulation! You found unknown RLUI datatype: " + ci.value_type + vbCrLf + "Please, contact author on HS forum because he is not aware of this one.")
        End Select

        If c IsNot Nothing Then
            If TypeOf c(0) Is TextBox Then
                If read_only = "true" Then DirectCast(c(0), TextBox).ReadOnly = True
                AddHandler DirectCast(c(0), TextBox).TextChanged, AddressOf Option_ValueChanged
            ElseIf TypeOf c(0) Is ComboBox Then
                Dim cmb = DirectCast(c(0), ComboBox)
                If nullable = "true" Then cmb.Items.Add(New SimpleListItem(" ", " "))
                cmb.Items.Add(New SimpleListItem("<Default>", "<Default>"))
                cmb.Text = "<Default>" : If defaults <> "" Then cmb.Text = defaults
                If read_only = "true" Then cmb.Enabled = False
                AddHandler cmb.SelectedIndexChanged, AddressOf Option_ValueChanged
            End If

            If c.Count = 2 Then
                If read_only = "true" Then c(1).Enabled = False
                AddHandler DirectCast(c(1), Button).Click, AddressOf BrowseValue_ButtonClick
            End If

            Return c
        Else
            Return {New Label With {.Visible = False}}
        End If
    End Function
    'Create romList for %RomName% RL module ini section
    Function createRomList(tab As TabControl, tbl_to_move_bottom As DBLayoutPanel, ini As INI_File_Class_Interface) As ListBox
        Dim rom_list = New ListBox With {.Top = 10, .Left = 10, .Height = 200, .Sorted = True}
        rom_list.Width = tab.TabPages(tab.TabPages.Count - 1).Width - 20
        rom_list.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        tab.TabPages(tab.TabPages.Count - 1).Controls.Add(rom_list)

        'Fill rom list
        If ini IsNot Nothing Then
            For Each rom_section In ini.IniListKey
                If rom_section.ToUpper = "SETTINGS" Then Continue For
                rom_list.Items.Add(rom_section)
            Next
        End If

        'Add/Remove Buttons
        Dim b1 As New Button With {.Left = 10, .Top = 180, .Width = 30, .Height = 30, .Text = "-", .Anchor = AnchorStyles.Top Or AnchorStyles.Left, .Tag = rom_list}
        Dim b2 As New Button With {.Left = 45, .Top = 180, .Width = 30, .Height = 30, .Text = "+", .Anchor = AnchorStyles.Top Or AnchorStyles.Left, .Tag = rom_list}
        AddHandler b1.Click, AddressOf add_remove_rom_button_click
        AddHandler b2.Click, AddressOf add_remove_rom_button_click
        tab.TabPages(tab.TabPages.Count - 1).Controls.Add(b1) : b1.BringToFront()
        tab.TabPages(tab.TabPages.Count - 1).Controls.Add(b2) : b2.BringToFront()

        'Alter controls table positions
        tbl_to_move_bottom.Dock = DockStyle.None
        tbl_to_move_bottom.Top = 220
        tbl_to_move_bottom.Left = 10
        tbl_to_move_bottom.Width = tab.TabPages(tab.TabPages.Count - 1).Width - 20
        tbl_to_move_bottom.Height = tab.TabPages(tab.TabPages.Count - 1).Height - 230
        tbl_to_move_bottom.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom

        rom_list.Tag = tbl_to_move_bottom
        tbl_to_move_bottom.Tag = rom_list
        AddHandler rom_list.SelectedIndexChanged, AddressOf romList_SelectedIndexChanged
        Return rom_list
    End Function


    '%romName% section - Browsing rom
    Sub romList_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim l = DirectCast(sender, ListBox)
        Dim tbl = DirectCast(l.Tag, DBLayoutPanel)

        refr = True
        For Each c As Control In tbl.Controls
            If TypeOf c Is Label Then Continue For

            Dim justReset As Boolean = False
            If l.SelectedIndex < 0 Then
                c.Enabled = False
                justReset = True
            Else
                c.Enabled = True
                Dim ci = DirectCast(c.Tag, control_info)
                Dim ini = ci.associated_ini
                If ini IsNot Nothing Then
                    Dim val = ini.IniReadValue(l.SelectedItem.ToString, ci.key_name)
                    If val = "" Then justReset = True Else c.Text = val
                Else
                    justReset = True
                End If
            End If

            If justReset Then
                If TypeOf c Is TextBox Then c.Text = ""
                If TypeOf c Is ComboBox Then c.Text = "<Default>"
            End If
        Next
        refr = False
    End Sub
    '%romName% section - add/remove rom buttons
    Sub add_remove_rom_button_click(sender As Object, e As EventArgs)
        Dim b = DirectCast(sender, Button)
        Dim romList = DirectCast(b.Tag, ListBox)

        If b.Text = "+" Then
            Dim tab = DirectCast(b.Parent, TabPage)
            Dim gamelist As New ListBox With {.Top = 10, .Width = 300, .IntegralHeight = False}
            gamelist.Left = CInt(tab.Width / 2 - 150)
            gamelist.Height = tab.Height - 20
            gamelist.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
            gamelist.Font = New Font(gamelist.Font.FontFamily.Name, 12, FontStyle.Bold)
            tab.Controls.Add(gamelist) : gamelist.BringToFront()
            Dim btn As New Button With {.Text = "OK", .Width = 60, .Height = 30, .Tag = gamelist}
            btn.Left = gamelist.Right - 60
            btn.Top = gamelist.Bottom - 30
            AddHandler btn.Click, AddressOf SelectGameToAdd_OKButtonClick
            tab.Controls.Add(btn) : btn.BringToFront()
            Dim btn2 As New Button With {.Text = "Cancel", .Width = 60, .Height = 30, .Tag = gamelist}
            btn2.Left = gamelist.Right - 130
            btn2.Top = gamelist.Bottom - 30
            AddHandler btn2.Click, AddressOf SelectGameToAdd_OKButtonClick
            tab.Controls.Add(btn2) : btn2.BringToFront()
            gamelist.Tag = New Control() {romList, gamelist, btn, btn2}
            AddHandler gamelist.DoubleClick, Sub() btn.PerformClick()

            Dim x = New Xml.XmlDocument
            x.Load(Class1.HyperspinPath + "\Databases\" + system + "\" + system + ".xml")
            gamelist.BeginUpdate()
            For Each node As Xml.XmlNode In x.SelectNodes("/menu/game")
                Dim g = node.Attributes.GetNamedItem("name")
                If g IsNot Nothing Then
                    Dim name = g.Value.Trim
                    If name <> "" Then
                        If Not romList.Items.OfType(Of String).Contains(name, StringComparer.CurrentCultureIgnoreCase) Then
                            gamelist.Items.Add(name)
                        End If
                    End If
                End If
            Next
            gamelist.EndUpdate()
        ElseIf b.Text = "-" Then
            If romList.SelectedIndex < 0 Then MsgBox("Please, choose a game to remove.") : Exit Sub
            Dim res = MsgBox("Are you sure to remove " + romList.SelectedItem.ToString + " from ini?", MsgBoxStyle.YesNo)
            If res = MsgBoxResult.Yes Then
                Dim tbl = DirectCast(romList.Tag, DBLayoutPanel)
                If tbl.Controls.Count >= 2 Then
                    Dim ctrl = tbl.Controls(1)
                    Dim ci = DirectCast(ctrl.Tag, control_info)
                    Dim ini = ci.associated_ini
                    If ini IsNot Nothing Then
                        ini.IniWriteValue(romList.SelectedItem.ToString, Nothing, Nothing)
                        romList.Items.Remove(romList.SelectedItem)
                    End If
                End If
            End If
        End If
    End Sub
    '%romName% section - add rom - OK button
    Sub SelectGameToAdd_OKButtonClick(o As Object, e As EventArgs)
        Dim b = DirectCast(o, Button)
        Dim l = DirectCast(b.Tag, ListBox)
        Dim arr = DirectCast(l.Tag, Control())
        If b.Text.ToUpper = "OK" Then
            Dim romlist = DirectCast(arr(0), ListBox)
            If l.SelectedIndex < 0 Then MsgBox("No game selected.") : Exit Sub
            romlist.Items.Add(l.SelectedItem.ToString)
            romlist.SelectedItem = l.SelectedItem.ToString
        End If

        Dim page = l.Parent
        page.Controls.Remove(arr(1))
        page.Controls.Remove(arr(2))
        page.Controls.Remove(arr(3))
    End Sub

    'Change tab button
    Sub ChangeTab_ButtonClick(sender As Object, e As EventArgs)
        Dim b = DirectCast(sender, Button)

        For Each c In b.Parent.Controls.OfType(Of Button)
            c.Font = New Font(c.Font, FontStyle.Regular)
            c.BackColor = SystemColors.ControlLight

            Dim t = DirectCast(c.Tag, TabControl)
            t.Visible = False
        Next

        b.Font = New Font(b.Font, FontStyle.Bold Or FontStyle.Underline)
        b.BackColor = SystemColors.ButtonFace
        DirectCast(b.Tag, TabControl).Visible = True
    End Sub

    'Browse value button
    Sub BrowseValue_ButtonClick(sender As Object, e As EventArgs)
        Dim b = DirectCast(sender, Button)

    End Sub

    'INI - Value changed
    Sub Option_ValueChanged(sender As Object, e As EventArgs)
        If refr Then Exit Sub
        Dim ctrl = DirectCast(sender, Control)
        Dim info = DirectCast(ctrl.Tag, control_info)

        Dim section = info.associated_ini_section

        'If it's %RomName% section - we need to alter ini section
        If ctrl.Parent.Tag IsNot Nothing Then
            Dim l = DirectCast(ctrl.Parent.Tag, ListBox)
            If l.SelectedIndex < 0 Then Exit Sub
            section = l.SelectedItem.ToString
        End If

        If info.associated_ini IsNot Nothing Then
            If TypeOf ctrl Is TextBox Then
                info.associated_ini.IniWriteValue(section, info.key_name, ctrl.Text)
            ElseIf TypeOf ctrl Is CheckBox Then
                info.associated_ini.IniWriteValue(section, info.key_name, DirectCast(ctrl, CheckBox).Checked.ToString.ToLower)
            ElseIf TypeOf ctrl Is combobox Then
                Dim cmb = DirectCast(ctrl, ComboBox)
                'Dim cmb_item = DirectCast(cmb.SelectedItem, SimpleListItem)
                Dim cmb_item = TryCast(cmb.SelectedItem, SimpleListItem)

                Dim val = ""
                If cmb_item Is Nothing Then val = cmb.Text.Trim Else val = cmb_item.value.Trim

                If val.ToUpper = "<DEFAULT>" Then
                    info.associated_ini.IniWriteValue(section, info.key_name, Nothing)
                Else
                    info.associated_ini.IniWriteValue(section, info.key_name, val)
                End If
            End If
        End If
    End Sub
    'INI - Save pending
    Public Sub save_pending_ini()
        For n = pending_ini_save.Count - 1 To 0 Step -1
            Dim ini = pending_ini_save(n)
            ini.save()
            pending_ini_save.Remove(ini)
        Next
    End Sub

    Class SimpleListItem
        Public descr As String
        Public value As String
        Sub New(_value As String, _descr As String)
            descr = _descr
            value = _value
        End Sub
        Public Overrides Function ToString() As String
            Return descr
        End Function
    End Class
    Partial Class DBLayoutPanel
        'Double boufered - Prevent panel from flicker
        Inherits TableLayoutPanel

        Public Sub New()
            SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.UserPaint, True)
        End Sub
    End Class
    Partial Class DBProgressBar
        'Double boufered progress bar, to avoid flicker
        Inherits ProgressBar

        Public txt As String = ""

        Public Sub New()
            SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.UserPaint, True)
        End Sub

        Protected Overrides Sub OnPaint(e As PaintEventArgs)
            'MyBase.OnPaint(e)
            Dim rect = ClientRectangle
            Dim g = e.Graphics

            ProgressBarRenderer.DrawHorizontalBar(g, rect)
            rect.Inflate(-3, -3)
            If Value > 0 Then
                'As we doing this ourselves we need to draw the chunks on the progress bar
                Dim clip = New Rectangle(rect.X, rect.Y, CInt(Math.Round((Value / Maximum) * rect.Width)), rect.Height)
                ProgressBarRenderer.DrawHorizontalChunks(g, clip)
            End If

            g.DrawString(txt, New Font("Arial", 8.25, FontStyle.Bold), Brushes.Black, New PointF(10, CSng(Height / 2) - 7))
        End Sub
    End Class
End Class

