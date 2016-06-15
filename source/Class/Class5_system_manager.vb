Imports Microsoft.VisualBasic.FileIO.FileSystem

Public Class Class5_system_manager
    Dim ini As New IniFileApi
    Dim emulators As New Dictionary(Of String, List(Of String))
    Dim modules As New Dictionary(Of String, List(Of String))
    Dim modules_supported_systems As New List(Of String)
    Dim hs_ini As New List(Of String)
    Dim mainMenu As New List(Of String)
    Dim checked As Boolean = False
    Private frm As New Form8_systemProperties
    Private Systems As New Dictionary(Of String, List(Of String))
    Private WithEvents Btn_add As Button = Form1.Button23
    Private WithEvents Btn_scan As Button = Form1.Button33
    Private WithEvents Btn_prop As Button = Form1.Button34
    Private WithEvents Btn_exclude As Button = Form1.Button22
    Private WithEvents Btn_startHS As Button = Form1.Button25
    Private WithEvents grid As DataGridView = Form1.DataGridView2

    'Constructor (just add handler here)
    Public Sub New()
        AddHandler Form8_systemProperties.paths_updated, AddressOf SystemPathUpdated
    End Sub

    'Main Scan Systems
    Private Sub scan() Handles Btn_scan.Click
        If Form1.Label23.BackColor <> Color.LightGreen Then MsgBox("HyperSpin path is incorrect or not set. Set HS path in 'Program Settings' tab.") : Exit Sub
        Systems.Clear()
        emulators.Clear()
        modules.Clear()
        modules_supported_systems.Clear()
        hs_ini.Clear()
        mainMenu.Clear()
        Form1.DataGridView2.Rows.Clear()
        Dim iniClass As New IniFile, iniclass2 As New IniFile
        Dim activeCount As Integer = 0

        'Load exclusion settings
        ini.path = Class1.confPath
        Dim tmp As String = ""
        Dim required_media_number As Integer = 0
        Dim dont_show_completed As Boolean = False
        Dim required_media_list As String() = Nothing
        tmp = ini.IniReadValue("SystemManager", "required_media_number")
        If tmp <> "" AndAlso IsNumeric(tmp) Then required_media_number = CInt(tmp)
        tmp = ini.IniReadValue("SystemManager", "dont_show_completed")
        If tmp <> "" And tmp <> "0" Then dont_show_completed = True
        tmp = ini.IniReadValue("SystemManager", "required_media_list")
        If tmp <> "" Then required_media_list = tmp.Split({","c}, StringSplitOptions.RemoveEmptyEntries)

        'Get HL Path
        iniClass.Load(Class1.HyperspinPath + "\Settings\Settings.ini")
        Dim HL_Path As String = iniClass.GetKeyValue("Main", "Hyperlaunch_Path").Trim
        If Not HL_Path = "" Then
            If HL_Path.ToUpper.EndsWith(".EXE") Then HL_Path = HL_Path.Substring(0, HL_Path.LastIndexOf("\") + 1)

            Dim t1 As Boolean = FileIO.FileSystem.FileExists(HL_Path + "\HyperLaunch.exe")
            Dim t2 As Boolean = FileIO.FileSystem.FileExists(HL_Path + "\RocketLauncher.exe")
            If Not t1 And Not t2 Then HL_Path = ""
        End If
        If Not HL_Path = "" Then iniclass2.Load(HL_Path + "\Settings\Global Emulators.ini")

        'Fill emulators list
        Dim ini2 As New IniFileApi
        ini2.path = HL_Path + "\Settings\Global Emulators.ini"
        For Each s As String In ini2.IniListKey
            Dim m_raw As String = ini2.IniReadValue(s, "Module").Trim
            Dim path As String = ini2.IniReadValue(s, "Emu_Path").Trim
            Dim m As String = m_raw
            If m.Contains("\") Then m = m.Substring(m.LastIndexOf("\") + 1)
            Dim l As New List(Of String)
            l.Add(m_raw) : l.Add(m) : l.Add(path)
            emulators.Add(s, l)
        Next

        'SCANING MODULES
        If Not HL_Path = "" Then
            Dim s As String = ""
            Dim modules_files As System.Collections.ObjectModel.ReadOnlyCollection(Of String) = FileIO.FileSystem.GetFiles(HL_Path + "\Modules\", FileIO.SearchOption.SearchAllSubDirectories, {"*.ahk"})
            For Each module_file As String In modules_files
                s = ""
                FileOpen(1, module_file, OpenMode.Input, OpenAccess.Read)
                Do While Not EOF(1)
                    s = LineInput(1)
                    If s.Trim.ToUpper.StartsWith("MSYSTEM") Then
                        s = s.Substring(s.IndexOf("=") + 1).Trim
                        Dim l As New List(Of String)
                        l.Add(module_file)
                        For Each supportedSys In s.Split({","}, StringSplitOptions.RemoveEmptyEntries)
                            supportedSys = supportedSys.Replace("""", "").Trim
                            l.Add(supportedSys.ToUpper)
                            If Not modules_supported_systems.Contains(supportedSys) Then modules_supported_systems.Add(supportedSys)
                        Next
                        Dim modul_fileName = module_file.Substring(module_file.LastIndexOf("\") + 1)
                        While modules.ContainsKey(modul_fileName)
                            modul_fileName = modul_fileName + "[a]"
                        End While
                        modules.Add(modul_fileName, l)
                    End If
                Loop
                FileClose(1)
            Next
        End If

        'Load MainMenu.xml
        Dim mainMenuXml As New Xml.XmlDocument
        mainMenuXml.Load(Class1.HyperspinPath + "\Databases\Main Menu\Main Menu.xml")
        activeCount = mainMenuXml.SelectNodes("/menu/game").Count
        For Each node As Xml.XmlNode In mainMenuXml.SelectNodes("/menu/game")
            Systems.Add(node.Attributes("name").Value.ToUpper, New List(Of String))
            Systems(node.Attributes("name").Value.ToUpper).Add(node.Attributes("name").Value)
            Systems(node.Attributes("name").Value.ToUpper).Add("%1%")
            mainMenu.Add(node.Attributes("name").Value)
        Next

        'Checking databases
        Dim sys As String = ""
        For Each Dir As String In FileIO.FileSystem.GetDirectories(Class1.HyperspinPath + "\Databases\")
            sys = Dir.Substring(Dir.LastIndexOf("\") + 1)
            If FileIO.FileSystem.FileExists(Dir + "\" + sys + ".xml") And Not sys.ToUpper = "MAIN MENU" Then
                If Not Systems.Keys.Contains(sys.ToUpper) Then
                    Systems.Add(sys.ToUpper, New List(Of String))
                    Systems(sys.ToUpper).Add(sys)
                End If
                Systems(sys.ToUpper).Add("%2%")
            End If
        Next

        'Checking main-menu themes
        Dim files As System.Collections.ObjectModel.ReadOnlyCollection(Of String) = FileIO.FileSystem.GetFiles(Class1.HyperspinPath + "\Media\Main Menu\Themes\", FileIO.SearchOption.SearchTopLevelOnly, {"*.zip"})
        For Each file As String In files
            sys = file.Substring(file.LastIndexOf("\") + 1)
            sys = sys.Substring(0, sys.LastIndexOf("."))
            If Not Systems.Keys.Contains(sys.ToUpper) Then
                Systems.Add(sys.ToUpper, New List(Of String))
                Systems(sys.ToUpper).Add(sys)
            End If
            Systems(sys.ToUpper).Add("%3%")
        Next

        'Checking default system themes
        For Each Dir As String In GetDirectories(Class1.HyperspinPath + "\Media\")
            sys = Dir.Substring(Dir.LastIndexOf("\") + 1)
            If Not sys.ToUpper = "MAIN MENU" Then
                If FileExists(Dir + "\Themes\default.zip") Then
                    If Not Systems.Keys.Contains(sys.ToUpper) Then
                        Systems.Add(sys.ToUpper, New List(Of String))
                        Systems(sys.ToUpper).Add(sys)
                    End If
                    Systems(sys.ToUpper).Add("%4%")
                ElseIf DirectoryExists(Dir + "\Themes") AndAlso GetFiles(Dir + "\Themes", FileIO.SearchOption.SearchTopLevelOnly, {"*.zip"}).Count > 0 Then
                    If Not Systems.Keys.Contains(sys.ToUpper) Then
                        Systems.Add(sys.ToUpper, New List(Of String))
                        Systems(sys.ToUpper).Add(sys)
                    End If
                    Systems(sys.ToUpper).Add("%4G%")
                End If
            End If
        Next

        'Checking default system wheels
        For Each file As String In FileIO.FileSystem.GetFiles(Class1.HyperspinPath + "\Media\Main Menu\Images\Wheel", FileIO.SearchOption.SearchTopLevelOnly, {"*.png", "*.jpg"})
            sys = file.Substring(file.LastIndexOf("\") + 1)
            sys = sys.Substring(0, sys.LastIndexOf("."))
            If Not Systems.Keys.Contains(sys.ToUpper) Then
                Systems.Add(sys.ToUpper, New List(Of String))
                Systems(sys.ToUpper).Add(sys)
            End If
            Systems(sys.ToUpper).Add("%8%")
        Next

        'Checking default system videos
        For Each file As String In FileIO.FileSystem.GetFiles(Class1.HyperspinPath + "\Media\Main Menu\Video", FileIO.SearchOption.SearchTopLevelOnly, {"*.flv", "*.mp4"})
            sys = file.Substring(file.LastIndexOf("\") + 1)
            sys = sys.Substring(0, sys.LastIndexOf("."))
            If Not Systems.Keys.Contains(sys.ToUpper) Then
                Systems.Add(sys.ToUpper, New List(Of String))
                Systems(sys.ToUpper).Add(sys)
            End If
            Systems(sys.ToUpper).Add("%9%")
        Next

        'Checking Hyperspin Setting File
        Dim c As Integer = 0
        Dim emuPath As String = ""
        Dim romPath As String = ""
        For Each file As String In FileIO.FileSystem.GetFiles(Class1.HyperspinPath + "\Settings\", FileIO.SearchOption.SearchTopLevelOnly, {"*.ini"})
            c += 1 : emuPath = ""
            sys = file.Substring(file.LastIndexOf("\") + 1)
            sys = sys.Substring(0, sys.LastIndexOf("."))
            If Not sys.ToUpper = "SETTINGS" And Not sys.ToUpper = "BETABRITE" And Not sys.ToUpper = "MAIN MENU" And Not sys.ToUpper = "GLOBAL SETTINGS" Then
                If Not Systems.Keys.Contains(sys.ToUpper) Then
                    Systems.Add(sys.ToUpper, New List(Of String))
                    Systems(sys.ToUpper).Add(sys)
                End If
                Systems(sys.ToUpper).Add("%5%")
                hs_ini.Add(file.Substring(file.LastIndexOf("\") + 1))

                iniClass.Load(file)
                If iniClass.GetKeyValue("EXE INFO", "hyperlaunch").Trim.ToUpper = "TRUE" Then
                    'HYPERLAUNCH INI
                    Systems(sys.ToUpper).Add("%5HL%")
                    If HL_Path = "" Then
                        Systems(sys.ToUpper).Add("%5HL_PATH_NOT_SET%")
                    Else
                        Dim emu As String = ""
                        If FileIO.FileSystem.FileExists(HL_Path + "\Settings\" + sys + "\Emulators.ini") Then
                            iniClass.Load(HL_Path + "\Settings\" + sys + "\Emulators.ini")
                            emu = iniClass.GetKeyValue("ROMS", "Default_Emulator").Trim
                            romPath = iniClass.GetKeyValue("ROMS", "Rom_Path").Trim
                            If romPath.StartsWith(".") Then romPath = HL_Path + "\" + romPath

                            If emu <> "" Then
                                'iniClass.Load(HL_Path + "\Settings\Global Emulators.ini")
                                emuPath = iniclass2.GetKeyValue(emu, "Emu_Path").Trim
                                If emuPath.StartsWith(".") Then emuPath = HL_Path + "\" + emuPath
                            End If
                        End If
                    End If
                Else
                    'HYPERSPIN INI
                    emuPath = iniClass.GetKeyValue("EXE INFO", "path").Trim
                    If emuPath <> "" AndAlso Not emuPath.EndsWith("\") Then emuPath = emuPath + "\"
                    If emuPath <> "" Then emuPath = emuPath + iniClass.GetKeyValue("EXE INFO", "exe").Trim
                    If emuPath <> "" And emuPath.StartsWith(".") Then emuPath = Class1.HyperspinPath + "\" + emuPath

                    romPath = iniClass.GetKeyValue("EXE INFO", "rompath").Trim
                    If romPath <> "" And romPath.StartsWith(".") Then romPath = Class1.HyperspinPath + "\" + romPath
                End If
                If FileIO.FileSystem.FileExists(emuPath) Then Systems(sys.ToUpper).Add("%6%")
                If FileIO.FileSystem.DirectoryExists(romPath) Then Systems(sys.ToUpper).Add("%7%")
            End If

            Form1.Label41.Text = "Scanning... " + c.ToString + " of " + files.Count.ToString : Form1.Label41.Refresh()
        Next

        'FILL GRID
        For Each l As List(Of String) In Systems.Values
            Dim r As DataGridViewRow = scan_FillRowSub(l, required_media_list, dont_show_completed, required_media_number)
            If r IsNot Nothing Then
                'add last_check_result
                tmp = ini.IniReadValue("LastCheckResult", l(0))
                If tmp <> "" Then
                    Dim arr() As String = tmp.Split({","c}, StringSplitOptions.RemoveEmptyEntries) 'total, rom, vid, thm, w, a1-a4, snd
                    Dim arr_int = arr.ToList.ConvertAll(Function(str) Int32.Parse(str))

                    For i As Integer = 1 To 4
                        r.Cells(i + 9).Value = arr_int(i).ToString + " \ " + arr_int(0).ToString
                        If arr_int(i) = arr_int(0) Then
                            r.Cells(i + 9).Style.BackColor = Form1.colorYES
                        ElseIf arr_int(i) = 0 Then
                            r.Cells(i + 9).Style.BackColor = Form1.colorNO
                        Else
                            r.Cells(i + 9).Style.BackColor = Color.Yellow
                        End If
                    Next
                    'Artworks handle
                    tmp = ""
                    Dim atLeastOneComplete As Boolean = False
                    Dim atLeastOneIncomplete As Boolean = False
                    For i As Integer = 5 To 8
                        If arr_int(i) <> 0 Then
                            If tmp <> "" Then tmp = tmp + ", "
                            tmp = tmp + arr_int(i).ToString
                            If arr_int(i) = arr_int(0) Then atLeastOneComplete = True Else atLeastOneIncomplete = True
                        End If
                    Next
                    If Not tmp = "" Then tmp = tmp + " \ " + arr_int(0).ToString Else tmp = "none"
                    r.Cells(14).Value = tmp
                    If atLeastOneIncomplete Then r.Cells(14).Style.BackColor = Color.Yellow
                    If Not atLeastOneIncomplete And atLeastOneComplete Then r.Cells(14).Style.BackColor = Form1.colorYES
                    If Not atLeastOneIncomplete And Not atLeastOneComplete Then r.Cells(14).Style.BackColor = Form1.colorNO
                End If
                    'add row to grid
                    Form1.DataGridView2.Rows.Add(r)
            End If
        Next
        Form1.Label41.Text = "System count: " + Form1.DataGridView2.Rows.Count.ToString + ", Active:" + activeCount.ToString
        checked = True
    End Sub

    Private Function scan_FillRowSub(l As List(Of String), Optional req_ml As String() = Nothing, Optional req_notCompleted As Boolean = False, Optional req_N As Integer = 0) As DataGridViewRow
        Dim x(9) As String
        x(0) = l(0)
        Dim fnd_n As Integer = 0
        If l.Contains("%1%") Then x(1) = "X" : fnd_n += 1 Else x(1) = ""
        If l.Contains("%2%") Then x(2) = "X" : fnd_n += 1 Else x(2) = ""
        If l.Contains("%3%") Then x(3) = "X" : fnd_n += 1 Else x(3) = ""
        If l.Contains("%4%") Then x(4) = "X" : fnd_n += 1 Else x(4) = ""
        If l.Contains("%4G%") Then x(4) = "Per Game" : fnd_n += 1
        If l.Contains("%5%") Then x(5) = "X" : fnd_n += 1 Else x(5) = ""
        If l.Contains("%5HL%") Then x(5) = "HL"
        If l.Contains("%5HL_PATH_NOT_SET%") Then x(5) = "INVALID HL PATH"
        If l.Contains("%6%") Then x(6) = "X" : fnd_n += 1 Else x(6) = ""
        If l.Contains("%7%") Then x(7) = "X" : fnd_n += 1 Else x(7) = ""
        If l.Contains("%8%") Then x(8) = "X" : fnd_n += 1 Else x(8) = ""
        If l.Contains("%9%") Then x(9) = "X" : fnd_n += 1 Else x(9) = ""

        Dim req As Boolean = True
        If req_ml IsNot Nothing Then
            For Each item As String In req_ml
                If item.Trim = "0" And x(1) = "" Then req = False : Exit For
                If item.Trim = "1" And x(2) = "" Then req = False : Exit For
                If item.Trim = "2" And x(3) = "" Then req = False : Exit For
                If item.Trim = "3" And x(4) = "" Then req = False : Exit For
                If item.Trim = "4" And x(8) = "" Then req = False : Exit For
                If item.Trim = "5" And x(9) = "" Then req = False : Exit For
                If item.Trim = "6" And x(5) = "" Or x(5) = "INVALID HL PATH" Then req = False : Exit For
                If item.Trim = "7" And x(6) = "" Then req = False : Exit For
                If item.Trim = "8" And x(7) = "" Then req = False : Exit For
            Next
        End If
        If req_notCompleted And fnd_n = 9 Then req = False

        If req And (req_N = 0 OrElse req_N < fnd_n) Then
            Dim r As New DataGridViewRow()
            r.CreateCells(grid, x)
            For i As Integer = 1 To 9
                If x(i) = "X" Or x(i) = "HL" Then
                    r.Cells(i).Style.BackColor = Form1.colorYES
                Else
                    r.Cells(i).Style.BackColor = Form1.colorNO
                    If x(i).ToUpper = "PER GAME" Then
                        r.Cells(i).Style.BackColor = Color.YellowGreen
                    End If
                    If x(i) = "INVALID HL PATH" Then
                        Dim f As Font = New Font(Control.DefaultFont.FontFamily, 6, FontStyle.Regular)
                        r.Cells(i).Style.Font = f
                    End If
                End If
            Next
            Return r
        Else
            Return Nothing
        End If
    End Function

    'Grid selection changed
    Private Sub grid_SelectionChanged(sender As Object, e As EventArgs) Handles grid.SelectionChanged
        If grid.SelectionMode <> DataGridViewSelectionMode.FullRowSelect Then Exit Sub
        If frm.Visible = False Then Exit Sub
        setPropertiesFormData()
    End Sub

    'setPropertiesFormData()
    Private Sub setPropertiesFormData()
        Dim data = New List(Of String)
        Dim modules_for_cur_system As New List(Of String)

        If grid.SelectedRows.Count > 0 Then
            Dim sys = grid.SelectedRows(0).Cells(0).Value.ToString.ToUpper
            If Systems(sys).Contains("%1%") Then data.Add("1") Else data.Add("") 'Main menu active
            If Systems(sys).Contains("%2%") Then data.Add("1") Else data.Add("") 'Database exist
            If Systems(sys).Contains("%3%") Then data.Add("1") Else data.Add("") 'Main Menu Theme Exist
            If Systems(sys).Contains("%4%") Then data.Add("1") Else data.Add("") 'System Theme Exist
            If Systems(sys).Contains("%5%") Then data.Add("1") Else data.Add("") 'HS Settings Exist
            If Systems(sys).Contains("%5HL%") Then data.Add("1") Else data.Add("") 'HS Set to use HyperLaunch for this system
            If Systems(sys).Contains("%5HL_PATH_NOT_SET%") Then data.Add("1") Else data.Add("") 'HL path is invalid
            If Systems(sys).Contains("%6%") Then data.Add("1") Else data.Add("") 'Emulator exists
            If Systems(sys).Contains("%7%") Then data.Add("1") Else data.Add("") 'Rom Path exists
            If Systems(sys).Contains("%8%") Then data.Add("1") Else data.Add("") 'System wheel
            If Systems(sys).Contains("%9%") Then data.Add("1") Else data.Add("") 'System video
            data.Add(sys)
            data.Add(grid.SelectedRows(0).Cells(0).Value.ToString())

            'SET modules
            For Each m As KeyValuePair(Of String, List(Of String)) In modules
                If m.Value.Contains(sys) Then
                    modules_for_cur_system.Add(m.Key)
                End If
            Next
        End If
        frm.data = data
        frm.modules = modules_for_cur_system
    End Sub

    'System path Updated Event from form_properties
    Private Sub SystemPathUpdated(sys As String)
        Dim ind As Integer = -1
        Dim system_grid_index As Integer = -1

        For i As Integer = 0 To grid.Rows.Count - 1
            If grid.Rows(i).Cells(0).Value.ToString.ToUpper = sys.ToUpper Then system_grid_index = i : Exit For
        Next
        If system_grid_index = -1 Then MsgBox("Can't find system '" + sys + "' in the grid.") : Exit Sub
        Dim r As DataGridViewRow = grid.Rows(system_grid_index)

        'Update mainMenu status
        Dim selectedSys = ""
        Dim sysNormalCase As String = r.Cells(0).Value.ToString
        Dim mainMenuXml As New Xml.XmlDocument
        mainMenuXml.Load(Class1.HyperspinPath + "\Databases\Main Menu\Main Menu.xml")
        Dim node = mainMenuXml.SelectSingleNode("/menu/game[@name='" + sysNormalCase + "']")
        'Dim node = mainMenuXml.SelectSingleNode("/menu/game[translate(@name, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')='" + sys + "']")
        If node IsNot Nothing Then
            'Adding new system to main menu
            Form1.ComboBox1.Tag = "REFRESHING"
            If Form1.ComboBox1.SelectedItem IsNot Nothing Then selectedSys = Form1.ComboBox1.SelectedItem.ToString

            mainMenu.Clear()
            Form1.ComboBox1.Items.Clear()
            For Each x As Xml.XmlNode In mainMenuXml.SelectNodes("/menu/game")
                mainMenu.Add(x.Attributes("name").Value)
                Form1.ComboBox1.Items.Add(x.Attributes("name").Value)
            Next
            If Not Systems(sys).Contains("%1%") Then Systems(sys).Add("%1%")

            If selectedSys <> "" Then Form1.ComboBox1.SelectedItem = selectedSys
            Form1.ComboBox1.Tag = "0"
        Else
            'Remove a system to main menu
            ind = Systems(sys).IndexOf("%1%")
            If ind >= 0 Then Systems(sys).RemoveAt(ind)
            If mainMenu.IndexOf(sysNormalCase) <> -1 Then mainMenu.Remove(sysNormalCase)

            If Form1.ComboBox1.SelectedItem IsNot Nothing Then selectedSys = Form1.ComboBox1.SelectedItem.ToString
            If selectedSys = sysNormalCase Then Form1.ComboBox1.SelectedIndex = -1
            Form1.ComboBox1.Tag = "REFRESHING"
            If Form1.ComboBox1.Items.IndexOf(sysNormalCase) >= 0 Then Form1.ComboBox1.Items.Remove(sysNormalCase)
            Form1.ComboBox1.Tag = "0"
        End If


        'HS ini, Rompath and emupaths
        Dim iniClass As New IniFile, iniclass2 As New IniFile
        iniClass.Load(Class1.HyperspinPath + "\Settings\Settings.ini")
        Dim HL_Path As String = iniClass.GetKeyValue("Main", "Hyperlaunch_Path").Trim
        If HL_Path.ToUpper.EndsWith("EXE") Then HL_Path = HL_Path.Substring(0, HL_Path.LastIndexOf("\"))
        If Not HL_Path = "" AndAlso (Not FileExists(HL_Path + "\HyperLaunch.exe") And Not FileExists(HL_Path + "\RocketLauncher.exe")) Then HL_Path = ""

        If Not FileIO.FileSystem.FileExists(Class1.HyperspinPath + "\Settings\" + sys + ".ini") Then
            ind = Systems(sys).IndexOf("%5%") : If ind >= 0 Then Systems(sys).RemoveAt(ind)
            ind = Systems(sys).IndexOf("%5HL%") : If ind >= 0 Then Systems(sys).RemoveAt(ind)
            ind = Systems(sys).IndexOf("%5HL_PATH_NOT_SET%") : If ind >= 0 Then Systems(sys).RemoveAt(ind)
        Else
            If Not Systems(sys).Contains("%5%") Then Systems(sys).Add("%5%")
            iniClass.Load(Class1.HyperspinPath + "\Settings\" + sys + ".ini")

            If iniClass.GetKeyValue("EXE INFO", "hyperlaunch").Trim.ToUpper = "TRUE" Then
                'Check Hyperlaunch paths
                If Not Systems(sys).Contains("%5HL%") Then Systems(sys).Add("%5HL%")
                If HL_Path = "" Then
                    If Not Systems(sys).Contains("%5HL_PATH_NOT_SET%") Then Systems(sys).Add("%5HL_PATH_NOT_SET%")
                    ind = Systems(sys).IndexOf("%6%")
                    If ind >= 0 Then Systems(sys).RemoveAt(ind)
                    ind = Systems(sys).IndexOf("%7%")
                    If ind >= 0 Then Systems(sys).RemoveAt(ind)
                Else
                    Dim emu As String = ""
                    Dim emuPath As String = ""
                    Dim romPath As String = ""
                    If FileIO.FileSystem.FileExists(HL_Path + "\Settings\" + sys + "\Emulators.ini") Then
                        iniClass.Load(HL_Path + "\Settings\" + sys + "\Emulators.ini")
                        emu = iniClass.GetKeyValue("ROMS", "Default_Emulator").Trim
                        romPath = iniClass.GetKeyValue("ROMS", "Rom_Path").Trim
                        If romPath.StartsWith(".") Then romPath = HL_Path + "\" + romPath

                        If emu <> "" Then
                            iniclass2.Load(HL_Path + "\Settings\Global Emulators.ini")
                            emuPath = iniclass2.GetKeyValue(emu, "Emu_Path").Trim
                            If emuPath.StartsWith(".") Then emuPath = HL_Path + "\" + emuPath
                        End If

                        If FileIO.FileSystem.FileExists(emuPath) Then
                            If Not Systems(sys).Contains("%6%") Then Systems(sys).Add("%6%")
                        Else
                            ind = Systems(sys).IndexOf("%6%") : If ind >= 0 Then Systems(sys).RemoveAt(ind)
                        End If
                        If FileIO.FileSystem.DirectoryExists(romPath) Then
                            If Not Systems(sys).Contains("%7%") Then Systems(sys).Add("%7%")
                        Else
                            ind = Systems(sys).IndexOf("%7%") : If ind >= 0 Then Systems(sys).RemoveAt(ind)
                        End If
                    Else
                        ind = Systems(sys).IndexOf("%6%") : If ind >= 0 Then Systems(sys).RemoveAt(ind)
                        ind = Systems(sys).IndexOf("%7%") : If ind >= 0 Then Systems(sys).RemoveAt(ind)
                    End If
                End If
            Else
                'Check Hyperspin paths
                Dim emuPath As String = iniClass.GetKeyValue("EXE INFO", "path").Trim
                If emuPath <> "" AndAlso Not emuPath.EndsWith("\") Then emuPath = emuPath + "\"
                If emuPath <> "" Then emuPath = emuPath + iniClass.GetKeyValue("EXE INFO", "exe").Trim
                If emuPath <> "" And emuPath.StartsWith(".") Then emuPath = Class1.HyperspinPath + "\" + emuPath

                Dim romPath As String = iniClass.GetKeyValue("EXE INFO", "rompath").Trim
                If romPath <> "" And romPath.StartsWith(".") Then romPath = Class1.HyperspinPath + "\" + romPath

                If FileIO.FileSystem.FileExists(emuPath) Then
                    If Not Systems(sys).Contains("%6%") Then Systems(sys).Add("%6%")
                Else
                    ind = Systems(sys).IndexOf("%6%") : If ind >= 0 Then Systems(sys).RemoveAt(ind)
                End If
                If FileIO.FileSystem.DirectoryExists(romPath) Then
                    If Not Systems(sys).Contains("%7%") Then Systems(sys).Add("%7%")
                Else
                    ind = Systems(sys).IndexOf("%7%") : If ind >= 0 Then Systems(sys).RemoveAt(ind)
                End If
            End If
        End If

        Dim r1 As DataGridViewRow = scan_FillRowSub(Systems(sys))
        For i As Integer = 0 To r.Cells.Count - 1
            r.Cells(i).Value = r1.Cells(i).Value
            r.Cells(i).Style = r1.Cells(i).Style
        Next
    End Sub

#Region "Region: Show forms / launch HS"
    'Show properties
    Private Sub show_properties() Handles Btn_prop.Click, grid.CellDoubleClick
        frm = New Form8_systemProperties
        frm.oneTimeInit(hs_ini, mainMenu, emulators)
        setPropertiesFormData()
        frm.Show()
    End Sub

    'Show add system dialog
    Private Sub addSystem() Handles Btn_add.Click
        If Not checked Then MsgBox("Please, run scan before adding new systems") : Exit Sub
        Dim f As New FormF_systemManager_addSystem
        f.systems = modules_supported_systems
        f.ShowDialog(Form1)
        Dim sys = f.value

        If Not sys = "" Then
            If Systems.Keys.Contains(sys.ToUpper) Then MsgBox("This system is already in the list. Just use properties dialog to configure it.") : Exit Sub

            Systems.Add(sys.ToUpper, New List(Of String))
            Systems(sys.ToUpper).Add(sys)

            Dim r As New DataGridViewRow()
            r.CreateCells(grid)
            r.Cells(0).Value = sys
            For i As Integer = 1 To 9
                r.Cells(i).Style.BackColor = Form1.colorNO
            Next
            Form1.DataGridView2.Rows.Add(r)
            Form1.DataGridView2.Rows(Form1.DataGridView2.Rows.Count - 1).Selected = True
            Form1.DataGridView2.FirstDisplayedScrollingRowIndex = Form1.DataGridView2.SelectedRows(0).Index
        End If
    End Sub

    'Show exclusion form
    Private Sub shoExclusionDialog() Handles Btn_exclude.Click
        Dim f As New FormF_systemManager_exclusions
        f.ShowDialog(Form1)
    End Sub

    'Start HS
    Private Sub startHS() Handles Btn_startHS.Click
        Dim hs As String = Class1.HyperspinPath + "\Hyperspin.exe"
        Process.Start(hs)
    End Sub
#End Region

#Region "Region: Drag'n'drop"
    Private Sub drag_enter(sender As Object, e As Windows.Forms.DragEventArgs) Handles grid.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Move
            grid.SelectionMode = DataGridViewSelectionMode.CellSelect
            grid.DefaultCellStyle.SelectionBackColor = Color.LightGoldenrodYellow
        End If
    End Sub

    Private Sub drag_over(sender As Object, e As Windows.Forms.DragEventArgs) Handles grid.DragOver
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
            Dim clientPoint As Point = grid.PointToClient(New Point(e.X, e.Y))
            Dim hitTest As DataGridView.HitTestInfo = grid.HitTest(clientPoint.X, clientPoint.Y)
            If hitTest.RowIndex >= 0 And hitTest.ColumnIndex >= 0 Then
                grid.Rows(hitTest.RowIndex).Cells(hitTest.ColumnIndex).Selected = True
            End If
        End If
    End Sub

    Private Sub drag_drop(sender As Object, e As System.EventArgs) Handles grid.DragLeave
        Dim r As Integer = -1
        grid.DefaultCellStyle.SelectionBackColor = SystemColors.Highlight
        If grid.SelectedCells.Count > 0 Then r = grid.SelectedCells(0).RowIndex
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        If r >= 0 Then grid.Rows(r).Selected = True
    End Sub

    Private Sub drag_drop(sender As Object, e As Windows.Forms.DragEventArgs) Handles grid.DragDrop
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            grid.DefaultCellStyle.SelectionBackColor = SystemColors.Highlight
            Dim clientPoint As Point = grid.PointToClient(New Point(e.X, e.Y))
            Dim hitTest As DataGridView.HitTestInfo = grid.HitTest(clientPoint.X, clientPoint.Y)

            Dim failed As New List(Of String)
            Dim files As String() = DirectCast(e.Data.GetData(DataFormats.FileDrop), String())
            If files.Count > 0 Then
                Select Case hitTest.ColumnIndex
                    Case 2 'Database
                        For Each file As String In files
                            If file.ToUpper.EndsWith(".XML") Then
                                Dim sys As String = file.ToUpper.Replace(".XML", "")
                                If Systems.Keys.Contains(sys) Then
                                    FileCopy(file, Class1.HyperspinPath + "\Databases\" + Systems(sys)(0) + "\" + Systems(sys)(0) + ".xml")
                                Else
                                    failed.Add(file)
                                End If
                            End If

                            If failed.Count = 1 Then
                                Dim sys As String = grid.Rows(hitTest.RowIndex).Cells(0).Value.ToString
                                FileCopy(failed(0), Class1.HyperspinPath + "\Databases\" + sys + "\" + sys + ".xml")
                                grid.Rows(hitTest.RowIndex).Cells(4).Style.BackColor = Form1.colorYES
                                If grid.Rows(hitTest.RowIndex).Cells(4).Value.ToString = "" Then
                                    grid.Rows(hitTest.RowIndex).Cells(4).Value = "X"
                                Else
                                    grid.Rows(hitTest.RowIndex).Cells(4).Value = "NEW"
                                End If
                            End If
                        Next
                    Case 3 'Main menu theme
                        For Each file As String In files
                            If file.ToUpper.EndsWith(".ZIP") Then
                                Dim sys As String = file.ToUpper.Replace(".ZIP", "")
                                If Systems.Keys.Contains(sys) Then
                                    FileCopy(file, Class1.HyperspinPath + "\Media\Main Menu\Themes\" + Systems(sys)(0) + ".zip")
                                Else
                                    failed.Add(file)
                                End If
                            End If
                        Next

                        If failed.Count = 1 Then
                            Dim sys As String = grid.Rows(hitTest.RowIndex).Cells(0).Value.ToString
                            Dim res As Microsoft.VisualBasic.MsgBoxResult = MsgBox("Following file were not recognized as system: " + failed(0).Substring(failed(0).LastIndexOf("\")) + vbCrLf + "Do you want to set it as " + sys + "?", MsgBoxStyle.YesNo)
                            If res = MsgBoxResult.Yes Then
                                FileCopy(failed(0), Class1.HyperspinPath + "\Media\Main Menu\Themes\" + sys + ".zip")
                                grid.Rows(hitTest.RowIndex).Cells(3).Style.BackColor = Form1.colorYES
                                If grid.Rows(hitTest.RowIndex).Cells(3).Value.ToString = "" Then
                                    grid.Rows(hitTest.RowIndex).Cells(3).Value = "X"
                                Else
                                    grid.Rows(hitTest.RowIndex).Cells(3).Value = "NEW"
                                End If
                            End If
                        Else
                            MsgBox("Following files were not recognized as system: " + vbCrLf + String.Join(vbCrLf, failed))
                        End If
                    Case 4 'System Theme
                        For Each file As String In files
                            If file.ToUpper.EndsWith(".ZIP") Then
                                Dim sys As String = file.ToUpper.Replace(".ZIP", "")
                                If Systems.Keys.Contains(sys) Then
                                    FileCopy(file, Class1.HyperspinPath + "\Media\" + Systems(sys)(0) + "\Themes\default.zip")
                                Else
                                    failed.Add(file)
                                End If
                            End If

                            If failed.Count = 1 Then
                                Dim sys As String = grid.Rows(hitTest.RowIndex).Cells(0).Value.ToString.ToUpper
                                FileCopy(failed(0), Class1.HyperspinPath + "\Media\" + Systems(sys)(0) + "\Themes\default.zip")
                                grid.Rows(hitTest.RowIndex).Cells(4).Style.BackColor = Form1.colorYES
                                If grid.Rows(hitTest.RowIndex).Cells(4).Value.ToString = "" Then
                                    grid.Rows(hitTest.RowIndex).Cells(4).Value = "X"
                                Else
                                    grid.Rows(hitTest.RowIndex).Cells(4).Value = "NEW"
                                End If
                            End If
                        Next
                End Select
            End If

            grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            If hitTest.RowIndex >= 0 And hitTest.ColumnIndex >= 0 Then grid.Rows(hitTest.RowIndex).Selected = True
        End If
    End Sub
#End Region
End Class
