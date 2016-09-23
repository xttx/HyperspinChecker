Imports System.ComponentModel
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
    Private status_Label As Label = Form1.Label41
    Private need_check_hl_media As CheckBox = Form1.CheckBox31
    Private WithEvents Btn_add As Button = Form1.Button23
    Private WithEvents Btn_scan As Button = Form1.Button33
    Private WithEvents Btn_prop As Button = Form1.Button34
    Private WithEvents Btn_exclude As Button = Form1.Button22
    Private WithEvents Btn_startHS As Button = Form1.Button25
    Private WithEvents grid As DataGridView = Form1.DataGridView2
    Private WithEvents bg_check As New BackgroundWorker() With {.WorkerReportsProgress = True}

    'Constructor (just add handler here)
    Public Sub New()
        AddHandler Form8_systemProperties.paths_updated, AddressOf SystemPathUpdated
        AddHandler Form8_systemProperties.dropped_updated, AddressOf drag_drop_action
    End Sub

    'Main Scan Systems
    Private Sub scan() Handles Btn_scan.Click
        If Form1.Label23.BackColor = Color.LightBlue Then Form1.TextBox14_TextChanged_sub_check()
        If Form1.Label23.BackColor <> Color.LightGreen Then MsgBox("HyperSpin path is incorrect or not set. Set HS path in 'Program Settings' tab.") : Exit Sub
        Form1.DataGridView2.Rows.Clear()
        bg_check.RunWorkerAsync()
    End Sub
    Private Sub scan_bg(o As Object, e As DoWorkEventArgs) Handles bg_check.DoWork
        Systems.Clear()
        emulators.Clear()
        modules.Clear()
        modules_supported_systems.Clear()
        hs_ini.Clear()
        mainMenu.Clear()
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

        Dim HL_Path As String = Class1.HyperlaunchPath
        Dim t1 As Boolean = FileIO.FileSystem.FileExists(HL_Path + "\HyperLaunch.exe")
        Dim t2 As Boolean = FileIO.FileSystem.FileExists(HL_Path + "\RocketLauncher.exe")
        If Not t1 And Not t2 Then HL_Path = ""
        If Not HL_Path = "" Then iniclass2.Load(HL_Path + "\Settings\Global Emulators.ini")

        'Fill emulators list
        status_Label.BeginInvoke(Sub() status_Label.Text = "[Task 1/10] Get emulators from HL/RL Global Emulators.ini...")
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
        status_Label.BeginInvoke(Sub() status_Label.Text = "[Task 2/10] Scaning HL/RL modules...")
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
        status_Label.BeginInvoke(Sub() status_Label.Text = "[Task 3/10] Scaning HyperSpin Main Menu.xml...")
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
        status_Label.BeginInvoke(Sub() status_Label.Text = "[Task 4/10] Scaning System Databases (XML)...")
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
        status_Label.BeginInvoke(Sub() status_Label.Text = "[Task 5/10] Scaning Main Menu Themes...")
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
        status_Label.BeginInvoke(Sub() status_Label.Text = "[Task 6/10] Scaning System Thenes...")
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
        status_Label.BeginInvoke(Sub() status_Label.Text = "[Task 7/10] Scaning System Wheels...")
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
        status_Label.BeginInvoke(Sub() status_Label.Text = "[Task 8/10] Scaning System Videos...")
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
            status_Label.BeginInvoke(Sub() status_Label.Text = "[Task 9/10] Scanning Hyperspin Settings.ini... " + c.ToString + " of " + files.Count.ToString)
        Next

        'Check for HL/RL Media
        If need_check_hl_media.Checked And HL_Path <> "" Then
            status_Label.BeginInvoke(Sub() status_Label.Text = "[Task 10/10] Check for HL/RL Media... ")
            For Each k In Systems.Keys
                sys = Systems(k)(0)
                If FileIO.FileSystem.DirectoryExists(HL_Path + "\Media\Artwork\" + sys + "\_Default") Then
                    If FileIO.FileSystem.GetFiles(HL_Path + "\Media\Artwork\" + sys + "\_Default", FileIO.SearchOption.SearchAllSubDirectories, {"*.png"}).Count > 0 Then Systems(k).Add("%HLMEDIA_ARTWORK%")
                End If

                If FileIO.FileSystem.FileExists(HL_Path + "\Media\Backgrounds\" + sys + "\_Default\default.png") Then Systems(k).Add("%HLMEDIA_BACKG%")

                If FileIO.FileSystem.DirectoryExists(HL_Path + "\Media\Bezels\" + sys + "\_Default") Then
                    If FileIO.FileSystem.GetFiles(HL_Path + "\Media\Bezels\" + sys + "\_Default", FileIO.SearchOption.SearchAllSubDirectories, {"Bezel.png"}).Count > 0 Then Systems(k).Add("%HLMEDIA_BEZEL%")
                End If
                If FileIO.FileSystem.DirectoryExists(HL_Path + "\Media\Fade\" + sys + "\_Default") Then
                    If FileIO.FileSystem.GetFiles(HL_Path + "\Media\Fade\" + sys + "\_Default", FileIO.SearchOption.SearchAllSubDirectories, {"Layer *.png"}).Count > 0 Then Systems(k).Add("%HLMEDIA_FADE%")
                End If
                If FileIO.FileSystem.DirectoryExists(HL_Path + "\Media\Guides\" + sys + "\_Default") Then
                    If FileIO.FileSystem.GetFiles(HL_Path + "\Media\Guides\" + sys + "\_Default", FileIO.SearchOption.SearchAllSubDirectories, {"*.*"}).Count > 0 Then Systems(k).Add("%HLMEDIA_GUIDE%")
                End If
                If FileIO.FileSystem.DirectoryExists(HL_Path + "\Media\Manuals\" + sys + "\_Default") Then
                    If FileIO.FileSystem.GetFiles(HL_Path + "\Media\Manuals\" + sys + "\_Default", FileIO.SearchOption.SearchAllSubDirectories, {"*.pdf", "*.txt"}).Count > 0 Then Systems(k).Add("%HLMEDIA_MANUAL%")
                End If
                If FileIO.FileSystem.DirectoryExists(HL_Path + "\Media\Music\" + sys + "\_Default") Then
                    If FileIO.FileSystem.GetFiles(HL_Path + "\Media\Music\" + sys + "\_Default", FileIO.SearchOption.SearchAllSubDirectories, {"*.mp3"}).Count > 0 Then Systems(k).Add("%HLMEDIA_MUSIC%")
                End If
                If FileIO.FileSystem.DirectoryExists(HL_Path + "\Media\Videos\" + sys + "\_Default") Then
                    If FileIO.FileSystem.GetFiles(HL_Path + "\Media\Videos\" + sys + "\_Default", FileIO.SearchOption.SearchAllSubDirectories, {"*.mp4", "*.avi"}).Count > 0 Then Systems(k).Add("%HLMEDIA_VIDEO%")
                End If

            Next
        End If

        'FILL GRID
        status_Label.BeginInvoke(Sub() status_Label.Text = "[Task Final] Filling grid... ")
        For Each l As List(Of String) In Systems.Values
            Dim r As DataGridViewRow = scan_FillRowSub(l, required_media_list, dont_show_completed, required_media_number)
            If r IsNot Nothing Then
                'add last_check_result
                tmp = ini.IniReadValue("LastCheckResult", l(0))
                If tmp <> "" Then
                    Dim arr() As String = tmp.Split({","c}, StringSplitOptions.RemoveEmptyEntries) 'total, rom, vid, thm, w, a1-a4, snd
                    Dim arr_int = arr.ToList.ConvertAll(Function(str) Int32.Parse(str))

                    For i As Integer = 1 To 4
                        r.Cells(i + 17).Value = arr_int(i).ToString + " \ " + arr_int(0).ToString
                        If arr_int(i) = arr_int(0) Then
                            r.Cells(i + 17).Style.BackColor = Class1.colorYES
                        ElseIf arr_int(i) = 0 Then
                            r.Cells(i + 17).Style.BackColor = Class1.colorNO
                        Else
                            r.Cells(i + 17).Style.BackColor = Color.Yellow
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
                    r.Cells(22).Value = tmp
                    If atLeastOneIncomplete Then r.Cells(22).Style.BackColor = Color.Yellow
                    If Not atLeastOneIncomplete And atLeastOneComplete Then r.Cells(22).Style.BackColor = Class1.colorYES
                    If Not atLeastOneIncomplete And Not atLeastOneComplete Then r.Cells(22).Style.BackColor = Class1.colorNO
                End If
                'add row to grid
                grid.BeginInvoke(Sub() grid.Rows.Add(r))
            End If
        Next
        'Form1.Label41.Text = "System count:  " + Form1.DataGridView2.Rows.Count.ToString + ", Active:" + activeCount.ToString
        status_Label.BeginInvoke(Sub() status_Label.Text = "System count: " + grid.Rows.Count.ToString + ", Active:" + activeCount.ToString)
        checked = True
    End Sub

    Private Function scan_FillRowSub(l As List(Of String), Optional req_ml As String() = Nothing, Optional req_notCompleted As Boolean = False, Optional req_N As Integer = 0) As DataGridViewRow
        Dim x(17) As String
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
        If l.Contains("%HLMEDIA_ARTWORK%") Then x(10) = "X" Else x(10) = ""
        If l.Contains("%HLMEDIA_BACKG%") Then x(11) = "X" Else x(11) = ""
        If l.Contains("%HLMEDIA_BEZEL%") Then x(12) = "X" Else x(12) = ""
        If l.Contains("%HLMEDIA_FADE%") Then x(13) = "X" Else x(13) = ""
        If l.Contains("%HLMEDIA_GUIDE%") Then x(14) = "X" Else x(14) = ""
        If l.Contains("%HLMEDIA_MANUAL%") Then x(15) = "X" Else x(15) = ""
        If l.Contains("%HLMEDIA_MUSIC%") Then x(16) = "X" Else x(16) = ""
        If l.Contains("%HLMEDIA_VIDEO%") Then x(17) = "X" Else x(17) = ""

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
            For i As Integer = 1 To 17
                If x(i) = "X" Or x(i) = "HL" Then
                    r.Cells(i).Style.BackColor = Class1.colorYES
                Else
                    r.Cells(i).Style.BackColor = Class1.colorNO
                    If x(i).ToUpper = "PER GAME" Then
                        r.Cells(i).Style.BackColor = Color.YellowGreen
                    End If
                    If x(i) = "INVALID HL PATH" Then
                        Dim f As Font = New Font(Control.DefaultFont.FontFamily, 6, FontStyle.Regular)
                        r.Cells(i).Style.Font = f
                    End If
                    If i >= 10 And i <= 17 And Not need_check_hl_media.Checked Then r.Cells(i).Value = "not checked" : r.Cells(i).Style.BackColor = Color.White
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
        'Dim modules_for_cur_system As New List(Of String)
        Dim modules_for_cur_system As New Dictionary(Of String, String)

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
                    modules_for_cur_system.Add(m.Key, m.Value(0))
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

        'Dim HL_Path As String = iniClass.GetKeyValue("Main", "Hyperlaunch_Path").Trim
        'If HL_Path.ToUpper.EndsWith("EXE") Then HL_Path = HL_Path.Substring(0, HL_Path.LastIndexOf("\"))
        'If Not HL_Path = "" AndAlso (Not FileExists(HL_Path + "\HyperLaunch.exe") And Not FileExists(HL_Path + "\RocketLauncher.exe")) Then HL_Path = ""
        Dim HL_Path As String = Class1.HyperlaunchPath
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

                            Dim m_raw As String = iniclass2.GetKeyValue(emu, "Module").Trim
                            Dim m As String = m_raw
                            If m.Contains("\") Then m = m.Substring(m.LastIndexOf("\") + 1)
                            Dim l As New List(Of String)
                            l.Add(m_raw) : l.Add(m) : l.Add(emuPath)
                            emulators(emu) = l
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
                ind = Systems(sys).IndexOf("%5HL%") : If ind >= 0 Then Systems(sys).RemoveAt(ind)
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
                r.Cells(i).Style.BackColor = Class1.colorNO
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
                Dim col = hitTest.ColumnIndex
                If col = 2 Or col = 3 Or col = 4 Or col = 8 Or col = 9 Then
                    grid.DefaultCellStyle.SelectionBackColor = Color.LightGoldenrodYellow
                Else
                    grid.DefaultCellStyle.SelectionBackColor = Color.Orange
                End If
                grid.Rows(hitTest.RowIndex).Cells(hitTest.ColumnIndex).Selected = True
            End If
        End If
    End Sub

    Private Sub drag_leave(sender As Object, e As System.EventArgs) Handles grid.DragLeave
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

            Dim files As String() = DirectCast(e.Data.GetData(DataFormats.FileDrop), String())
            If hitTest.RowIndex >= 0 And files.Count > 0 Then
                Dim sys As String = grid.Rows(hitTest.RowIndex).Cells(0).Value.ToString
                drag_drop_action(sys, hitTest.ColumnIndex, files)
            End If

            grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            If hitTest.RowIndex >= 0 And hitTest.ColumnIndex >= 0 Then grid.Rows(hitTest.RowIndex).Selected = True
        End If
    End Sub
    Private Sub drag_drop_action(sys As String, action As Integer, files() As String)
        Select Case action
            Case 2 'Database
                If files.Count > 1 Then MsgBox("You can not set multiple files as database.") : Exit Select
                Dim file = files(0)
                Dim file_no_path = file.Substring(file.LastIndexOf("\") + 1)
                Dim file_no_ext = file_no_path.Substring(0, file_no_path.LastIndexOf("."))

                If Not file_no_path.ToUpper.EndsWith(".XML") Then
                    Dim res = MsgBox("You filename have different extension then .xml. Are you sure you want to set it as database (it will be renamed to .xml)?", MsgBoxStyle.YesNo)
                    If res = MsgBoxResult.No Then Exit Select
                End If
                If file_no_ext.ToUpper <> sys.ToUpper And Systems.Keys.Contains(file_no_ext.ToUpper) Then
                    Dim res = MsgBox("You filename correspond to system:" + vbCrLf + file_no_ext + vbCrLf + "but you dragged it on:" + vbCrLf + sys + vbCrLf + vbCrLf + "Do you want to use it on " + sys + " (press Yes) or on " + file_no_ext + " (press No)?", MsgBoxStyle.YesNoCancel)
                    If res = MsgBoxResult.Cancel Then Exit Select
                    If res = MsgBoxResult.No Then sys = file_no_ext
                End If
                Dim dest = (Class1.HyperspinPath + "\Databases\" + sys + "\" + sys + ".xml").Replace("\\", "\")
                If file.ToUpper = dest.ToUpper Then MsgBox("Can't copy file to itself.") : Exit Select
                If Not DirectoryExists(Class1.HyperspinPath + "\Databases\" + sys) Then CreateDirectory(Class1.HyperspinPath + "\Databases\" + sys)
                If FileExists(dest) Then
                    Dim fname = sys + " " + DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss.fff") + ".xml"
                    fname = IO.Path.GetDirectoryName(dest) + "\" + fname
                    MoveFile(dest, fname)
                End If
                FileCopy(file, dest)

                For i As Integer = 0 To grid.Rows.Count - 1
                    If grid.Rows(i).Cells(0).Value.ToString.ToUpper = sys.ToUpper Then
                        grid.Rows(i).Cells(2).Value = "NEW"
                        grid.Rows(i).Cells(2).Style.BackColor = Class1.colorYES
                        If Not Systems(sys.ToUpper).Contains("%2%") Then Systems(sys.ToUpper).Add("%2%")
                        Exit For
                    End If
                Next
            Case 3 'Main menu theme
                If files.Count > 1 Then MsgBox("You can not set multiple files as main menu theme.") : Exit Select
                Dim file = files(0)
                Dim file_no_path = file.Substring(file.LastIndexOf("\") + 1)
                Dim file_no_ext = file_no_path.Substring(0, file_no_path.LastIndexOf("."))
                If Not file_no_path.ToUpper.EndsWith(".ZIP") Then
                    Dim res = MsgBox("You filename have different extension then .zip. Unzipped themes not allowed yet.")
                    Exit Select
                End If
                If file_no_ext.ToUpper <> sys.ToUpper And Systems.Keys.Contains(file_no_ext.ToUpper) Then
                    Dim res = MsgBox("You filename correspond to system:" + vbCrLf + file_no_ext + vbCrLf + "but you dragged it on:" + vbCrLf + sys + vbCrLf + vbCrLf + "Do you want to use it on " + sys + " (press Yes) or on " + file_no_ext + " (press No)?", MsgBoxStyle.YesNoCancel)
                    If res = MsgBoxResult.Cancel Then Exit Select
                    If res = MsgBoxResult.No Then sys = file_no_ext
                End If
                Dim dest = (Class1.HyperspinPath + "\Media\Main Menu\Themes\" + sys + ".zip").Replace("\\", "\")
                If file.ToUpper = dest.ToUpper Then MsgBox("Can't copy file to itself.") : Exit Select
                If Not DirectoryExists(Class1.HyperspinPath + "\Media\Main Menu\Themes") Then CreateDirectory(Class1.HyperspinPath + "\Media\Main Menu\Themes")
                If FileExists(dest) Then
                    Dim fname = sys + " " + DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss.fff") + ".zip"
                    fname = IO.Path.GetDirectoryName(dest) + "\" + fname
                    MoveFile(dest, fname)
                End If
                FileCopy(file, dest)

                For i As Integer = 0 To grid.Rows.Count - 1
                    If grid.Rows(i).Cells(0).Value.ToString.ToUpper = sys.ToUpper Then
                        grid.Rows(i).Cells(3).Value = "NEW"
                        grid.Rows(i).Cells(3).Style.BackColor = Class1.colorYES
                        If Not Systems(sys.ToUpper).Contains("%3%") Then Systems(sys.ToUpper).Add("%3%")
                        Exit For
                    End If
                Next
            Case 4 'System Theme
                Dim not_zip_message_shown As Boolean = False
                For Each file As String In files
                    If Not file.ToUpper.EndsWith(".ZIP") Then
                        If Not not_zip_message_shown Then
                            MsgBox("At least one of dragged files was not .zip. All no zip files will be ignored.")
                            not_zip_message_shown = True
                        End If
                        Continue For
                    End If

                    Dim file_no_path = file.Substring(file.LastIndexOf("\") + 1)
                    Dim file_no_ext = file_no_path.Substring(0, file_no_path.LastIndexOf("."))
                    Dim dest = (Class1.HyperspinPath + "\Media\" + sys + "\Themes\" + file_no_path).Replace("\\", "\")
                    If file.ToUpper = dest.ToUpper Then MsgBox("Can't copy file to itself. Abort.") : Exit Select
                    If Not DirectoryExists(Class1.HyperspinPath + "\Media\" + sys + "\Themes") Then CreateDirectory(Class1.HyperspinPath + "\Media\" + sys + "\Themes")
                    If FileExists(dest) Then
                        Dim fname = file_no_ext + " " + DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss.fff") + ".zip"
                        fname = IO.Path.GetDirectoryName(dest) + "\" + fname
                        MoveFile(dest, fname)
                    End If
                    FileCopy(file, dest)

                    For i As Integer = 0 To grid.Rows.Count - 1
                        If grid.Rows(i).Cells(0).Value.ToString.ToUpper = sys.ToUpper Then
                            grid.Rows(i).Cells(4).Value = "NEW"
                            grid.Rows(i).Cells(4).Style.BackColor = Class1.colorYES
                            If Not Systems(sys.ToUpper).Contains("%4%") Then Systems(sys.ToUpper).Add("%4%")
                            Exit For
                        End If
                    Next
                Next
            Case 8 'Main Menu Wheel
                If files.Count > 1 Then MsgBox("You can not set multiple files as main menu wheel.") : Exit Select
                Dim file = files(0)
                Dim file_no_path = file.Substring(file.LastIndexOf("\") + 1)
                Dim file_no_ext = file_no_path.Substring(0, file_no_path.LastIndexOf("."))
                If Not file_no_path.ToUpper.EndsWith(".PNG") Then
                    Dim res = MsgBox("You filename have different extension then .png. Aborting.")
                    Exit Select
                End If
                If file_no_ext.ToUpper <> sys.ToUpper And Systems.Keys.Contains(file_no_ext.ToUpper) Then
                    Dim res = MsgBox("You filename correspond to system:" + vbCrLf + file_no_ext + vbCrLf + "but you dragged it on:" + vbCrLf + sys + vbCrLf + vbCrLf + "Do you want to use it on " + sys + " (press Yes) or on " + file_no_ext + " (press No)?", MsgBoxStyle.YesNoCancel)
                    If res = MsgBoxResult.Cancel Then Exit Select
                    If res = MsgBoxResult.No Then sys = file_no_ext
                End If
                Dim dest = (Class1.HyperspinPath + "\Media\Main Menu\Images\Wheel\" + sys + ".png").Replace("\\", "\")
                If file.ToUpper = dest.ToUpper Then MsgBox("Can't copy file to itself.") : Exit Select
                If Not DirectoryExists(Class1.HyperspinPath + "\Media\Main Menu\Images\Wheel") Then CreateDirectory(Class1.HyperspinPath + "\Media\Main Menu\Images\Wheel")
                If FileExists(dest) Then
                    Dim fname = sys + " " + DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss.fff") + ".png"
                    fname = IO.Path.GetDirectoryName(dest) + "\" + fname
                    MoveFile(dest, fname)
                End If
                FileCopy(file, dest)

                For i As Integer = 0 To grid.Rows.Count - 1
                    If grid.Rows(i).Cells(0).Value.ToString.ToUpper = sys.ToUpper Then
                        grid.Rows(i).Cells(8).Value = "NEW"
                        grid.Rows(i).Cells(8).Style.BackColor = Class1.colorYES
                        If Not Systems(sys.ToUpper).Contains("%8%") Then Systems(sys.ToUpper).Add("%8%")
                        Exit For
                    End If
                Next
            Case 9 'Main Menu Video
                If files.Count > 1 Then MsgBox("You can not set multiple files as main menu video.") : Exit Select
                Dim file = files(0)
                Dim file_no_path = file.Substring(file.LastIndexOf("\") + 1)
                Dim file_no_ext = file_no_path.Substring(0, file_no_path.LastIndexOf("."))
                If Not file_no_path.ToUpper.EndsWith(".MP4") Then
                    Dim res = MsgBox("You filename have different extension then .png. Aborting.")
                    Exit Select
                End If
                If file_no_ext.ToUpper <> sys.ToUpper And Systems.Keys.Contains(file_no_ext.ToUpper) Then
                    Dim res = MsgBox("You filename correspond to system:" + vbCrLf + file_no_ext + vbCrLf + "but you dragged it on:" + vbCrLf + sys + vbCrLf + vbCrLf + "Do you want to use it on " + sys + " (press Yes) or on " + file_no_ext + " (press No)?", MsgBoxStyle.YesNoCancel)
                    If res = MsgBoxResult.Cancel Then Exit Select
                    If res = MsgBoxResult.No Then sys = file_no_ext
                End If
                Dim dest = (Class1.HyperspinPath + "\Media\Main Menu\Video\" + sys + ".mp4").Replace("\\", "\")
                If file.ToUpper = dest.ToUpper Then MsgBox("Can't copy file to itself.") : Exit Select
                If Not DirectoryExists(Class1.HyperspinPath + "\Media\Main Menu\Video") Then CreateDirectory(Class1.HyperspinPath + "\Media\Main Menu\Video")
                If FileExists(dest) Then
                    Dim fname = sys + " " + DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss.fff") + ".mp4"
                    fname = IO.Path.GetDirectoryName(dest) + "\" + fname
                    MoveFile(dest, fname)
                End If
                FileCopy(file, dest)

                For i As Integer = 0 To grid.Rows.Count - 1
                    If grid.Rows(i).Cells(0).Value.ToString.ToUpper = sys.ToUpper Then
                        grid.Rows(i).Cells(9).Value = "NEW"
                        grid.Rows(i).Cells(9).Style.BackColor = Class1.colorYES
                        If Not Systems(sys.ToUpper).Contains("%9%") Then Systems(sys.ToUpper).Add("%9%")
                        Exit For
                    End If
                Next
        End Select
        If frm.Visible Then setPropertiesFormData()
    End Sub
#End Region
End Class
