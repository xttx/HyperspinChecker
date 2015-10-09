﻿Imports WindowsApplication1.Language
Imports Microsoft.VisualBasic.FileIO
Imports System.Runtime.InteropServices

Public Class Form1
    'QRS #49 Thanks for your help!! Problem is solved. I just checked the last cue file and the quotations was missing in the name inside the cue file Example: Supreme Warrior (USA) (Disc 1).bin instead of "Supreme Warrior (USA) (Disc 1).bin"
    'aupton #63 When selecting "Move unneeded roms or media to subfolder" the app thows a JIT error.
    'MadCoder #85 Well, I had an idea about the fix ISO for DC...I going to add Dreamcast CDI to ISO ability, and ISO fixer.
    'RandomName1 #91 - Would it be possible to strip tags from rom descriptions automatically (except disc numbers) when creating an xml from a rom folder?
    'Obiwantje #103 - updateble HS ini files
    'desrop69 #105 - case sensitive crc
    'Acidnine #108 - When I check the master snes .xml db file it says there's an error loading it. When I tell it to fix it, it replaces all the ampersands '&' with '&amp;' in the <game name=""> tag. Is this wanted? because now I have to rename my roms with ampersands in them to match the name tag.
    'tastyratz #112 - I encountered a bug when trying to create a cue file for a saturn bin which does not have one. It gave me the error, created a 0kb file, and left the file locked but unfinished. Any subsequent attempts show the file in use error below. This error mentions the F:\ path but I do not have an F:\ drive, maybe related?
    'tastyratz #112 - I want to be able to treat iso, cue like bin, cue because all of my sega cd collection is iso & cue. Would that be the right field (handle together) for entry?
    'tastyratz #112 - Finally another bug: When matching an iso/cue set in subfolder it tells me it cannot find the original iso filename - however it does rename and edit the iso/cue just fine (just not the folder). If I click it a second time it renames the folder without errors.
    'brownvim #118 - hide clones in a MAME list. I want to check if im missing any unique roms but so many are red (clones).
    'Andyman #121 - Is there a way to show which games are clones in MAME? That would be a tremendous help. 
    'baezl #128 - How do you guys handle multiple disc games (Sega CD)? What is your structure so we can use this kick ass tool which I freaking love!!!! I noticed the current database has Disc1 and Disc2 separate? On some games, (Sega CD again), I get "Length Less Than Zero" or "File already open".
    'rpdmatt #129 - After choosing the system I can try to check for roms in the Summary page and it will populate everything from my database xml file but won't pick up on the roms that I have. I have tried going to the HyperSpin system settings and changing the 'Rom Path:' but it still doesn't affect the outcome. Also, if I change the 'Video Path:' to something else I cannot get the outcome to change. It seems like the user settings on the HyperSpin system settings tab don't affect the actual check on the Summary tab. Not sure if this is just the case since the new HyperLaunch HQ or not. Has anyone been able to change the 'Rom Path:' and have it affect the check? 
    'potts43 #137 - Hyperpause media...? Controller Manual Guides Fades etc... It would be great to have columns for all of these folders 
    'zero dreams #146 - When the check button is selected, it does a check and returns to the top of the list at the beginning of the database. I would love it if there was an option to turn this off. I'm basically looking for something like how HLHQ audits games and the list remains static during and after the audit. 
    'Pike #147, Akelarre #149 - is it possible to have another button ("duplicate" or whatever you want), which is keeping the original art file and make a copy of it with the new name ? I'm asking this because when you have localized games not from original HS XML database but added to it, all art files already exist with the original rom, and we just need to copy-rename these files.
    'potts43 #151 - multibin games related
    'tanin #152 - error
    'potts43 #154 - If possible have two way renaming. At the moment you can rename roms and artwork from the xml but for systems like future pinball and WinUAEloader that change regularly it would be good to point the DB xml to an updated rom and update the xml to match that newer rom name.
    'when switching rom path (HL / HS), and then recheck without reselect system, path is not updated
    'when changing custom rom path in matcher, sometimes got error (try use MAME path)
    'HL multiple path support (move unneeded - done)
    'Rename media when edit romname
    'To see if hs path is set, system manager refers to Label23.color, which can be not necessary green but purple, if path is ok, but config.save clicked.

    'DONE TODO thrasherx #119, Turranius #160 - Any chance you could filter the MAME BIOS files from the 'move unneeded roms' option? I consolidated my MAME roms and it took my BIOS files with it. 
    'MOSTLY DONE TODO flames #153 - print check table on printer or exel export
    'DONE TODO potts43 #154 - On the renamer page have a totals number always present.eg.  Unmatched/total  Matched/total  Both/total
    'MOSTLY DONE TODO potts43 #154 - Be able to add/copy/delete xml entries. This way you could add a blank entry and complete it and then save it.
    'DONE TODO potts43 #154 - Be able to edit both name and description in the xml. Even better to have all fields. This way you can have custom Xmls made easy.
    'DONE TODO Yardley #159 - option to move themes in the "Move unneeded roms or media to subfolder" options. Seems simple to add this feature. 
    'DONE TODO Added default icon
    'DONE TODO Ability to change HyperLaunch path in HS Settings.ini.
    'DONE TODO Ability to change HyperLaunch / HyperLaunchHQ ini to make HS easier to move from one drive/folder to another
    'DONE TODO Speed up loading time
    'DONE TODO Iso Checker now tracks wrong formated GDI (no fix though)

    'TODO option to rename audio files associated with .cue (ME)
    'TODO It appears to be setting the "video" folder for Visual Pinball to the previous system a "Check" was ran on. In the Matcher tab, even when I tell it to use a custom folder and point it to the correct Visual Pinball video folder, once I perform the "Associate" action, it's still copying the video files out of my Visual Pinball video folder and moving them to the Gamecube video folder (note: Gamecube was the last system I had ran a "Check" on). So, the Matcher seems to be using the custom folder fine to actually find the video files, but when it performs an "Associate", something is getting confused. Not a big deal, but thought I'd pass the info along. I've tried using both the System.ini and default Hyperspin Video location settings and happens for both. Let me know if you have any questions on my setup & thanks for the great tool.  (Signet145 #47)
    'TODO undo the renaming INSIDE the .cue. The .cue and .bin files both get renamed back to the correct original file name, but inside the cue file does not. (windowlicker #27)
    'TODO undo message box displays a couple of extra backslashes, but it does not affect anything. (windowlicker #27)
    'TODO REQUEST [MadCoder] Dreamcast CDI to ISO ability, and ISO fixer
    'TODO Restore .cues from backup
    'TODO in button20(markAsFound) and TextBox4(matcher filelist) WE ALSO NEED TO CHECK IF A FILE WITH ROMNAME AND APPROPRIATE EXTENSION EXISTS IN SUBDIR
    'TODO check and improve collect media dialog
    'TODO check/verify all possibilities while associating in matcher (copy, move, in/from different folders, and check listBoxex behaviour)
    'TODO when move roms in subfolder in NOT subfoldered mode, filelist don't reflect changes
    'DONE ADDED notes
    'DONE ADDED tool to move roms found in .txt list to subfolder
    'DONE ADDED tool to remove clones in HSxmlDB
    'DONE ADDED HLv3 support (but all extensions set to *)
    'DONE TODO REQUEST [RandomName1] - Would it be possible to strip tags from rom descriptions automatically (except disc numbers) when creating an xml from a rom folder? 
    'DONE TODO REQUEST [goldorakiller] - make a datfile from hyperlist files and reverse too
    'DONE TODO REQUEST [goldorakiller] - Add filters for mame. Like ini files that are in "folders" folder in mame.
    'DONE TODO REQUEST [goldorakiller] - Add cropping for filtered using MAME folders.
    'DONE TODO remove .cue/.bin errors from list if this error has been fixed
    'DONE TODO move unneded to subfolder for other media than roms
    'DONE TODO Check is very slow
    'DONE TODO folder to xml
    'DONE TODO mark files (in current folder) as 'found'
    'DONE TODO check if file exists while just copy from another folder but name is good
    'DONE TODO subfoldered mode limitation for artwork other then roms
    'DONE TODO ASSOCIATE move video and other artwork in rom folder
    'DONE TODO check if file already exist, while ASSOCIATING
    'DONE TODO restrict extension in matcher filelist to match selected media, and not only roms
    'DONE TODO "MARK AS FOUND" is enable with empty filelist
    'DONE TODO don't rename files in file textbox if just copying without renaming in source dir
    'DONE TODO mark associated as found to prevent associate two times, and mark DEASSOCIATED as UNFOUND
    'DONE TODO change HS path somewhere, and save config
    'DONE TODO repair wrong formated XML (with unescaped &)
    'DONE TODO handle mdf/mds cue/bin and other in matcher
    'DONE TODO read/save in config ALL setings
    'DONE TODO when using 'other directory' in matcher, and recheck rom, dir is reseted.
    'DONE TODO "MARK AS FOUND" should only be enabled when in foreign directory
    'DONE TODO check if rom path exist (in check ISO e.x.)
    'DONE TODO check all '...' buttons
    'DONE TODO check video in default HS video folder
    'DONE TODO move roms in subfolder in NOT subfoldered mode

#Region "Declarations"
    Public xmlPath As String = ""
    Private ini As New IniFileApi
    Private xml_class As Class2_xml
    Private matcher_class As Class3_matcher
    Private system_manager_class As Class5_system_manager
    Private isoChecker As ISOChecker
    Private clrmame_class As Class4_clrmamepro
    Public editor_delete_command_list As New List(Of DataGridViewRow)
    Public editor_insert_command_list As New List(Of DataGridViewRow)
    Public editor_update_command_list As New Dictionary(Of DataGridViewRow, String)
    Friend subfoldered As Boolean = False
    Friend subfoldered2 As Boolean = False
    Friend undo As New List(Of List(Of String))
    Friend colorNO As Color = Color.OrangeRed
    Friend colorYES As Color = Color.LightGreen
    Friend colorPAR As Color = Color.FromArgb(&HFF, &HB0, &HFF, &HB0)
    Private romExtensions() As String = {""}
    Private oldCombo4Value As Integer = 0
    Private useParentVids, useParentThemes As Boolean
    Private WithEvents myContextMenu As New ToolStripDropDownMenu 'checking for roms in other folders
    Friend WithEvents myContextMenu2 As New ToolStripDropDownMenu 'matcher options
    Private WithEvents myContextMenu3 As New ToolStripDropDownMenu 'move unneeded
    Private WithEvents myContextMenu4 As New ToolStripDropDownMenu 'main table columns hide/show
    Public WithEvents myContextMenu5 As New ToolStripDropDownMenu 'folder2xml options
    Friend WithEvents myContextMenu6 As New ToolStripDropDownMenu 'convert to clrmamepro
    'Public WithEvents myContextMenu7 As New ToolStripDropDownMenu 'autorenamer
    Friend WithEvents RadioStrip1 As New RadioButton
    Friend WithEvents RadioStrip2 As New RadioButton
    'Friend WithEvents CheckStrip1 As New CheckBox
    Friend WithEvents TextStrip1 As New TextBox
    Friend WithEvents ButtonStrip1 As New Button With {.Text = "UNDO last operation."}
    Dim WithEvents CheckStrip2_0 As New CheckBox With {.Name = "00", .Text = "Show Database Entry", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = True, .Enabled = False}
    Dim WithEvents CheckStrip2_1 As New CheckBox With {.Name = "01", .Text = "Show Romname", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = True, .Enabled = False}
    Dim WithEvents CheckStrip2_2 As New CheckBox With {.Name = "02", .Text = "Show Rom", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = True}
    Dim WithEvents CheckStrip2_3 As New CheckBox With {.Name = "03", .Text = "Show Video", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = True}
    Dim WithEvents CheckStrip2_4 As New CheckBox With {.Name = "04", .Text = "Show Theme", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = True}
    Dim WithEvents CheckStrip2_5 As New CheckBox With {.Name = "05", .Text = "Show Wheel", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = True}
    Dim WithEvents CheckStrip2_6 As New CheckBox With {.Name = "06", .Text = "Show Artwork1", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = True}
    Dim WithEvents CheckStrip2_7 As New CheckBox With {.Name = "07", .Text = "Show Artwork2", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = True}
    Dim WithEvents CheckStrip2_8 As New CheckBox With {.Name = "08", .Text = "Show Artwork3", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = True}
    Dim WithEvents CheckStrip2_9 As New CheckBox With {.Name = "09", .Text = "Show Artwork4", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = True}
    Dim WithEvents CheckStrip2_10 As New CheckBox With {.Name = "10", .Text = "Show Sounds", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = True}
    Dim WithEvents CheckStrip2_11 As New CheckBox With {.Name = "11", .Text = "Show CRC", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = False}
    Dim WithEvents CheckStrip2_12 As New CheckBox With {.Name = "12", .Text = "Show Manufacturer", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = False}
    Dim WithEvents CheckStrip2_13 As New CheckBox With {.Name = "13", .Text = "Show Year", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = False}
    Dim WithEvents CheckStrip2_14 As New CheckBox With {.Name = "14", .Text = "Show Genre", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = False}
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> Private Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function
#End Region

#Region "Main Form Actions (loadForm, system select, check)"

    'FORM LOAD
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Class1.Log("Initializing...")
        Class1.Log("Creating menus")
        DataGridView1.Columns.Add("col0", "Database entry")
        DataGridView1.Columns.Add("col1", "RomName")
        DataGridView1.Columns.Add("col2", "Rom")
        DataGridView1.Columns.Add("col3", "Vid")
        DataGridView1.Columns.Add("col4", "Thm")
        DataGridView1.Columns.Add("col5", "Whl")
        DataGridView1.Columns.Add("col6", "Art1")
        DataGridView1.Columns.Add("col7", "Art2")
        DataGridView1.Columns.Add("col8", "Art3")
        DataGridView1.Columns.Add("col9", "Art4")
        DataGridView1.Columns.Add("col10", "Snd")
        DataGridView1.Columns.Add("col11", "crc")
        DataGridView1.Columns.Add("col12", "Manufacturer")
        DataGridView1.Columns.Add("col13", "Year")
        DataGridView1.Columns.Add("col14", "Genre")
        DataGridView1.Columns(0).Width = 300 : DataGridView1.Columns(0).ReadOnly = False
        DataGridView1.Columns(1).Width = 100 : DataGridView1.Columns(1).ReadOnly = True
        DataGridView1.Columns(2).Width = 50 : DataGridView1.Columns(2).ReadOnly = True
        DataGridView1.Columns(3).Width = 50 : DataGridView1.Columns(3).ReadOnly = True
        DataGridView1.Columns(4).Width = 50 : DataGridView1.Columns(4).ReadOnly = True
        DataGridView1.Columns(5).Width = 50 : DataGridView1.Columns(5).ReadOnly = True
        DataGridView1.Columns(6).Width = 50 : DataGridView1.Columns(6).ReadOnly = True
        DataGridView1.Columns(7).Width = 50 : DataGridView1.Columns(7).ReadOnly = True
        DataGridView1.Columns(8).Width = 50 : DataGridView1.Columns(8).ReadOnly = True
        DataGridView1.Columns(9).Width = 50 : DataGridView1.Columns(9).ReadOnly = True
        DataGridView1.Columns(10).Width = 50 : DataGridView1.Columns(10).ReadOnly = True
        DataGridView1.Columns(11).Width = 75 : DataGridView1.Columns(11).ReadOnly = False : DataGridView1.Columns(11).Visible = False
        DataGridView1.Columns(12).Width = 180 : DataGridView1.Columns(12).ReadOnly = False : DataGridView1.Columns(12).Visible = False
        DataGridView1.Columns(13).Width = 65 : DataGridView1.Columns(13).ReadOnly = False : DataGridView1.Columns(13).Visible = False
        DataGridView1.Columns(14).Width = 120 : DataGridView1.Columns(14).ReadOnly = False : DataGridView1.Columns(14).Visible = False
        Dim t As Type = DataGridView1.GetType
        Dim pi As Reflection.PropertyInfo = t.GetProperty("DoubleBuffered", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)
        pi.SetValue(DataGridView1, True, Nothing)

        DataGridView2.Columns.Add("col0", "System")
        DataGridView2.Columns.Add("col1", "Main Menu")
        DataGridView2.Columns.Add("col2", "Database")
        DataGridView2.Columns.Add("col3", "MainMenu Theme")
        DataGridView2.Columns.Add("col4", "System Theme")
        DataGridView2.Columns.Add("col5", "Settings")
        DataGridView2.Columns.Add("col6", "Emulator Path")
        DataGridView2.Columns.Add("col7", "Rom Path")
        DataGridView2.Columns(0).Width = 250
        DataGridView2.Columns(1).Width = 80 : DataGridView2.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView2.Columns(2).Width = 80 : DataGridView2.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView2.Columns(3).Width = 80 : DataGridView2.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView2.Columns(4).Width = 80 : DataGridView2.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView2.Columns(5).Width = 80 : DataGridView2.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView2.Columns(6).Width = 80 : DataGridView2.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView2.Columns(7).Width = 80 : DataGridView2.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        myContextMenu.Items.Add("check for missing Roms")
        myContextMenu.Items.Add(New ToolStripSeparator)
        myContextMenu.Items.Add("check for missing Video")
        myContextMenu.Items.Add(New ToolStripSeparator)
        myContextMenu.Items.Add("check for missing Wheels")
        myContextMenu.Items.Add("check for missing Artwork1")
        myContextMenu.Items.Add("check for missing Artwork2")
        myContextMenu.Items.Add("check for missing Artwork3")
        myContextMenu.Items.Add("check for missing Artwork4")

        myContextMenu3.Items.Add("move unneeded Roms")
        myContextMenu3.Items.Add(New ToolStripSeparator)
        myContextMenu3.Items.Add("move unneeded Video")
        myContextMenu3.Items.Add(New ToolStripSeparator)
        myContextMenu3.Items.Add("move unneeded Wheels")
        myContextMenu3.Items.Add("move unneeded Artwork1")
        myContextMenu3.Items.Add("move unneeded Artwork2")
        myContextMenu3.Items.Add("move unneeded Artwork3")
        myContextMenu3.Items.Add("move unneeded Artwork4")
        myContextMenu3.Items.Add("move unneeded Themes")
        myContextMenu3.Items.Add("move unneeded Sound")

        'myContextMenu7.Items.Add("Autorenamer")

        RadioStrip1.AutoSize = True
        RadioStrip2.AutoSize = True
        'CheckStrip1.AutoSize = True
        RadioStrip1.Checked = True
        'CheckStrip1.Checked = True
        'CheckStrip1.Enabled = False
        RadioStrip1.Text = "Use default HyperSpin folder"
        RadioStrip2.Text = "Use custom folder"
        'CheckStrip1.Text = "Copy to default HyperSpin folder"
        Dim striplabel As New Label With {.Text = "Overide extentions:"}
        RadioStrip1.BackColor = Color.FromArgb(0, 255, 0, 0)
        RadioStrip2.BackColor = Color.FromArgb(0, 255, 0, 0)
        'CheckStrip1.BackColor = Color.FromArgb(0, 255, 0, 0)
        striplabel.BackColor = Color.FromArgb(0, 255, 0, 0)

        Dim CheckStripHost1 As New ToolStripControlHost(RadioStrip1)
        Dim CheckStripHost2 As New ToolStripControlHost(RadioStrip2)
        'Dim CheckStripHost3 As New ToolStripControlHost(CheckStrip1)
        Dim CheckStripHost4 As New ToolStripControlHost(striplabel)
        Dim CheckStripHost5 As New ToolStripControlHost(TextStrip1)
        Dim CheckStripHost6 As New ToolStripControlHost(ButtonStrip1)
        CheckStripHost1.AutoSize = True
        CheckStripHost2.AutoSize = True
        'CheckStripHost3.AutoSize = True
        CheckStripHost4.AutoSize = True
        CheckStripHost5.AutoSize = False
        myContextMenu2.Items.Add(CheckStripHost1)
        myContextMenu2.Items.Add(CheckStripHost2)
        myContextMenu2.Items.Add(New ToolStripSeparator)
        'myContextMenu2.Items.Add(CheckStripHost3)
        myContextMenu2.Items.Add(New ToolStripSeparator)
        myContextMenu2.Items.Add(CheckStripHost4)
        myContextMenu2.Items.Add(CheckStripHost5)
        myContextMenu2.Items.Add(New ToolStripSeparator)
        myContextMenu2.Items.Add(CheckStripHost6)
        myContextMenu2.Items.Add(New ToolStripSeparator)
        Dim tooltip1 As New ToolTip()
        tooltip1.ShowAlways = True
        'tooltip1.SetToolTip(CheckStrip1, "Check this, if you want to copy roms or media from custom folder, to default hyperspin folder when you click 'associate'.")
        Dim tooltip2 As New ToolTip()
        tooltip2.ShowAlways = True
        tooltip2.SetToolTip(Button20_markAsFound, "Temporary mark all matching media found in this folder as 'found', in the main grid.")

        Dim CheckStripHost2_0 As New ToolStripControlHost(CheckStrip2_0)
        Dim CheckStripHost2_1 As New ToolStripControlHost(CheckStrip2_1)
        Dim CheckStripHost2_2 As New ToolStripControlHost(CheckStrip2_2)
        Dim CheckStripHost2_3 As New ToolStripControlHost(CheckStrip2_3)
        Dim CheckStripHost2_4 As New ToolStripControlHost(CheckStrip2_4)
        Dim CheckStripHost2_5 As New ToolStripControlHost(CheckStrip2_5)
        Dim CheckStripHost2_6 As New ToolStripControlHost(CheckStrip2_6)
        Dim CheckStripHost2_7 As New ToolStripControlHost(CheckStrip2_7)
        Dim CheckStripHost2_8 As New ToolStripControlHost(CheckStrip2_8)
        Dim CheckStripHost2_9 As New ToolStripControlHost(CheckStrip2_9)
        Dim CheckStripHost2_10 As New ToolStripControlHost(CheckStrip2_10)
        Dim CheckStripHost2_11 As New ToolStripControlHost(CheckStrip2_11)
        Dim CheckStripHost2_12 As New ToolStripControlHost(CheckStrip2_12)
        Dim CheckStripHost2_13 As New ToolStripControlHost(CheckStrip2_13)
        Dim CheckStripHost2_14 As New ToolStripControlHost(CheckStrip2_14)
        myContextMenu4.Items.Add(CheckStripHost2_0)
        myContextMenu4.Items.Add(CheckStripHost2_1)
        myContextMenu4.Items.Add(New ToolStripSeparator)
        myContextMenu4.Items.Add(CheckStripHost2_2)
        myContextMenu4.Items.Add(CheckStripHost2_3)
        myContextMenu4.Items.Add(CheckStripHost2_4)
        myContextMenu4.Items.Add(CheckStripHost2_5)
        myContextMenu4.Items.Add(CheckStripHost2_6)
        myContextMenu4.Items.Add(CheckStripHost2_7)
        myContextMenu4.Items.Add(CheckStripHost2_8)
        myContextMenu4.Items.Add(CheckStripHost2_9)
        myContextMenu4.Items.Add(CheckStripHost2_10)
        myContextMenu4.Items.Add(CheckStripHost2_11)
        myContextMenu4.Items.Add(CheckStripHost2_12)
        myContextMenu4.Items.Add(CheckStripHost2_13)
        myContextMenu4.Items.Add(CheckStripHost2_14)
        myContextMenu4.Items.Add(New ToolStripSeparator)
        myContextMenu4.Items.Add(Preset_Checker)
        myContextMenu4.Items.Add(Preset_Editor)
        myContextMenu4.Items.Add(New ToolStripSeparator)
        myContextMenu4.Items.Add(Save_current_cols_conf_as_startup)

        Panel1.Left = -100000
        ComboBox12.SelectedIndex = 0
        ComboBox13.SelectedIndex = 0
        ComboBox14.SelectedIndex = 0
        ComboBox15.SelectedIndex = 1
        ComboBox16.SelectedIndex = 0
        ComboBox17.SelectedIndex = 2
        Panel1.BackColor = Color.FromArgb(0, 255, 0, 0)
        Dim CheckHost11 As New ToolStripControlHost(Panel1)
        myContextMenu5.Items.Add(CheckHost11)

        Panel2.Left = -100000
        Panel2.BackColor = Color.FromArgb(0, 255, 0, 0)
        Dim CheckHost22 As New ToolStripControlHost(Panel2)
        myContextMenu6.Items.Add(CheckHost22)

        ComboBox4.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0

        Class1.Log("Loading config")
        If Not Microsoft.VisualBasic.FileIO.FileSystem.FileExists(Class1.confPath) Then
            Dim fd As New FolderBrowserDialog
            fd.Description = "Select hyperspin folder"
            fd.ShowDialog()
            FileOpen(1, Class1.confPath, OpenMode.Output)
            PrintLine(1, "[main]")
            PrintLine(1, "HyperSpin_Path = " + fd.SelectedPath)
            PrintLine(1, "rename_just_cue = 1")
            PrintLine(1, "search_cue_for = bin, iso")
            PrintLine(1, "useHLv3 = 0")
            PrintLine(1, "")
            PrintLine(1, "[handle_together]")
            PrintLine(1, "mds, mds")
            FileClose(1)
        End If

        Dim v As Boolean
        Dim s As String = ""
        Dim ini2 As New IniFileApi
        ini2.IniFile(Class1.confPath)
        TextBox14.Text = ini2.IniReadValue("main", "HyperSpin_Path")
        TextBox16.Text = ini2.IniReadValue("main", "search_cue_for")
        If Not ini2.IniReadValue("main", "rename_just_CUE") = "1" Then CheckBox6.Checked = False
        If ini2.IniReadValue("main", "usehlv3") = "1" Then CheckBox26.Checked = True
        Dim i As Integer = 1
        Do While ini2.IniReadValue("handle_together", i.ToString) <> ""
            ListBox4.Items.Add(ini2.IniReadValue("handle_together", i.ToString))
            i += 1
        Loop
        For c As Integer = 0 To DataGridView1.ColumnCount - 1
            s = ini2.IniReadValue("Main_Table_Columns_Config", "Col_" + c.ToString + "_visible")
            If s <> "" Then
                If s = "0" Then v = False Else v = True
                DataGridView1.Columns(c).Visible = v
            End If

            s = ini2.IniReadValue("Main_Table_Columns_Config", "Col_" + c.ToString + "_Width")
            If s <> "" Then DataGridView1.Columns(c).Width = CInt(s)
        Next
        s = ini2.IniReadValue("Main", "Main_window_size")
        If s <> "" Then
            Me.Width = CInt(s.Split({"x"c})(0))
            Me.Height = CInt(s.Split({"x"c})(1))
        End If

        If Not FileSystem.FileExists(Class1.HyperspinPath + "\Databases\Main Menu\Main Menu.xml") Then
            MsgBox("Can't find '" + Class1.HyperspinPath + "\Databases\Main Menu\Main Menu.xml'. Check hyperspin path under 'settings' tab, or in the config.conf")
        Else
            If Not FileSystem.FileExists(Class1.HyperspinPath + "\Settings\Settings.ini") Then
                MsgBox("Can't find '" + Class1.HyperspinPath + "\Settings\Settings.ini'. Check hyperspin path under 'settings' tab, or in the config.conf")
            End If
        End If

        Class1.Log("Init classes")
        xml_class = New Class2_xml
        matcher_class = New Class3_matcher
        isoChecker = New ISOChecker
        clrmame_class = New Class4_clrmamepro
        system_manager_class = New Class5_system_manager
        Class1.Log("Done init")
    End Sub

    'Form closing - this is handled to save window size in config
    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        ini.IniFile(Class1.confPath)
        ini.IniWriteValue("Main", "Main_window_size", Me.Width.ToString + "x" + Me.Height.ToString)
    End Sub

    'Main Check
    Private Sub check(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Class1.data.Clear()
        Class1.romlist.Clear()
        Class1.romFoundlist.Clear()
        Label2.Text = "..."
        Label2.Refresh()
        Button2_moveUnneeded.Enabled = False
        Button4.Enabled = False
        Button5_Associate.Enabled = False
        Button18.Enabled = False
        DataGridView1.Rows.Clear()
        ComboBox4.SelectedIndex = 0
        ComboBox3.SelectedIndex = -1
        If ComboBox1.SelectedIndex < 0 Then MsgBox("Please, select a system.") : Exit Sub

        retrieve_rom_path_from_HL()
        If Class1.romPath = "" Then MsgBox("Can't retrive rom path.") : Exit Sub

        Dim a(10) As String
        Dim romName As String
        Dim tempStr As String
        'Dim tempPath As String
        Dim x As New Xml.XmlDocument
        Try
            x.Load(xmlPath)
        Catch ex As Exception
            Dim res = MsgBox(ex.Message + vbCrLf + "An error ocured while loading a database xml. Should i try to fix it (I will make a backup :) )?", MsgBoxStyle.YesNo, "Corrupted XML.")
            If res = MsgBoxResult.Yes Then xml_class.xml_repare(xmlPath) Else Exit Sub
            Try
                x.Load(xmlPath)
            Catch ex2 As Exception
                MsgBox("Repair unsuccessful, sorry.") : Exit Sub
            End Try
        End Try

        ProgressBar1.Value = 0
        ProgressBar1.Maximum = x.SelectNodes("/menu/game").Count
        Dim a_crc, a_manufacturer, a_year, a_genre, a_cloneof As String
        For Each node As Xml.XmlNode In x.SelectNodes("/menu/game")
            If node.SelectSingleNode("crc") IsNot Nothing Then a_crc = node.SelectSingleNode("crc").InnerText Else a_crc = ""
            If node.SelectSingleNode("manufacturer") IsNot Nothing Then a_manufacturer = node.SelectSingleNode("manufacturer").InnerText Else a_manufacturer = ""
            If node.SelectSingleNode("year") IsNot Nothing Then a_year = node.SelectSingleNode("year").InnerText Else a_year = ""
            If node.SelectSingleNode("genre") IsNot Nothing Then a_genre = node.SelectSingleNode("genre").InnerText Else a_genre = ""
            If node.SelectSingleNode("cloneof") IsNot Nothing Then a_cloneof = node.SelectSingleNode("cloneof").InnerText Else a_cloneof = ""

            a(0) = node.SelectSingleNode("description").InnerText
            romName = node.Attributes.GetNamedItem("name").Value
            Class1.romlist.Add(romName.ToLower)

            a(2) = ""
            If Not Class1.romPath.Contains("|") Then 'Handle HL multiple paths
                'tempPath = Class1.romPath + romName
                If tryToFindRom(Class1.romPath, romExtensions, romName) Then a(2) = "YES"
            Else 'Multiple paths routine
                For Each temppath2 As String In Class1.romPath.Split({"|"c}, StringSplitOptions.RemoveEmptyEntries)
                    If Not temppath2.EndsWith("\") Then temppath2 = temppath2 + "\"
                    'temppath2 = temppath2 + romName
                    If tryToFindRom(temppath2, romExtensions, romName) Then a(2) = "YES" : Exit For
                Next
            End If
            If a(2) = "" Then a(2) = "NO" Else Class1.romFoundlist.Add(romName.ToLower)

            a(3) = "NO"
            Dim useVidFromParent As Boolean = False
            If CheckBox11.Checked Then
                If FileSystem.FileExists(Class1.videoPath + romName + ".flv") Then a(3) = "YES" Else a(3) = "NO"
                If a(3) = "NO" And CheckBox23.Checked And a_cloneof <> "" Then
                    If FileSystem.FileExists(Class1.videoPath + a_cloneof + ".flv") Then a(3) = "YES" : useVidFromParent = True
                End If
            End If
            If CheckBox12.Checked And a(3) = "NO" Then
                If FileSystem.FileExists(Class1.videoPath + romName + ".mp4") Then a(3) = "YES" Else a(3) = "NO"
                If a(3) = "NO" And CheckBox23.Checked And a_cloneof <> "" Then
                    If FileSystem.FileExists(Class1.videoPath + a_cloneof + ".mp4") Then a(3) = "YES" : useVidFromParent = True
                End If
            End If
            If CheckBox13.Checked And a(3) = "NO" Then
                If FileSystem.FileExists(Class1.videoPath + romName + ".png") Then a(3) = "YES" Else a(3) = "NO"
                If a(3) = "NO" And CheckBox23.Checked And a_cloneof <> "" Then
                    If FileSystem.FileExists(Class1.videoPath + a_cloneof + ".png") Then a(3) = "YES" : useVidFromParent = True
                End If
            End If

            Dim useThemeFromParent As Boolean = False
            Dim p As String = Class1.HyperspinPath + "Media\" + ComboBox1.SelectedItem.ToString
            If FileSystem.FileExists(p + "\Themes\" + romName + ".zip") Then a(4) = "YES" Else a(4) = "NO"
            If a(4) = "NO" And CheckBox24.Checked And a_cloneof <> "" Then
                If FileSystem.FileExists(p + "\Themes\" + a_cloneof + ".zip") Then a(4) = "YES" : useThemeFromParent = True
            End If
            If FileSystem.FileExists(p + "\Images\Wheel\" + romName + ".png") Or FileSystem.FileExists(p + "\Images\Wheel\" + romName + ".jpg") Then a(5) = "YES" Else a(5) = "NO"
            If FileSystem.FileExists(p + "\Images\Artwork1\" + romName + ".png") Or FileSystem.FileExists(p + "\Images\Artwork1\" + romName + ".jpg") Then a(6) = "YES" Else a(6) = "NO"
            If FileSystem.FileExists(p + "\Images\Artwork2\" + romName + ".png") Or FileSystem.FileExists(p + "\Images\Artwork2\" + romName + ".jpg") Then a(7) = "YES" Else a(7) = "NO"
            If FileSystem.FileExists(p + "\Images\Artwork3\" + romName + ".png") Or FileSystem.FileExists(p + "\Images\Artwork3\" + romName + ".jpg") Then a(8) = "YES" Else a(8) = "NO"
            If FileSystem.FileExists(p + "\Images\Artwork4\" + romName + ".png") Or FileSystem.FileExists(p + "\Images\Artwork4\" + romName + ".jpg") Then a(9) = "YES" Else a(9) = "NO"
            If FileSystem.FileExists(p + "\Sound\Background Music\" + romName + ".mp3") Then a(10) = "YES" Else a(10) = "NO"

            Dim r As New DataGridViewRow
            r.CreateCells(DataGridView1, {a(0), romName, a(2), a(3), a(4), a(5), a(6), a(7), a(8), a(9), a(10), a_crc, a_manufacturer, a_year, a_genre})
            tempStr = ""
            For i As Integer = 2 To 10
                If a(i) = "YES" Then r.Cells(i).Style.BackColor = colorYES Else r.Cells(i).Style.BackColor = colorNO
                tempStr = tempStr + a(i).Substring(0, 1)
            Next
            If useVidFromParent And a(3) = "YES" Then r.Cells(3).Style.BackColor = colorPAR
            If useThemeFromParent And a(4) = "YES" Then r.Cells(4).Style.BackColor = colorPAR
            Class1.data.Add(tempStr)
            DataGridView1.Rows.Add(r)
            ProgressBar1.Value = ProgressBar1.Value + 1 : If ProgressBar1.Value Mod 50 = 0 Then ProgressBar1.Refresh()
        Next
        ProgressBar1.Value = 0
        Button2_moveUnneeded.Enabled = True
        Button5_Associate.Enabled = True
        Button4.Enabled = True
        Button18.Enabled = True
        If AlowEditToolStripMenuItem.Checked Then Button21.Enabled = True
        Label2.Text = "Total: " + DataGridView1.Rows.Count.ToString
    End Sub

    Private Function tryToFindRom(ByVal temppath As String, ByVal romExtensions() As String, ByVal romname As String) As Boolean
        Dim result As Boolean = False
        For Each ext As String In romExtensions
            If Dir(temppath + romname + "." + ext.Trim) <> "" Then result = True : Exit For
        Next
        If result Then Return True

        result = False
        For Each ext As String In romExtensions
            Try
                If Dir(temppath + "\" + romname + "\" + romname + "." + ext.Trim) <> "" Then result = True : Exit For
            Catch ex As Exception
            End Try
        Next
        If result Then Return True
        Return False
    End Function

    'System select
    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Button2_moveUnneeded.Enabled = False
        Button4.Enabled = False
        Button5_Associate.Enabled = False
        Button18.Enabled = False
        Button21.Enabled = False

        ComboBox3.SelectedIndex = -1
        DataGridView1.Rows.Clear()
        If ComboBox1.SelectedIndex < 0 Then Exit Sub
        xmlPath = Class1.HyperspinPath + "\Databases\" + ComboBox1.SelectedItem.ToString + "\" + ComboBox1.SelectedItem.ToString + ".xml"
        If Not Microsoft.VisualBasic.FileIO.FileSystem.FileExists(xmlPath) Then
            MsgBox("Database file: '" & xmlPath & "' not found.")
            Exit Sub
        End If

        Dim iniFile As String = Class1.HyperspinPath + "\Settings\" + ComboBox1.SelectedItem.ToString + ".ini"
        If Not Microsoft.VisualBasic.FileIO.FileSystem.FileExists(iniFile) Then
            MsgBox("System settings file: '" & iniFile & "' not found.")
            Exit Sub
        End If

        Dim s As String
        'Dim found As Integer = 0
        Dim category As String = ""
        useParentVids = False : useParentThemes = False
        FileOpen(1, iniFile, OpenMode.Input)
        Do While Not EOF(1)
            s = LineInput(1).Trim
            If s.Trim.StartsWith("[") Then category = s
            If s.StartsWith("ROMPATH", StringComparison.InvariantCultureIgnoreCase) And category.ToLower.Contains("exe info") Then
                Class1.romPath = s.Substring(s.IndexOf("=") + 1).Trim
                TextBox1.Text = Class1.romPath : TextBox1.Enabled = True : Button14.Enabled = True
                TextBox17.Text = Class1.romPath
                If Class1.romPath.StartsWith(".") Then Class1.romPath = Class1.HyperspinPath + Class1.romPath
                If Not Class1.romPath.EndsWith("\") Then Class1.romPath = Class1.romPath + "\"
                'found += 1
            End If
            If s.StartsWith("romextension", StringComparison.InvariantCultureIgnoreCase) And category.ToLower.Contains("exe info") Then
                romExtensions = s.Substring(s.IndexOf("=") + 1).Trim.Split(","c)
                TextBox3.Text = s.Substring(s.IndexOf("=") + 1).Trim : TextBox3.Enabled = True
                'found += 1
            End If
            If s.StartsWith("path", StringComparison.InvariantCultureIgnoreCase) And category.ToUpper.Contains("VIDEO DEFAULTS") Then
                Class1.videoPath = s.Substring(s.IndexOf("=") + 1).Trim
                Class1.videoPathOrig = s.Substring(s.IndexOf("=") + 1).Trim
                TextBox2.Text = Class1.videoPathOrig : TextBox2.Enabled = True : Button11.Enabled = True

                If Not FileSystem.DirectoryExists(Class1.videoPath) Or RadioButton11.Checked Then Class1.videoPath = Class1.HyperspinPath + "Media\" + ComboBox1.SelectedItem.ToString + "\Video\"
                If Class1.videoPath.StartsWith(".") Then Class1.videoPath = Class1.HyperspinPath + Class1.videoPath
                If Not Class1.videoPath.EndsWith("\") Then Class1.videoPath = Class1.videoPath + "\"
                'found += 1
            End If
            If s.StartsWith("use_parent_vids", StringComparison.InvariantCultureIgnoreCase) And category.ToUpper.Contains("THEMES") Then
                If s.Substring(s.IndexOf("=") + 1).Trim.ToUpper = "TRUE" Then useParentVids = True
            End If
            If s.StartsWith("use_parent_themes", StringComparison.InvariantCultureIgnoreCase) And category.ToUpper.Contains("THEMES") Then
                If s.Substring(s.IndexOf("=") + 1).Trim.ToUpper = "TRUE" Then useParentThemes = True
            End If
            'If found = 3 Then Exit Do
        Loop
        FileClose(1)

        If CheckBox22.Checked Then CheckBox23.Checked = useParentVids
        If CheckBox22.Checked Then CheckBox24.Checked = useParentThemes
        'Button28.Text = "Remove all clones from current database" + vbCrLf + "(" + ComboBox1.SelectedItem.ToString + " selected)"
        RemoveClonesFromCurrentDBToolStripMenuItem.Text = "Remove all clones from current database (" + ComboBox1.SelectedItem.ToString + " selected)"
    End Sub

    Private Sub retrieve_rom_path_from_HL()
        'HLv3 Thing
        Dim s As String = ""
        If CheckBox26.Checked And ComboBox1.SelectedIndex >= 0 Then
            'Dim HLPath As String = ""
            Dim HLPath As String = TextBox18.Text
            If HLPath.ToUpper.EndsWith("EXE") Then HLPath = HLPath.Substring(0, HLPath.LastIndexOf("\"))
            'FileOpen(1, Class1.HyperspinPath + "\Settings\Settings.ini", OpenMode.Input)
            'Do While Not EOF(1)
            's = LineInput(1).Trim
            'If s.StartsWith("hyperlaunch_path", StringComparison.InvariantCultureIgnoreCase) Then
            'HLPath = s.Substring(s.IndexOf("=") + 1).Trim : Exit Do
            'End If
            'Loop
            'FileClose(1)
            If HLPath = "" Then MsgBox("Hyperlaunch_path in Settings.ini is not set. Rom path from HS .ini will be used.") : Exit Sub
            If HLPath.StartsWith(".") Then MsgBox("Relative hyperlaunch_path in Settings.ini is not supported. Rom path from HS .ini will be used.") : Exit Sub

            If Not HLPath.EndsWith("\") Then HLPath = HLPath + "\"
            Dim HLiniFile As String = HLPath + "Settings\" + ComboBox1.SelectedItem.ToString + "\Emulators.ini"
            If Not Microsoft.VisualBasic.FileIO.FileSystem.FileExists(HLiniFile) Then
                MsgBox("HL system settings file: '" & HLiniFile & "' not found. Rom path from HS .ini will be used.")
                Exit Sub
            End If

            Dim HLRomPath As String = ""
            FileOpen(1, HLiniFile, OpenMode.Input)
            Do While Not EOF(1)
                s = LineInput(1).Trim
                If s.StartsWith("Rom_Path", StringComparison.InvariantCultureIgnoreCase) Then
                    HLRomPath = s.Substring(s.IndexOf("=") + 1).Trim : Exit Do
                End If
            Loop
            FileClose(1)
            If HLRomPath = "" Then
                MsgBox("Rom_Path not set in '" & HLiniFile & "'. Rom path from HS .ini will be used.")
                Exit Sub
            End If

            Class1.romPath = ""
            For Each _hlRomPath As String In HLRomPath.Split({"|"c}, StringSplitOptions.RemoveEmptyEntries)
                If _hlRomPath.StartsWith(".") Then _hlRomPath = HLPath + _hlRomPath
                If Not _hlRomPath.EndsWith("\") Then _hlRomPath = _hlRomPath + "\"
                Class1.romPath = Class1.romPath + _hlRomPath + "|"
            Next
            Class1.romPath = Class1.romPath.Substring(0, Class1.romPath.Length - 1)
            romExtensions = {"*"}
        End If
    End Sub
#End Region

#Region "Main Form Commands (check Missing In Another Folder, Move Unneeded)"
    'checkMissingInAnotherFolder
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        myContextMenu.Show(Cursor.Position.X, Cursor.Position.Y)
    End Sub

    'checkMissingInAnotherFolder SUBItem click
    Private Sub contextMenuCheckRoms(ByVal sender As Object, ByVal e As ToolStripItemClickedEventArgs) Handles myContextMenu.ItemClicked
        If e.ClickedItem Is myContextMenu.Items(0) Then Class1.i = 0
        If e.ClickedItem Is myContextMenu.Items(2) Then Class1.i = 1
        If e.ClickedItem Is myContextMenu.Items(4) Then Class1.i = 3
        If e.ClickedItem Is myContextMenu.Items(5) Then Class1.i = 4
        If e.ClickedItem Is myContextMenu.Items(6) Then Class1.i = 5
        If e.ClickedItem Is myContextMenu.Items(7) Then Class1.i = 6
        If e.ClickedItem Is myContextMenu.Items(8) Then Class1.i = 7
        Form2.Show()
    End Sub

    'MoveUnneeded button
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2_moveUnneeded.Click
        myContextMenu3.Show(Cursor.Position.X, Cursor.Position.Y)
    End Sub

    'MoveUnneeded SUBItems click
    Private Sub contextMenuMoveRoms(ByVal sender As Object, ByVal e As ToolStripItemClickedEventArgs) Handles myContextMenu3.ItemClicked
        Dim mediaID As Integer = 0
        Dim mediaPath As String = Class1.HyperspinPath + "Media\" + ComboBox1.SelectedItem.ToString
        If e.ClickedItem Is myContextMenu3.Items(0) Then mediaID = 0 : mediaPath = Class1.romPath
        If e.ClickedItem Is myContextMenu3.Items(2) Then mediaID = 1 : mediaPath = Class1.videoPath
        If e.ClickedItem Is myContextMenu3.Items(4) Then mediaID = 2 : mediaPath = mediaPath + "\Images\Wheel\"
        If e.ClickedItem Is myContextMenu3.Items(5) Then mediaID = 3 : mediaPath = mediaPath + "\Images\Artwork1\"
        If e.ClickedItem Is myContextMenu3.Items(6) Then mediaID = 4 : mediaPath = mediaPath + "\Images\Artwork2\"
        If e.ClickedItem Is myContextMenu3.Items(7) Then mediaID = 5 : mediaPath = mediaPath + "\Images\Artwork3\"
        If e.ClickedItem Is myContextMenu3.Items(8) Then mediaID = 6 : mediaPath = mediaPath + "\Images\Artwork4\"
        If e.ClickedItem Is myContextMenu3.Items(9) Then mediaID = 7 : mediaPath = mediaPath + "\Themes\"
        If e.ClickedItem Is myContextMenu3.Items(10) Then mediaID = 8 : mediaPath = mediaPath + "\Sound\Background Music\"

        'Exclusion list
        Dim exclusionList As New List(Of String)
        If mediaID = 7 Then exclusionList.Add("default")
        If mediaID = 0 Then
            If Microsoft.VisualBasic.FileIO.FileSystem.FileExists(".\Exclusions.txt") Then
                Dim s As String = ""
                Dim go As Boolean = False
                FileOpen(1, ".\Exclusions.txt", OpenMode.Input, OpenAccess.Read)
                Do While Not EOF(1)
                    s = LineInput(1)
                    If go Then
                        exclusionList.Add(s.ToUpper)
                    End If
                    If s.Trim.ToUpper = "[" + ComboBox1.SelectedItem.ToString.ToUpper + "]" Then go = True
                Loop
                FileClose(1)
            End If
        End If


        'Move Files
        Dim romWoExt As String = ""
        For Each _mediaPath In mediaPath.Split({"|"c}, StringSplitOptions.RemoveEmptyEntries)
            If Not FileSystem.DirectoryExists(_mediaPath) Then MsgBox("Path not exist: " + vbCrLf + _mediaPath) : Continue For
            For Each rom As String In Microsoft.VisualBasic.FileIO.FileSystem.GetFiles(_mediaPath)
                Class1.Log("Checking file: " + rom)
                rom = rom.Substring(rom.LastIndexOf("\") + 1)
                romWoExt = rom.Substring(0, rom.LastIndexOf("."))
                If exclusionList.Contains(romWoExt.ToUpper) Then Class1.Log(rom + " is found on exclusion list. Skipped") : Continue For
                If Not Class1.romlist.Contains(romWoExt.ToLower) Then
                    Microsoft.VisualBasic.FileIO.FileSystem.MoveFile(_mediaPath + "\" + rom, _mediaPath + "\Unneeded\" + rom)
                End If
            Next
        Next
        If mediaID <> 0 Then MsgBox("Done.") : Exit Sub

        'Move Directories (roms only)
        Dim wildcards() As String
        If TextBox3.Text = "" Then ReDim wildcards(0) : wildcards(0) = "*.*" Else wildcards = TextBox3.Text.Split(","c)
        For i = 0 To wildcards.Length - 1
            wildcards(i) = "*." + wildcards(i).Trim
        Next

        For Each _mediaPath In mediaPath.Split({"|"c}, StringSplitOptions.RemoveEmptyEntries)
            For Each dir As String In Microsoft.VisualBasic.FileIO.FileSystem.GetDirectories(_mediaPath)
                Class1.Log("Checking folder: " + dir)
                dir = dir.Substring(dir.LastIndexOf("\") + 1)
                If dir = "Unneeded" Then Continue For
                If exclusionList.Contains(dir.ToUpper) Then Class1.Log(dir + " folder found on exclusion list. Skipped") : Continue For
                If Not Class1.romlist.Contains(dir.ToLower) Then
                    Microsoft.VisualBasic.FileIO.FileSystem.MoveDirectory(_mediaPath + "\" + dir, _mediaPath + "\Unneeded\" + dir)
                Else
                    If Microsoft.VisualBasic.FileIO.FileSystem.GetFiles(_mediaPath + "\" + dir, FileIO.SearchOption.SearchTopLevelOnly, wildcards).Count = 0 Then
                        Microsoft.VisualBasic.FileIO.FileSystem.MoveDirectory(_mediaPath + "\" + dir, _mediaPath + "\Unneeded\" + dir)
                    End If
                End If
            Next
        Next
        MsgBox("Done.")
    End Sub
#End Region

#Region "System Settings"
    'HS path changing
    Private Sub TextBox14_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox14.TextChanged
        Class1.data.Clear()
        Class1.romlist.Clear()
        Class1.romFoundlist.Clear()
        ComboBox1.Items.Clear()
        DataGridView1.Rows.Clear()
        If FileSystem.DirectoryExists(TextBox14.Text) Then
            Class1.HyperspinPath = TextBox14.Text
            If Not Class1.HyperspinPath.EndsWith("\") Then Class1.HyperspinPath = Class1.HyperspinPath + "\"

            If Not Microsoft.VisualBasic.FileIO.FileSystem.FileExists(Class1.HyperspinPath + "\Databases\Main Menu\Main Menu.xml") Then
                Label23.BackColor = Color.Orange
                Label23.Text = "Can't find \Databases\Main Menu\Main Menu.xml." : Exit Sub
            Else
                Label23.BackColor = Color.LightGreen
                Label23.Text = "\Databases\Main Menu\Main Menu.xml FOUND"
            End If

            Dim x As New Xml.XmlDocument
            x.Load(Class1.HyperspinPath + "\Databases\Main Menu\Main Menu.xml")
            For Each node As Xml.XmlNode In x.SelectNodes("/menu/game")
                ComboBox1.Items.Add(node.Attributes.GetNamedItem("name").Value)
            Next

            'Get hyperlaunch path from settings.ini
            If FileSystem.FileExists(Class1.HyperspinPath + "\Settings\Settings.ini") Then
                ini.path = Class1.HyperspinPath + "\Settings\Settings.ini"
                TextBox18.Text = ini.IniReadValue("main", "Hyperlaunch_Path")
            End If
        Else
            Label23.BackColor = Color.Orange
            Label23.Text = "Directory not exist." : Exit Sub
        End If
    End Sub

    'Save Config
    Private Sub Button20_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click 'Save Config
        ini.IniFile(Class1.confPath)
        ini.IniWriteValue("MAIN", "HyperSpin_Path", TextBox14.Text)
        ini.IniWriteValue("MAIN", "rename_just_cue", DirectCast(IIf(CheckBox6.Checked, "1", "0"), String))
        ini.IniWriteValue("MAIN", "search_cue_for", TextBox16.Text)
        ini.IniWriteValue("MAIN", "useHLv3", DirectCast(IIf(CheckBox26.Checked, "1", "0"), String))
        For i As Integer = 1 To ListBox4.Items.Count
            ini.IniWriteValue("handle_together", i.ToString, ListBox4.Items(i - 1).ToString)
        Next

        'FileOpen(1, Application.StartupPath + "\Config.conf", OpenMode.Output)
        'PrintLine(1, TextBox14.Text)
        'PrintLine(1, "rename_just_cue = " + DirectCast(IIf(CheckBox6.Checked, "1", "0"), String))
        'PrintLine(1, "search_cue_for = " + TextBox16.Text)
        'PrintLine(1, "useHLv3 = " + DirectCast(IIf(CheckBox26.Checked, "1", "0"), String))
        'PrintLine(1, "[handle_together]")
        'For Each h In ListBox4.Items
        'PrintLine(1, h)
        'Next
        'FileClose(1)
        Label23.BackColor = Color.LightBlue
        Label23.Text = "Config.conf Saved"
    End Sub

    'add extension to PAIR_EXTENSION list
    Private Sub Button5_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If TextBox15.Text.Contains(",") Then
            ListBox4.Items.Add(TextBox15.Text)
        Else
            MsgBox("extensions must be separated by "",""")
        End If
    End Sub

    'remove from PAIR_EXTENSION list
    Private Sub Button27_Click(sender As System.Object, e As System.EventArgs) Handles Button27.Click
        If ListBox4.SelectedIndex >= 0 Then
            ListBox4.Items.RemoveAt(ListBox4.SelectedIndex)
        End If
    End Sub

    'Video path switch - hs\media\system\video or from ini
    Private Sub RadioButton10_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton10.CheckedChanged, RadioButton11.CheckedChanged
        If ComboBox1.SelectedIndex < 0 Then Exit Sub
        If RadioButton10.Checked Then
            Class1.videoPath = Class1.videoPathOrig
            If Not FileSystem.DirectoryExists(Class1.videoPath) Then Class1.videoPath = Class1.HyperspinPath + "Media\" + ComboBox1.SelectedItem.ToString + "\Video\"
        End If
        If RadioButton11.Checked Then
            Class1.videoPath = Class1.HyperspinPath + "Media\" + ComboBox1.SelectedItem.ToString + "\Video\"
        End If
    End Sub

    'videoPath change
    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged 'change video path
        If FileSystem.DirectoryExists(TextBox2.Text) And RadioButton10.Checked Then
            Class1.videoPath = TextBox2.Text
        Else
            Class1.videoPath = Class1.videoPathOrig
        End If
    End Sub

    'Inline Editor enable/disable OLD code
    Private Sub CheckBox17_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Exit Sub
        'DataGridView1.ReadOnly = Not CheckBox17.Checked
        'If CheckBox17.Checked Then
        'Button21.Text = "UPDATE HyperSpin Database"
        'Else
        'Button21.Text = "Delete selected rom from database"
        'End If
        'For i As Integer = 1 To 10
        'DataGridView1.Columns(i).ReadOnly = True
        'Next
    End Sub

    'Change checkbox 'use HS settings for clones/parents' 
    Private Sub CheckBox22_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox22.CheckedChanged
        If CheckBox22.Checked Then
            CheckBox23.Enabled = False : CheckBox24.Enabled = False
            CheckBox23.Checked = useParentVids : CheckBox24.Checked = useParentThemes
        Else
            CheckBox23.Enabled = True : CheckBox24.Enabled = True
        End If
    End Sub

    'Hyperlaunch path changed
    Private Sub TextBox18_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox18.TextChanged
        Dim ok As Boolean = False
        If FileSystem.DirectoryExists(TextBox18.Text) Then
            If FileSystem.FileExists(TextBox18.Text + "\HyperLaunch.exe") Then
                ok = True
            End If
        End If
        If ok Then TextBox18.BackColor = colorYES Else TextBox18.BackColor = colorNO
    End Sub

    'Update Hyperspin INI
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If Not TextBox18.Text.EndsWith("\") Then TextBox18.Text = TextBox18.Text + "\"
        TextBox18.Text = TextBox18.Text.Replace("\\", "\").Replace("\\", "\")
        ini.path = Class1.HyperspinPath + "\Settings\Settings.ini"
        ini.IniWriteValue("main", "Hyperlaunch_Path", TextBox18.Text)
    End Sub
#End Region

#Region "Menus"
    'File / Exit
    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Application.Exit()
    End Sub

    'Table / Allow Edit
    Private Sub AlowEditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AlowEditToolStripMenuItem.Click
        AlowEditToolStripMenuItem.Checked = Not AlowEditToolStripMenuItem.Checked
        DataGridView1.ReadOnly = Not AlowEditToolStripMenuItem.Checked
        If AlowEditToolStripMenuItem.Checked And Button2_moveUnneeded.Enabled Then
            Button21.Enabled = True
            CommitDbEditionsToolStripMenuItem.Enabled = True
        Else
            Button21.Enabled = False
            CommitDbEditionsToolStripMenuItem.Enabled = False
        End If

        For i As Integer = 2 To 10
            DataGridView1.Columns(i).ReadOnly = True
        Next
    End Sub

    'Table / Reorder / No Reorder
    Private Sub NoReorderinsertedLinesAddedToTheEndToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NoReorderinsertedLinesAddedToTheEndToolStripMenuItem.Click
        NoReorderinsertedLinesAddedToTheEndToolStripMenuItem.Checked = True
        ReorderAlphabetycallyToolStripMenuItem.Checked = False
        ReorderAlphabeticallyByDescriptionToolStripMenuItem.Checked = False
        ReorderAsSeenInTheCheckTableToolStripMenuItem.Checked = False
    End Sub
    'Table / Reorder / Reorder alphabetically by romname
    Private Sub ReorderAlphabetycallyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReorderAlphabetycallyToolStripMenuItem.Click
        NoReorderinsertedLinesAddedToTheEndToolStripMenuItem.Checked = False
        ReorderAlphabetycallyToolStripMenuItem.Checked = True
        ReorderAlphabeticallyByDescriptionToolStripMenuItem.Checked = False
        ReorderAsSeenInTheCheckTableToolStripMenuItem.Checked = False
    End Sub
    'Table / Reorder / Reorder alphabetically by description
    Private Sub ReorderAlphabeticallyByDescriptionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReorderAlphabeticallyByDescriptionToolStripMenuItem.Click
        NoReorderinsertedLinesAddedToTheEndToolStripMenuItem.Checked = False
        ReorderAlphabetycallyToolStripMenuItem.Checked = False
        ReorderAlphabeticallyByDescriptionToolStripMenuItem.Checked = True
        ReorderAsSeenInTheCheckTableToolStripMenuItem.Checked = False
    End Sub
    'Table / Reorder / Reorder as seen in table
    Private Sub ReorderAsSeenInTheCheckTableToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReorderAsSeenInTheCheckTableToolStripMenuItem.Click
        NoReorderinsertedLinesAddedToTheEndToolStripMenuItem.Checked = False
        ReorderAlphabetycallyToolStripMenuItem.Checked = False
        ReorderAlphabeticallyByDescriptionToolStripMenuItem.Checked = False
        ReorderAsSeenInTheCheckTableToolStripMenuItem.Checked = True
    End Sub

    'Table / ShowHide columns
    Private Sub FilterColumns_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles F0.Click, F1.Click, F2.Click, F3.Click, F4.Click, F5.Click, F6.Click, F7.Click, F8.Click, F9.Click, F10.Click, F11.Click, F12.Click, F13.Click, F14.Click
        Dim tsmi As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
        If tsmi.Checked Then tsmi.Checked = False Else tsmi.Checked = True
        Dim i As Integer = CInt(tsmi.Name.Substring(1))

        For Each c In myContextMenu4.Items
            Dim ch As ToolStripControlHost = TryCast(c, ToolStripControlHost)
            If ch IsNot Nothing Then
                If CInt(DirectCast(ch.Control, CheckBox).Name) = i Then
                    DirectCast(ch.Control, CheckBox).Checked = tsmi.Checked : Exit For
                End If
            End If
        Next
    End Sub
    'Table / ShowHide columns / Preset Checker
    Private Sub FPresetChecker_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FPresetChecker.Click
        ShowHidePresets(sender, New ToolStripItemClickedEventArgs(myContextMenu4.Items(myContextMenu4.Items.Count - 2)))
    End Sub
    'Table / ShowHide columns / Preset Editor
    Private Sub FPresetEditor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FPresetEditor.Click
        ShowHidePresets(sender, New ToolStripItemClickedEventArgs(myContextMenu4.Items(myContextMenu4.Items.Count - 1)))
    End Sub

    'Table / Update (commit) DB --- handled in Class2_xml

    'Table / Excel csv export
    Private Sub ExportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExportToolStripMenuItem.Click
        If DataGridView1.Rows.Count = 0 Then MsgBox("Table is empty. Nothing to export.") : Exit Sub

        Dim sep As String = ""
        Dim rep As String = ""
        If SemicolonExcelDefaultToolStripMenuItem.Checked Then sep = "; " : rep = "," Else sep = ", " : rep = ";"

        Dim fd As New SaveFileDialog
        fd.AddExtension = True
        fd.CheckPathExists = True
        fd.FileName = "Hyperspin Check.csv"
        'fd.Filter = "Comma Separated values(*.csv)|*.csv|All files (*.*)|*.*"
        fd.Filter = "Comma Separated values(*.csv)|*.csv"
        fd.ShowDialog()

        Dim s As String = ""
        FileOpen(1, fd.FileName, OpenMode.Output, OpenAccess.Write)
        s = ""
        For Each col As DataGridViewColumn In DataGridView1.Columns
            s = s + col.HeaderText.Replace(sep.Trim, rep) + sep
        Next
        s = s.Substring(0, s.Length - 2)
        PrintLine(1, s)

        For Each r As DataGridViewRow In DataGridView1.Rows
            s = ""
            For Each col As DataGridViewColumn In DataGridView1.Columns
                s = s + r.Cells(col.Index).Value.ToString.Replace(sep.Trim, rep) + sep
            Next
            s = s.Substring(0, s.Length - 2)
            PrintLine(1, s)
        Next
        FileClose(1)
    End Sub

    'Table / Excel export separator change
    Private Sub SemicolonExcelDefaultToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SemicolonExcelDefaultToolStripMenuItem.Click, CommaFormatDefaultToolStripMenuItem.Click
        Dim m As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
        If m.Name.ToLower.StartsWith("semicolon") Then
            CommaFormatDefaultToolStripMenuItem.Checked = False
            SemicolonExcelDefaultToolStripMenuItem.Checked = True
        Else
            CommaFormatDefaultToolStripMenuItem.Checked = True
            SemicolonExcelDefaultToolStripMenuItem.Checked = False
        End If
    End Sub

    'Matcher / Autorenamer
    Private Sub AutorenamerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AutorenamerToolStripMenuItem.Click
        If DataGridView1.Rows.Count = 0 Then MsgBox("Please, make ""check"" in summary page, before using this future.") : Exit Sub
        If ComboBox3.SelectedIndex < 0 Then MsgBox("Select media you want autorename (roms/video/artwork/wheels/themes...).") : Exit Sub
        RadioButton2.Checked = True
        RadioButton5.Checked = True
        If ListBox1.Items.Count = 1 Then
            If ListBox1.Items(0).ToString.StartsWith("No missing") Then MsgBox(ListBox1.Items(0).ToString + " found. Nothing to do.") : Exit Sub
        End If
        If ListBox2.Items.Count = 0 Then MsgBox("There is no unmatched media. Nothing to do.") : Exit Sub
        Form6_autorenamer.Show()
    End Sub

    'Matcher / Associate option in HS folder click
    Private Sub AssocOption_fileInHsFolder_copy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AssocOption_fileInHsFolder_copy.Click, AssocOption_fileInHsFolder_move.Click
        Dim i As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
        AssocOption_fileInHsFolder_copy.Checked = False
        AssocOption_fileInHsFolder_move.Checked = False
        i.Checked = True
    End Sub

    'Macher / Associate option in different folder click
    Private Sub AssocOption_fileInDiffFolder_copy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AssocOption_fileInDiffFolder_copy.Click, AssocOption_fileInDiffFolder_copyToHS.Click, AssocOption_fileInDiffFolder_move.Click, AssocOption_fileInDiffFolder_moveToHS.Click
        Dim i As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
        AssocOption_fileInDiffFolder_copy.Checked = False
        AssocOption_fileInDiffFolder_copyToHS.Checked = False
        AssocOption_fileInDiffFolder_move.Checked = False
        AssocOption_fileInDiffFolder_moveToHS.Checked = False
        i.Checked = True
    End Sub

    'Tools / Show HL 3rd party paths
    Private Sub CheckHyperLaunch3rdPartyPathsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckHyperLaunch3rdPartyPathsToolStripMenuItem.Click
        If TextBox18.BackColor <> colorYES Then MsgBox("You Hyperlaunch path is not correctly set. Check 'Hyperspin system settings' tab") : Exit Sub
        Dim f As New FormA_hyperlaunch_3rd_party_paths : f.Show()
    End Sub

    'Tools / Show Genres/favorites manager
    Private Sub GenresFavoritesManagerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GenresFavoritesManagerToolStripMenuItem.Click
        If ComboBox1.SelectedIndex < 0 Then MsgBox("Please, select a system.") : Exit Sub
        Form4_genres_favorites.Show()
    End Sub

    'Tools / Show DB statistic
    Private Sub ShowDatabaseStatisticToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ShowDatabaseStatisticToolStripMenuItem.Click
        Form9_database_statistic.Show()
    End Sub

    'Tools / DIFF Options 1
    Private Sub CompareAgainstCurrentDBToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CompareAgainstCurrentDBToolStripMenuItem.Click, ChooseBothFilesToCompareToolStripMenuItem.Click
        CompareAgainstCurrentDBToolStripMenuItem.Checked = False
        ChooseBothFilesToCompareToolStripMenuItem.Checked = False
        DirectCast(sender, ToolStripMenuItem).Checked = True
    End Sub
    'Tools / DIFF Options 2
    Private Sub CompareGameromNameToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CompareGameromNameToolStripMenuItem.Click, CompareDescriptionTagusualyFullNameToolStripMenuItem.Click
        CompareGameromNameToolStripMenuItem.Checked = False
        CompareDescriptionTagusualyFullNameToolStripMenuItem.Checked = False
        DirectCast(sender, ToolStripMenuItem).Checked = True
    End Sub

    'Tools / Dual folder ops
    Private Sub DualFolderOperationsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DualFolderOperationsToolStripMenuItem.Click
        Form7_dualFolderOperations.Show()
    End Sub

    'Tools / Show video downloader
    Private Sub VideoDownloaderToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VideoDownloaderToolStripMenuItem.Click
        Form5_videoDownloader.Show()
    End Sub

    'Tools / PCSX2 Create indexes
    Private Sub PCSX2CreateIndexFilesForCompressedIsoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PCSX2CreateIndexFilesForCompressedIsoToolStripMenuItem.Click
        Dim f As New FormB_PCSX2_createIndex
        f.Show()
    End Sub

    'Tools / Mame romset reducer
    Private Sub MAMERomsetReducerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MAMERomsetReducerToolStripMenuItem.Click
        Dim f As New FormC_mameRomListBuilder
        f.Show(Me)
    End Sub
#End Region

#Region "Table headers and painting"
    'Table Filter
    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        Dim counter1 As Integer = 0
        Dim counter2 As Integer = 0
        Button18.Text = "Crop XML to only found games"
        'SendMessage(DataGridView1.Parent.Handle, 11, False, 0)

        'MAME FOLDERS
        If ComboBox4.SelectedIndex = 9 Then
            If DataGridView1.RowCount = 0 Then MsgBox("Please, fill the grid by using 'CHECK' button before applying filter.") : ComboBox4.SelectedIndex = oldCombo4Value : Exit Sub
            Dim f As New OpenFileDialog : f.Title = "Open mame folder (.ini) or .txt with filelist"
            f.Filter = "ini Files (*.ini)|*.ini|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
            f.ShowDialog()
            If f.FileName = "" Then ComboBox4.SelectedIndex = oldCombo4Value : Exit Sub
            If Not FileSystem.FileExists(f.FileName) Then MsgBox("Can't open file") : ComboBox4.SelectedIndex = oldCombo4Value : Exit Sub
            Class1.askVar1 = f.FileName
            Form3.ShowDialog()
            If Form3.filter.Count = 0 Then MsgBox("No games selected") : ComboBox4.SelectedIndex = oldCombo4Value : Exit Sub

            oldCombo4Value = ComboBox4.SelectedIndex
            ProgressBar1.Value = 0
            ProgressBar1.Maximum = DataGridView1.Rows.Count
            Label2.Text = "Filtering..." : Label2.Refresh()
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                DataGridView1.Rows(i).Visible = True
                If Form3.filter.Contains(DataGridView1.Rows(i).Cells(1).Value.ToString.ToLower) Then
                    counter1 = counter1 + 1
                    DataGridView1.Rows(i).Visible = True
                Else
                    counter2 = counter2 + 1
                    DataGridView1.Rows(i).Visible = False
                End If
                ProgressBar1.Value = i : If ProgressBar1.Value Mod 10 = 0 Then ProgressBar1.Refresh()
            Next
            Button18.Text = "Crop XML to only filtered games"
            Label2.Text = "Showing: " + counter1.ToString + " of total " + DataGridView1.Rows.Count.ToString
            Exit Sub
        End If
        oldCombo4Value = ComboBox4.SelectedIndex

        'Regular filter
        ProgressBar1.Value = 0
        ProgressBar1.Maximum = DataGridView1.Rows.Count
        Label2.Text = "Filtering..." : Label2.Refresh()
        Dim col As Integer = ComboBox4.SelectedIndex + 1
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If col = 1 Then
                DataGridView1.Rows(i).Visible = True
            Else
                If DataGridView1.Rows(i).Cells(col).Value.ToString = "NO" Then
                    counter1 = counter1 + 1
                    DataGridView1.Rows(i).Visible = True
                Else
                    counter2 = counter2 + 1
                    DataGridView1.Rows(i).Visible = False
                End If
            End If
            ProgressBar1.Value = i : If ProgressBar1.Value Mod 10 = 0 Then ProgressBar1.Refresh()
        Next

        ProgressBar1.Value = 0
        If col = 1 Then
            Label2.Text = "Total: " + DataGridView1.Rows.Count.ToString
        Else
            Label2.Text = "Missing: " + counter1.ToString + " of " + DataGridView1.Rows.Count.ToString
        End If
        'SendMessage(DataGridView1.Parent.Handle, 11, True, 0)
    End Sub

    'Column Headers
    Private Sub DataGridView1_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.ColumnHeaderMouseClick
        If e.Button <> Windows.Forms.MouseButtons.Right Then Exit Sub
        myContextMenu4.Show(Cursor.Position.X, Cursor.Position.Y)
    End Sub

    'Show/hide columns checked_change
    Private Sub ShowHideCol(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckStrip2_0.CheckedChanged, CheckStrip2_1.CheckedChanged, CheckStrip2_2.CheckedChanged, CheckStrip2_3.CheckedChanged, CheckStrip2_4.CheckedChanged, CheckStrip2_5.CheckedChanged, CheckStrip2_6.CheckedChanged, CheckStrip2_7.CheckedChanged, CheckStrip2_8.CheckedChanged, CheckStrip2_9.CheckedChanged, CheckStrip2_10.CheckedChanged, CheckStrip2_11.CheckedChanged, CheckStrip2_12.CheckedChanged, CheckStrip2_13.CheckedChanged, CheckStrip2_14.CheckedChanged
        Dim cb As CheckBox = DirectCast(sender, CheckBox)
        Dim i As Integer = CInt(cb.Name)
        DirectCast(ShowHideColumnsToolStripMenuItem.DropDownItems("F" + i.ToString), ToolStripMenuItem).Checked = cb.Checked

        DataGridView1.Columns(i).Visible = cb.Checked
    End Sub

    'Column headers presets
    Private Sub ShowHidePresets(ByVal sender As Object, ByVal e As ToolStripItemClickedEventArgs) Handles myContextMenu4.ItemClicked
        If e.ClickedItem.Text = Preset_Checker Then
            CheckStrip2_2.Checked = True
            CheckStrip2_3.Checked = True
            CheckStrip2_4.Checked = True
            CheckStrip2_5.Checked = True
            CheckStrip2_6.Checked = True
            CheckStrip2_7.Checked = True
            CheckStrip2_8.Checked = True
            CheckStrip2_9.Checked = True
            CheckStrip2_10.Checked = True
            CheckStrip2_11.Checked = False
            CheckStrip2_12.Checked = False
            CheckStrip2_13.Checked = False
            CheckStrip2_14.Checked = False
        End If
        If e.ClickedItem.Text = Preset_Editor Then
            CheckStrip2_2.Checked = False
            CheckStrip2_3.Checked = False
            CheckStrip2_4.Checked = False
            CheckStrip2_5.Checked = False
            CheckStrip2_6.Checked = False
            CheckStrip2_7.Checked = False
            CheckStrip2_8.Checked = False
            CheckStrip2_9.Checked = False
            CheckStrip2_10.Checked = False
            CheckStrip2_11.Checked = True
            CheckStrip2_12.Checked = True
            CheckStrip2_13.Checked = True
            CheckStrip2_14.Checked = True
        End If
        If e.ClickedItem.Text = Save_current_cols_conf_as_startup Then
            ini.IniFile(Class1.confPath)
            Dim v As String = ""
            For i As Integer = 0 To DataGridView1.ColumnCount - 1
                If DataGridView1.Columns(i).Visible Then v = "1" Else v = "0"
                ini.IniWriteValue("Main_Table_Columns_Config", "Col_" + i.ToString + "_visible", v)
                ini.IniWriteValue("Main_Table_Columns_Config", "Col_" + i.ToString + "_width", DataGridView1.Columns(i).Width.ToString)
            Next
        End If
    End Sub

    'CELL PAINTING to paint selected cell withing 'row selection mode'
    Private Sub DataGridView1_CellPainting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs) Handles DataGridView1.CellPainting
        If DataGridView1.CurrentCell IsNot Nothing Then
            If AlowEditToolStripMenuItem.Checked And (DataGridView1.CurrentCell.ColumnIndex = e.ColumnIndex And DataGridView1.CurrentCell.RowIndex = e.RowIndex) Then
                'e.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit
                Dim gridLinePen As New Pen(Me.DataGridView1.GridColor)
                Dim newRect As New Rectangle(e.CellBounds.X + 1, e.CellBounds.Y + 1, e.CellBounds.Width - 4, e.CellBounds.Height - 4)

                If e.ColumnIndex = 0 Or e.ColumnIndex > 10 Then
                    e.Graphics.FillRectangle(New SolidBrush(e.CellStyle.SelectionBackColor), e.CellBounds)
                    e.Graphics.DrawRectangle(Pens.Yellow, newRect)
                    e.Graphics.DrawLine(New Pen(Brushes.LightYellow), e.CellBounds.Left, e.CellBounds.Bottom - 1, e.CellBounds.Right - 1, e.CellBounds.Bottom - 1)
                    e.Graphics.DrawRectangle(Pens.LightYellow, e.CellBounds.X, e.CellBounds.Y, e.CellBounds.Width - 2, e.CellBounds.Height - 2)
                    If (e.Value IsNot Nothing) Then
                        e.Graphics.DrawString(CStr(e.Value), e.CellStyle.Font, New SolidBrush(e.CellStyle.SelectionForeColor), e.CellBounds.X + 1, e.CellBounds.Y + 4, StringFormat.GenericDefault)
                    End If
                ElseIf e.ColumnIndex = 1 Then
                    e.Graphics.FillRectangle(New SolidBrush(e.CellStyle.SelectionBackColor), e.CellBounds)
                    e.Graphics.DrawRectangle(Pens.Red, newRect)
                    e.Graphics.DrawLine(New Pen(Brushes.Red), e.CellBounds.Left, e.CellBounds.Bottom - 1, e.CellBounds.Right - 1, e.CellBounds.Bottom - 1)
                    e.Graphics.DrawRectangle(Pens.Red, e.CellBounds.X, e.CellBounds.Y, e.CellBounds.Width - 2, e.CellBounds.Height - 2)
                    If (e.Value IsNot Nothing) Then
                        e.Graphics.DrawString(CStr(e.Value), e.CellStyle.Font, New SolidBrush(e.CellStyle.SelectionForeColor), e.CellBounds.X + 1, e.CellBounds.Y + 4, StringFormat.GenericDefault)
                    End If
                Else
                    e.Graphics.FillRectangle(New SolidBrush(e.CellStyle.SelectionBackColor), e.CellBounds)
                    e.Graphics.DrawRectangle(Pens.Yellow, newRect)
                    e.Graphics.DrawLine(gridLinePen, e.CellBounds.Left, e.CellBounds.Bottom - 1, e.CellBounds.Right - 1, e.CellBounds.Bottom - 1)
                    If (e.Value IsNot Nothing) Then
                        e.Graphics.DrawString(CStr(e.Value), e.CellStyle.Font, New SolidBrush(e.CellStyle.SelectionForeColor), e.CellBounds.X + 1, e.CellBounds.Y + 4, StringFormat.GenericDefault)
                    End If
                End If
                e.Handled = True
            End If
        End If
    End Sub
#End Region

#Region "All the '...' buttons (open file browser), Notes, Autorenamer"
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim fd As New FolderBrowserDialog : fd.Description = "Select folder" : fd.ShowDialog()
        TextBox4.Text = fd.SelectedPath
    End Sub '...
    Private Sub Button36_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button36.Click
        Dim fd As New FolderBrowserDialog : fd.Description = "Select hyperlaunch folder" : fd.ShowDialog()
        TextBox18.Text = fd.SelectedPath
    End Sub '...
    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Dim fd As New FolderBrowserDialog : fd.Description = "Select rom folder" : fd.ShowDialog()
        TextBox1.Text = fd.SelectedPath
    End Sub '...
    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Dim fd As New FolderBrowserDialog : fd.Description = "Select Video folder" : fd.ShowDialog()
        TextBox2.Text = fd.SelectedPath
    End Sub '...
    Private Sub Button2_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim fd As New FolderBrowserDialog : fd.Description = "Select hyperspin root folder (where hyperspin.exe is)" : fd.ShowDialog()
        TextBox14.Text = fd.SelectedPath
    End Sub '...

    'Set HL folder to \Hyperspin\Hyperlaunch
    Private Sub Button37_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button37.Click
        TextBox18.Text = (Class1.HyperspinPath + "\Hyperlaunch\").Replace("\\", "\").Replace("\\", "\")
    End Sub

    'Open Notes
    Private Sub Button30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button30.Click
        System.Diagnostics.Process.Start("notepad.exe", ".\notes.txt")
    End Sub
    'Show autorenamer
    'Private Sub contextMenuAutorenamer(ByVal sender As Object, ByVal e As ToolStripItemClickedEventArgs) Handles myContextMenu7.ItemClicked
    'If DataGridView1.Rows.Count = 0 Then MsgBox("Please, make ""check"" in summary page, before using this future.") : Exit Sub
    'If ComboBox3.SelectedIndex < 0 Then MsgBox("Select media you want autorename (roms/video/artwork/wheels/themes...).") : Exit Sub
    'RadioButton2.Checked = True
    'RadioButton5.Checked = True
    'If ListBox1.Items.Count = 1 Then
    'If ListBox1.Items(0).ToString.StartsWith("No missing") Then MsgBox(ListBox1.Items(0).ToString + " found. Nothing to do.") : Exit Sub
    'End If
    'If ListBox2.Items.Count = 0 Then MsgBox("There is no unmatched media. Nothing to do.") : Exit Sub
    'Form6_autorenamer.Show()
    'End Sub

    'Show Database statistic
#End Region
End Class