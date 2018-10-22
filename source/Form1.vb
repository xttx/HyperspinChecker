Imports WindowsApplication1.Language
Imports Microsoft.VisualBasic.FileIO
Imports System.Runtime.InteropServices
Imports System.ComponentModel

Public Class Form1
    'TODO Add 'Reset columns size' to both tables context menu
    'TODO Use file extentions from hyperlaunch when check/list roms in matcher
    'TODO After autorenaming something, and then reopen autorenamer and collect info, crash on Form6_autorenamer.vb-line145 cause file not found
    'REQUEST [desrop69 #105] case sensitive crc
    'REQUEST [Acidnine #108] When I check the master snes .xml db file it says there's an error loading it. When I tell it to fix it, it replaces all the ampersands '&' with '&amp;' in the <game name=""> tag. Is this wanted? because now I have to rename my roms with ampersands in them to match the name tag.
    'REQUEST [tastyratz #112] I encountered a bug when trying to create a cue file for a saturn bin which does not have one. It gave me the error, created a 0kb file, and left the file locked but unfinished. Any subsequent attempts show the file in use error below. This error mentions the F:\ path but I do not have an F:\ drive, maybe related?
    'REQUEST [tastyratz #112] I want to be able to treat iso, cue like bin, cue because all of my sega cd collection is iso & cue. Would that be the right field (handle together) for entry?
    'REQUEST [tastyratz #112] Finally another bug: When matching an iso/cue set in subfolder it tells me it cannot find the original iso filename - however it does rename and edit the iso/cue just fine (just not the folder). If I click it a second time it renames the folder without errors.
    'REQUEST [brownvim #118] hide clones in a MAME list. I want to check if im missing any unique roms but so many are red (clones).
    'REQUEST [Andyman #121] Is there a way to show which games are clones in MAME? That would be a tremendous help. 
    'REQUEST [baezl #128] How do you guys handle multiple disc games (Sega CD)? What is your structure so we can use this kick ass tool which I freaking love!!!! I noticed the current database has Disc1 and Disc2 separate? On some games, (Sega CD again), I get "Length Less Than Zero" or "File already open".
    'REQUEST [rpdmatt #129] After choosing the system I can try to check for roms in the Summary page and it will populate everything from my database xml file but won't pick up on the roms that I have. I have tried going to the HyperSpin system settings and changing the 'Rom Path:' but it still doesn't affect the outcome. Also, if I change the 'Video Path:' to something else I cannot get the outcome to change. It seems like the user settings on the HyperSpin system settings tab don't affect the actual check on the Summary tab. Not sure if this is just the case since the new HyperLaunch HQ or not. Has anyone been able to change the 'Rom Path:' and have it affect the check? 
    'REQUEST [potts43 #137] Hyperpause media...? Controller Manual Guides Fades etc... It would be great to have columns for all of these folders 
    'REQUEST [zero dreams #146] When the check button is selected, it does a check and returns to the top of the list at the beginning of the database. I would love it if there was an option to turn this off. I'm basically looking for something like how HLHQ audits games and the list remains static during and after the audit. 
    'REQUEST [Pike #147, Akelarre #149] is it possible to have another button ("duplicate" or whatever you want), which is keeping the original art file and make a copy of it with the new name ? I'm asking this because when you have localized games not from original HS XML database but added to it, all art files already exist with the original rom, and we just need to copy-rename these files.
    'REQUEST [potts43 #151] multibin games related
    'REQUEST [tanin #152] error
    'REQUEST [potts43 #154] If possible have two way renaming. At the moment you can rename roms and artwork from the xml but for systems like future pinball and WinUAEloader that change regularly it would be good to point the DB xml to an updated rom and update the xml to match that newer rom name.
    'Rename media when edit romname in main table

    'TODO option to rename audio files associated with .cue (ME)
    'TODO It appears to be setting the "video" folder for Visual Pinball to the previous system a "Check" was ran on. In the Matcher tab, even when I tell it to use a custom folder and point it to the correct Visual Pinball video folder, once I perform the "Associate" action, it's still copying the video files out of my Visual Pinball video folder and moving them to the Gamecube video folder (note: Gamecube was the last system I had ran a "Check" on). So, the Matcher seems to be using the custom folder fine to actually find the video files, but when it performs an "Associate", something is getting confused. Not a big deal, but thought I'd pass the info along. I've tried using both the System.ini and default Hyperspin Video location settings and happens for both. Let me know if you have any questions on my setup & thanks for the great tool.  (Signet145 #47)
    'TODO undo the renaming INSIDE the .cue. The .cue and .bin files both get renamed back to the correct original file name, but inside the cue file does not. (windowlicker #27)
    'TODO undo message box displays a couple of extra backslashes, but it does not affect anything. (windowlicker #27)
    'TODO Restore .cues from backup
    'TODO in button20(markAsFound) and TextBox4(matcher filelist) WE ALSO NEED TO CHECK IF A FILE WITH ROMNAME AND APPROPRIATE EXTENSION EXISTS IN SUBDIR
    'TODO check and improve collect media dialog
    'TODO check/verify all possibilities while associating in matcher (copy, move, in/from different folders, and check listBoxex behaviour)
    'TODO when move roms in subfolder in NOT subfoldered mode, filelist don't reflect changes

    'TODO - 'handle-together' in 'move unneeded to subfolder'

    'TODO - RL fade preview and editor
    'TODO - move mode when in another directory, in subfoldered mode is missing (code is missing)
    'TODO - copy (duplicate) mode in the same folder, just rename file in filelist instead of duplicating
    'TODO - renaming history to file
    'TODO - better renaming log
    'TODO - drag & drop in system manager
    'TODO - while in edit mode in the cell, click Or right click on another cell - error - can't reproduce
    'TODO - create db from files/folders - disable fill crc by default if folder is selected
    'TODO - create db from files/folders - automatically suggest output name, based on output folder And output filename
    'TODO - when renaming in another folder (copy or move), and only show unmatched files, files not disappear when renaming
    'TODO - Copy rom from another folder - ask to rename files in every archive
    'TODO - subfoldered mode - matched/unmatched should have option (enabled by default) to match a FILE inside a FOLDER, Not only folder name
    'TODO - move copy/move/rename/duplicate controls from options menu to main matcher window
    'TODO - Hide CHECK button (And other controls) when checking
    'TODO - deploy emulator pack and customize emu video and controls and module settings, and paths

    'TODO - Mark items in table as found, when renaming in autorenamer
    'TODO - Mark items in listboxes as found, when renaming in autorenamer
    'TODO - more parsing expressions in regex constructor regex parsing

    'TODO - Wrong matcher stats update when copying (duplicated)
    'TODO - Move checkboxes of matcher file operation options (copy/move...) to matcher main tab
    'TODO - Wheel/box creator
    'TODO - Youtube downloader not working
    'TODO - While trying to test an expression, in regex constructor - replace\((".*?"),(".*?")\) - unexpected behaviour: multiple singles was collapsing to one, after entering first closing paranthesis. (have to enter expression by type for this to happens)

    'MOSTLY DONE TODO relative HL path not working ('Hyperlaunch path changed, 'HS path changing)
    'MOSTLY DONE TODO flames #153 - print check table on printer or exel export
    'MOSTLY DONE TODO potts43 #154 - Be able to add/copy/delete xml entries. This way you could add a blank entry and complete it and then save it.
    'DONE TODO potts43 #154 - On the renamer page have a totals number always present.eg.  Unmatched/total  Matched/total  Both/total
    'DONE TODO thrasherx #119, Turranius #160 - Any chance you could filter the MAME BIOS files from the 'move unneeded roms' option? I consolidated my MAME roms and it took my BIOS files with it. 
    'DONE TODO potts43 #154 - Be able to edit both name and description in the xml. Even better to have all fields. This way you can have custom Xmls made easy.
    'DONE TODO Yardley #159 - option to move themes in the "Move unneeded roms or media to subfolder" options. Seems simple to add this feature. 
    'DONE TODO Added default icon
    'DONE TODO Ability to change HyperLaunch path in HS Settings.ini.
    'DONE TODO Ability to change HyperLaunch / HyperLaunchHQ ini to make HS easier to move from one drive/folder to another
    'DONE TODO Speed up loading time
    'DONE TODO Iso Checker now tracks wrong formated GDI (no fix though)

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
    'DONE TODO HL multiple path support (in 'move unneeded' - this is already done)
    'DONE TODO repare system manager (hl path including filename not properly handled)
    'DONE TODO add to system manager MainMenuWheel, MainMenuVideo
    'DONE TODO move dir2dat options to form
    'DONE TODO REQUEST [QRS, post #49] I just checked the last cue file and the quotations was missing in the name inside the cue file Example: Supreme Warrior (USA) (Disc 1).bin instead of "Supreme Warrior (USA) (Disc 1).bin"
    'DONE TODO REQUEST [aupton, post #63] When selecting "Move unneeded roms or media to subfolder" the app thows a JIT error. (added additional check for roms without extension)
    'DONE TODO REQUEST [MadCoder, post #85] Well, I had an idea about the fix ISO for DC...I going to add Dreamcast CDI to ISO ability, and ISO fixer. (Not skilled enough to write my own CD images converter. Added bath converter using UltraISO)
    'DONE TODO REQUEST [RandomName1, post #91] Would it be possible to strip tags from rom descriptions automatically (except disc numbers) when creating an xml from a rom folder? (added long ago, but was broken. Also added exception option for disk numbers)
    'DONE TODO REQUEST [Obiwantje #103] updateble HS ini files
    'DONE TODO After saving config (label became violet), scan says HS path not found
    'DONE TODO When loading form, HL path says "please select system", therefore, this goes into path variable ??? can't reproduce!!! I don't know how, but r.cells(0).style.backcolor = Class1.colorYES cause this
    'DONE TODO Autorenamer crc mode have problems with zips
    'DONE TODO When update emu / rom paths and/or use hl in sys properties, and reopen system properties without scan - paths not updated sometimes
    'DONE TODO When update rompath for selected system, and not using HL, rompath is not updated on main check
    'DONE TODO when switching rom path (HL / HS), and then recheck without reselect system, path is not updated
    'DONE TODO when changing custom rom path in matcher, sometimes got error (try use MAME path) (was it multipath issue?)
    'DONE TODO To see if hs path is set, system manager refers to Label23.color, which can be not necessary green but purple, if path is ok, but config.save clicked.
    'DONE TODO REQUEST [MadCoder] Dreamcast CDI to ISO ability, and ISO fixer
    'DONE TODO can't change hl module in sys properties. It's just not updated.


    'DONE TODO save last check result and show it in system manager table
    'DONE TODO fix auto selection change in listbox when renaming
    'DONE TODO after showing any defaultly hided column and saving table config, after program restart, column Is shown, but menu item does not checked
    'DONE TODO ajustable font size in matcher listboxes
    'DONE TODO hotkey (ctrl+enter or F2) to enter in cell edit mode
    'DONE TODO ctrl+c should copy a single cell, not an entire line in the table, when in edit mode
    'DONE TODO MSU auto audio track and related files renaming
    'DONE TODO adding entry in 'handle theese extension togeether' list in options does not clear textbox
    'DONE TODO filter in game list in matcher not working
    'DONE TODO find-as-you-type in matcher listboxes
    'DONE TODO If a disk does not exist (i.e. network drive), and multiple paths are set in RL - continue in another path, not abort
    'DONE TODO RocketLauncher media renaming
    'DONE TODO Save association matcher options
    'DONE TODO Autorenamer - better word removing. Sometimes word is removed in the midle of word (snowmobiLES). Sometimes does not at all (wii, spyro - down of THE dragon)
    'DONE TODO Save regexp options
    'DONE TODO Quick stats, at right of the main table, under buttons - how many items was found and how many still missing of each category
    'DONE TODO add "NOT" to regexp options (letter, number)
    'DONE TODO When using Roms->Db way autofilter (the inversed one), with show unmatched only for roms and db, can cause NullReference in matcher line 279 (.ListBox1.SelectedIndex)
    'DONE TODO Sort systems list alphabetically - WARNING, some code depends on system order, including add system in system manager and unpack system

    'MOSTLY DONE TODO - 2way autofilter (making it work both way in same time will confuse matcher when listboxes selection will change after association)

#Region "Declarations"
    Structure check_param
        Dim x As Xml.XmlDocument
        Dim sys As String
        Dim path As String
    End Structure
    Public xmlPath As String = ""
    Private ini As New IniFileApi
    Private xml_class As Class2_xml
    Public matcher_class As Class3_matcher
    Private system_manager_class As Class5_system_manager
    Private isoChecker As ISOChecker
    Private clrmame_class As Class4_clrmamepro
    Public system_list As New DataTable
    Public editor_delete_command_list As New List(Of DataGridViewRow)
    Public editor_insert_command_list As New List(Of DataGridViewRow)
    Public editor_update_command_list As New Dictionary(Of DataGridViewRow, String)
    Friend subfoldered As Boolean = False
    Friend subfoldered2 As Boolean = False
    Friend undo As New List(Of List(Of String))
    Friend undo_humanReadable As New List(Of List(Of String))
    Private romExtensions() As String = {""}
    Private oldCombo4Value As Integer = 0
    Private useParentVids, useParentThemes As Boolean
    Private WithEvents bg_check As New BackgroundWorker() With {.WorkerReportsProgress = True}
    Friend WithEvents myContextMenu As New ToolStripDropDownMenu 'checking for roms in other folders
    Friend WithEvents myContextMenu2 As New ToolStripDropDownMenu 'matcher options
    Friend WithEvents myContextMenu3 As New ToolStripDropDownMenu With {.Name = "myContextMenu3"} 'move unneeded
    Friend WithEvents myContextMenu4 As New ToolStripDropDownMenu With {.Name = "myContextMenu4"} 'main table columns hide/show
    Friend WithEvents myContextMenu5 As New ToolStripDropDownMenu 'folder2xml options
    Friend WithEvents myContextMenu6 As New ToolStripDropDownMenu 'convert to clrmamepro
    Friend WithEvents myContextMenu7 As New ToolStripDropDownMenu 'main table context menu
    Friend WithEvents myContextMenu8 As New ToolStripDropDownMenu 'system manager table context menu
    Friend WithEvents RadioStrip1 As New RadioButton
    Friend WithEvents RadioStrip2 As New RadioButton
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
    Dim WithEvents CheckStrip2_11 As New CheckBox With {.Name = "11", .Text = "Show HL Artwork", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = False}
    Dim WithEvents CheckStrip2_12 As New CheckBox With {.Name = "12", .Text = "Show HL Backgrounds", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = False}
    Dim WithEvents CheckStrip2_13 As New CheckBox With {.Name = "13", .Text = "Show HL Bezels", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = False}
    Dim WithEvents CheckStrip2_14 As New CheckBox With {.Name = "14", .Text = "Show HL Fade", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = False}
    Dim WithEvents CheckStrip2_15 As New CheckBox With {.Name = "15", .Text = "Show HL Guides", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = False}
    Dim WithEvents CheckStrip2_16 As New CheckBox With {.Name = "16", .Text = "Show HL Manuals", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = False}
    Dim WithEvents CheckStrip2_17 As New CheckBox With {.Name = "17", .Text = "Show HL Music", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = False}
    Dim WithEvents CheckStrip2_18 As New CheckBox With {.Name = "18", .Text = "Show HL Videos", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = False}
    Dim WithEvents CheckStrip2_19 As New CheckBox With {.Name = "19", .Text = "Show CRC", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = False}
    Dim WithEvents CheckStrip2_20 As New CheckBox With {.Name = "20", .Text = "Show Manufacturer", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = False}
    Dim WithEvents CheckStrip2_21 As New CheckBox With {.Name = "21", .Text = "Show Year", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = False}
    Dim WithEvents CheckStrip2_22 As New CheckBox With {.Name = "22", .Text = "Show Genre", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = False}

    Dim WithEvents CheckStrip3_1 As New CheckBox With {.Name = "01", .Text = "Main Menu Database Entry", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = True, .Enabled = True}
    Dim WithEvents CheckStrip3_2 As New CheckBox With {.Name = "02", .Text = "Database", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = True, .Enabled = True}
    Dim WithEvents CheckStrip3_3 As New CheckBox With {.Name = "03", .Text = "Main Menu Theme", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = True, .Enabled = True}
    Dim WithEvents CheckStrip3_4 As New CheckBox With {.Name = "04", .Text = "System Theme", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = True, .Enabled = True}
    Dim WithEvents CheckStrip3_5 As New CheckBox With {.Name = "05", .Text = "Settings", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = True, .Enabled = True}
    Dim WithEvents CheckStrip3_6 As New CheckBox With {.Name = "06", .Text = "Emulator path", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = True, .Enabled = True}
    Dim WithEvents CheckStrip3_7 As New CheckBox With {.Name = "07", .Text = "Rom Path", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = True, .Enabled = True}
    Dim WithEvents CheckStrip3_8 As New CheckBox With {.Name = "08", .Text = "Main Menu Wheel", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = False, .Enabled = True}
    Dim WithEvents CheckStrip3_9 As New CheckBox With {.Name = "09", .Text = "Main Menu Video", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = False, .Enabled = True}
    Dim WithEvents CheckStrip3_10 As New CheckBox With {.Name = "10", .Text = "HL Artwork", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = False, .Enabled = True}
    Dim WithEvents CheckStrip3_11 As New CheckBox With {.Name = "11", .Text = "HL Background", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = False, .Enabled = True}
    Dim WithEvents CheckStrip3_12 As New CheckBox With {.Name = "12", .Text = "HL Bezel", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = False, .Enabled = True}
    Dim WithEvents CheckStrip3_13 As New CheckBox With {.Name = "13", .Text = "HL Fade", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = False, .Enabled = True}
    Dim WithEvents CheckStrip3_14 As New CheckBox With {.Name = "14", .Text = "HL Guide", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = False, .Enabled = True}
    Dim WithEvents CheckStrip3_15 As New CheckBox With {.Name = "15", .Text = "HL Manual", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = False, .Enabled = True}
    Dim WithEvents CheckStrip3_16 As New CheckBox With {.Name = "16", .Text = "HL Music", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = False, .Enabled = True}
    Dim WithEvents CheckStrip3_17 As New CheckBox With {.Name = "17", .Text = "HL Video", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = False, .Enabled = True}
    Dim WithEvents CheckStrip3_18 As New CheckBox With {.Name = "18", .Text = "Roms", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = False, .Enabled = True}
    Dim WithEvents CheckStrip3_19 As New CheckBox With {.Name = "19", .Text = "Videos", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = False, .Enabled = True}
    Dim WithEvents CheckStrip3_20 As New CheckBox With {.Name = "20", .Text = "Themes", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = False, .Enabled = True}
    Dim WithEvents CheckStrip3_21 As New CheckBox With {.Name = "21", .Text = "Wheels", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = False, .Enabled = True}
    Dim WithEvents CheckStrip3_22 As New CheckBox With {.Name = "22", .Text = "Artworks", .BackColor = Color.FromArgb(0, 255, 0, 0), .Checked = False, .Enabled = True}
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> Private Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function
#End Region

#Region "Main Form Actions (loadForm, system select, main check)"
    'FORM LOAD
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Test zipping block
        'Dim zc As New SevenZip.SevenZipCompressor
        'zc.ArchiveFormat = SevenZip.OutArchiveFormat.SevenZip
        'zc.CompressionMethod = SevenZip.CompressionMethod.Default
        'zc.CompressionLevel = SevenZip.CompressionLevel.High
        'Dim files As New Dictionary(Of String, String)
        'files.Add("VGhlIHBhdGgtbGVuZ3RoIHByb2JsZW0gaXMgY29tcGxldGVseSB1bnJlbGF0ZWQgdG8gdGhlIEZBVCBsaW1pdGF0aW9uLgoKTlRGUyBjYW4gYWN0dWFsbHkgc3VwcG9ydCBwYXRobmFtZXMgdXAgdG8gMzIsMDAwIGNoYXJhY3RlcnMgaW4gbGVuZ3RoLgoKTWljcm9zb2Z0J3MgdG9vbHMsIG9uIHRoZSBvdGhlciBoYW5kIChleHBsb3Jlci5leGUsIGRpciwgZXRjLikgYXJlIGNvZGVkIHRvIHRoZSBBTlNJIHNwZWMgZm9yIHdyaXRpbmcgcGF0aHMsIHNvIHRoZXkgY3JhcC1vdXQgYXQgMjYwIGNoYXJhY3RlcnMuICBJJ20gbm90IGV2ZW4gc3VyZSBpZiB0aGF0IGZhaWx1cmUgd2lsbCBzaG93IHVwIGluIEZpbGVtb24uICBJIHRoaW5rIHRoZSBjYWxsIGNyYXBzIG91dCBiZWZvcmUgaXQgZXZlbiBnZXRzIGRvd24gdG8gdGhlIGZpbHRlciBkcml2ZXIuCgpBbnl3YXksIHRoZXJlJ3MgYSBmZXcgdG9vbHMgeW91IGNhbiBnZXQgdG8gd29yayBhcm91bmQgdGhlIDI2MCBjaGFyYWN0ZXIgbGltaXQgdGhhdCBoYXMgYmVlbiB3aXRoIHVzIHNpbmNlIDE5OTQuICBEZWxpbW9uLCBXaW4zMkV4cGxvcmVyLCBhbmQgYWJvdXQgdGhlIE9OTFkgdG9vbCB0aGF0IGlzIGNhcGFibGUgb2YgYXJjaGl2aW5nIG9yIHppcHBpbmcg.txt", "D:\1\Planetside.Software.Terragen.Professional.v4.1.21.X64-AMPED.rar")
        'files.Add("2", "D:\1\Planetside.Software.Terragen.Professional.v4.1.21.X64-AMPED.rar")
        'files.Add("3", "D:\1\Planetside.Software.Terragen.Professional.v4.1.21.X64-AMPED.rar")
        'files.Add("4", "D:\1\Planetside.Software.Terragen.Professional.v4.1.21.X64-AMPED.rar")
        'files.Add("5", "D:\1\Planetside.Software.Terragen.Professional.v4.1.21.X64-AMPED.rar")
        'zc.CompressFileDictionary(files, "D:\Documents\My_Progs\HyperspinChecker\git_MSVS\HyperspinChecker\source\bin\Debug\SystemPackages\test.zip")

        Class1.Log("Initializing...")
        ComboBox7.Left = TextBox4.Left
        ComboBox7.Width = TextBox4.Width - 10
        ComboBox8.SelectedIndex = 0
        ComboBox9.SelectedIndex = 0
        system_list.Columns.Add("system")

        If FileSystem.DirectoryExists(".\Localization") Then
            For Each file In FileSystem.GetFiles(".\Localization", SearchOption.SearchTopLevelOnly, {"*.ini"})
                Dim f As String = file.Substring(file.LastIndexOf("\") + 1)
                f = f.Substring(0, f.LastIndexOf("."))
                ComboBox13.Items.Add(f)
            Next
        End If

        'Init Table Columns
        Form1_Load_Sub_initTableColumns()

        'Init Context Menus
        Class1.Log("Creating menus")
        Form1_Load_Sub_initContextMenus()

        'Convert current DB to clrmame pro dat - options panel
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

        'Load config
        Dim v As Boolean
        Dim s As String = ""
        Dim ini2 As New IniFileApi
        ini2.IniFile(Class1.confPath)
        TextBox14.Text = ini2.IniReadValue("main", "HyperSpin_Path")
        TextBox16.Text = ini2.IniReadValue("main", "search_cue_for")
        TextBox28.Text = ini2.IniReadValue("3rd_Party_Tools", "UltraISO")
        If Not ini2.IniReadValue("main", "rename_just_CUE") = "1" Then CheckBox6.Checked = False
        If ini2.IniReadValue("main", "archives_rename_inside") = "1" Then CheckBox28.Checked = True
        If ini2.IniReadValue("main", "archives_remove_unneeded") = "1" Then CheckBox29.Checked = True
        If ini2.IniReadValue("main", "usehlv3") = "1" Then CheckBox26.Checked = True
        If ini2.IniReadValue("main", "sort_systems") = "1" Then CheckBox32.Checked = True
        If ini2.IniReadValue("main", "check_hl_game_media") = "1" Then CheckBox30.Checked = True
        If ini2.IniReadValue("main", "check_hl_system_media") = "1" Then CheckBox31.Checked = True
        Dim i As Integer = 1
        Do While ini2.IniReadValue("handle_together", i.ToString) <> ""
            ListBox4.Items.Add(ini2.IniReadValue("handle_together", i.ToString))
            i += 1
        Loop
        'Main table startup config
        For c As Integer = 0 To DataGridView1.ColumnCount - 1
            s = ini2.IniReadValue("Main_Table_Columns_Config", "Col_" + c.ToString + "_visible")
            If s <> "" Then
                If s = "0" Then v = False Else v = True
                DirectCast(myContextMenu4.Controls(Format(c, "00")), CheckBox).Checked = v
            End If

            s = ini2.IniReadValue("Main_Table_Columns_Config", "Col_" + c.ToString + "_Width")
            If s <> "" Then DataGridView1.Columns(c).Width = CInt(s)
        Next
        'System table startup config
        For c As Integer = 1 To DataGridView2.ColumnCount - 1
            s = ini2.IniReadValue("System_Table_Columns_Config", "Col_" + c.ToString + "_visible")
            If s <> "" Then
                If s = "0" Then v = False Else v = True
                DirectCast(myContextMenu8.Controls(Format(c, "00")), CheckBox).Checked = v
            End If

            s = ini2.IniReadValue("System_Table_Columns_Config", "Col_" + c.ToString + "_Width")
            If s <> "" Then DataGridView2.Columns(c).Width = CInt(s)
        Next
        'Main_window_size
        s = ini2.IniReadValue("Main", "Main_window_size")
        If s <> "" Then
            Me.Width = CInt(s.Split({"x"c})(0))
            Me.Height = CInt(s.Split({"x"c})(1))
        End If
        'Freeze HL Path
        s = ini2.IniReadValue("Main", "Freeze_HL_path")
        If s <> "" Then
            CheckBox21.Checked = True
            TextBox18.Text = s
        End If
        'Compression
        s = ini2.IniReadValue("Main", "Compression_type")
        If s <> "" Then ComboBox10.Text = s Else ComboBox10.SelectedIndex = 0
        s = ini2.IniReadValue("Main", "Compression_method")
        If s <> "" Then ComboBox11.Text = s Else ComboBox11.SelectedIndex = 0
        s = ini2.IniReadValue("Main", "Compression_level")
        If s <> "" Then ComboBox12.Text = s Else ComboBox12.SelectedIndex = 5
        s = ini2.IniReadValue("Main", "Compression_temp")
        If s <> "" Then TextBox30.Text = s

        'Matcher Font Size
        s = ini2.IniReadValue("Main", "MatcherFontSize")
        Dim fs As Single
        If Single.TryParse(s, fs) Then NumericUpDown1.Value = CInt(fs)

        'Top menu - Matcher options
        s = ini2.IniReadValue("Main", "AssocOption_fileInHsFolder")
        If s <> "" Then
            For Each t As ToolStripMenuItem In AssocOption_fileInHsFolder.DropDownItems
                If t.Name.ToUpper = s.ToUpper Then t.PerformClick() : Exit For
            Next
        End If
        s = ini2.IniReadValue("Main", "AssocOption_fileInDiffFolder")
        If s <> "" Then
            For Each t As ToolStripMenuItem In AssocOption_fileInDiffFolder.DropDownItems
                If t.Name.ToUpper = s.ToUpper Then t.PerformClick() : Exit For
            Next
        End If

        'Regular expression - active and presets
        s = ini2.IniReadValue("RegEx", "active")
        If s <> "" Then Class3_matcher.autofilter_regex = s
        s = ini2.IniReadValue("RegEx", "opt_strip_br")
        If s.ToUpper = "TRUE" Then Class3_matcher.autofilter_regex_options(0) = True
        s = ini2.IniReadValue("RegEx", "opt_strip_pr")
        If s.ToUpper = "TRUE" Then Class3_matcher.autofilter_regex_options(1) = True
        s = ini2.IniReadValue("RegEx", "opt_out_group")
        If IsNumeric(s) Then Class3_matcher.autofilter_regex_opt_outGroup = CInt(s)
        For Each k In ini2.IniListKey("RegEx")
            If k.ToLower.StartsWith("preset") Then
                Dim p = ini2.IniReadValue("RegEx", k).Trim.Split({"^^^"}, StringSplitOptions.RemoveEmptyEntries)
                If p.Count = 5 Then Class3_matcher.autofilter_regex_presets.Add(p(0).Trim, p(1).Trim + "^^^" + p(2).Trim + "^^^" + p(3).Trim + "^^^" + p(4).Trim)
            End If
        Next

        'Localization
        s = ini2.IniReadValue("Main", "Language")
        If s.Trim = "" Then
            ComboBox13.SelectedIndex = 0
        Else
            For t As Integer = 0 To ComboBox13.Items.Count - 1
                If s.Trim.ToUpper = ComboBox13.Items(t).ToString.Trim.ToUpper Then ComboBox13.SelectedIndex = t : Exit For
            Next
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
    Private Sub Form1_Load_Sub_initTableColumns()
        'Main table columns set up
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
        DataGridView1.Columns.Add("col11", "HL Artwork")
        DataGridView1.Columns.Add("col12", "HL Backgrounds")
        DataGridView1.Columns.Add("col13", "HL Bezels")
        DataGridView1.Columns.Add("col14", "HL Fade")
        DataGridView1.Columns.Add("col15", "HL Guides")
        DataGridView1.Columns.Add("col16", "HL Manuals")
        DataGridView1.Columns.Add("col17", "HL Music")
        DataGridView1.Columns.Add("col18", "HL Videos")
        DataGridView1.Columns.Add("col19", "crc")
        DataGridView1.Columns.Add("col20", "Manufacturer")
        DataGridView1.Columns.Add("col21", "Year")
        DataGridView1.Columns.Add("col22", "Genre")
        DataGridView1.Columns.Add("col23", "Rating")
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
        DataGridView1.Columns(11).Width = 50 : DataGridView1.Columns(11).ReadOnly = True : DataGridView1.Columns(11).Visible = False
        DataGridView1.Columns(12).Width = 50 : DataGridView1.Columns(12).ReadOnly = True : DataGridView1.Columns(12).Visible = False
        DataGridView1.Columns(13).Width = 50 : DataGridView1.Columns(13).ReadOnly = True : DataGridView1.Columns(13).Visible = False
        DataGridView1.Columns(14).Width = 50 : DataGridView1.Columns(14).ReadOnly = True : DataGridView1.Columns(14).Visible = False
        DataGridView1.Columns(15).Width = 50 : DataGridView1.Columns(15).ReadOnly = True : DataGridView1.Columns(15).Visible = False
        DataGridView1.Columns(16).Width = 50 : DataGridView1.Columns(16).ReadOnly = True : DataGridView1.Columns(16).Visible = False
        DataGridView1.Columns(17).Width = 50 : DataGridView1.Columns(17).ReadOnly = True : DataGridView1.Columns(17).Visible = False
        DataGridView1.Columns(18).Width = 50 : DataGridView1.Columns(18).ReadOnly = True : DataGridView1.Columns(18).Visible = False
        DataGridView1.Columns(19).Width = 75 : DataGridView1.Columns(19).ReadOnly = False : DataGridView1.Columns(19).Visible = False
        DataGridView1.Columns(20).Width = 180 : DataGridView1.Columns(20).ReadOnly = False : DataGridView1.Columns(20).Visible = False
        DataGridView1.Columns(21).Width = 65 : DataGridView1.Columns(21).ReadOnly = False : DataGridView1.Columns(21).Visible = False
        DataGridView1.Columns(22).Width = 120 : DataGridView1.Columns(22).ReadOnly = False : DataGridView1.Columns(22).Visible = False
        DataGridView1.Columns(23).Width = 120 : DataGridView1.Columns(23).ReadOnly = False : DataGridView1.Columns(23).Visible = False
        Dim t As Type = DataGridView1.GetType
        Dim pi As Reflection.PropertyInfo = t.GetProperty("DoubleBuffered", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)
        pi.SetValue(DataGridView1, True, Nothing)

        'System manager table columns set up
        DataGridView2.Columns.Add("col0", "System")
        DataGridView2.Columns.Add("col1", "Main Menu")
        DataGridView2.Columns.Add("col2", "Database")
        DataGridView2.Columns.Add("col3", "MainMenu Theme")
        DataGridView2.Columns.Add("col4", "System Theme")
        DataGridView2.Columns.Add("col5", "Settings")
        DataGridView2.Columns.Add("col6", "Emulator Path")
        DataGridView2.Columns.Add("col7", "Rom Path")
        DataGridView2.Columns.Add("col8", "Main Menu Wheel")
        DataGridView2.Columns.Add("col9", "Main Menu Video")
        DataGridView2.Columns.Add("col10", "HL Artworks")
        DataGridView2.Columns.Add("col11", "HL Background")
        DataGridView2.Columns.Add("col12", "HL Bezel")
        DataGridView2.Columns.Add("col13", "HL Fade")
        DataGridView2.Columns.Add("col14", "HL Guide")
        DataGridView2.Columns.Add("col15", "HL Manual")
        DataGridView2.Columns.Add("col16", "HL Music")
        DataGridView2.Columns.Add("col17", "HL Video")
        DataGridView2.Columns.Add("col18", "Roms")
        DataGridView2.Columns.Add("col19", "Videos")
        DataGridView2.Columns.Add("col20", "Themes")
        DataGridView2.Columns.Add("col21", "Wheels")
        DataGridView2.Columns.Add("col22", "Artworks")
        DataGridView2.Columns(0).Width = 250
        DataGridView2.Columns(1).Width = 80 : DataGridView2.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView2.Columns(2).Width = 80 : DataGridView2.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView2.Columns(3).Width = 80 : DataGridView2.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView2.Columns(4).Width = 80 : DataGridView2.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView2.Columns(5).Width = 80 : DataGridView2.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView2.Columns(6).Width = 80 : DataGridView2.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView2.Columns(7).Width = 80 : DataGridView2.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView2.Columns(8).Width = 90 : DataGridView2.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView2.Columns(9).Width = 90 : DataGridView2.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView2.Columns(8).DisplayIndex = 3 : DataGridView2.Columns(8).Visible = False
        DataGridView2.Columns(9).DisplayIndex = 5 : DataGridView2.Columns(9).Visible = False
        DataGridView2.Columns(10).Width = 80 : DataGridView2.Columns(10).Visible = False : DataGridView2.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView2.Columns(11).Width = 80 : DataGridView2.Columns(11).Visible = False : DataGridView2.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView2.Columns(12).Width = 80 : DataGridView2.Columns(12).Visible = False : DataGridView2.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView2.Columns(13).Width = 80 : DataGridView2.Columns(13).Visible = False : DataGridView2.Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView2.Columns(14).Width = 80 : DataGridView2.Columns(14).Visible = False : DataGridView2.Columns(14).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView2.Columns(15).Width = 80 : DataGridView2.Columns(15).Visible = False : DataGridView2.Columns(15).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView2.Columns(16).Width = 80 : DataGridView2.Columns(16).Visible = False : DataGridView2.Columns(16).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView2.Columns(17).Width = 80 : DataGridView2.Columns(17).Visible = False : DataGridView2.Columns(17).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView2.Columns(18).Width = 90 : DataGridView2.Columns(18).Visible = False
        DataGridView2.Columns(19).Width = 90 : DataGridView2.Columns(19).Visible = False
        DataGridView2.Columns(20).Width = 90 : DataGridView2.Columns(20).Visible = False
        DataGridView2.Columns(21).Width = 90 : DataGridView2.Columns(21).Visible = False
        DataGridView2.Columns(22).Width = 90 : DataGridView2.Columns(22).Visible = False
    End Sub
    Private Sub Form1_Load_Sub_initContextMenus()
        'Context menu - check missing rom or media in another folder
        myContextMenu.Items.Add("check for missing Roms")
        myContextMenu.Items.Add(New ToolStripSeparator)
        myContextMenu.Items.Add("check for missing Video")
        myContextMenu.Items.Add(New ToolStripSeparator)
        myContextMenu.Items.Add("check for missing Wheels")
        myContextMenu.Items.Add("check for missing Artwork1")
        myContextMenu.Items.Add("check for missing Artwork2")
        myContextMenu.Items.Add("check for missing Artwork3")
        myContextMenu.Items.Add("check for missing Artwork4")
        myContextMenu.Items.Add("check for missing Themes")

        'Context menu - move unneeded to sub folder
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


        'Matcher options
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

        'Main table columns headers context menu
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
        Dim CheckStripHost2_15 As New ToolStripControlHost(CheckStrip2_15)
        Dim CheckStripHost2_16 As New ToolStripControlHost(CheckStrip2_16)
        Dim CheckStripHost2_17 As New ToolStripControlHost(CheckStrip2_17)
        Dim CheckStripHost2_18 As New ToolStripControlHost(CheckStrip2_18)
        Dim CheckStripHost2_19 As New ToolStripControlHost(CheckStrip2_19)
        Dim CheckStripHost2_20 As New ToolStripControlHost(CheckStrip2_20)
        Dim CheckStripHost2_21 As New ToolStripControlHost(CheckStrip2_21)
        Dim CheckStripHost2_22 As New ToolStripControlHost(CheckStrip2_22)
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
        myContextMenu4.Items.Add(CheckStripHost2_15)
        myContextMenu4.Items.Add(CheckStripHost2_16)
        myContextMenu4.Items.Add(CheckStripHost2_17)
        myContextMenu4.Items.Add(CheckStripHost2_18)
        myContextMenu4.Items.Add(CheckStripHost2_19)
        myContextMenu4.Items.Add(CheckStripHost2_20)
        myContextMenu4.Items.Add(CheckStripHost2_21)
        myContextMenu4.Items.Add(CheckStripHost2_22)
        myContextMenu4.Items.Add(New ToolStripSeparator)
        myContextMenu4.Items.Add(Preset_Checker)
        myContextMenu4.Items.Add(Preset_Editor)
        myContextMenu4.Items.Add(Preset_HL)
        myContextMenu4.Items.Add(New ToolStripSeparator)
        myContextMenu4.Items.Add(Save_current_cols_conf_as_startup)

        'System Manager table columns headers context menu
        Dim CheckStripHost3_1 As New ToolStripControlHost(CheckStrip3_1)
        Dim CheckStripHost3_2 As New ToolStripControlHost(CheckStrip3_2)
        Dim CheckStripHost3_3 As New ToolStripControlHost(CheckStrip3_3)
        Dim CheckStripHost3_4 As New ToolStripControlHost(CheckStrip3_4)
        Dim CheckStripHost3_5 As New ToolStripControlHost(CheckStrip3_5)
        Dim CheckStripHost3_6 As New ToolStripControlHost(CheckStrip3_6)
        Dim CheckStripHost3_7 As New ToolStripControlHost(CheckStrip3_7)
        Dim CheckStripHost3_8 As New ToolStripControlHost(CheckStrip3_8)
        Dim CheckStripHost3_9 As New ToolStripControlHost(CheckStrip3_9)
        Dim CheckStripHost3_10 As New ToolStripControlHost(CheckStrip3_10)
        Dim CheckStripHost3_11 As New ToolStripControlHost(CheckStrip3_11)
        Dim CheckStripHost3_12 As New ToolStripControlHost(CheckStrip3_12)
        Dim CheckStripHost3_13 As New ToolStripControlHost(CheckStrip3_13)
        Dim CheckStripHost3_14 As New ToolStripControlHost(CheckStrip3_14)
        Dim CheckStripHost3_15 As New ToolStripControlHost(CheckStrip3_15)
        Dim CheckStripHost3_16 As New ToolStripControlHost(CheckStrip3_16)
        Dim CheckStripHost3_17 As New ToolStripControlHost(CheckStrip3_17)
        Dim CheckStripHost3_18 As New ToolStripControlHost(CheckStrip3_18)
        Dim CheckStripHost3_19 As New ToolStripControlHost(CheckStrip3_19)
        Dim CheckStripHost3_20 As New ToolStripControlHost(CheckStrip3_20)
        Dim CheckStripHost3_21 As New ToolStripControlHost(CheckStrip3_21)
        Dim CheckStripHost3_22 As New ToolStripControlHost(CheckStrip3_22)
        myContextMenu8.Items.Add(CheckStripHost3_1)
        myContextMenu8.Items.Add(CheckStripHost3_2)
        myContextMenu8.Items.Add(CheckStripHost3_3)
        myContextMenu8.Items.Add(CheckStripHost3_4)
        myContextMenu8.Items.Add(CheckStripHost3_5)
        myContextMenu8.Items.Add(CheckStripHost3_6)
        myContextMenu8.Items.Add(CheckStripHost3_7)
        myContextMenu8.Items.Add(CheckStripHost3_8)
        myContextMenu8.Items.Add(CheckStripHost3_9)
        myContextMenu8.Items.Add(CheckStripHost3_10)
        myContextMenu8.Items.Add(CheckStripHost3_11)
        myContextMenu8.Items.Add(CheckStripHost3_12)
        myContextMenu8.Items.Add(CheckStripHost3_13)
        myContextMenu8.Items.Add(CheckStripHost3_14)
        myContextMenu8.Items.Add(CheckStripHost3_15)
        myContextMenu8.Items.Add(CheckStripHost3_16)
        myContextMenu8.Items.Add(CheckStripHost3_17)
        myContextMenu8.Items.Add(CheckStripHost3_18)
        myContextMenu8.Items.Add(CheckStripHost3_19)
        myContextMenu8.Items.Add(CheckStripHost3_20)
        myContextMenu8.Items.Add(CheckStripHost3_21)
        myContextMenu8.Items.Add(CheckStripHost3_22)
        myContextMenu8.Items.Add(New ToolStripSeparator)
        myContextMenu8.Items.Add(Preset_SysMngr_Default)
        myContextMenu8.Items.Add(Preset_SysMngr_Checker)
        myContextMenu8.Items.Add(Preset_SysMngr_Manager)
        myContextMenu8.Items.Add(Preset_SysMngr_HLMedia)
        myContextMenu8.Items.Add(Preset_SysMngr_Full)
        myContextMenu8.Items.Add(New ToolStripSeparator)
        myContextMenu8.Items.Add(Save_current_cols_conf_as_startup)
    End Sub

    'Form closing - this is handled to save window size in config
    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        ini.IniFile(Class1.confPath)
        ini.IniWriteValue("Main", "Main_window_size", Me.Width.ToString + "x" + Me.Height.ToString)
        Application.Exit()
    End Sub

    'Main Check
    Private Sub check(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Class1.data.Clear()
        Class1.data_crc.Clear()
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

        'Get rom path(s)
        retrieve_rom_path_from_HL()

        'Check rom path(s)
        Dim romPath As String = ""
        If Class1.romPath = "" Or Class1.romPath = "\" Then MsgBox("Can't retrive rom path.") : Exit Sub
        If Not Class1.romPath.Contains("|") Then 'Handle HL multiple paths
            'Single path
            If Class1.romPath.Substring(1, 2) = ":\" Then
                Dim p_drv As String = Class1.romPath.Substring(0, 3).ToUpper
                Dim ex As Boolean = FileIO.FileSystem.Drives.Any(Function(drv) drv.Name = p_drv And drv.DriveType <> IO.DriveType.CDRom)
                If Not ex Then MsgBox("Rompath refers to drive " + p_drv + ", but it does not exist or is CDRom." + vbCrLf + "You can fix rom paths in system manager.") : Exit Sub
            End If
            romPath = Class1.romPath
        Else
            'Multiple paths routine
            Dim DriveNoExistMessageWasShown As Boolean = False
            For Each temppath2 As String In Class1.romPath.Split({"|"c}, StringSplitOptions.RemoveEmptyEntries)
                If Not temppath2.EndsWith("\") Then temppath2 = temppath2 + "\"
                If temppath2.Substring(1, 2) = ":\" Then
                    Dim p_drv As String = temppath2.Substring(0, 3).ToUpper
                    Dim ex As Boolean = FileIO.FileSystem.Drives.Any(Function(drv) drv.Name = p_drv And drv.DriveType <> IO.DriveType.CDRom)
                    If Not ex Then
                        If Not DriveNoExistMessageWasShown Then
                            MsgBox("One of rompaths refers to drive " + p_drv + ", but it does not exist or is CDRom." + vbCrLf + "You can fix rom paths in system manager.")
                            DriveNoExistMessageWasShown = True
                        End If
                    Else
                        romPath += temppath2 + "|"
                    End If
                ElseIf temppath2.StartsWith("\\") Then
                    'Network path
                    Label2.Text = "Checking network path..." : Label2.Refresh()
                    If IO.Directory.Exists(temppath2) Then
                        romPath += temppath2 + "|"
                    End If
                End If
            Next

            If romPath.EndsWith("|") Then romPath = romPath.Substring(0, romPath.Length - 1)
            If romPath.Trim = "" Then MsgBox("No suitable rom paths found.") : Exit Sub
        End If

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

        Label2.Text = "..."
        ProgressBar1.Value = 0
        ProgressBar1.Maximum = x.SelectNodes("/menu/game").Count
        Dim param As New check_param With {.x = x, .sys = ComboBox1.SelectedItem.ToString, .path = romPath}
        bg_check.RunWorkerAsync(param)
    End Sub
    Private Sub check_bg(sender As Object, e As DoWorkEventArgs) Handles bg_check.DoWork
        Dim a(18) As String
        Dim romName As String
        Dim tempStr As String
        Dim progress As Integer = 0
        Dim param = DirectCast(e.Argument, check_param)
        Dim x As Xml.XmlDocument = param.x
        Dim counters() As Integer = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}
        Dim a_crc, a_manufacturer, a_year, a_genre, a_cloneof, a_rating As String
        a(11) = "Not checked" : a(12) = "Not checked" : a(13) = "Not checked" : a(14) = "Not checked" : a(15) = "Not checked" : a(16) = "Not checked" : a(17) = "Not checked" : a(18) = "Not checked"
        For Each node As Xml.XmlNode In x.SelectNodes("/menu/game")
            If node.SelectSingleNode("crc") IsNot Nothing Then a_crc = node.SelectSingleNode("crc").InnerText Else a_crc = ""
            If node.SelectSingleNode("manufacturer") IsNot Nothing Then a_manufacturer = node.SelectSingleNode("manufacturer").InnerText Else a_manufacturer = ""
            If node.SelectSingleNode("year") IsNot Nothing Then a_year = node.SelectSingleNode("year").InnerText Else a_year = ""
            If node.SelectSingleNode("genre") IsNot Nothing Then a_genre = node.SelectSingleNode("genre").InnerText Else a_genre = ""
            If node.SelectSingleNode("rating") IsNot Nothing Then a_rating = node.SelectSingleNode("rating").InnerText Else a_rating = ""
            If node.SelectSingleNode("cloneof") IsNot Nothing Then a_cloneof = node.SelectSingleNode("cloneof").InnerText Else a_cloneof = ""

            'Get romname and description
            a(0) = node.SelectSingleNode("description").InnerText
            romName = node.Attributes.GetNamedItem("name").Value
            Class1.romlist.Add(romName.ToLower)

            'Check rom
            a(2) = ""
            If Not param.path.Contains("|") Then
                'Simple path
                If tryToFindRom(param.path, romExtensions, romName) Then a(2) = "YES"
            Else
                'HL Multiple paths routine
                For Each temppath2 As String In param.path.Split({"|"c}, StringSplitOptions.RemoveEmptyEntries)
                    If Not temppath2.EndsWith("\") Then temppath2 = temppath2 + "\"
                    If tryToFindRom(temppath2, romExtensions, romName) Then a(2) = "YES" : Exit For
                Next
            End If
            If a(2) = "" Then a(2) = "NO" Else Class1.romFoundlist.Add(romName.ToLower) : counters(1) += 1

            'Check video
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
            If a(3) = "YES" Then counters(2) += 1

            'Check theme
            Dim useThemeFromParent As Boolean = False
            Dim p As String = Class1.HyperspinPath + "Media\" + param.sys
            If FileSystem.FileExists(p + "\Themes\" + romName + ".zip") Then a(4) = "YES" Else a(4) = "NO"
            If a(4) = "NO" And CheckBox24.Checked And a_cloneof <> "" Then
                If FileSystem.FileExists(p + "\Themes\" + a_cloneof + ".zip") Then a(4) = "YES" : useThemeFromParent = True
            End If
            If a(4) = "YES" Then counters(3) += 1

            'Check Wheels and artworks
            If FileSystem.FileExists(p + "\Images\Wheel\" + romName + ".png") Or FileSystem.FileExists(p + "\Images\Wheel\" + romName + ".jpg") Then a(5) = "YES" : counters(4) += 1 Else a(5) = "NO"
            If FileSystem.FileExists(p + "\Images\Artwork1\" + romName + ".png") Or FileSystem.FileExists(p + "\Images\Artwork1\" + romName + ".jpg") Then a(6) = "YES" : counters(5) += 1 Else a(6) = "NO"
            If FileSystem.FileExists(p + "\Images\Artwork2\" + romName + ".png") Or FileSystem.FileExists(p + "\Images\Artwork2\" + romName + ".jpg") Then a(7) = "YES" : counters(6) += 1 Else a(7) = "NO"
            If FileSystem.FileExists(p + "\Images\Artwork3\" + romName + ".png") Or FileSystem.FileExists(p + "\Images\Artwork3\" + romName + ".jpg") Then a(8) = "YES" : counters(7) += 1 Else a(8) = "NO"
            If FileSystem.FileExists(p + "\Images\Artwork4\" + romName + ".png") Or FileSystem.FileExists(p + "\Images\Artwork4\" + romName + ".jpg") Then a(9) = "YES" : counters(8) += 1 Else a(9) = "NO"
            If FileSystem.FileExists(p + "\Sound\Background Music\" + romName + ".mp3") Then a(10) = "YES" : counters(9) += 1 Else a(10) = "NO"

            'Check HL/RL Media
            If CheckBox30.Checked Then
                If FileSystem.DirectoryExists(Class1.HyperlaunchPath + "\Media") Then
                    a(11) = "NO" : a(12) = "NO" : a(13) = "NO" : a(14) = "NO" : a(15) = "NO" : a(16) = "NO" : a(17) = "NO" : a(18) = "NO"
                    Dim media_fld As String = Class1.HyperlaunchPath + "\Media\"
                    If FileSystem.DirectoryExists(media_fld + "Artwork\" + param.sys + "\" + romName) Then
                        If FileSystem.GetFiles(media_fld + "Artwork\" + param.sys + "\" + romName, SearchOption.SearchAllSubDirectories, {"*.png"}).Count > 0 Then a(11) = "YES"
                    End If
                    If FileSystem.DirectoryExists(media_fld + "Backgrounds\" + param.sys + "\" + romName) Then
                        If FileSystem.GetFiles(media_fld + "Backgrounds\" + param.sys + "\" + romName, SearchOption.SearchAllSubDirectories, {romName + ".png", "Layer 1*.png"}).Count > 0 Then a(12) = "YES"
                    End If
                    If FileSystem.DirectoryExists(media_fld + "Bezels\" + param.sys + "\" + romName) Then
                        If FileSystem.GetFiles(media_fld + "Bezels\" + param.sys + "\" + romName, SearchOption.SearchAllSubDirectories, {"Bezel.png"}).Count > 0 Then a(13) = "YES"
                    End If
                    If FileSystem.DirectoryExists(media_fld + "Fade\" + param.sys + "\" + romName) Then
                        If FileSystem.GetFiles(media_fld + "Fade\" + param.sys + "\" + romName, SearchOption.SearchAllSubDirectories, {"Layer 1*.png", "Layer 1*.jpg"}).Count > 0 Then a(14) = "YES"
                    End If
                    If FileSystem.DirectoryExists(media_fld + "Guides\" + param.sys + "\" + romName) Then
                        If FileSystem.GetFiles(media_fld + "Guides\" + param.sys + "\" + romName, SearchOption.SearchAllSubDirectories, {"*.*"}).Count > 0 Then a(15) = "YES"
                    End If
                    If FileSystem.DirectoryExists(media_fld + "Manuals\" + param.sys + "\" + romName) Then
                        If FileSystem.GetFiles(media_fld + "Manuals\" + param.sys + "\" + romName, SearchOption.SearchAllSubDirectories, {"*.pdf", "*.txt"}).Count > 0 Then a(16) = "YES"
                    End If
                    If FileSystem.DirectoryExists(media_fld + "Music\" + param.sys + "\" + romName) Then
                        If FileSystem.GetFiles(media_fld + "Music\" + param.sys + "\" + romName, SearchOption.SearchAllSubDirectories, {romName + ".m3u"}).Count > 0 Then a(17) = "YES"
                    End If
                    If FileSystem.DirectoryExists(media_fld + "Videos\" + param.sys + "\" + romName) Then
                        If FileSystem.GetFiles(media_fld + "Videos\" + param.sys + "\" + romName, SearchOption.SearchAllSubDirectories, {"*.avi", "*.mp4"}).Count > 0 Then a(18) = "YES"
                    End If
                End If
            End If


            'Add row
            Dim r As New DataGridViewRow
            counters(0) += 1
            r.CreateCells(DataGridView1, {a(0), romName, a(2), a(3), a(4), a(5), a(6), a(7), a(8), a(9), a(10), a(11), a(12), a(13), a(14), a(15), a(16), a(17), a(18), a_crc, a_manufacturer, a_year, a_genre, a_rating})
            tempStr = ""
            For i As Integer = 2 To 18
                If a(i) = "YES" Then r.Cells(i).Style.BackColor = Class1.colorYES Else r.Cells(i).Style.BackColor = Class1.colorNO
                If i >= 11 And i <= 18 And Not CheckBox30.Checked Then r.Cells(i).Style.BackColor = Color.White
                tempStr = tempStr + a(i).Substring(0, 1)
            Next
            If useVidFromParent And a(3) = "YES" Then r.Cells(3).Style.BackColor = Class1.colorPAR
            If useThemeFromParent And a(4) = "YES" Then r.Cells(4).Style.BackColor = Class1.colorPAR
            Class1.data.Add(tempStr)
            Class1.data_crc.Add(a_crc.PadLeft(8, "0"c).ToUpper)
            DataGridView1.BeginInvoke(Sub() DataGridView1.Rows.Add(r))
            progress += 1 : If progress Mod 50 = 0 Then bg_check.ReportProgress(progress)
        Next

        'Update stat labels
        Me.BeginInvoke(Sub() Label_Stat0.Text = "Total: " + counters(0).ToString)
        Dim arr_labels = {Label_Stat1, Label_Stat2, Label_Stat3, Label_Stat4, Label_Stat5, Label_Stat6, Label_Stat7}
        Dim arr_labels_txt = {"Roms", "Videos", "Artwork1", "Artwork2", "Artwork3", "Artwork4", "Wheel"}
        Dim arr_counter_indices = {1, 2, 5, 6, 7, 8, 4}
        For i As Integer = 0 To arr_labels.Length - 1
            Dim ind = i
            Me.BeginInvoke(Sub() arr_labels(ind).Text = arr_labels_txt(ind) + ": " + counters(arr_counter_indices(ind)).ToString)
            If counters(arr_counter_indices(ind)) = 0 Then
                Me.BeginInvoke(Sub() arr_labels(ind).BackColor = Class1.colorNO)
            ElseIf counters(arr_counter_indices(ind)) = counters(0) Then
                Me.BeginInvoke(Sub() arr_labels(ind).BackColor = Class1.colorYES)
            Else
                Me.BeginInvoke(Sub() arr_labels(ind).BackColor = Color.Yellow)
            End If
        Next

        'save lastCheckResult
        ini.path = Class1.confPath
        ini.IniWriteValue("LastCheckResult", param.sys, String.Join(",", counters))
    End Sub
    Private Sub check_bg_complete() Handles bg_check.RunWorkerCompleted
        ProgressBar1.Value = 0
        Button2_moveUnneeded.Enabled = True
        Button5_Associate.Enabled = True
        Button4.Enabled = True
        Button18.Enabled = True
        If AlowEditToolStripMenuItem.Checked Then Button21.Enabled = True
        Label2.Text = "Total: " + DataGridView1.Rows.Count.ToString
    End Sub
    Private Sub check_bg_progress(o As Object, e As ProgressChangedEventArgs) Handles bg_check.ProgressChanged
        ProgressBar1.Value = e.ProgressPercentage
    End Sub

    Private Function tryToFindRom(ByVal temppath As String, ByVal romExtensions() As String, ByVal romname As String) As Boolean
        Dim result As Boolean = False
        For Each ext As String In romExtensions
            If Dir(temppath + romname + "." + ext.Trim) <> "" Then result = True : Exit For
            'If FileSystem.FileExists(temppath + romname + "." + ext.Trim) Then result = True : Exit For
        Next
        If result Then Return True

        result = False
        For Each ext As String In romExtensions
            Try
                If Dir(temppath + "\" + romname + "\" + romname + "." + ext.Trim) <> "" Then result = True : Exit For
                'If FileSystem.FileExists(temppath + "\" + romname + "\" + romname + "." + ext.Trim) Then result = True : Exit For
            Catch ex As Exception
            End Try
        Next
        If result Then Return True
        Return False
    End Function

    'System select
    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Tag.ToString = "REFRESHING" Then Exit Sub

        Button2_moveUnneeded.Enabled = False
        Button4.Enabled = False
        Button5_Associate.Enabled = False
        Button18.Enabled = False
        Button21.Enabled = False
        Class1.HyperspinIniCursysEmuExe = ""
        Class1.HyperspinIniCursysEmuPath = ""
        Class1.HyperspinIniCursysEmuExist = False

        Label_Stat0.BackColor = Color.Transparent : Label_Stat0.Text = "Total: N/A"
        Label_Stat1.BackColor = Color.Transparent : Label_Stat1.Text = "Roms: N/A"
        Label_Stat2.BackColor = Color.Transparent : Label_Stat2.Text = "Videos: N/A"
        Label_Stat3.BackColor = Color.Transparent : Label_Stat3.Text = "Artwork1: N/A"
        Label_Stat4.BackColor = Color.Transparent : Label_Stat4.Text = "Artwork2: N/A"
        Label_Stat5.BackColor = Color.Transparent : Label_Stat5.Text = "Artwork3: N/A"
        Label_Stat6.BackColor = Color.Transparent : Label_Stat6.Text = "Artwork4: N/A"
        Label_Stat7.BackColor = Color.Transparent : Label_Stat7.Text = "Wheel: N/A"


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
            If s.StartsWith("path", StringComparison.InvariantCultureIgnoreCase) And category.ToUpper.Contains("exe info") Then
                Class1.HyperspinIniCursysEmuPath = s.Substring(s.IndexOf("=") + 1)
            End If
            If s.StartsWith("exe", StringComparison.InvariantCultureIgnoreCase) And category.ToUpper.Contains("exe info") Then
                Class1.HyperspinIniCursysEmuExe = s.Substring(s.IndexOf("=") + 1)
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

        If FileSystem.FileExists(Class1.HyperspinIniCursysEmuPath + "\" + Class1.HyperspinIniCursysEmuExe) Then Class1.HyperspinIniCursysEmuExist = True
        retrieve_rom_path_from_HL(False)

        If CheckBox22.Checked Then CheckBox23.Checked = useParentVids
        If CheckBox22.Checked Then CheckBox24.Checked = useParentThemes
        'Button28.Text = "Remove all clones from current database" + vbCrLf + "(" + ComboBox1.SelectedItem.ToString + " selected)"
        RemoveClonesFromCurrentDBToolStripMenuItem.Text = "Remove all clones from current database (" + ComboBox1.SelectedItem.ToString + " selected)"
    End Sub

    Private Sub retrieve_rom_path_from_HL(Optional messages As Boolean = True)
        'HLv3 Thing
        Dim s As String = ""
        If CheckBox26.Checked And ComboBox1.SelectedIndex >= 0 Then
            'Dim HLPath As String = TextBox18.Text
            Dim HLPath As String = Class1.HyperlaunchPath
            'If HLPath.ToUpper.EndsWith("EXE") Then HLPath = HLPath.Substring(0, HLPath.LastIndexOf("\"))
            'FileOpen(1, Class1.HyperspinPath + "\Settings\Settings.ini", OpenMode.Input)
            'Do While Not EOF(1)
            's = LineInput(1).Trim
            'If s.StartsWith("hyperlaunch_path", StringComparison.InvariantCultureIgnoreCase) Then
            'HLPath = s.Substring(s.IndexOf("=") + 1).Trim : Exit Do
            'End If
            'Loop
            'FileClose(1)
            If HLPath = "" Then
                If messages Then MsgBox("Hyperlaunch_path in Settings.ini is not set. Rom path from HS .ini will be used.")
                Exit Sub
            End If

            If HLPath.StartsWith(".") Then
                If messages Then MsgBox("Relative hyperlaunch_path in Settings.ini is not supported. Rom path from HS .ini will be used.")
                Exit Sub
            End If


            If Not HLPath.EndsWith("\") Then HLPath = HLPath + "\"
            Dim HLiniFile As String = HLPath + "Settings\" + ComboBox1.SelectedItem.ToString + "\Emulators.ini"
            If Not Microsoft.VisualBasic.FileIO.FileSystem.FileExists(HLiniFile) Then
                If messages Then MsgBox("HL system settings file: '" & HLiniFile & "' not found. Rom path from HS .ini will be used.")
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
                If messages Then MsgBox("Rom_Path not set in '" & HLiniFile & "'. Rom path from HS .ini will be used.")
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
        If e.ClickedItem Is myContextMenu.Items(9) Then Class1.i = 2
        Form2_checkMissingInOtherFolders.Show()
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
        If mediaID = 7 Then exclusionList.Add("DEFAULT")
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
                If rom.Contains(".") Then
                    romWoExt = rom.Substring(0, rom.LastIndexOf("."))
                Else
                    romWoExt = rom
                End If
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

#Region "System properties and Program Settings"
    'HS path changing
    Private Sub TextBox14_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox14.TextChanged
        Class1.data.Clear()
        Class1.data_crc.Clear()
        Class1.romlist.Clear()
        Class1.romFoundlist.Clear()
        ComboBox1.Items.Clear()
        DataGridView1.Rows.Clear()
        TextBox14_TextChanged_sub_check()
    End Sub
    Public Sub TextBox14_TextChanged_sub_check()
        If FileSystem.DirectoryExists(TextBox14.Text) Then
            Class1.HyperspinPath = TextBox14.Text
            If Not Class1.HyperspinPath.EndsWith("\") Then Class1.HyperspinPath = Class1.HyperspinPath + "\"

            If Not FileSystem.FileExists(Class1.HyperspinPath + "\Databases\Main Menu\Main Menu.xml") Then
                Label23.BackColor = Color.Orange
                Label23.Text = "Can't find \Databases\Main Menu\Main Menu.xml." : Exit Sub
            Else
                Label23.BackColor = Color.LightGreen
                Label23.Text = "\Databases\Main Menu\Main Menu.xml FOUND"
            End If

            Dim x As New Xml.XmlDocument
            'ComboBox1.BeginUpdate()
            system_list.Clear()
            x.Load(Class1.HyperspinPath + "\Databases\Main Menu\Main Menu.xml")
            For Each node As Xml.XmlNode In x.SelectNodes("/menu/game")
                system_list.Rows.Add({node.Attributes.GetNamedItem("name").Value})
                ComboBox1.Items.Add(node.Attributes.GetNamedItem("name").Value)
            Next
            'ComboBox1.DataSource = system_list
            'ComboBox1.DisplayMember = "system"
            'ComboBox1.ValueMember = "system"
            'ComboBox1.EndUpdate()

            'Get hyperlaunch path from settings.ini
            If Not CheckBox21.Checked AndAlso FileSystem.FileExists(Class1.HyperspinPath + "\Settings\Settings.ini") Then
                ini.path = Class1.HyperspinPath + "\Settings\Settings.ini"
                TextBox18.Text = ini.IniReadValue("main", "Hyperlaunch_Path")
            End If
        Else
            Label23.BackColor = Color.Orange
            Label23.Text = "Directory not exist." : Exit Sub
        End If
    End Sub

    'Hyperlaunch path changed
    Private Sub TextBox18_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox18.TextChanged
        Dim ok As Boolean = False
        If TextBox18.Text.StartsWith(".") Then TextBox18.Text = IO.Path.GetFullPath(Class1.HyperspinPath + "\" + TextBox18.Text) : Exit Sub
        Dim str As String = TextBox18.Text


        If str.ToUpper.EndsWith("EXE") Then
            Class1.HyperlaunchPath = str.Substring(0, str.LastIndexOf("\"))
            Class1.HyperlaunchExeName = str.Substring(str.LastIndexOf("\") + 1)
            If FileSystem.FileExists(str) Then
                ok = True
            End If
        Else
            Class1.HyperlaunchPath = str
            Class1.HyperlaunchExeName = ""
            If FileSystem.DirectoryExists(str) Then
                If FileSystem.FileExists(str + "\HyperLaunch.exe") Or FileSystem.FileExists(str + "\RocketLauncher.exe") Then
                    ok = True
                End If
            End If
        End If
        If ok Then TextBox18.BackColor = Class1.colorYES Else TextBox18.BackColor = Class1.colorNO
    End Sub

    'Save Config
    Private Sub Button20_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click 'Save Config
        ini.IniFile(Class1.confPath)
        ini.IniWriteValue("MAIN", "HyperSpin_Path", TextBox14.Text)
        ini.IniWriteValue("MAIN", "rename_just_cue", DirectCast(IIf(CheckBox6.Checked, "1", "0"), String))
        ini.IniWriteValue("MAIN", "search_cue_for", TextBox16.Text)
        ini.IniWriteValue("MAIN", "useHLv3", DirectCast(IIf(CheckBox26.Checked, "1", "0"), String))
        ini.IniWriteValue("MAIN", "sort_systems", DirectCast(IIf(CheckBox32.Checked, "1", "0"), String))
        ini.IniWriteValue("MAIN", "check_hl_game_media", DirectCast(IIf(CheckBox30.Checked, "1", "0"), String))
        ini.IniWriteValue("MAIN", "check_hl_system_media", DirectCast(IIf(CheckBox31.Checked, "1", "0"), String))
        ini.IniWriteValue("MAIN", "archives_rename_inside", DirectCast(IIf(CheckBox28.Checked, "1", "0"), String))
        ini.IniWriteValue("MAIN", "archives_remove_unneeded", DirectCast(IIf(CheckBox29.Checked, "1", "0"), String))
        ini.IniWriteValue("3rd_Party_Tools", "UltraISO", TextBox28.Text.Trim)
        For i As Integer = 1 To ListBox4.Items.Count
            ini.IniWriteValue("handle_together", i.ToString, ListBox4.Items(i - 1).ToString)
        Next
        If CheckBox21.Checked Then
            ini.IniWriteValue("Main", "Freeze_HL_path", TextBox18.Text)
        Else
            ini.IniWriteValue("Main", "Freeze_HL_path", "")
        End If

        'Compression
        ini.IniWriteValue("Main", "Compression_type", ComboBox10.SelectedItem.ToString)
        ini.IniWriteValue("Main", "Compression_method", ComboBox11.SelectedItem.ToString)
        ini.IniWriteValue("Main", "Compression_level", ComboBox12.SelectedItem.ToString)
        ini.IniWriteValue("Main", "Compression_temp", TextBox30.Text)

        'Matcher font size
        ini.IniWriteValue("Main", "MatcherFontSize", ListBox1.Font.Size.ToString)

        'Language
        If ComboBox13.SelectedIndex = 0 Then
            ini.IniWriteValue("Main", "Language", "")
        Else
            ini.IniWriteValue("Main", "Language", ComboBox13.SelectedItem.ToString)
        End If

        'Top menu - Matcher options
        For Each t As ToolStripItem In AssocOption_fileInHsFolder.DropDownItems
            If DirectCast(t, ToolStripMenuItem).Checked Then ini.IniWriteValue("Main", "AssocOption_fileInHsFolder", t.Name)
        Next
        For Each t As ToolStripItem In AssocOption_fileInDiffFolder.DropDownItems
            If DirectCast(t, ToolStripMenuItem).Checked Then ini.IniWriteValue("Main", "AssocOption_fileInDiffFolder", t.Name)
        Next

        Label23.BackColor = Color.LightBlue
        Label23.Text = "Config.conf Saved"
    End Sub

    'add extension to PAIR_EXTENSION list
    Private Sub Button5_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If TextBox15.Text.Contains(",") Then
            ListBox4.Items.Add(TextBox15.Text)
            TextBox15.Text = ""
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

    'Change checkbox 'use HS settings for clones/parents' 
    Private Sub CheckBox22_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox22.CheckedChanged
        If CheckBox22.Checked Then
            CheckBox23.Enabled = False : CheckBox24.Enabled = False
            CheckBox23.Checked = useParentVids : CheckBox24.Checked = useParentThemes
        Else
            CheckBox23.Enabled = True : CheckBox24.Enabled = True
        End If
    End Sub

    'Update Hyperspin INI
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If Not TextBox18.Text.ToUpper.EndsWith("EXE") Then
            If Not TextBox18.Text.EndsWith("\") Then TextBox18.Text = TextBox18.Text + "\"
        End If
        TextBox18.Text = TextBox18.Text.Replace("\\", "\").Replace("\\", "\")
        ini.path = Class1.HyperspinPath + "\Settings\Settings.ini"
        If IO.File.Exists(TextBox18.Text + "RocketLauncher.exe") Then
            ini.IniWriteValue("main", "Hyperlaunch_Path", TextBox18.Text + "RocketLauncher.exe")
        ElseIf IO.File.Exists(TextBox18.Text + "HyperLaunch.exe") Then
            ini.IniWriteValue("main", "Hyperlaunch_Path", TextBox18.Text)
        End If


        TextBox1.Text = TextBox1.Text.Trim
        TextBox3.Text = TextBox3.Text.Trim
        TextBox2.Text = TextBox2.Text.Trim
        If Not TextBox1.Text = "" AndAlso Not TextBox1.Text.EndsWith("\") Then TextBox1.Text = TextBox1.Text + "\"
        If Not TextBox2.Text = "" AndAlso Not TextBox2.Text.EndsWith("\") Then TextBox2.Text = TextBox2.Text + "\"

        If ComboBox1.SelectedIndex >= 0 Then
            ini.path = Class1.HyperspinPath + "\Settings\" + ComboBox1.SelectedItem.ToString + ".ini"
            ini.IniWriteValue("exe info", "rompath", TextBox1.Text)
            ini.IniWriteValue("exe info", "romextension", TextBox3.Text)
            ini.IniWriteValue("video defaults", "path", TextBox2.Text)
        End If
    End Sub

    'Update UltraISO path
    Private Sub TextBox28_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox28.TextChanged
        If FileSystem.FileExists(TextBox28.Text) Then TextBox28.BackColor = Class1.colorYES Else TextBox28.BackColor = Class1.colorNO
    End Sub

    'Language change
    Private Sub ComboBox13_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox13.SelectedIndexChanged
        Dim iniLang As New IniFileApi
        iniLang.path = ".\Localization\" + ComboBox13.SelectedItem.ToString + ".ini"

        Language.strings.Clear()
        For Each k In iniLang.IniListKey("Main")
            Dim val As String = iniLang.IniReadValue("Main", k).Replace("{CRLF}", vbCrLf)
            Language.strings.Add(k, val)
        Next

        Language.localize(Me)

        'For Each k In iniLang.IniListKey("Main")
        '    'For Each k In iniLang.IniListKey("Main")
        '    Dim s() = k.Split({"."c})
        '    Dim val As String = iniLang.IniReadValue("Main", k).Replace("{CRLF}", vbCrLf)

        '    Dim tmp As Control = Nothing
        '    Dim tmp_menu_item As ToolStripItem = Nothing
        '    For Each c In s
        '        If c.ToUpper = "FORM1" Then
        '            tmp = Me
        '        ElseIf c.ToUpper = "FORM2_CHECKMISSINGINOTHERFOLDERS" Then
        '            tmp = Form2_checkMissingInOtherFolders
        '        ElseIf c.ToUpper = "FORM3" Then
        '            tmp = Form3_mameFoldersFilter
        '        ElseIf c.ToUpper = "Form4_genres_favorites".ToUpper Then
        '            tmp = Form4_genres_favorites
        '        ElseIf c.ToUpper = "Form4_genres_favorites_preview".ToUpper Then
        '            tmp = Form4_genres_favorites_preview
        '        ElseIf c.ToUpper = "Form5_videoDownloader".ToUpper Then
        '            tmp = Form5_videoDownloader
        '        ElseIf c.ToUpper = "Form6_autorenamer".ToUpper Then
        '            tmp = Form6_autorenamer
        '        ElseIf c.ToUpper = "Form7_dualFolderOperations".ToUpper Then
        '            tmp = Form7_dualFolderOperations
        '        ElseIf c.ToUpper = "Form8_systemProperties".ToUpper Then
        '            tmp = Form8_systemProperties
        '        ElseIf c.ToUpper = "Form9_database_statistic".ToUpper Then
        '            tmp = Form9_database_statistic
        '        ElseIf c.ToUpper = "FormA_hyperlaunch_3rd_party_paths".ToUpper Then
        '            tmp = FormA_hyperlaunch_3rd_party_paths
        '        ElseIf c.ToUpper = "FormB_PCSX2_createIndex".ToUpper Then
        '            tmp = FormB_PCSX2_createIndex
        '        ElseIf c.ToUpper = "FormC_mameRomListBuilder".ToUpper Then
        '            tmp = FormC_mameRomListBuilder
        '        ElseIf c.ToUpper = "FormD_matcher_autofilter_constructor".ToUpper Then
        '            tmp = FormD_matcher_autofilter_constructor
        '        ElseIf c.ToUpper = "FormE_Create_database_XML_from_folder".ToUpper Then
        '            tmp = FormE_Create_database_XML_from_folder
        '        ElseIf c.ToUpper = "FormF_createNewHL_system".ToUpper Then
        '            tmp = FormF_createNewHL_system
        '        ElseIf c.ToUpper = "FormF_systemManager_addSystem".ToUpper Then
        '            tmp = FormF_systemManager_addSystem
        '        ElseIf c.ToUpper = "FormF_systemManager_exclusions".ToUpper Then
        '            tmp = FormF_systemManager_exclusions
        '        ElseIf c.ToUpper = "FormG_associationTables".ToUpper Then
        '            tmp = FormG_associationTables
        '        Else
        '            If tmp_menu_item IsNot Nothing Then
        '                tmp_menu_item = DirectCast(tmp_menu_item, ToolStripMenuItem).DropDownItems(c)
        '            Else
        '                If TypeOf tmp Is MenuStrip Then
        '                    tmp_menu_item = DirectCast(tmp, MenuStrip).Items(c)
        '                Else
        '                    tmp = tmp.Controls(c)
        '                End If
        '            End If

        '            If tmp.Name.ToUpper = "SPLITCONTAINER1" Then
        '                Dim tmp2 = DirectCast(tmp, SplitContainer)
        '                tmp2.Panel1.Name = "Panel1"
        '                tmp2.Panel2.Name = "Panel2"
        '            End If
        '        End If
        '    Next

        '    If tmp_menu_item IsNot Nothing Then
        '        tmp_menu_item.Text = val
        '    Else
        '        tmp.Text = val
        '    End If
        'Next
    End Sub

    'Matcher lists font size
    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged
        Dim v As Single = CSng(NumericUpDown1.Value)
        If v = 8 Then v = 8.25
        ListBox1.Font = New Font(ListBox1.Font.FontFamily, v)
        ListBox2.Font = New Font(ListBox2.Font.FontFamily, v)
    End Sub

    'Sort system alphabetically
    Private Sub CheckBox32_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox32.CheckedChanged
        Dim cur_sys = ""
        If ComboBox1.SelectedItem IsNot Nothing Then cur_sys = ComboBox1.SelectedItem.ToString

        ComboBox1.Sorted = CheckBox32.Checked

        'Refresh system list, if going from sorted to non_sorted
        If Not CheckBox32.Checked Then
            ComboBox1.Items.Clear()
            For Each r As DataRow In system_list.Rows
                ComboBox1.Items.Add(r.Item(0))
            Next
        End If

        If cur_sys <> "" Then ComboBox1.SelectedItem = cur_sys
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
    Private Sub FilterColumns_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles F0.Click, F1.Click, F2.Click, F3.Click, F4.Click, F5.Click, F6.Click, F7.Click, F8.Click, F9.Click, F10.Click, F19.Click, F20.Click, F21.Click, F22.Click
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
        'If ListBox1.Items.Count = 1 Then
        'We don't use this "no missing" message any more
        '    If DirectCast(ListBox1.Items(0), DataRowView).Item(0).ToString.StartsWith("No missing") Then MsgBox(ListBox1.Items(0).ToString + " found. Nothing to do.") : Exit Sub
        'End If
        If ListBox1.Items.Count = 0 Then MsgBox("No missing media found. Nothing to do.") : Exit Sub
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
    'Macher / Show filters
    Private Sub ShowFiltersToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ShowFiltersToolStripMenuItem.Click, Button29.Click
        If ShowFiltersToolStripMenuItem.Checked = False Then ShowFiltersToolStripMenuItem.Checked = True Else ShowFiltersToolStripMenuItem.Checked = False

        If ShowFiltersToolStripMenuItem.Checked Then
            TextBox26.Visible = True
            TextBox27.Visible = True
            ListBox1.Height = ListBox1.Height - 25
            ListBox2.Height = ListBox2.Height - 25
        Else
            TextBox26.Text = ""
            TextBox27.Text = ""
            TextBox26.Visible = False
            TextBox27.Visible = False
            ListBox1.Height = ListBox1.Height + 25
            ListBox2.Height = ListBox2.Height + 25
        End If
    End Sub
    'Macher / Autofilter
    Private Sub AutofilterToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AutofilterToolStripMenuItem.Click, AutofilterRomDBToolStripMenuItem.Click
        Dim t = DirectCast(sender, ToolStripMenuItem)

        If t.Checked = False Then
            AutofilterToolStripMenuItem.Checked = False
            AutofilterRomDBToolStripMenuItem.Checked = False
            t.Checked = True

            CheckBox33.Checked = AutofilterToolStripMenuItem.Checked
            CheckBox34.Checked = AutofilterRomDBToolStripMenuItem.Checked
        Else
            t.Checked = False
            CheckBox33.Checked = False
            CheckBox34.Checked = False
        End If
    End Sub
    'Macher / Autofilter regex constructor
    Private Sub AutofilterRegexConstructorToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AutofilterRegexConstructorToolStripMenuItem.Click
        If FormD_matcher_autofilter_constructor.Visible Then Exit Sub
        Dim f As New FormD_matcher_autofilter_constructor
        f.ShowDialog(Me)

        'Resave current regex and preset to ini
        ini.IniFile(Class1.confPath)
        ini.IniWriteValue("RegEx", "active", Class3_matcher.autofilter_regex)
        ini.IniWriteValue("RegEx", "opt_strip_br", Class3_matcher.autofilter_regex_options(0).ToString.ToUpper)
        ini.IniWriteValue("RegEx", "opt_strip_pr", Class3_matcher.autofilter_regex_options(1).ToString.ToUpper)
        ini.IniWriteValue("RegEx", "opt_out_group", Class3_matcher.autofilter_regex_opt_outGroup.ToString)
        Dim c As Integer = 1
        For Each k In Class3_matcher.autofilter_regex_presets
            ini.IniWriteValue("RegEx", "preset" + c.ToString, k.Key + "^^^" + k.Value)
            c += 1
        Next
    End Sub
    'Undo History
    Private Sub UndoHistoryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UndoHistoryToolStripMenuItem.Click
        Dim frm As New FormH_undoHistory
        frm.Show()
    End Sub

    'Tools / Association tables
    Private Sub AssociationTablesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AssociationTablesToolStripMenuItem.Click
        Dim frm As New FormG_associationTables : frm.Show()
    End Sub
    'Tools / Show HL 3rd party paths
    Private Sub CheckHyperLaunch3rdPartyPathsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckHyperLaunch3rdPartyPathsToolStripMenuItem.Click
        If TextBox18.BackColor <> Class1.colorYES Then MsgBox("You Hyperlaunch path is not correctly set. Check 'Hyperspin system settings' tab") : Exit Sub
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
    'Tools / Pack System
    Private Sub PackSystemToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PackSystemToolStripMenuItem.Click
        Dim f As New FormJa_PackSystem
        f.Show(Me)
    End Sub
    'Tools / UnPack System
    Private Sub UnpackSystemToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UnpackSystemToolStripMenuItem.Click
        Dim f As New FormJb_UnPackSystem
        f.Show(Me)
    End Sub

    'About
    Private Sub AboutToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem1.Click
        Dim a As New FormZ_about
        a.ShowDialog(Me)
    End Sub
#End Region

#Region "Table headers, columns show/hide and painting"
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
            Form3_mameFoldersFilter.ShowDialog()
            If Form3_mameFoldersFilter.filter.Count = 0 Then MsgBox("No games selected") : ComboBox4.SelectedIndex = oldCombo4Value : Exit Sub

            oldCombo4Value = ComboBox4.SelectedIndex
            ProgressBar1.Value = 0
            ProgressBar1.Maximum = DataGridView1.Rows.Count
            Label2.Text = "Filtering..." : Label2.Refresh()
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                DataGridView1.Rows(i).Visible = True
                If Form3_mameFoldersFilter.filter.Contains(DataGridView1.Rows(i).Cells(1).Value.ToString.ToLower) Then
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

    'Main Table - Column Headers context menu (show / hide columns and presets)
    Private Sub DataGridView1_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.ColumnHeaderMouseClick
        If e.Button <> Windows.Forms.MouseButtons.Right Then Exit Sub
        myContextMenu4.Show(Cursor.Position.X, Cursor.Position.Y)
    End Sub

    'System Manager - Column Headers context menu (show / hide columns and presets)
    Private Sub DataGridView2_ColumnHeaderMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView2.ColumnHeaderMouseClick
        If e.Button <> Windows.Forms.MouseButtons.Right Then Exit Sub
        myContextMenu8.Show(Cursor.Position.X, Cursor.Position.Y)
    End Sub

    'Main Table - Show/hide columns checked_change
    Private Sub ShowHideCol(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
        CheckStrip2_0.CheckedChanged, CheckStrip2_1.CheckedChanged, CheckStrip2_2.CheckedChanged,
        CheckStrip2_3.CheckedChanged, CheckStrip2_4.CheckedChanged, CheckStrip2_5.CheckedChanged,
        CheckStrip2_6.CheckedChanged, CheckStrip2_7.CheckedChanged, CheckStrip2_8.CheckedChanged,
        CheckStrip2_9.CheckedChanged, CheckStrip2_10.CheckedChanged, CheckStrip2_11.CheckedChanged,
        CheckStrip2_12.CheckedChanged, CheckStrip2_13.CheckedChanged, CheckStrip2_14.CheckedChanged,
        CheckStrip2_15.CheckedChanged, CheckStrip2_16.CheckedChanged, CheckStrip2_17.CheckedChanged,
        CheckStrip2_18.CheckedChanged, CheckStrip2_19.CheckedChanged, CheckStrip2_20.CheckedChanged,
        CheckStrip2_21.CheckedChanged, CheckStrip2_22.CheckedChanged

        Dim cb As CheckBox = DirectCast(sender, CheckBox)
        Dim i As Integer = CInt(cb.Name)
        If ShowHideColumnsToolStripMenuItem.DropDownItems("F" + i.ToString) IsNot Nothing Then
            DirectCast(ShowHideColumnsToolStripMenuItem.DropDownItems("F" + i.ToString), ToolStripMenuItem).Checked = cb.Checked
        Else
            DirectCast(ShowHideColumnsToolStripMenuItemHL.DropDownItems("F" + i.ToString), ToolStripMenuItem).Checked = cb.Checked
        End If
        DataGridView1.Columns(i).Visible = cb.Checked
    End Sub

    'System Manager - Show/hide columns checked_change
    Private Sub ShowHideCol_sysMgr(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckStrip3_1.CheckedChanged,
        CheckStrip3_2.CheckedChanged, CheckStrip3_3.CheckedChanged, CheckStrip3_4.CheckedChanged, CheckStrip3_5.CheckedChanged,
        CheckStrip3_6.CheckedChanged, CheckStrip3_7.CheckedChanged, CheckStrip3_8.CheckedChanged, CheckStrip3_9.CheckedChanged,
        CheckStrip3_10.CheckedChanged, CheckStrip3_11.CheckedChanged, CheckStrip3_12.CheckedChanged, CheckStrip3_13.CheckedChanged,
        CheckStrip3_14.CheckedChanged, CheckStrip3_15.CheckedChanged, CheckStrip3_16.CheckedChanged, CheckStrip3_17.CheckedChanged,
        CheckStrip3_18.CheckedChanged, CheckStrip3_19.CheckedChanged, CheckStrip3_20.CheckedChanged, CheckStrip3_21.CheckedChanged,
        CheckStrip3_22.CheckedChanged

        Dim cb As CheckBox = DirectCast(sender, CheckBox)
        Dim i As Integer = CInt(cb.Name)
        DataGridView2.Columns(i).Visible = cb.Checked
    End Sub

    'Main Table - Column headers presets
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
            CheckStrip2_15.Checked = False
            CheckStrip2_16.Checked = False
            CheckStrip2_17.Checked = False
            CheckStrip2_18.Checked = False
            CheckStrip2_19.Checked = False
            CheckStrip2_20.Checked = False
            CheckStrip2_21.Checked = False
            CheckStrip2_22.Checked = False
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
            CheckStrip2_11.Checked = False
            CheckStrip2_12.Checked = False
            CheckStrip2_13.Checked = False
            CheckStrip2_14.Checked = False
            CheckStrip2_15.Checked = False
            CheckStrip2_16.Checked = False
            CheckStrip2_17.Checked = False
            CheckStrip2_18.Checked = False
            CheckStrip2_19.Checked = True
            CheckStrip2_20.Checked = True
            CheckStrip2_21.Checked = True
            CheckStrip2_22.Checked = True
        End If
        If e.ClickedItem.Text = Preset_HL Then
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
            CheckStrip2_15.Checked = True
            CheckStrip2_16.Checked = True
            CheckStrip2_17.Checked = True
            CheckStrip2_18.Checked = True
            CheckStrip2_19.Checked = False
            CheckStrip2_20.Checked = False
            CheckStrip2_21.Checked = False
            CheckStrip2_22.Checked = False
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

    'System Manager - Column headers presets
    Private Sub ShowHidePresets_sysMgr(ByVal sender As Object, ByVal e As ToolStripItemClickedEventArgs) Handles myContextMenu8.ItemClicked
        If e.ClickedItem.Text = Preset_SysMngr_Default Then
            CheckStrip3_1.Checked = True
            CheckStrip3_2.Checked = True
            CheckStrip3_3.Checked = True
            CheckStrip3_4.Checked = True
            CheckStrip3_5.Checked = True
            CheckStrip3_6.Checked = True
            CheckStrip3_7.Checked = True
            CheckStrip3_8.Checked = False
            CheckStrip3_9.Checked = False
            CheckStrip3_10.Checked = False
            CheckStrip3_11.Checked = False
            CheckStrip3_12.Checked = False
            CheckStrip3_13.Checked = False
            CheckStrip3_14.Checked = False
            CheckStrip3_15.Checked = False
            CheckStrip3_16.Checked = False
            CheckStrip3_17.Checked = False
            CheckStrip3_18.Checked = False
            CheckStrip3_19.Checked = False
            CheckStrip3_20.Checked = False
            CheckStrip3_21.Checked = False
            CheckStrip3_22.Checked = False
        End If
        If e.ClickedItem.Text = Preset_SysMngr_Checker Then
            CheckStrip3_1.Checked = True
            CheckStrip3_2.Checked = True
            CheckStrip3_3.Checked = True
            CheckStrip3_4.Checked = True
            CheckStrip3_5.Checked = False
            CheckStrip3_6.Checked = False
            CheckStrip3_7.Checked = False
            CheckStrip3_8.Checked = True
            CheckStrip3_9.Checked = True
            CheckStrip3_10.Checked = False
            CheckStrip3_11.Checked = False
            CheckStrip3_12.Checked = False
            CheckStrip3_13.Checked = False
            CheckStrip3_14.Checked = False
            CheckStrip3_15.Checked = False
            CheckStrip3_16.Checked = False
            CheckStrip3_17.Checked = False
            CheckStrip3_18.Checked = False
            CheckStrip3_19.Checked = False
            CheckStrip3_20.Checked = False
            CheckStrip3_21.Checked = False
            CheckStrip3_22.Checked = False
        End If
        If e.ClickedItem.Text = Preset_SysMngr_Manager Then
            CheckStrip3_1.Checked = True
            CheckStrip3_2.Checked = True
            CheckStrip3_3.Checked = True
            CheckStrip3_4.Checked = False
            CheckStrip3_5.Checked = True
            CheckStrip3_6.Checked = True
            CheckStrip3_7.Checked = True
            CheckStrip3_8.Checked = False
            CheckStrip3_9.Checked = False
            CheckStrip3_10.Checked = False
            CheckStrip3_11.Checked = False
            CheckStrip3_12.Checked = False
            CheckStrip3_13.Checked = False
            CheckStrip3_14.Checked = False
            CheckStrip3_15.Checked = False
            CheckStrip3_16.Checked = False
            CheckStrip3_17.Checked = False
            CheckStrip3_18.Checked = False
            CheckStrip3_19.Checked = False
            CheckStrip3_20.Checked = False
            CheckStrip3_21.Checked = False
            CheckStrip3_22.Checked = False
        End If
        If e.ClickedItem.Text = Preset_SysMngr_HLMedia Then
            CheckStrip3_1.Checked = False
            CheckStrip3_2.Checked = False
            CheckStrip3_3.Checked = False
            CheckStrip3_4.Checked = False
            CheckStrip3_5.Checked = False
            CheckStrip3_6.Checked = False
            CheckStrip3_7.Checked = False
            CheckStrip3_8.Checked = False
            CheckStrip3_9.Checked = False
            CheckStrip3_10.Checked = True
            CheckStrip3_11.Checked = True
            CheckStrip3_12.Checked = True
            CheckStrip3_13.Checked = True
            CheckStrip3_14.Checked = True
            CheckStrip3_15.Checked = True
            CheckStrip3_16.Checked = True
            CheckStrip3_17.Checked = True
            CheckStrip3_18.Checked = False
            CheckStrip3_19.Checked = False
            CheckStrip3_20.Checked = False
            CheckStrip3_21.Checked = False
            CheckStrip3_22.Checked = False
        End If
        If e.ClickedItem.Text = Preset_SysMngr_Full Then
            CheckStrip3_1.Checked = True
            CheckStrip3_2.Checked = True
            CheckStrip3_3.Checked = True
            CheckStrip3_4.Checked = False
            CheckStrip3_5.Checked = True
            CheckStrip3_6.Checked = True
            CheckStrip3_7.Checked = True
            CheckStrip3_8.Checked = False
            CheckStrip3_9.Checked = False
            CheckStrip3_10.Checked = True
            CheckStrip3_11.Checked = True
            CheckStrip3_12.Checked = True
            CheckStrip3_13.Checked = True
            CheckStrip3_14.Checked = True
            CheckStrip3_15.Checked = True
            CheckStrip3_16.Checked = True
            CheckStrip3_17.Checked = True
            CheckStrip3_18.Checked = True
            CheckStrip3_19.Checked = True
            CheckStrip3_20.Checked = True
            CheckStrip3_21.Checked = True
            CheckStrip3_22.Checked = True
        End If
        If e.ClickedItem.Text = Save_current_cols_conf_as_startup Then
            ini.IniFile(Class1.confPath)
            Dim v As String = ""
            For i As Integer = 0 To DataGridView2.ColumnCount - 1
                If DataGridView2.Columns(i).Visible Then v = "1" Else v = "0"
                ini.IniWriteValue("System_Table_Columns_Config", "Col_" + i.ToString + "_visible", v)
                ini.IniWriteValue("System_Table_Columns_Config", "Col_" + i.ToString + "_width", DataGridView2.Columns(i).Width.ToString)
            Next
        End If
    End Sub

    'Cell context menu (start game / Fade preview)
    Dim romname_to_start As String = ""
    Dim context_menu_row As DataGridViewRow = Nothing
    Private Sub DataGridView1_CellMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        If e.Button <> Windows.Forms.MouseButtons.Right Then Exit Sub
        If e.RowIndex < 0 Then Exit Sub
        myContextMenu7.Items.Clear()
        romname_to_start = DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString
        myContextMenu7.Items.Add("Start '" + romname_to_start + "' using HyperSpin settings")
        myContextMenu7.Items.Add("Start '" + romname_to_start + "' using HyperLaunch settings")
        myContextMenu7.Items.Add("-")
        myContextMenu7.Items.Add("Rocket Launcher Fade Preview")

        context_menu_row = DataGridView1.Rows(e.RowIndex)
        If context_menu_row.Cells(2).Value.ToString.ToUpper <> "YES" Then
            myContextMenu7.Items(0).Enabled = False : myContextMenu7.Items(1).Enabled = False
        End If
        If Not Class1.HyperspinIniCursysEmuExist Then myContextMenu7.Items(0).Enabled = False

        myContextMenu7.Show(Cursor.Position.X, Cursor.Position.Y)
    End Sub

    'Cell context menu (start game / Fade preview) item click
    Private Sub contextMenuStartGame(ByVal sender As Object, ByVal e As ToolStripItemClickedEventArgs) Handles myContextMenu7.ItemClicked
        If myContextMenu7.Items.IndexOf(e.ClickedItem) = 0 Then
            'Start with HyperSpin settings
            'TODO
        ElseIf myContextMenu7.Items.IndexOf(e.ClickedItem) = 1 Then
            'Start with HyperLaunch settings
            Dim HLexe As String = "HyperLaunch.exe"
            If Class1.HyperlaunchExeName <> "" Then HLexe = Class1.HyperlaunchExeName
            Dim pStart As New ProcessStartInfo(Class1.HyperlaunchPath + "\" + HLexe, "-s """ + ComboBox1.SelectedItem.ToString + """ -r """ + romname_to_start + """")
            Process.Start(pStart)
        ElseIf myContextMenu7.Items.IndexOf(e.ClickedItem) = 3 Then
            Dim descr = context_menu_row.Cells(0).Value.ToString
            Dim publisher = context_menu_row.Cells(20).Value.ToString
            Dim year = context_menu_row.Cells(21).Value.ToString
            Dim genre = context_menu_row.Cells(22).Value.ToString
            Dim rating = context_menu_row.Cells(23).Value.ToString
            Dim f As New FormI_RLFadePreview(ComboBox1.SelectedItem.ToString, romname_to_start, {descr, year, genre, publisher, rating})
            f.Show()
        End If
    End Sub

    'Cell ctrl+c in edit mode - copy cell content to clipboard (instead of while line)
    Private Sub DataGridView1_KeyDown(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyDown
        If AlowEditToolStripMenuItem.Checked AndAlso DataGridView1.CurrentCell IsNot Nothing Then
            If e.Control AndAlso e.KeyCode = Keys.C Then
                Clipboard.SetText(DataGridView1.CurrentCell.Value.ToString)
                e.Handled = True
            ElseIf e.Control AndAlso (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Return) Then
                DataGridView1.BeginEdit(False)
                e.Handled = True
            End If
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

#Region "All the '...' buttons (open file browser), Notes, Autorenamer, Autofilter toggles"
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
    Private Sub Button19_Click(sender As System.Object, e As System.EventArgs) Handles Button19.Click
        Dim fd As New OpenFileDialog
        fd.Filter = "EXE Files|*.exe"
        fd.RestoreDirectory = True
        If TextBox28.Text.Trim <> "" Then fd.InitialDirectory = TextBox28.Text.Trim
        fd.ShowDialog()
        TextBox28.Text = fd.FileName
    End Sub

    'Set HL folder to \Hyperspin\Hyperlaunch
    Private Sub Button37_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button37.Click
        'TextBox18.Text = (Class1.HyperspinPath + "\Hyperlaunch\").Replace("\\", "\").Replace("\\", "\")
        TextBox18.Text = (Class1.HyperspinPath + "\RocketLauncher\").Replace("\\", "\").Replace("\\", "\")
    End Sub

    'Open Notes
    Private Sub Button30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button30.Click
        System.Diagnostics.Process.Start("notepad.exe", ".\notes.txt")
    End Sub

    'Iso Converter output to orig dir
    Private Sub CheckBox18_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox18.CheckedChanged
        TextBox29.Enabled = Not CheckBox18.Checked
    End Sub

    'Autofilter toggles in matcher
    Private Sub CheckBox33_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox33.Click
        AutofilterToolStripMenuItem.PerformClick()
    End Sub
    Private Sub CheckBox34_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox34.Click
        AutofilterRomDBToolStripMenuItem.PerformClick()
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

    Private Sub localization_create(ctrl As Control, sw As IO.StreamWriter)
        Dim control_path As String = ctrl.Name
        Dim tmp As Control = ctrl
        Do While tmp.Parent IsNot Nothing
            tmp = tmp.Parent
            control_path = tmp.Name + "." + control_path
        Loop

        'sw.WriteLine("//" + control_path + " - Labels")
        For Each l As Label In ctrl.Controls.OfType(Of Label)
            sw.WriteLine(control_path + "." + l.Name + " = " + l.Text.Replace(vbCrLf, "{CRLF}"))
        Next
        'sw.WriteLine("//" + control_path + " - Buttons")
        For Each b As Button In ctrl.Controls.OfType(Of Button)
            sw.WriteLine(control_path + "." + b.Name + " = " + b.Text)
        Next
        'sw.WriteLine("//" + control_path + " - Checkboxes")
        For Each c As CheckBox In ctrl.Controls.OfType(Of CheckBox)
            sw.WriteLine(control_path + "." + c.Name + " = " + c.Text)
        Next
        'sw.WriteLine("//" + control_path + " - RadioButtons")
        For Each r As RadioButton In ctrl.Controls.OfType(Of RadioButton)
            sw.WriteLine(control_path + "." + r.Name + " = " + r.Text)
        Next
        sw.WriteLine("")

        For Each tabControl In ctrl.Controls.OfType(Of TabControl)
            For Each tabpage As TabPage In tabControl.TabPages
                sw.WriteLine(control_path + "." + tabControl.Name + "." + tabpage.Name + " = " + tabpage.Text)
                localization_create(tabpage, sw)
                sw.WriteLine("")
            Next
        Next

        For Each spl In ctrl.Controls.OfType(Of SplitContainer)
            spl.Panel1.Name = "Panel1"
            spl.Panel2.Name = "Panel2"
            localization_create(spl.Panel1, sw)
            sw.WriteLine("")
            localization_create(spl.Panel2, sw)
            sw.WriteLine("")
        Next

        For Each panel In ctrl.Controls.OfType(Of Panel)
            localization_create(panel, sw)
            sw.WriteLine("")
        Next

        For Each grp In ctrl.Controls.OfType(Of GroupBox)
            sw.WriteLine(control_path + "." + grp.Name + " = " + grp.Text)
            localization_create(grp, sw)
            sw.WriteLine("")
        Next

        For Each item In ctrl.Controls.OfType(Of MenuStrip)
            For Each item2 In item.Items
                If TypeOf item2 Is ToolStripSeparator Then Continue For
                Dim item2a = DirectCast(item2, ToolStripMenuItem)
                sw.WriteLine(control_path + "." + item.Name + "." + item2a.Name + " = " + item2a.Text)
                For Each item3 In item2a.DropDownItems
                    If TypeOf item3 Is ToolStripSeparator Then Continue For
                    Dim item3a = DirectCast(item3, ToolStripMenuItem)
                    sw.WriteLine(control_path + "." + item.Name + "." + item2a.Name + "." + item3a.Name + " = " + item3a.Text)
                    For Each item4 In item3a.DropDownItems
                        If TypeOf item4 Is ToolStripSeparator Then Continue For
                        Dim item4a = DirectCast(item4, ToolStripMenuItem)
                        sw.WriteLine(control_path + "." + item.Name + "." + item2a.Name + "." + item3a.Name + "." + item4a.Name + " = " + item4a.Text)
                        For Each item5 In item4a.DropDownItems
                            If TypeOf item5 Is ToolStripSeparator Then Continue For
                            Dim item5a = DirectCast(item5, ToolStripMenuItem)
                            sw.WriteLine(control_path + "." + item.Name + "." + item2a.Name + "." + item3a.Name + "." + item4a.Name + "." + item5a.Name + " = " + item5a.Text)
                        Next
                    Next
                Next
            Next
        Next
    End Sub
End Class
