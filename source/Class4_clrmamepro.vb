Imports System.Text.RegularExpressions

Public Class Class4_clrmamepro
    Dim WithEvents button_convert_to As Button = Form1.Button22
    Dim WithEvents button_convert_from As Button = Form1.Button23
    Dim WithEvents button_convert_to_GO As Button = Form1.Button24

    Dim WithEvents button_convert_tosec_to_HS As Button = Form1.Button38
    Dim WithEvents button_convert_redump_to_HS As Button = Form1.Button39
    Dim WithEvents button_convert_nointro_std_to_HS As Button = Form1.Button40
    Dim WithEvents button_convert_nointro_xml_to_HS As Button = Form1.Button41
    Dim WithEvents button_convert_mess_xml_to_HS As Button = Form1.Button42
    Dim WithEvents button_convert_new_clrmame_to_HSDB As ToolStripMenuItem = Form1.ConvertFromClrmameprodatToHyperSpinxmlToolStripMenuItem
    Dim WithEvents button_convert_old_clrmame_to_HSDB As ToolStripMenuItem = Form1.ConvertFromOldClrmameprodattosecNointroStdDatToHyperSpinxmlToolStripMenuItem
    Dim WithEvents button_convert_mess_to_HSDB As ToolStripMenuItem = Form1.ConvertMessToHs

    'Convert HSxml to clrMamePro show options
    Private Sub convert_to_click() Handles button_convert_to.Click
        Form1.ComboBox6.Items.Clear()
        Form1.ComboBox6.Text = ""
        For Each s As String In Form1.TextBox3.Text.Split(","c)
            s = s.Trim
            If s.ToLower <> "zip" Then
                Form1.ComboBox6.Items.Add(s)
            End If
        Next
        If Form1.ComboBox6.Items.Count > 0 Then Form1.ComboBox6.SelectedIndex = 0
        Form1.myContextMenu6.Show(Cursor.Position.X, Cursor.Position.Y)
    End Sub

    'Convert HSxml to clrMamePro
    Private Sub convert_to() Handles button_convert_to_GO.Click
        If Form1.ComboBox6.Text = "" Then MsgBox("The rom extension field must be filled!") : Exit Sub
        If Form1.ComboBox1.SelectedIndex < 0 Then MsgBox("You have to select a system.") : Exit Sub
        Form1.myContextMenu6.Hide()

        Dim crc As String
        Dim romExt As String
        Dim romName As String
        Dim x As New Xml.XmlDocument
        Dim xdat As New Xml.XmlDocument
        Dim NotAllCrcFilled As Boolean = False
        romExt = Form1.ComboBox6.Text
        If Not romExt.StartsWith(".") Then romExt = "." + romExt

        'Creating datafile header
        Dim datafileElem As Xml.XmlElement = xdat.CreateElement("datafile") : xdat.AppendChild(datafileElem)
        Dim headerElem As Xml.XmlElement = xdat.CreateElement("header") : datafileElem.AppendChild(headerElem)
        Dim header_nameElem As Xml.XmlElement = xdat.CreateElement("name")
        header_nameElem.InnerText = Form1.ComboBox1.SelectedItem.ToString
        headerElem.AppendChild(header_nameElem)
        Dim header_descriptionElem As Xml.XmlElement = xdat.CreateElement("description")
        header_descriptionElem.InnerText = Form1.ComboBox1.SelectedItem.ToString
        headerElem.AppendChild(header_descriptionElem)
        Dim header_categoryElem As Xml.XmlElement = xdat.CreateElement("category")
        header_categoryElem.InnerText = Form1.TextBox19.Text
        headerElem.AppendChild(header_categoryElem)
        Dim header_versionElem As Xml.XmlElement = xdat.CreateElement("version")
        header_versionElem.InnerText = Form1.TextBox20.Text
        headerElem.AppendChild(header_versionElem)
        Dim header_dateElem As Xml.XmlElement = xdat.CreateElement("date")
        header_dateElem.InnerText = Format(Date.Now.Day, "00") + "/" + Format(Date.Now.Month, "00") + "/" + Date.Now.Year.ToString
        headerElem.AppendChild(header_dateElem)
        Dim header_authorElem As Xml.XmlElement = xdat.CreateElement("author")
        header_authorElem.InnerText = Form1.TextBox21.Text
        headerElem.AppendChild(header_authorElem)
        Dim header_emailElem As Xml.XmlElement = xdat.CreateElement("email")
        header_emailElem.InnerText = Form1.TextBox22.Text
        headerElem.AppendChild(header_emailElem)
        Dim header_homepageElem As Xml.XmlElement = xdat.CreateElement("homepage")
        header_homepageElem.InnerText = Form1.TextBox23.Text
        headerElem.AppendChild(header_homepageElem)
        Dim header_urlElem As Xml.XmlElement = xdat.CreateElement("url")
        header_urlElem.InnerText = Form1.TextBox24.Text
        headerElem.AppendChild(header_urlElem)
        Dim header_commentElem As Xml.XmlElement = xdat.CreateElement("comment")
        header_commentElem.InnerText = Form1.TextBox25.Text
        headerElem.AppendChild(header_commentElem)
        Dim header_clrmameproElem As Xml.XmlElement = xdat.CreateElement("clrmamepro")
        header_clrmameproElem.InnerText = ""
        headerElem.AppendChild(header_clrmameproElem)


        x.Load(Form1.xmlPath)
        For Each node As Xml.XmlNode In x.SelectNodes("/menu/game")
            romName = node.Attributes.GetNamedItem("name").Value
            Dim dat_gameElem As Xml.XmlElement = xdat.CreateElement("game")
            dat_gameElem.SetAttribute("name", romName)
            datafileElem.AppendChild(dat_gameElem)

            Dim dat_gameDescrElem As Xml.XmlElement = xdat.CreateElement("description")
            dat_gameDescrElem.InnerText = node.SelectSingleNode("description").InnerText
            dat_gameElem.AppendChild(dat_gameDescrElem)
            Dim dat_gameYearElem As Xml.XmlElement = xdat.CreateElement("year")
            dat_gameYearElem.InnerText = node.SelectSingleNode("year").InnerText
            dat_gameElem.AppendChild(dat_gameYearElem)
            Dim dat_gameManufacturerElem As Xml.XmlElement = xdat.CreateElement("manufacturer")
            dat_gameManufacturerElem.InnerText = node.SelectSingleNode("manufacturer").InnerText
            dat_gameElem.AppendChild(dat_gameManufacturerElem)

            crc = node.SelectSingleNode("crc").InnerText
            If crc = "" Then NotAllCrcFilled = True
            Dim dat_gameRomElem As Xml.XmlElement = xdat.CreateElement("rom")
            dat_gameRomElem.SetAttribute("name", romName + romExt)
            dat_gameRomElem.SetAttribute("crc", crc)
            dat_gameElem.AppendChild(dat_gameRomElem)
        Next

        If NotAllCrcFilled Then
            If MsgBox("One ore more CRC field in HS Database are empty. Dat file will be useless. Do you want to continue?") = MsgBoxResult.No Then Exit Sub
        End If

        Dim fs As New SaveFileDialog
        fs.Title = "Select folder to put DAT to."
        fs.Filter = "ClrMamePro dat Files | *.dat"
        fs.RestoreDirectory = True
        fs.InitialDirectory = Class1.HyperspinPath + "\Databases"
        fs.ShowDialog()
        If fs.FileName = "" Then Exit Sub
        Dim w As Xml.XmlWriter = Xml.XmlWriter.Create(fs.FileName, New Xml.XmlWriterSettings With {.Indent = True, .NewLineHandling = Xml.NewLineHandling.None})
        xdat.Save(w) : w.Close()
    End Sub

    'Convert clrMamePro to HSxml
    Private Sub convert_from() Handles button_convert_from.Click, button_convert_new_clrmame_to_HSDB.Click
        Dim fb As New OpenFileDialog
        fb.Title = "Open clrMamePro dat or xml"
        fb.Filter = "ClrMamePro/Redump/No-Intro xml|*.xml|ClrMamePro/Redump/No-Intro dat|*.dat|All files|*.*"
        fb.ShowDialog()
        If fb.FileName = "" Then MsgBox("No file selected.") : Exit Sub
        If Not FileIO.FileSystem.FileExists(fb.FileName) Then MsgBox("File """ + fb.FileName + """ does not exist.") : Exit Sub

        Dim xHS As New Xml.XmlDocument
        Dim xdat As New Xml.XmlDocument
        Try
            xdat.Load(fb.FileName)
        Catch ex As Exception
            MsgBox("Sorry, only new clrmamepro dat format (xml) is supported.") : Exit Sub
        End Try

        Dim menuElem As Xml.XmlElement = xHS.CreateElement("menu") : xHS.AppendChild(menuElem)
        For Each node As Xml.XmlNode In xdat.SelectNodes("/datafile/game")
            Dim elemDescription As Xml.XmlElement = xHS.CreateElement("description")
            Dim elemCrc As Xml.XmlElement = xHS.CreateElement("crc")
            Dim elemManufacturer As Xml.XmlElement = xHS.CreateElement("manufacturer")
            Dim elemYear As Xml.XmlElement = xHS.CreateElement("year")
            Dim elemGenre As Xml.XmlElement = xHS.CreateElement("genre")
            Dim cloneof As Xml.XmlElement = xHS.CreateElement("cloneof")

            Dim gameElem As Xml.XmlElement = xHS.CreateElement("game")
            gameElem.SetAttribute("name", node.Attributes.GetNamedItem("name").Value)

            If node.SelectNodes("description").Count > 0 Then
                elemDescription.InnerText = node.SelectSingleNode("description").InnerText
            End If
            If node.SelectNodes("year").Count > 0 Then
                elemYear.InnerText = node.SelectSingleNode("year").InnerText
            End If
            If node.SelectNodes("manufacturer").Count > 0 Then
                elemManufacturer.InnerText = node.SelectSingleNode("manufacturer").InnerText
            End If
            If node.SelectNodes("rom").Count = 1 Then
                If node.SelectSingleNode("rom").Attributes.GetNamedItem("crc") IsNot Nothing Then
                    elemCrc.InnerText = node.SelectSingleNode("rom").Attributes.GetNamedItem("crc").Value
                End If
            End If
            If node.Attributes.GetNamedItem("cloneof") IsNot Nothing Then
                If node.Attributes.GetNamedItem("cloneof").Value <> "" Then
                    cloneof.InnerText = node.Attributes.GetNamedItem("cloneof").Value
                End If
            End If

            gameElem.AppendChild(elemDescription)
            gameElem.AppendChild(elemCrc)
            gameElem.AppendChild(elemManufacturer)
            gameElem.AppendChild(elemYear)
            gameElem.AppendChild(elemGenre)
            gameElem.AppendChild(cloneof)
            menuElem.AppendChild(gameElem)
        Next

        Dim fs As New SaveFileDialog
        fs.Title = "Select folder to put HS xml DB to."
        fs.Filter = "Hyperspin database files|*.xml"
        fs.RestoreDirectory = True
        fs.InitialDirectory = Class1.HyperspinPath + "\Databases"
        fs.ShowDialog()
        If fs.FileName = "" Then Exit Sub
        Dim w As Xml.XmlWriter = Xml.XmlWriter.Create(fs.FileName, New Xml.XmlWriterSettings With {.Indent = True, .NewLineHandling = Xml.NewLineHandling.None})
        xHS.Save(w) : w.Close()
    End Sub

    'Convert tosec.dat to HS xml
    Private Sub tosec_to_hs() Handles button_convert_tosec_to_HS.Click, button_convert_old_clrmame_to_HSDB.Click
        Dim fb As New OpenFileDialog
        fb.Title = "Open Tosec dat"
        fb.Filter = "Tosec / no-intro dat|*.dat|All files|*.*"
        fb.ShowDialog()
        If fb.FileName = "" Then MsgBox("No file selected.") : Exit Sub
        If Not FileIO.FileSystem.FileExists(fb.FileName) Then MsgBox("File """ + fb.FileName + """ does not exist.") : Exit Sub

        Dim xHS As New Xml.XmlDocument
        Dim menuElem As Xml.XmlElement = xHS.CreateElement("menu") : xHS.AppendChild(menuElem)

        Dim dat_file As String = IO.File.ReadAllText(fb.FileName)
        Dim games() As String = Regex.Split(dat_file, "game", RegexOptions.IgnoreCase)
        For Each game As String In games
            If game.ToLower.Contains("clrmamepro") Then Continue For
            Try
                Dim gamename As String = Regex.Match(game, "name\s*""(.*)""", RegexOptions.IgnoreCase).Groups(1).Value
                Dim description As String = Regex.Match(game, "description\s*""(.*)""", RegexOptions.IgnoreCase).Groups(1).Value
                Dim romname As String = Regex.Match(game, "rom\s*" + Regex.Escape("(") + ".*name\s*""(.*)""", RegexOptions.IgnoreCase).Groups(1).Value
                Dim crc As String = Regex.Match(game, "rom\s*" + Regex.Escape("(") + ".*crc\s*(\S*)", RegexOptions.IgnoreCase).Groups(1).Value

                Dim elemDescription As Xml.XmlElement = xHS.CreateElement("description")
                Dim elemCrc As Xml.XmlElement = xHS.CreateElement("crc")
                Dim elemManufacturer As Xml.XmlElement = xHS.CreateElement("manufacturer")
                Dim elemYear As Xml.XmlElement = xHS.CreateElement("year")
                Dim elemGenre As Xml.XmlElement = xHS.CreateElement("genre")

                Dim gameElem As Xml.XmlElement = xHS.CreateElement("game")
                gameElem.SetAttribute("name", gamename)
                elemDescription.InnerText = description
                elemCrc.InnerText = crc

                gameElem.AppendChild(elemDescription)
                gameElem.AppendChild(elemCrc)
                gameElem.AppendChild(elemManufacturer)
                gameElem.AppendChild(elemYear)
                gameElem.AppendChild(elemGenre)
                menuElem.AppendChild(gameElem)
            Catch ex As Exception
                MsgBox(ex.Message)
                Exit Sub
            End Try
        Next

        Dim fs As New SaveFileDialog
        fs.Title = "Select folder to put HS xml DB to."
        fs.Filter = "Hyperspin database files|*.xml"
        fs.RestoreDirectory = True
        fs.InitialDirectory = Class1.HyperspinPath + "\Databases"
        fs.ShowDialog()
        If fs.FileName = "" Then Exit Sub
        Dim w As Xml.XmlWriter = Xml.XmlWriter.Create(fs.FileName, New Xml.XmlWriterSettings With {.Indent = True, .NewLineHandling = Xml.NewLineHandling.None})
        xHS.Save(w) : w.Close()
    End Sub

    'Convert redump.dat to HS xml
    Private Sub redump_to_hs() Handles button_convert_redump_to_HS.Click
        convert_from()
    End Sub

    'Convert no-intro std .dat to HS xml
    Private Sub nointro_std_to_hs() Handles button_convert_nointro_std_to_HS.Click
        tosec_to_hs()
    End Sub

    'Convert no-intro xml .dat to HS xml
    Private Sub nointro_xml_to_hs() Handles button_convert_nointro_xml_to_HS.Click
        convert_from()
    End Sub

    'Convert mess xml to HS xml
    Private Sub mess_to_hs() Handles button_convert_mess_xml_to_HS.Click, button_convert_mess_to_HSDB.Click
        Dim fb As New OpenFileDialog
        fb.Title = "Open clrMamePro dat or xml"
        fb.Filter = "MESS softwarelist xml|*.xml|All files|*.*"
        fb.ShowDialog()
        If fb.FileName = "" Then MsgBox("No file selected.") : Exit Sub
        If Not FileIO.FileSystem.FileExists(fb.FileName) Then MsgBox("File """ + fb.FileName + """ does not exist.") : Exit Sub

        Dim xml_class As New Class2_xml()
        xml_class.xml_repare(fb.FileName)

        Dim xHS As New Xml.XmlDocument
        Dim xdat As New Xml.XmlDocument
        Try
            xdat.Load(fb.FileName)
        Catch ex As Exception
            MsgBox("Sorry, only new clrmamepro dat format (xml) is supported.") : Exit Sub
        End Try

        Dim menuElem As Xml.XmlElement = xHS.CreateElement("menu") : xHS.AppendChild(menuElem)
        For Each node As Xml.XmlNode In xdat.SelectNodes("/datafile/game")
            Dim elemDescription As Xml.XmlElement = xHS.CreateElement("description")
            Dim elemCrc As Xml.XmlElement = xHS.CreateElement("crc")
            Dim elemManufacturer As Xml.XmlElement = xHS.CreateElement("manufacturer")
            Dim elemYear As Xml.XmlElement = xHS.CreateElement("year")
            Dim elemGenre As Xml.XmlElement = xHS.CreateElement("genre")
            Dim cloneof As Xml.XmlElement = xHS.CreateElement("cloneof")

            Dim gameElem As Xml.XmlElement = xHS.CreateElement("game")
            gameElem.SetAttribute("name", node.Attributes.GetNamedItem("name").Value)

            If node.SelectNodes("description").Count > 0 Then
                elemDescription.InnerText = node.SelectSingleNode("description").InnerText
            End If
            If node.SelectNodes("year").Count > 0 Then
                elemYear.InnerText = node.SelectSingleNode("year").InnerText
            End If
            If node.SelectNodes("manufacturer").Count > 0 Then
                elemManufacturer.InnerText = node.SelectSingleNode("manufacturer").InnerText
            End If
            If node.SelectNodes("rom").Count = 1 Then
                If node.SelectSingleNode("rom").Attributes.GetNamedItem("crc") IsNot Nothing Then
                    elemCrc.InnerText = node.SelectSingleNode("rom").Attributes.GetNamedItem("crc").Value
                End If
            End If
            If node.Attributes.GetNamedItem("cloneof") IsNot Nothing Then
                If node.Attributes.GetNamedItem("cloneof").Value <> "" Then
                    cloneof.InnerText = node.Attributes.GetNamedItem("cloneof").Value
                End If
            End If

            gameElem.AppendChild(elemDescription)
            gameElem.AppendChild(elemCrc)
            gameElem.AppendChild(elemManufacturer)
            gameElem.AppendChild(elemYear)
            gameElem.AppendChild(elemGenre)
            gameElem.AppendChild(cloneof)
            menuElem.AppendChild(gameElem)
        Next

        Dim fs As New SaveFileDialog
        fs.Title = "Select folder to put HS xml DB to."
        fs.Filter = "Hyperspin database files|*.xml"
        fs.RestoreDirectory = True
        fs.InitialDirectory = Class1.HyperspinPath + "\Databases"
        fs.ShowDialog()
        If fs.FileName = "" Then Exit Sub
        Dim w As Xml.XmlWriter = Xml.XmlWriter.Create(fs.FileName, New Xml.XmlWriterSettings With {.Indent = True, .NewLineHandling = Xml.NewLineHandling.None})
        xHS.Save(w) : w.Close()
    End Sub
End Class
