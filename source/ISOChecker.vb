Imports Microsoft.VisualBasic.FileIO
Public Class ISOChecker
    Private isofix(10) As Integer
    Private cueWoBin_filename_mismatch As New List(Of String)
    Private cueWoBin_filename_mismatch_binFname As New List(Of String)
    Private WithEvents Button21_checkISO As Button = Form1.Button21_checkISO

    Public Sub Check() Handles Button21_checkISO.Click
        If Form1.ComboBox2.SelectedIndex = 0 And Form1.CheckBox15.Checked Then
            Dim m As String
            m = "You have selected autodetect .bin file data mode and bps. If .cue file will need to be created the program will try to autodetect .bin format." + vbCrLf
            m = m + "Keep in mind, that autodetection WILL NOT WORK for systems with non standard filesystem (Panasonic 3DO, NEC PC-FX, Nintendo GameCube)" + vbCrLf
            m = m + "though, all theese systems are usually ISO, but please be advised."
            Dim ret As MsgBoxResult = MsgBox(m, MsgBoxStyle.OkCancel)
            If ret = MsgBoxResult.Cancel Then Exit Sub
        End If
        For i As Integer = 0 To 10
            isofix(i) = 0
        Next
        Form1.ListBox3.Items.Clear() : Form1.ListBox3.Refresh()
        'If ComboBox1.SelectedIndex < 0 Then MsgBox("Please, select a system.") : Exit Sub
        If Not FileSystem.DirectoryExists(Form1.TextBox17.Text) Then MsgBox("Path: """ + Form1.TextBox17.Text + """ does not exist.") : Exit Sub
        Class1.Log("checking folder: " + Form1.TextBox17.Text)
        Button2_Click_1_subFilesCheck(Form1.TextBox17.Text)
        For Each d As String In FileSystem.GetDirectories(Form1.TextBox17.Text)
            Class1.Log("checking folder: " + d)
            Button2_Click_1_subFilesCheck(d)
            Form1.ListBox3.Refresh()
        Next
        Form1.ListBox3.Items.Add("All done. " + Form1.ListBox3.Items.Count.ToString + " warnings.")
    End Sub 'Check ISO

    Private Sub Button2_Click_1_subFilesCheck(ByVal dir As String)
        'MDF wo MDS / MDS wo MDF
        Dim mdfWOmds As New List(Of String), mdsWOmdf As New List(Of String)
        If Form1.CheckBox1.Checked Then
            Dim mdsList As System.Collections.ObjectModel.ReadOnlyCollection(Of String)
            Dim mdfList As System.Collections.ObjectModel.ReadOnlyCollection(Of String)
            Dim mdsArr As New List(Of String)
            Dim mdfArr As New List(Of String)
            Dim mdsArrRightCase As New List(Of String)
            Dim mdfArrRightCase As New List(Of String)
            mdsList = FileSystem.GetFiles(dir, SearchOption.SearchTopLevelOnly, {"*.mds"})
            mdfList = FileSystem.GetFiles(dir, SearchOption.SearchTopLevelOnly, {"*.mdf"})
            For Each mds In mdsList
                mds = mds.Substring(mds.LastIndexOf("\") + 1)
                mds = mds.Substring(0, mds.LastIndexOf("."))
                mdsArr.Add(mds.ToLower)
                mdsArrRightCase.Add(mds)
            Next
            For Each mdf In mdfList
                mdf = mdf.Substring(mdf.LastIndexOf("\") + 1)
                mdf = mdf.Substring(0, mdf.LastIndexOf("."))
                mdfArr.Add(mdf.ToLower)
                mdfArrRightCase.Add(mdf)
            Next

            Dim i As Integer = 0
            For Each mds In mdsArr
                If Not mdfArr.Contains(mds) Then mdsWOmdf.Add(mdsArrRightCase(i) + ".mds")
                i += 1
            Next
            i = 0
            For Each mdf In mdfArr
                If Not mdsArr.Contains(mdf) Then mdfWOmds.Add(mdfArrRightCase(i) + ".mdf")
                i += 1
            Next

            'MDF / MDS FIX
            If mdsWOmdf.Count > 0 And mdfWOmds.Count > 0 And Form1.CheckBox14.Checked Then
                Class1.askVar1 = dir
                Button2_Click_1_subFilesFix1(mdsWOmdf, mdfWOmds)
            End If

            'Show Result
            For Each mds In mdsWOmdf
                Form1.ListBox3.Items.Add("MDS without MDF: " + mds + ".mds")
            Next
            For Each mdf In mdfWOmds
                Form1.ListBox3.Items.Add("MDF without MDS: " + mdf + ".mds")
            Next
        End If

        'Cue / Bin problems
        If Form1.CheckBox2.Checked Or Form1.CheckBox4.Checked Or Form1.CheckBox8.Checked Then
            Dim binfilelist As New List(Of String)
            Dim wavefilelist As New List(Of String)
            Dim binWocue As New List(Of String)
            Dim cueWoBin As New List(Of String)
            Dim cueWoBin_file_in_cue As New List(Of String)
            cueWoBin_filename_mismatch.Clear()
            cueWoBin_filename_mismatch_binFname.Clear()
            Dim cueList As System.Collections.ObjectModel.ReadOnlyCollection(Of String)
            'cueList = FileSystem.GetFiles(dir, SearchOption.SearchTopLevelOnly, {"*.cue"}).OrderBy(Function(fi) fi)
            'Dim cueList = From file In FileSystem.GetFiles(dir, SearchOption.SearchTopLevelOnly, {"*.cue"}) Order By file Select file
            cueList = FileSystem.GetFiles(dir, SearchOption.SearchTopLevelOnly, {"*.cue"})

            For Each cue As String In cueList
                Class1.Log("checking CUE: " + cue.Substring(cue.LastIndexOf("\")))
                Dim line As String
                Dim binFilename As String = ""
                Dim waveFilename As String = ""
                Dim cueFileName As String = cue.Substring(cue.LastIndexOf("\") + 1)
                FileOpen(1, cue, OpenMode.Input)
                Dim msgshown As Boolean = False
                Dim msgshown2 As Boolean = False
                Dim msgshown3 As Boolean = False
                Do While Not EOF(1)
                    line = LineInput(1)
                    If line.IndexOf("file", System.StringComparison.InvariantCultureIgnoreCase) >= 0 Then
                        If line.IndexOf("binary", System.StringComparison.InvariantCultureIgnoreCase) >= 0 And binFilename = "" Then
                            binFilename = line.Substring(line.IndexOf("""") + 1)
                            binFilename = binFilename.Substring(0, binFilename.LastIndexOf(""""))
                        Else
                            'Check audio format and presence
                            If line.IndexOf("wave", System.StringComparison.InvariantCultureIgnoreCase) < 0 And line.IndexOf("binary", System.StringComparison.InvariantCultureIgnoreCase) < 0 Then
                                If Not msgshown And Form1.CheckBox5.Checked Then msgshown = True : Form1.ListBox3.Items.Add(".CUE contain audio other then WAVE: " + cueFileName)
                            End If
                            waveFilename = line.Substring(line.IndexOf("""") + 1)
                            waveFilename = waveFilename.Substring(0, waveFilename.LastIndexOf(""""))
                            If Not waveFilename.Trim.EndsWith(".wav", StringComparison.InvariantCultureIgnoreCase) And Not waveFilename.Trim.EndsWith(".bin", StringComparison.InvariantCultureIgnoreCase) Then
                                If Not msgshown2 And Form1.CheckBox5.Checked Then msgshown2 = True : Form1.ListBox3.Items.Add(".CUE reference other format then .wav: " + cueFileName)
                            End If
                            If waveFilename.Contains("\") Then
                                waveFilename = waveFilename.Substring(waveFilename.LastIndexOf("\") + 1)
                                If Not msgshown3 And Form1.CheckBox4.Checked Then msgshown3 = True : Form1.ListBox3.Items.Add(".CUE audio reference contain file path: " + cueFileName)
                            End If
                            If Form1.CheckBox9.Checked Then
                                If Not FileSystem.FileExists(dir + "\" + waveFilename) Then
                                    Form1.ListBox3.Items.Add("missing audio: " + waveFilename + " in .CUE: " + cueFileName)
                                End If
                            End If
                            wavefilelist.Add(waveFilename.ToLower)
                        End If
                    End If
                Loop
                FileClose(1)

                'Malformed .cue
                If binFilename = "" Then If Form1.CheckBox4.Checked Then Form1.ListBox3.Items.Add(".CUE with no defined imagefile: " + cueFileName)
                If binFilename.Contains("\") Then
                    binFilename = binFilename.Substring(binFilename.LastIndexOf("\") + 1)
                    If Form1.CheckBox4.Checked Then Form1.ListBox3.Items.Add(".CUE contain file path: " + cueFileName)
                End If
                binfilelist.Add(binFilename.ToLower)

                'CUE without BIN
                Dim v1 As Boolean = FileSystem.FileExists(dir + "\" + binFilename)
                Dim v2 As Boolean = binFilename.Substring(0, binFilename.LastIndexOf(".")).ToLower = cueFileName.Substring(0, cueFileName.LastIndexOf(".")).ToLower
                'Dim v2 As Boolean = FileSystem.FileExists(dir + "\" + cueFileName.Substring(0, cueFileName.LastIndexOf(".")) + ".bin")
                If Not v1 Then cueWoBin.Add(cue.Substring(cue.LastIndexOf("\") + 1)) : cueWoBin_file_in_cue.Add(binFilename)
                If v1 And Not v2 Then cueWoBin_filename_mismatch.Add(cue.Substring(cue.LastIndexOf("\") + 1)) : cueWoBin_filename_mismatch_binFname.Add(binFilename)
                If Form1.CheckBox8.Checked Then
                    If Not dir.EndsWith("\") Then dir = dir + "\"
                    Dim c As String = cue.Substring(cue.LastIndexOf("\") + 1) + " (" + dir + ")"
                    If Form1.RadioButton8.Checked Or Form1.RadioButton9.Checked Then If Not v1 Then Form1.ListBox3.Items.Add(".CUE without .BIN (provided match): " + c)
                    If Form1.RadioButton7.Checked Or Form1.RadioButton9.Checked Then If Not v2 Then Form1.ListBox3.Items.Add(".CUE without .BIN (filename match): " + c)
                End If
            Next

            'create BIN wo CUE list
            For Each bin As String In FileSystem.GetFiles(dir, SearchOption.SearchTopLevelOnly, {"*.bin"})
                bin = bin.Substring(bin.LastIndexOf("\") + 1)
                If Not binfilelist.Contains(bin.ToLower) And Not wavefilelist.Contains(bin.ToLower) Then
                    If Autodetect_bin_format(dir + "\" + bin) <> "AUDIO" Then
                        binWocue.Add(bin)
                    End If
                End If
            Next

            'CUE without BIN FIX
            If Form1.CheckBox20.Checked Then
                Dim indexesToRemove As New List(Of Integer)
                For i As Integer = 0 To cueWoBin.Count - 1
                    If binWocue.Count > 0 Then
                        If Not dir.EndsWith("\") Then dir = dir + "\"
                        Dim ret As Integer = Button2_Click_1_subFilesFix3(dir + cueWoBin(i), cueWoBin_file_in_cue(i), binWocue)
                        If ret <> -1 Then
                            binWocue.RemoveAt(ret) : indexesToRemove.Add(i)
                        End If
                    End If
                Next
                For i As Integer = indexesToRemove.Count - 1 To 0 Step -1
                    cueWoBin.RemoveAt(indexesToRemove(i))
                Next
            End If

            'CUE / BIN filenames mismatch FIX
            If Form1.CheckBox19.Checked Then
                For i As Integer = 0 To cueWoBin_filename_mismatch.Count - 1
                    Button2_Click_1_subFilesFix4(dir, cueWoBin_filename_mismatch(i), cueWoBin_filename_mismatch_binFname(i))
                Next
            End If

            'BIN without CUE
            If Form1.CheckBox2.Checked Then
                For Each bin As String In binWocue
                    If Form1.CheckBox15.Checked Then
                        'If bin.Length > 100 Then Class1.askVar1 = "..." + bin.Substring(bin.Length - 98) Else Class1.askVar1 = bin
                        Class1.askVar1 = bin
                        If Button2_Click_1_subFilesFix2(dir, bin, cueWoBin) = -1 Then Form1.ListBox3.Items.Add(".BIN without .CUE: " + bin + " (" + dir + ")")
                    Else
                        Form1.ListBox3.Items.Add(".BIN without .CUE: " + bin + " (" + dir + ")")
                    End If
                Next
            End If

            'AUDIO without CUE
            If Form1.CheckBox7.Checked Then
                For Each audio As String In FileSystem.GetFiles(dir, SearchOption.SearchTopLevelOnly, {"*.WAV", "*.MP3", "*.OGG", "*.FLAC"})
                    audio = audio.Substring(audio.LastIndexOf("\") + 1)
                    If Not wavefilelist.Contains(audio.ToLower) Then Form1.ListBox3.Items.Add("Audiofile not referenced in .CUE: " + dir + "\" + audio)
                Next
            End If
        End If

        'GDI
        If Form1.CheckBox25.Checked Then
            Dim gdiList As System.Collections.ObjectModel.ReadOnlyCollection(Of String)
            gdiList = FileSystem.GetFiles(dir, SearchOption.SearchTopLevelOnly, {"*.gdi"})
            For Each gdi As String In gdiList
                Dim fName As String = gdi.Substring(gdi.LastIndexOf("\") + 1)
                If My.Computer.FileSystem.GetFileInfo(gdi).Length > 307200 Then Form1.ListBox3.Items.Add("Strange GDI, maybe it is just renamed .CDI: " + fName + " (" + dir + ")")
            Next
        End If
    End Sub 'Check ISO SUB

    Private Sub Button2_Click_1_subFilesFix1(ByVal mdsWOmdf As List(Of String), ByVal mdfWOmds As List(Of String))
        Dim f1WOext, f2WOext, f1Ext, f2Ext As String
        If mdsWOmdf.Count = 1 And mdfWOmds.Count = 1 Then
            Class1.askVar2 = mdfWOmds(0)
            Class1.askVar3 = mdsWOmdf(0)
            If isofix(0) = 0 Then
                ISOFixAsk_MdfMds.ShowDialog()
                If Class1.askResponse > 100 Then isofix(0) = Class1.askResponse - 100 : Class1.askResponse = Class1.askResponse - 100
            Else
                Class1.askResponse = isofix(0)
            End If
            f1Ext = Class1.askVar2.Substring(Class1.askVar2.LastIndexOf(".") + 1)
            f2Ext = Class1.askVar3.Substring(Class1.askVar3.LastIndexOf(".") + 1)
            f1WOext = Class1.askVar2.Substring(0, Class1.askVar2.LastIndexOf("."))
            f2WOext = Class1.askVar3.Substring(0, Class1.askVar3.LastIndexOf("."))
            If Class1.askResponse = 1 Then
                Microsoft.VisualBasic.FileIO.FileSystem.RenameFile(Class1.askVar1 + "\" + Class1.askVar2, f2WOext + "." + f1Ext)
            ElseIf Class1.askResponse = 2 Then
                Microsoft.VisualBasic.FileIO.FileSystem.RenameFile(Class1.askVar1 + "\" + Class1.askVar3, f1WOext + "." + f2Ext)
            End If
            If Class1.askResponse <> 3 Then mdfWOmds.Clear() : mdsWOmdf.Clear()
        ElseIf mdsWOmdf.Count = 1 And mdfWOmds.Count > 1 Then
            Class1.askVar2 = mdsWOmdf(0) : Class1.askList = mdfWOmds
            If isofix(1) <> 100000 Then
                ISOFixAsk_MdfMds_ext.ShowDialog()
                If Class1.askResponse >= 100000 Then isofix(1) = 100000 : Class1.askResponse = Class1.askResponse - 100000
            Else
                Class1.askResponse = 0
            End If
            If Class1.askResponse > 0 Then
                Class1.askVar3 = mdfWOmds(Class1.askResponse - 1)
                f1Ext = Class1.askVar2.Substring(Class1.askVar2.LastIndexOf(".") + 1)
                'f2Ext = Class1.askVar3.Substring(Class1.askVar3.LastIndexOf(".") + 1)
                'f1WOext = Class1.askVar2.Substring(0, Class1.askVar2.LastIndexOf("."))
                f2WOext = Class1.askVar3.Substring(0, Class1.askVar3.LastIndexOf("."))
                Microsoft.VisualBasic.FileIO.FileSystem.RenameFile(Class1.askVar1 + "\" + Class1.askVar2, f2WOext + "." + f1Ext)
                mdsWOmdf.Clear()
                mdfWOmds.RemoveAt(Class1.askResponse - 1)
            End If
        ElseIf mdsWOmdf.Count > 1 And mdfWOmds.Count = 1 Then
            Class1.askVar2 = mdfWOmds(0) : Class1.askList = mdsWOmdf
            If isofix(1) <> 100000 Then
                ISOFixAsk_MdfMds_ext.ShowDialog()
                If Class1.askResponse >= 100000 Then isofix(1) = 100000 : Class1.askResponse = Class1.askResponse - 100000
            Else
                Class1.askResponse = 0
            End If
            If Class1.askResponse > 0 Then
                Class1.askVar3 = mdsWOmdf(Class1.askResponse - 1)
                f1Ext = Class1.askVar2.Substring(Class1.askVar2.LastIndexOf(".") + 1)
                f2WOext = Class1.askVar3.Substring(0, Class1.askVar3.LastIndexOf("."))
                Microsoft.VisualBasic.FileIO.FileSystem.RenameFile(Class1.askVar1 + "\" + Class1.askVar2, f2WOext + "." + f1Ext)
                mdfWOmds.Clear()
                mdsWOmdf.RemoveAt(Class1.askResponse - 1)
            End If
        Else
            Class1.askList = mdfWOmds
            Dim mdsToRemove As New List(Of String)
            Dim mdfToRemove As New List(Of String)
            For Each mds In mdsWOmdf
                Class1.askVar2 = mds
                If isofix(1) <> 100000 Then
                    ISOFixAsk_MdfMds_ext.ShowDialog()
                    If Class1.askResponse >= 100000 Then isofix(1) = 100000 : Class1.askResponse = Class1.askResponse - 100000
                Else
                    Class1.askResponse = 0
                End If
                If Class1.askResponse > 0 Then
                    Class1.askVar3 = mdfWOmds(Class1.askResponse - 1)
                    f1Ext = Class1.askVar2.Substring(Class1.askVar2.LastIndexOf(".") + 1)
                    f2WOext = Class1.askVar3.Substring(0, Class1.askVar3.LastIndexOf("."))
                    Microsoft.VisualBasic.FileIO.FileSystem.RenameFile(Class1.askVar1 + "\" + Class1.askVar2, f2WOext + "." + f1Ext)
                    mdsToRemove.Add(mds)
                    mdfWOmds.RemoveAt(Class1.askResponse - 1)
                End If
            Next
            For Each mds In mdsToRemove
                mdsWOmdf.Remove(mds)
            Next

            Class1.askList = mdsWOmdf
            For Each mdf In mdfWOmds
                Class1.askVar2 = mdf
                If isofix(1) <> 100000 Then
                    ISOFixAsk_MdfMds_ext.ShowDialog()
                    If Class1.askResponse > 100000 Then isofix(1) = 100000 : Class1.askResponse = Class1.askResponse - 100000
                Else
                    Class1.askResponse = 0
                End If
                If Class1.askResponse > 0 Then
                    Class1.askVar3 = mdsWOmdf(Class1.askResponse - 1)
                    f1Ext = Class1.askVar2.Substring(Class1.askVar2.LastIndexOf(".") + 1)
                    f2WOext = Class1.askVar3.Substring(0, Class1.askVar3.LastIndexOf("."))
                    Microsoft.VisualBasic.FileIO.FileSystem.RenameFile(Class1.askVar1 + "\" + Class1.askVar2, f2WOext + "." + f1Ext)
                    mdfToRemove.Add(mdf)
                    mdsWOmdf.RemoveAt(Class1.askResponse - 1)
                End If
            Next
            For Each mdf In mdfToRemove
                mdfWOmds.Remove(mdf)
            Next
        End If
    End Sub 'Check ISO FIX mdf/mds

    Private Function Button2_Click_1_subFilesFix2(ByVal dir As String, ByVal fname As String, ByVal l As List(Of String)) As Integer
        '-1=create, -2=do nothing, -3=always associate for single cue, +100000 = always
        If l.Count = 0 Then
            If isofix(2) = 0 Then
                ISOFixAsk_binWoCue.ShowDialog()
                If Class1.askResponse > 99000 Then Class1.askResponse = Class1.askResponse - 100000 : isofix(2) = Class1.askResponse
            Else
                Class1.askResponse = isofix(2)
            End If
        ElseIf l.Count = 1 Then
            If isofix(3) = 0 Then
                Class1.askVar2 = l(0)
                ISOFixAsk_binWoCue_ext.ShowDialog()
                If Class1.askResponse > 99000 Then Class1.askResponse = Class1.askResponse - 100000 : isofix(3) = Class1.askResponse
            Else
                Class1.askResponse = isofix(3)
            End If
        Else
            If isofix(4) = 0 Then
                Class1.askList = l
                ISOFixAsk_binWoCue_extPlus.ShowDialog()
                If Class1.askResponse > 99000 Then Class1.askResponse = Class1.askResponse - 100000 : isofix(4) = Class1.askResponse
            Else
                Class1.askResponse = isofix(4)
            End If
        End If
        If Class1.askResponse = -2 Then Return -1
        If Class1.askResponse = -3 Then Class1.askResponse = 0
        If Class1.askResponse >= 0 Then
            Class1.associate_rewriteCue(dir + "\" + l(Class1.askResponse), fname)
            If Not dir.EndsWith("\") Then dir = dir + "\"
            Form1.ListBox3.Items.Remove(".CUE without .BIN (provided match): " + l(Class1.askResponse) + " (" + dir + ")")
            Dim FnameWoExt = fname.Substring(0, fname.LastIndexOf("."))
            Dim cueFnameWoExt = l(Class1.askResponse).Substring(0, l(Class1.askResponse).LastIndexOf("."))
            If cueFnameWoExt.ToLower = FnameWoExt.ToLower Then
                Form1.ListBox3.Items.Remove(".CUE without .BIN (filename match): " + l(Class1.askResponse) + " (" + dir + ")")
            End If
            Return Class1.askResponse
        Else
            Dim audio As New List(Of String)
            For Each a As String In FileSystem.GetFiles(dir, SearchOption.SearchTopLevelOnly, {"*.WAV"})
                audio.Add(a.Substring(a.LastIndexOf("\") + 1))
            Next
            audio.Sort(System.StringComparer.InvariantCultureIgnoreCase)

            Dim binWoExt = fname.Substring(0, fname.LastIndexOf("."))
            If FileSystem.FileExists(dir + "\" + binWoExt + ".cue") Then MsgBox("File """ + binWoExt + ".cue"" already exist!") : Return -1
            FileOpen(1, dir + "\" + binWoExt + ".cue", OpenMode.Output)
            PrintLine(1, "FILE """ + fname + """ BINARY")
            '"MODE1/2048", "MODE1/2352", "MODE2/2336", "MODE2/2352"
            If Form1.ComboBox2.SelectedIndex = 0 Then
                PrintLine(1, "  TRACK 01 MODE" + Autodetect_bin_format(dir + "\" + fname))
            Else
                PrintLine(1, "  TRACK 01 " + Form1.ComboBox2.SelectedItem.ToString.Substring(0, 10))
            End If
            PrintLine(1, "    INDEX 01 00:00:00")
            Dim index As Integer = 2
            For Each a As String In audio
                PrintLine(1, "FILE """ + a + """ WAVE")
                PrintLine(1, "  TRACK " + Format(index, "00") + " AUDIO")
                PrintLine(1, "    INDEX 00 00:00:00")
                PrintLine(1, "    INDEX 01 00:02:00")
                index += 1
            Next
            FileClose(1)
            Return -2
        End If
    End Function 'Check ISO FIX .bin wo .cue

    Private Function Button2_Click_1_subFilesFix3(ByVal cue As String, ByVal inCueBinFilename As String, ByVal binList As List(Of String)) As Integer
        Dim cueFname As String = cue.Substring(cue.LastIndexOf("\") + 1)
        Dim cueFnameWoExt = cueFname.Substring(0, cueFname.LastIndexOf(".")).ToLower
        Class1.askVar1 = cue
        Class1.askVar2 = inCueBinFilename
        Class1.askVar3 = binList(0)
        Dim same As Boolean = False
        Dim dir As String = cue.Substring(0, cue.LastIndexOf("\"))
        If Not dir.EndsWith("\") Then dir = dir + "\"
        If binList.Count = 1 Then
            If isofix(5) = 0 Then
                ISOFixAsk_cueWoBin.ShowDialog()
                If Class1.askResponse > 100 Then isofix(5) = Class1.askResponse - 100 : Class1.askResponse = Class1.askResponse - 100
            Else
                Class1.askResponse = isofix(5)
            End If
            If Class1.askResponse = 1 Then
                Microsoft.VisualBasic.FileIO.FileSystem.RenameFile(dir + "\" + Class1.askVar3, Class1.askVar2)
                'same = cueFnameWoExt = Class1.askVar2.Substring(0, Class1.askVar2.LastIndexOf(".")).ToLower
                Form1.ListBox3.Items.Remove(".CUE without .BIN (provided match): " + cueFname + " (" + dir + ")")
                'If same Then Form1.ListBox3.Items.Remove(".CUE without .BIN (filename match): " + cueFname + " (" + dir + ")")
                ifSame(cueFname, Class1.askVar2, ".CUE without .BIN (filename match): " + cueFname + " (" + dir + ")")
                Return 0
            ElseIf Class1.askResponse = 2 Then
                Class1.associate_rewriteCue(Class1.askVar1, Class1.askVar3)
                'same = cueFnameWoExt = Class1.askVar3.Substring(0, Class1.askVar3.LastIndexOf(".")).ToLower
                Form1.ListBox3.Items.Remove(".CUE without .BIN (provided match): " + cueFname + " (" + dir + ")")
                'If same Then Form1.ListBox3.Items.Remove(".CUE without .BIN (filename match): " + cueFname + " (" + dir + ")")
                ifSame(cueFname, Class1.askVar3, ".CUE without .BIN (filename match): " + cueFname + " (" + dir + ")")
                Return 0
            End If
            Return -1
        Else
            Class1.askList = binList
            If isofix(6) <> 100000 Then
                ISOFixAsk_cueWoBin_ext.ShowDialog()
                If Class1.askResponse >= 100000 Then isofix(6) = 100000 : Class1.askResponse = Class1.askResponse - 100000
            Else
                Class1.askResponse = 0
            End If
            If Class1.askResponse > 10000 Then
                Microsoft.VisualBasic.FileIO.FileSystem.RenameFile(dir + "\" + binList(Class1.askResponse - 10001), Class1.askVar2)
                'same = cueFnameWoExt = Class1.askVar2.Substring(0, Class1.askVar2.LastIndexOf(".")).ToLower
                Form1.ListBox3.Items.Remove(".CUE without .BIN (provided match): " + cueFname + " (" + dir + ")")
                'If same Then Form1.ListBox3.Items.Remove(".CUE without .BIN (filename match): " + cueFname + " (" + dir + ")")
                ifSame(cueFname, Class1.askVar2, ".CUE without .BIN (filename match): " + cueFname + " (" + dir + ")")
                Return Class1.askResponse - 10001
            Else
                Dim binFileName As String = binList(Class1.askResponse - 1)
                Class1.associate_rewriteCue(Class1.askVar1, binFileName)
                'same = cueFnameWoExt = binFileName.Substring(0, binFileName.LastIndexOf(".")).ToLower
                Form1.ListBox3.Items.Remove(".CUE without .BIN (provided match): " + cueFname + " (" + dir + ")")
                'If same Then Form1.ListBox3.Items.Remove(".CUE without .BIN (filename match): " + cueFname + " (" + dir + ")")
                ifSame(cueFname, binFileName, ".CUE without .BIN (filename match): " + cueFname + " (" + dir + ")")
                Return Class1.askResponse - 1
            End If
            Return -1
        End If
    End Function 'Check ISO FIX .CUE wo .BIN

    Private Function ifSame(ByVal cueFname As String, ByVal binFname As String, Optional ByVal remove As String = "") As Boolean
        Dim same As Boolean
        Dim cueFnameWoExt = cueFname.Substring(0, cueFname.LastIndexOf("."))
        Dim binFnameWoExt = binFname.Substring(0, binFname.LastIndexOf("."))

        same = cueFnameWoExt.ToLower = binFnameWoExt.ToLower
        If remove <> "" Then Form1.ListBox3.Items.Remove(remove)
        Dim res As Integer = cueWoBin_filename_mismatch.IndexOf(cueFname)
        If same Then
            If res <> -1 Then cueWoBin_filename_mismatch.RemoveAt(res) : cueWoBin_filename_mismatch_binFname.RemoveAt(res)
        Else
            If res = -1 Then cueWoBin_filename_mismatch.Add(cueFname) : cueWoBin_filename_mismatch_binFname.Add(binFname)
        End If
        Return same
    End Function

    Private Sub Button2_Click_1_subFilesFix4(ByVal dir As String, ByVal cueFname As String, ByVal binFname As String)
        Dim f1WOext, f2WOext, f1Ext, f2Ext As String
        f1Ext = cueFname.Substring(cueFname.LastIndexOf(".") + 1)
        f2Ext = binFname.Substring(binFname.LastIndexOf(".") + 1)
        f1WOext = cueFname.Substring(0, cueFname.LastIndexOf("."))
        f2WOext = binFname.Substring(0, binFname.LastIndexOf("."))

        Class1.askVar1 = dir
        Class1.askVar2 = cueFname
        Class1.askVar3 = binFname
        If isofix(7) = 0 Then
            ISOFixAsk_cueWoBinFileMismatch.ShowDialog()
            If Class1.askResponse > 99000 Then Class1.askResponse = Class1.askResponse - 100000 : isofix(7) = Class1.askResponse
        Else
            Class1.askResponse = isofix(7)
        End If
        If Class1.askResponse = 3 Then Exit Sub
        Try
            If Class1.askResponse = 1 Then
                Microsoft.VisualBasic.FileIO.FileSystem.RenameFile(dir + "\" + binFname, f1WOext + "." + f2Ext)
                Class1.associate_rewriteCue(dir + "\" + cueFname, f1WOext + "." + f2Ext)
            End If
            If Class1.askResponse = 2 Then
                Microsoft.VisualBasic.FileIO.FileSystem.RenameFile(dir + "\" + cueFname, f2WOext + "." + f1Ext)
            End If
            If Not dir.EndsWith("\") Then dir = dir + "\"
            Form1.ListBox3.Items.Remove(".CUE without .BIN (filename match): " + cueFname + " (" + dir + ")")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub 'CUE BIN filenames mismatch fix

    Private Function Autodetect_bin_format(ByVal f As String) As String
        Dim datamode As Integer = 0
        Dim bytePerSector As Integer = 0
        Dim header() As Byte = {0, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 0, 0}
        Dim header_iso() As Byte = {1, &H43, &H44, &H30, &H30, &H31, 1, 0}
        Dim content() As Byte = {1, &H43, &H44, &H30, &H30, &H31, 1, 0}
        Dim streamBinary As New IO.FileStream(f, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.None, 36000)
        Dim readerInput As IO.BinaryReader = New IO.BinaryReader(streamBinary)
        Dim inputData As Byte() = readerInput.ReadBytes(65536)
        streamBinary.Close()
        readerInput.Close()
        If inputData.Take(13).SequenceEqual(header) Then
            datamode = inputData(15)
            bytePerSector = 2352
            'Form1.ListBox3.Items.Add(f + " - is - BIN (" + datamode.ToString + "/" + bytePerSector.ToString + ")")
        ElseIf inputData.Skip(&H930).Take(13).SequenceEqual(header) Then
            datamode = inputData(&H93F)
            bytePerSector = 2352
            'Form1.ListBox3.Items.Add(f + " - is - BIN (" + datamode.ToString + "/" + bytePerSector.ToString + ")")
        ElseIf inputData.Skip(&H8000).Take(8).SequenceEqual(header_iso) Then
            datamode = 1
            bytePerSector = 2048
            'Form1.ListBox3.Items.Add(f + " - is - ISO (1/2048)")
        Else
            datamode = 0
            bytePerSector = 0
            'Form1.ListBox3.Items.Add(f + " - is - .bin AUDIO")
            Return "AUDIO"
        End If
        Return datamode.ToString + "/" + bytePerSector.ToString
    End Function
End Class
