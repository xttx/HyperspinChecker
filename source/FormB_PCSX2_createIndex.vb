Imports Microsoft.VisualBasic.FileIO.FileSystem
Public Class FormB_PCSX2_createIndex
    'Private waitHandle As New System.Threading.AutoResetEvent(False)
    Private waitHandle As Boolean = False
    'Private curIsoFile As String = ""
    'Private curIsoPath As String = ""

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim fb As New OpenFileDialog
        If TextBox1.Text <> "" Then fb.FileName = TextBox1.Text
        fb.ShowDialog()
        TextBox1.Text = fb.FileName
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Dim fb As New FolderBrowserDialog
        If TextBox2.Text <> "" Then fb.SelectedPath = TextBox2.Text
        fb.ShowDialog()
        TextBox2.Text = fb.SelectedPath
    End Sub

    'START
    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        If Not FileExists(TextBox1.Text) Then MsgBox("File '" + TextBox1.Text + "' does not exist.") : Exit Sub
        If Not DirectoryExists(TextBox2.Text) Then MsgBox("Directory '" + TextBox2.Text + "' does not exist.") : Exit Sub

        enableBoxes(False)
        TextBox7.Text = ""
        Dim di As New IO.DirectoryInfo(TextBox2.Text)
        Dim fi() = di.GetFiles(TextBox6.Text)

        log("Found " + fi.Count.ToString + " files")

        Dim startinfo As New ProcessStartInfo
        startinfo.FileName = TextBox1.Text

        Dim threadWait As Threading.Thread = New Threading.Thread(AddressOf threadWaitSub)

        For Each f As IO.FileInfo In fi
            If FileExists(f.FullName + TextBox5.Text.Replace("*", "")) Then
                log("Tmp for " + f.Name + " found. Skipping.")
            Else
                log("Creating index for " + f.Name + "...")
                startinfo.Arguments = TextBox3.Text.Replace("%ISO_FILENAME%", """" + f.FullName + """")
                Dim p As Process = Process.Start(startinfo)

                Do While Not FileExists(f.FullName + TextBox5.Text.Replace("*", ""))
                    waitHandle = False
                    threadWait = New Threading.Thread(AddressOf threadWaitSub)
                    threadWait.Start()
                    Do While waitHandle = False
                        Application.DoEvents()
                    Loop
                Loop

                'Wait another couple of seconds just to be sure
                waitHandle = False
                threadWait = New Threading.Thread(AddressOf threadWaitSub)
                threadWait.Start()
                Do While waitHandle = False
                    Application.DoEvents()
                Loop

                p.Kill()
                log("Created")
            End If
        Next
        log("All done.")
        enableBoxes(True)
    End Sub

    Private Function evaluate_variables(v As String) As String
        Return ""
    End Function

    Private Sub threadWaitSub()
        Threading.Thread.Sleep(CInt(NumericUpDown1.Value) * 1000)
        waitHandle = True
    End Sub

    Private Sub log(s As String)
        TextBox7.AppendText(s + vbCrLf)
    End Sub

    Private Sub enableBoxes(e As Boolean)
        TextBox1.Enabled = e
        TextBox2.Enabled = e
        TextBox3.Enabled = e
        TextBox4.Enabled = e
        TextBox5.Enabled = e
        TextBox6.Enabled = e
    End Sub
End Class