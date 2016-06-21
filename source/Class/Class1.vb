Imports Microsoft.VisualBasic.FileIO
Public Class Class1
    Public Shared i As Integer = 0
    Public Shared data As New List(Of String)
    Public Shared data_crc As New List(Of String)
    Public Shared romPath As String = ""
    Public Shared videoPath As String = ""
    Public Shared videoPathOrig As String = ""
    Public Shared HyperspinPath As String = ""
    Public Shared HyperspinIniCursysEmuExe As String = ""
    Public Shared HyperspinIniCursysEmuPath As String = ""
    Public Shared HyperspinIniCursysEmuExist As Boolean = False
    Public Shared HyperlaunchPath As String = ""
    Public Shared HyperlaunchExeName As String = ""
    Public Shared romlist As New List(Of String)
    Public Shared romFoundlist As New ArrayList
    Public Shared askVar1 As String = "", askVar2 As String = "", askVar3 As String = ""
    Public Shared askList As New List(Of String), askResponse As Integer = 0
    Public Shared confPath As String = Application.StartupPath + ".\Config.conf"
    Public Shared colorNO As Color = Color.OrangeRed
    Public Shared colorYES As Color = Color.LightGreen
    Public Shared colorPAR As Color = Color.FromArgb(&HFF, &HB0, &HFF, &HB0)

    Public Shared Sub Log(ByVal s As String)
        FileOpen(1, ".\HyperCheckerLog.txt", OpenMode.Append)
        PrintLine(1, DateTime.Now.Hour.ToString + ":" + DateTime.Now.Minute.ToString + ":" + DateTime.Now.Second.ToString + " - " + s)
        FileClose(1)
    End Sub

    Public Shared Sub associate_rewriteCue(ByVal filename As String, ByVal replaceImageBy As String)
        Dim line As String
        Dim firstBinaryHandled As Boolean = False
        Dim list As New List(Of String)
        FileOpen(1, filename, OpenMode.Input)
        Do While Not EOF(1)
            line = LineInput(1)
            If line.IndexOf("file", System.StringComparison.InvariantCultureIgnoreCase) >= 0 Then
                If line.IndexOf("binary", System.StringComparison.InvariantCultureIgnoreCase) >= 0 And Not firstBinaryHandled Then
                    Dim tmp As String = line.Substring(line.IndexOf("""") + 1)
                    tmp = tmp.Substring(0, tmp.LastIndexOf(""""))

                    'if filename contains path, and this path exist, than we have nothing to rename
                    If line.Contains("\") Then
                        If FileSystem.FileExists(line) Then Exit Sub
                    End If
                    line = line.Replace(tmp, replaceImageBy)
                    firstBinaryHandled = True
                End If
            End If
            list.Add(line)
        Loop
        FileClose(1)

        'do backup
        Dim backupN As Integer = 0
        If FileSystem.FileExists(filename + ".backup") Then
            Do While Microsoft.VisualBasic.FileIO.FileSystem.FileExists(filename + ".backup" + backupN.ToString)
                backupN += 1
            Loop
            FileSystem.CopyFile(filename, filename + ".backup" + backupN.ToString)
        Else
            FileSystem.CopyFile(filename, filename + ".backup")
        End If

        'write .cue file
        FileOpen(1, filename, OpenMode.Output)
        For Each s As String In list
            PrintLine(1, s)
        Next
        FileClose(1)
    End Sub 'Associate SUB

    Public Class XmlTextWriterEE
        Inherits Xml.XmlTextWriter
        Private ElementStack As New System.Collections.Stack

        Sub New(ByVal a As String)
            MyBase.New(a, System.Text.Encoding.Default)
        End Sub

        Public Overloads Overrides Sub WriteStartElement(ByVal prefix As String, ByVal localName As String, ByVal ns As String)
            Me.ElementStack.Push(localName)
            MyBase.WriteStartElement(prefix, localName, ns)
        End Sub

        Public Overrides Sub WriteEndElement()
            Dim localName As String = CStr(ElementStack.Pop)
            MyBase.WriteRaw("</" + localName + ">" + vbCrLf)
        End Sub

        Public Overrides Sub WriteFullEndElement()
            Dim localName As String = CStr(ElementStack.Pop)
            MyBase.WriteRaw("</" + localName + ">" + vbCrLf)
        End Sub
    End Class
End Class
