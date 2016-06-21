Imports SevenZip
Imports Microsoft.VisualBasic.FileIO

Public Class Class7_archives
    Dim z As SevenZip.SevenZipExtractor
    Dim extArr() As String = {"7Z", "ZIP", "RAR"}
    Shared recompress_temp As String = Application.StartupPath + "\tmp"

    Public Sub setFile(f As String)
        z = New SevenZip.SevenZipExtractor(f)
    End Sub

    Public Function isArchive(fileName As String) As Boolean
        For Each ext As String In extArr
            If fileName.ToUpper.EndsWith("." + ext.ToUpper) Then Return True
        Next
        Return False
    End Function

    Public Function crc_list() As List(Of String)
        Dim l As New List(Of String)
        If z IsNot Nothing Then
            For Each f In z.ArchiveFileData
                l.Add(Hex(f.Crc).ToLower)
            Next
        End If
        Return l
    End Function

    Public Function ArchiveFileData() As List(Of ArchiveFileInfo)
        Return z.ArchiveFileData.ToList
    End Function
    Public Shared Sub set_recompress_temp(dir As String)
        recompress_temp = dir
    End Sub
    Public Sub delete_temp()
        FileSystem.DeleteDirectory(recompress_temp, DeleteDirectoryOption.DeleteAllContents)
    End Sub
    Public Enum answer
        yes
        no
        ask
        ask_once
        useDefaultForm1Settings
    End Enum

    Public Shared rename As answer = answer.useDefaultForm1Settings
    Public Shared keep_only_one As answer = answer.useDefaultForm1Settings
    Public Shared lastRespons_rename As MsgBoxResult = MsgBoxResult.Retry
    Public Shared lastRespons_keepOne As MsgBoxResult = MsgBoxResult.Retry
    Public Function renameInArchiveIfNeeded(newArchiveFileWoExtension As String, Optional crc As String = "") As Boolean
        Dim gameName As String = FileSystem.GetFileInfo(newArchiveFileWoExtension).Name
        If gameName.Contains(".") Then gameName = gameName.Substring(0, gameName.LastIndexOf("."))

        Class1.Log("Archive Handler - Archive name is - " + z.FileName)
        Dim rename_needed As Boolean = False
        Dim keep_only_one_needed As Boolean = False
        If z.ArchiveFileData.Count = 0 Then
            Class1.Log("Archive Handler - Archive empty!")
            Return False
        End If
        If crc <> "" Then
            For i As Integer = 0 To z.ArchiveFileData.Count - 1
                Dim za = z.ArchiveFileData(i)
                If Hex(za.Crc).ToUpper = crc.ToUpper Then
                    Dim fname As String = za.FileName
                    If fname.Contains("\") Then fname = fname.Substring(fname.LastIndexOf("\"))
                    If fname.Contains(".") Then fname = fname.Substring(0, fname.LastIndexOf("."))
                    If fname.ToUpper <> gameName.ToUpper Then rename_needed = True
                    Class1.Log("Archive Handler - CRC supplied, rename needed")
                End If
            Next
        ElseIf z.ArchiveFileData.Count = 1 Then
            Dim za = z.ArchiveFileData(0)
            crc = Hex(za.Crc).ToUpper

            Dim fname As String = za.FileName
            If fname.Contains("\") Then fname = fname.Substring(fname.LastIndexOf("\"))
            If fname.Contains(".") Then fname = fname.Substring(0, fname.LastIndexOf("."))
            If fname.ToUpper <> gameName.ToUpper Then rename_needed = True
            Class1.Log("Archive Handler - CRC NOT supplied, rename needed")
        Else
            rename_needed = True
            Class1.Log("Archive Handler - CRC NOT supplied, multiple files, _ASSUMING_ rename needed")
        End If
        If z.ArchiveFileData.Count > 1 Then
            Class1.Log("Archive Handler - Multiple files keep_only_one is needed")
            keep_only_one_needed = True
        End If
        If Not rename_needed And keep_only_one_needed Then
            Class1.Log("Archive Handler - not rename_needed neither keep_only_one is needed, exiting")
            Return False
        End If

        Dim frm As Form1 = DirectCast(Application.OpenForms("Form1"), Form1)
        Dim rename_in_archive As Boolean = False
        Dim onlyKeepOne As Boolean = False
        If rename_needed Then
            If rename = answer.useDefaultForm1Settings Then
                rename_in_archive = frm.CheckBox28.Checked
            ElseIf rename = answer.yes Then
                rename_in_archive = True
            ElseIf rename = answer.ask Or rename = answer.ask_once Then
                If rename = answer.ask Then lastRespons_rename = MsgBoxResult.Retry
                If lastRespons_rename = MsgBoxResult.Retry Then lastRespons_rename = MsgBox("At least one of files, INSIDE an ARCHIVE have incorrect name, do you want to rename files while copying?", MsgBoxStyle.YesNo)
                If lastRespons_rename = MsgBoxResult.Yes Then rename_in_archive = True
                Class1.Log("Archive Handler - Asked to allow rename files inside archive. Response - " + lastRespons_rename.ToString)
            End If
        End If
        If keep_only_one_needed Then
            If keep_only_one = answer.useDefaultForm1Settings Then
                onlyKeepOne = frm.CheckBox29.Checked
            ElseIf keep_only_one = answer.yes Then
                onlyKeepOne = True
            ElseIf keep_only_one = answer.ask Or keep_only_one = answer.ask_once Then
                If keep_only_one = answer.ask Then lastRespons_keepOne = MsgBoxResult.Retry
                lastRespons_keepOne = MsgBox("At least one of archive contains more than one files. Do you want to remove other files and only keep needed file?", MsgBoxStyle.YesNo)
                If lastRespons_keepOne = MsgBoxResult.Yes Then onlyKeepOne = True
                Class1.Log("Archive Handler - Asked to allow keep_only_one. Response - " + lastRespons_keepOne.ToString)
            End If
        End If

        If crc = "" And rename_in_archive Then rename_in_archive = False : Class1.Log("Archive Handler - Settings rename_in_archive Is enabled, but archive contains multiple files, And no crc was supplied. Renaming In archive Is impossible.")
        If crc = "" And onlyKeepOne Then onlyKeepOne = False : Class1.Log("Archive Handler - Settings only keep one Is enabled, but no crc was supplied. Keep_one Is impossible.")

        If rename_in_archive Or onlyKeepOne Then
            renameInArchive(newArchiveFileWoExtension, gameName, crc, onlyKeepOne)
            Return True
        Else
            Class1.Log("Archive Handler - nothing done, exiting.")
            Return False
        End If
    End Function

    Public Sub renameInArchive(newfileWoExtension As String, gamename As String, Optional crc As String = "", Optional onlyKeepOne As Boolean = False)
        Dim filelist = z.ArchiveFileData.ToList
        Dim f As String = z.FileName
        Dim dir As String = FileSystem.GetFileInfo(f).Directory.FullName

        Dim fExt As String = ""
        Dim fWoExt As String = ""
        Dim index_to_rename As Integer = -1

        Dim arr(z.ArchiveFileData.Count - 1) As String
        For i As Integer = 0 To z.ArchiveFileData.Count - 1
            Dim za = z.ArchiveFileData(i)
            If Not crc = "" AndAlso Hex(za.Crc).ToUpper = crc.ToUpper Then
                If onlyKeepOne Then
                    index_to_rename = 0
                    arr = {recompress_temp + "\" + za.FileName} : Exit For
                Else
                    index_to_rename = i
                End If
            End If
            arr(i) = recompress_temp + "\" + za.FileName
        Next
        z.ExtractArchive(recompress_temp)

        If index_to_rename >= 0 Then
            Dim ext As String = ""
            If arr(index_to_rename).Contains(".") Then ext = arr(index_to_rename).Substring(arr(index_to_rename).LastIndexOf("."))
            Dim newName As String = recompress_temp + "\" + gamename + ext
            FileSystem.MoveFile(arr(index_to_rename), newName, True)
            arr(index_to_rename) = newName
        End If

        'cast string to enum
        Dim arch_ext As String = ""
        Dim cFormat As OutArchiveFormat
        Dim cMethod As CompressionMethod
        Dim clevel As CompressionLevel
        Dim cFormatStr As String = ""
        Dim cMethodStr As String = ""
        Dim clevelStr As String = ""
        Try
            Dim frm As Form1 = DirectCast(Application.OpenForms("Form1"), Form1)
            If f.ToUpper.EndsWith(".ZIP") Then
                arch_ext = ".zip"
                cFormat = OutArchiveFormat.Zip
                cMethod = CompressionMethod.Default
            ElseIf f.ToUpper.EndsWith(".7Z") Then
                arch_ext = ".7z"
                cFormat = OutArchiveFormat.SevenZip
                cMethod = CompressionMethod.Default
            Else
                cFormatStr = frm.ComboBox10.SelectedItem.ToString()
                cMethodStr = frm.ComboBox10.SelectedItem.ToString()
                Select Case cFormatStr.ToUpper
                    Case "7Z"
                        arch_ext = "7z"
                    Case "BZIP2"
                        arch_ext = "bz2"
                    Case "GZIP"
                        arch_ext = "gz"
                    Case "TAR"
                        arch_ext = "tar"
                    Case "WIM"
                        arch_ext = "wim"
                    Case "XZ"
                        arch_ext = "xz"
                End Select
                cFormat = DirectCast([Enum].Parse(GetType(OutArchiveFormat), cFormatStr), OutArchiveFormat)
                cMethod = DirectCast([Enum].Parse(GetType(CompressionMethod), cMethodStr), CompressionMethod)
            End If

            clevelStr = frm.ComboBox12.SelectedItem.ToString
            clevel = DirectCast([Enum].Parse(GetType(CompressionLevel), clevelStr), CompressionLevel)
        Catch ex As Exception
            MsgBox("Can't parse compression settings:" + vbCrLf + cFormatStr + "," + cMethodStr + "," + clevelStr) : Exit Sub
        End Try

        Dim zc As New SevenZipCompressor
        zc.ArchiveFormat = cFormat
        zc.CompressionMethod = cMethod
        zc.CompressionLevel = clevel

        Try
            zc.CompressFiles(newfileWoExtension + arch_ext, arr)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
