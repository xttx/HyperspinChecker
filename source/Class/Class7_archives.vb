Imports SevenZip
Imports Microsoft.VisualBasic.FileIO

Public Class Class7_archives
    Dim z As SevenZip.SevenZipExtractor
    Dim extArr() As String = {"7Z", "ZIP", "RAR"}
    Shared recompress_temp As String = Application.StartupPath + "\tmp"

    Public Sub setFile(f As String)
        'z = New SevenZip.SevenZipExtractor("O:\(DC)Rayman 2 The Great Escape[RUS].rar")
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

    Public Sub renameInArchive(f As String, newfileWoExtension As String, gamename As String, crc As String, Optional onlyKeepOne As Boolean = False)
        z = New SevenZip.SevenZipExtractor(f)
        Dim filelist = z.ArchiveFileData.ToList
        Dim dir As String = FileSystem.GetFileInfo(f).Directory.FullName

        'Dim arr = FileSystem.GetFiles("O:\7zipSharpTest\compress", SearchOption.SearchAllSubDirectories).ToArray
        'arr = arr.Concat(FileSystem.GetDirectories("O:\7zipSharpTest\compress").ToArray).ToArray
        Dim fExt As String = ""
        Dim fWoExt As String = ""
        'Dim arr() As String = {""}
        Dim index_to_rename As Integer = -1
        'If Not onlyKeepOne Then
        '    arr = z.ArchiveFileNames.ToArray
        '    For i As Integer = 0 To arr.Count - 1
        '        arr(i) = recompress_temp + "\" + arr(i)
        '        If Hex(z.ArchiveFileData(i).Crc).ToUpper = crc.ToUpper Then index_to_rename = i
        '    Next
        'Else
        '    For Each i In z.ArchiveFileData
        '        If Hex(i.Crc).ToUpper = crc.ToUpper Then
        '            index_to_rename = 0
        '            arr = {recompress_temp + "\" + i.FileName} : Exit For
        '        End If
        '    Next
        'End If

        'arr = z.ArchiveFileNames.ToArray
        Dim arr(z.ArchiveFileData.Count - 1) As String
        For i As Integer = 0 To z.ArchiveFileData.Count - 1
            Dim za = z.ArchiveFileData(i)
            If Hex(za.Crc).ToUpper = crc.ToUpper Then
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

        Dim ext As String = ""
        If arr(index_to_rename).Contains(".") Then ext = arr(index_to_rename).Substring(arr(index_to_rename).LastIndexOf("."))
        Dim newName As String = recompress_temp + "\" + gamename + ext
        FileSystem.MoveFile(arr(index_to_rename), newName, True)
        arr(index_to_rename) = newName

        'cast string to enum
        Dim arch_ext As String = ""
        Dim cFormat As OutArchiveFormat
        Dim cMethod As CompressionMethod
        Dim clevel As CompressionLevel
        Try
            If f.ToUpper.EndsWith(".ZIP") Then
                arch_ext = ".zip"
                cFormat = OutArchiveFormat.Zip
                cMethod = CompressionMethod.Default
            ElseIf f.ToUpper.EndsWith(".7Z") Then
                arch_ext = ".7z"
                cFormat = OutArchiveFormat.SevenZip
                cMethod = CompressionMethod.Default
            Else
                Dim frm As String = Form1.ComboBox10.SelectedItem.ToString
                Select Case frm.ToUpper
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
                cFormat = DirectCast([Enum].Parse(GetType(OutArchiveFormat), frm), OutArchiveFormat)
                cMethod = DirectCast([Enum].Parse(GetType(CompressionMethod), frm), CompressionMethod)
            End If
            clevel = DirectCast([Enum].Parse(GetType(CompressionLevel), Form1.ComboBox12.SelectedItem.ToString), CompressionLevel)
        Catch ex As Exception
            Dim str As String = Form1.ComboBox10.SelectedItem.ToString + ","
            str += Form1.ComboBox11.SelectedItem.ToString + "," + Form1.ComboBox12.SelectedItem.ToString
            MsgBox("Can't parse compression settings:" + vbCrLf + str) : Exit Sub
        End Try

        Dim zc As New SevenZipCompressor
        zc.ArchiveFormat = cFormat
        zc.CompressionMethod = cMethod
        zc.CompressionLevel = clevel

        'If crcToKeep <> "" Then
        'For i As Integer = 0 To arr.Count - 1
        'If arr(i).Contains("\") Then arr(i) = arr(i).Substring(arr(i).LastIndexOf("\") + 1)
        'Next
        'End If

        Try
            zc.CompressFiles(newfileWoExtension + arch_ext, arr)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        'FileSystem.DeleteDirectory(Dir + "\tmp", DeleteDirectoryOption.DeleteAllContents)
    End Sub
End Class
