Imports System.IO
Imports System.Security.Cryptography

Public Class Class6_hash
    'Get CRC32
    Public Shared Function GetCRC32(ByVal sFileName As String) As String
        Try
            Dim FS As FileStream = New FileStream(sFileName, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
            Dim CRC32Result As Integer = &HFFFFFFFF
            Dim Buffer(4096) As Byte
            Dim ReadSize As Integer = 4096
            Dim Count As Integer = FS.Read(Buffer, 0, ReadSize)
            Dim CRC32Table(256) As Integer
            Dim DWPolynomial As Integer = &HEDB88320
            Dim DWCRC As Long
            Dim i As Integer, j As Integer, n As Integer

            'Create CRC32 Table
            For i = 0 To 255
                DWCRC = i
                For j = 8 To 1 Step -1
                    If CBool(DWCRC And 1) Then
                        DWCRC = ((DWCRC And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
                        DWCRC = DWCRC Xor DWPolynomial
                    Else
                        DWCRC = ((DWCRC And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
                    End If
                Next j
                CRC32Table(i) = CInt(DWCRC)
            Next i

            'Calcualting CRC32 Hash
            Do While (Count > 0)
                For i = 0 To Count - 1
                    n = (CRC32Result And &HFF) Xor Buffer(i)
                    CRC32Result = ((CRC32Result And &HFFFFFF00) \ &H100) And &HFFFFFF
                    CRC32Result = CRC32Result Xor CRC32Table(n)
                Next i
                Count = FS.Read(Buffer, 0, ReadSize)
            Loop
            FS.Close()
            Return Hex(Not (CRC32Result))
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Shared Function hash_generator(ByVal hash_type As String, ByVal file_name As String) As String
        If Not My.Computer.FileSystem.FileExists(file_name) Then Return ""

        Dim hashValue() As Byte
        Dim fileStream As FileStream = File.OpenRead(file_name)
        fileStream.Position = 0

        If hash_type.ToLower = "md5" Then
            Dim hash = MD5.Create
            hashValue = hash.ComputeHash(fileStream)
        ElseIf hash_type.ToLower = "sha1" Then
            Dim hash = SHA1.Create()
            hashValue = hash.ComputeHash(fileStream)
        ElseIf hash_type.ToLower = "sha256" Then
            Dim hash = SHA256.Create()
            hashValue = hash.ComputeHash(fileStream)
        Else
            'MsgBox("Unknown type of hash : " & hash_type, MsgBoxStyle.Critical)
            fileStream.Close()
            Return ""
        End If

        Dim hash_hex = PrintByteArray(hashValue)
        fileStream.Close()
        Return hash_hex
    End Function

    Private Shared Function PrintByteArray(ByVal array() As Byte) As String
        Dim hex_value As String = ""

        ' We traverse the array of bytes
        Dim i As Integer
        For i = 0 To array.Length - 1
            ' We convert each byte in hexadecimal
            hex_value += array(i).ToString("X2")
        Next i

        ' We return the string in lowercase
        Return hex_value.ToLower
    End Function
End Class
