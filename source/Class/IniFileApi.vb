Imports System.Runtime.InteropServices
Imports System.Text

Public Class IniFileApi
    Public path As String
    <DllImport("kernel32.dll", SetLastError:=True)> Private Shared Function WritePrivateProfileString(ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Boolean
    End Function
    '<DllImport("kernel32.dll", SetLastError:=True)> Private Shared Function GetPrivateProfileString(ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As StringBuilder, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    'End Function
    <DllImport("kernel32.dll", EntryPoint:="GetPrivateProfileStringW", SetLastError:=True, CharSet:=CharSet.Unicode, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
    Private Shared Function GetPrivateProfileString(ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    End Function

    Public Sub IniFile(INIPath As String)
        path = INIPath
    End Sub

    Public Sub IniWriteValue(Section As String, Key As String, Value As String)
        WritePrivateProfileString(Section, Key, Value, path)
    End Sub

    Public Function IniReadValue(Section As String, Key As String) As String
        'Dim temp As New StringBuilder(255)
        Dim temp As New String(" "c, 255)
        Dim i As Integer = GetPrivateProfileString(Section, Key, "", temp, 255, path)
        Dim l As List(Of String) = temp.Split(Chr(0)).ToList

        'Dim t As String = temp.ToString().Trim()
        Dim t As String = ""
        If l.Count > 0 Then t = l(0) Else t = ""

        If t.Contains("//") Then
            If t.IndexOf("//") = 0 Then t = "" Else t = t.Substring(0, t.IndexOf("//"))
        End If
        'If t = "" Then MsgBox("Ini section=" + Section + "; key=" + Key + " not found. Exiting", MsgBoxStyle.Exclamation, "Hyperspin Checker") : Environment.Exit(1)
        Return t.Trim
    End Function

    Public Function IniListKey(Optional section As String = Nothing) As String()
        'Dim temp As New String(" "c, 4096)
        Dim temp As New String(" "c, 65536)
        'Dim i As Integer = GetPrivateProfileString(section, Nothing, "", temp, 4096, path)
        Dim i As Integer = GetPrivateProfileString(section, Nothing, "", temp, 65536, path)
        Dim l As List(Of String) = temp.Split(Chr(0)).ToList
        l.RemoveRange(l.Count - 2, 2)
        Return l.ToArray
    End Function
End Class
