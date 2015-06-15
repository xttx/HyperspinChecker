Imports System.Runtime.InteropServices
Imports System.Text

Public Class IniFileApi
    Public path As String
    <DllImport("kernel32.dll", SetLastError:=True)> Private Shared Function WritePrivateProfileString(ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Boolean
    End Function
    <DllImport("kernel32.dll", SetLastError:=True)> Private Shared Function GetPrivateProfileString(ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As StringBuilder, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    End Function

    Public Sub IniFile(INIPath As String)
        path = INIPath
    End Sub

    Public Sub IniWriteValue(Section As String, Key As String, Value As String)
        WritePrivateProfileString(Section, Key, Value, path)
    End Sub

    Public Function IniReadValue(Section As String, Key As String) As String
        Dim temp As New StringBuilder(255)
        Dim i As Integer = GetPrivateProfileString(Section, Key, "", temp, 255, path)

        Dim t As String = temp.ToString().Trim()
        If t.Contains("//") Then
            If t.IndexOf("//") = 0 Then t = "" Else t = t.Substring(0, t.IndexOf("//"))
        End If
        'If t = "" Then MsgBox("Ini section=" + Section + "; key=" + Key + " not found. Exiting", MsgBoxStyle.Exclamation, "Hyperspin Checker") : Environment.Exit(1)
        Return t
    End Function
End Class
