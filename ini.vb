Module ini

    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Int32, ByVal lpFileName As String) As Int32
    Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Int32

    Public Function load(ByVal filename As String, ByVal section As String, ByVal key As String, ByRef value As String) As Boolean
        ' Load config from specific INI file, Loaded -> True / Failed -> False
        Dim buffer As String = New String(" ", 1000)
        If GetPrivateProfileString(section, key, "", buffer, 1000, filename) = 0 Then
            value = ""
            Return False
        Else
            Dim i As Long = InStr(buffer, vbNullChar)
            value = Trim(Left(buffer, (i - 1)))
            Return True
        End If
    End Function

    Public Function save(ByVal filename As String, ByVal section As String, ByVal key As String, ByVal value As String) As Boolean
        ' Save config to specific INI file, Saved -> True / Failed -> False
        If WritePrivateProfileString(section, key, value, filename) = 0 Then
            Return False
        Else Return True
        End If
    End Function

End Module