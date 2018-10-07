Imports System
Imports System.IO
Imports System.Collections
Imports System.Runtime.InteropServices
Public Class BackUp
    'All Volume Control functions''''''
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Integer) As Short
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> Private Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function
    Const WM_APPCOMMAND As UInteger = &H319
    Const APPCOMMAND_VOLUME_UP As UInteger = &HA
    Const APPCOMMAND_VOLUME_DOWN As UInteger = &H9
    Const APPCOMMAND_VOLUME_MUTE As UInteger = &H8
    'All Volume Control functions''''''

    Public username As String = ((Environ$("Username")))
    Dim BackUpLocation As String = "C:\Users\Danyil\Documents\Year 2 Backups"
    Dim BaseFileNamePath As String = "C:\Users\Danyil\Documents\BackupFileID.txt"
    'The messages to look for.
    Private Const WM_DEVICECHANGE As Integer = &H219
    Private Const DBT_DEVICEARRIVAL As Integer = &H8000
    Private Const DBT_DEVICEREMOVECOMPLETE As Integer = &H8004
    Private Const DBT_DEVTYP_VOLUME As Integer = &H2  '
    'Get the information about the detected volume.
    Private Structure DEV_BROADCAST_VOLUME
        Dim Dbcv_Size As Integer
        Dim Dbcv_Devicetype As Integer
        Dim Dbcv_Reserved As Integer
        Dim Dbcv_Unitmask As Integer
        Dim Dbcv_Flags As Short
    End Structure

    Protected Overrides Sub WndProc(ByRef M As System.Windows.Forms.Message)
        'These are the required subclassing codes for detecting device based removal and arrival.
        If M.Msg = WM_DEVICECHANGE Then
            Select Case M.WParam
                'Check if a device was added.
                Case DBT_DEVICEARRIVAL
                    Dim DevType As Integer = Runtime.InteropServices.Marshal.ReadInt32(M.LParam, 4)
                    If DevType = DBT_DEVTYP_VOLUME Then
                        Dim Vol As New DEV_BROADCAST_VOLUME
                        Vol = Runtime.InteropServices.Marshal.PtrToStructure(M.LParam, GetType(DEV_BROADCAST_VOLUME))
                        If Vol.Dbcv_Flags = 0 Then
                            For i As Integer = 0 To 20
                                If Math.Pow(2, i) = Vol.Dbcv_Unitmask Then
                                    Dim Usb As String = Chr(65 + i) + ":\"
                                    '''''''BackUp Code will now start'''''''
                                    Dim msg = "USB Detected, Do you want to backup?"
                                    Dim title = "Backup?"
                                    Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2 Or _
                                 MsgBoxStyle.Question
                                    Dim response = MsgBox(msg, style, title)
                                    If response = MsgBoxResult.Yes Then
                                        CheckingFile()
                                        GenerateID()
                                        Dim objStreamReader As StreamReader
                                        Dim strLine As String
                                        objStreamReader = New StreamReader(BaseFileNamePath)
                                        'Read the first line of text.

                                        strLine = objStreamReader.ReadLine
                                        objStreamReader.Close()
                                        Dim IDAddress As String = BackUpLocation & "/" & strLine
                                        Try
                                            My.Computer.FileSystem.CopyDirectory(Usb.ToString & "USB", IDAddress, True) 'Here I decide for full back ups or just the crucial.
                                            MsgBox("Complete! Time:" & TimeOfDay)
                                            If Directory.Exists(BackUpLocation) = False Then
                                                Directory.CreateDirectory(BackUpLocation)
                                            End If
                                        Catch ex As Exception
                                            MsgBox(ex.Message)
                                        End Try
                                    Else
                                    End If
                                    '''''''Backup Code will finish'''''''''
                                    Exit For
                                End If
                            Next
                        End If
                    End If

                    'Check if the message was for the removal of a device.
                Case DBT_DEVICEREMOVECOMPLETE
                    Dim DevType As Integer = Runtime.InteropServices.Marshal.ReadInt32(M.LParam, 4)
                    If DevType = DBT_DEVTYP_VOLUME Then
                        Dim Vol As New DEV_BROADCAST_VOLUME
                        Vol = Runtime.InteropServices.Marshal.PtrToStructure(M.LParam, GetType(DEV_BROADCAST_VOLUME))
                        If Vol.Dbcv_Flags = 0 Then
                            For i As Integer = 0 To 20

                                If Math.Pow(2, i) = Vol.Dbcv_Unitmask Then
                                    Dim Usb As String = Chr(65 + i) + ":\"
                                    ''Scenario where the USB is removed
                                    Exit For
                                End If
                            Next
                        End If
                    End If
            End Select
        End If
        MyBase.WndProc(M)
    End Sub

    Public Sub GenerateID()
        Dim objStreamReader As StreamReader
        Dim strLine As String
        'Pass the file path and the file name to the StreamReader constructor.
        objStreamReader = New StreamReader(BaseFileNamePath)
        'Read the first line of text.
        strLine = objStreamReader.ReadLine
        objStreamReader.Close()
        Dim newInt As String
        Dim str As String = strLine   'Function extract number from string and increments by 1. 
        Dim ParsedInt As String       'Numbers without the text
        Try
            For Each c As Char In str
                If IsNumeric(c) Then
                    ParsedInt = ParsedInt & c
                    Convert.ToInt32(ParsedInt)
                    newInt = ParsedInt + 1
                    Dim cleanString As String = Replace(strLine, ParsedInt, "") 'Text without the numbers 
                    Dim CompleteID As String = cleanString & newInt
                    Dim objWriter As New System.IO.StreamWriter(BaseFileNamePath) 'Will now record


                    If CompleteID = "Backup10" Then
                        MessageBox.Show("Limit reached. Clearing all backups...")  ''Delete all back ups here
                        My.Computer.FileSystem.DeleteDirectory(BackUpLocation, FileIO.UIOption.AllDialogs, FileIO.RecycleOption.SendToRecycleBin)
                        My.Computer.FileSystem.CreateDirectory(BackUpLocation)
                        objWriter.Close()
                        newInt = 0
                    End If

                    objWriter.Write(CompleteID)
                    objWriter.Close()
                End If
            Next
        Catch ex As Exception
        End Try

        Try
            If File.Exists(BaseFileNamePath) = False Then
                File.Create(BaseFileNamePath)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub CheckingFile()
        Dim text As String = System.IO.File.ReadAllText(BaseFileNamePath)
        If text.Length = 0 Then
            Dim objWriter As New System.IO.StreamWriter(BaseFileNamePath)
            objWriter.WriteLine("Backup0")
            objWriter.Close()
        End If
    End Sub

    Private Sub BackUp_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        setStartup()

        If File.Exists(BackUpLocation) = False Then
            File.Create(BackUpLocation)

        End If
        If File.Exists(BaseFileNamePath) = False Then
            File.Create(BaseFileNamePath)
        End If
    
    End Sub

    Private Sub tmrKeyListen_Tick(sender As Object, e As EventArgs) Handles tmrKeyListen.Tick
        If GetAsyncKeyState(Keys.F9) Then
            SendMessage(Me.Handle, WM_APPCOMMAND, &H30292, APPCOMMAND_VOLUME_UP * &H10000) 'Up

        ElseIf GetAsyncKeyState(Keys.F10) Then
            SendMessage(Me.Handle, WM_APPCOMMAND, &H30292, APPCOMMAND_VOLUME_DOWN * &H10000) ' Down

        ElseIf GetAsyncKeyState(Keys.F11) Then
            SendMessage(Me.Handle, WM_APPCOMMAND, &H200EB0, APPCOMMAND_VOLUME_MUTE * &H10000) 'Mute
        End If
    End Sub

End Class
