Imports System.ComponentModel
Imports System.IO

Module Systems
    Public ReadOnly t_Powershell As String = "P"
    Public ReadOnly t_Command As String = "C"
    Public ReadOnly t_System As String = "S"
End Module

Module Module1
    ReadOnly v_Path As String = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments) & "\ServiceService"
    'ReadOnly v_FileName As String = Reflection.Assembly.GetEntryAssembly().Location
    Private v_FileWatcher As FileSystemWatcher = New FileSystemWatcher(v_Path)
    Private v_FileLastProcessed As DateTime = DateTime.MinValue
    Private v_FileLock As New Object

    Function GetOldPath(NewFilename As String) As String
        Try
            Return String.Concat(Path.GetFullPath(NewFilename), Path.GetFileNameWithoutExtension(NewFilename), "_bak", Path.GetExtension(NewFilename))
        Catch ex As Exception
            Log(ex)
            Return Nothing
        End Try
    End Function

    Sub Log(p_Command As String, p_Systems As String)
        Try
            Using sw As New StreamWriter(Path.Combine(v_Path, "Service.log"), True, System.Text.Encoding.UTF8)
                sw.WriteLine(String.Format("{1}{0}{3}{0}{2}",
                                           vbTab,
                                           DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                                           p_Command,
                                           p_Systems))
            End Using
        Catch ex As Exception
            'Fatal
        End Try
    End Sub

    Sub Log(p_Ex As Exception)
        Try
            If p_Ex IsNot Nothing Then
                Log(p_Ex.Message, t_System)
                Log(p_Ex.StackTrace, t_System)
            End If
        Catch ex As Exception
            'Fatal
        End Try
    End Sub

    Sub Main()
        Try
            If Not Directory.Exists(v_Path) Then
                Directory.CreateDirectory(v_Path)
                Log("Path created", t_System)
            End If

            v_FileWatcher.NotifyFilter = NotifyFilters.Attributes Or
                                           NotifyFilters.CreationTime Or
                                           NotifyFilters.DirectoryName Or
                                           NotifyFilters.FileName Or
                                           NotifyFilters.LastAccess Or
                                           NotifyFilters.LastWrite Or
                                           NotifyFilters.Security Or
                                           NotifyFilters.Size

            AddHandler v_FileWatcher.Changed, AddressOf OnFile
            AddHandler v_FileWatcher.Created, AddressOf OnFile
            AddHandler v_FileWatcher.Error, AddressOf OnError

            v_FileWatcher.Filter = "*.txt"
            v_FileWatcher.IncludeSubdirectories = False
            v_FileWatcher.EnableRaisingEvents = True

            Log("Started", t_System)

            Console.ReadLine()
        Catch ex As Exception
            Log(ex)
        End Try
    End Sub

    Private Sub OnError(sender As Object, e As ErrorEventArgs)
        Try
            Log(e.GetException())
        Catch ex As Exception
            Log(ex)
        End Try
    End Sub

    Private Sub OnFile(sender As Object, e As FileSystemEventArgs)
        Try
            Dim LastChange As DateTime = File.GetLastWriteTime(e.FullPath)
            SyncLock v_FileLock
                If Not v_FileLastProcessed = LastChange Then
                    Log(String.Format("Trigger for {0}", e.FullPath), t_System)
                    v_FileLastProcessed = LastChange
                    Select Case Path.GetFileNameWithoutExtension(e.FullPath)
                        Case t_Command
                            RunCMD()
                        Case t_Powershell
                            RunPowerShell()
                        Case t_System
                            RunSystem()
                    End Select
                End If
            End SyncLock
        Catch ex As Exception
            Log(ex)
        End Try
    End Sub

    Private Sub RunCMD()
        Try
            Using sr As New StreamReader(Path.Combine(v_Path, String.Format("{0}.txt", t_Command)))
                Dim line As String
                Do
                    line = sr.ReadLine
                    If line IsNot Nothing Then
                        Log(line, t_Command)
                        If line.Length > 0 Then
                            Dim Proz As New Process
                            Proz.StartInfo.FileName = "cmd"
                            Proz.StartInfo.Arguments = String.Format("/c {0}", line)
                            Proz.StartInfo.UseShellExecute = False
                            Proz.StartInfo.RedirectStandardOutput = True
                            Proz.StartInfo.RedirectStandardError = True
                            Proz.StartInfo.CreateNoWindow = True

                            Proz = Process.Start(Proz.StartInfo)

                            Log(Proz.StandardOutput.ReadToEnd, t_Command)
                            Log(Proz.StandardError.ReadToEnd, t_Command)
                            Proz.WaitForExit()
                        End If
                    End If
                Loop Until line Is Nothing
            End Using
        Catch ex As Exception
            Log(ex)
        End Try
    End Sub

    Private Sub RunPowerShell()
        Try
            Using sr As New StreamReader(Path.Combine(v_Path, String.Format("{0}.txt", t_Powershell)))
                Dim line As String
                Do
                    line = sr.ReadLine
                    If line IsNot Nothing Then
                        Log(line, t_Powershell)
                        If line.Length > 0 Then
                            Dim Proz As New Process
                            Proz.StartInfo.FileName = "powershell"
                            Proz.StartInfo.Arguments = String.Format("-Command {0}", line)
                            Proz.StartInfo.UseShellExecute = False
                            Proz.StartInfo.RedirectStandardOutput = True
                            Proz.StartInfo.RedirectStandardError = True
                            Proz.StartInfo.CreateNoWindow = True

                            Proz = Process.Start(Proz.StartInfo)

                            Log(Proz.StandardOutput.ReadToEnd, t_Powershell)
                            Log(Proz.StandardError.ReadToEnd, t_Powershell)
                            Proz.WaitForExit()
                        End If
                    End If
                Loop Until line Is Nothing
            End Using
        Catch ex As Exception
            Log(ex)
        End Try
    End Sub

    Sub RunSystem()
        Try
            Using sr As New StreamReader(Path.Combine(v_Path, String.Format("{0}.txt", t_System)))
                Dim line As String
                Dim Splitted As String()
                Do
                    line = sr.ReadLine
                    If line IsNot Nothing Then
                        Log(line, t_System)
                        Splitted = line.Split("#")
                        If File.Exists(Splitted(0)) Then
                            If Splitted.Length = 1 Then
                                Log(String.Format("Starte {0} ohne Parameter", Splitted(0)), t_System)
                                Process.Start(Splitted(0))
                            Else
                                Log(String.Format("Starte {0} mit Parameter {1}", Splitted(0), Splitted(1)), t_System)
                                Process.Start(Splitted(0), Splitted(1))
                            End If
                        End If
                    End If
                Loop Until line Is Nothing
            End Using
        Catch ex As Exception
            Log(ex)
        End Try
    End Sub
End Module
