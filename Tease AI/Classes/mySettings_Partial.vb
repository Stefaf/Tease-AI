﻿'===========================================================================================
'
'                                           MySettings
'
' This file contains functions to extend the mySettings-Object with the capabilities to
' store and load the current user.config-file to/from a designated filepath.
' 
' The Local usersetting-file is duplicated on saving into the application-subdirectory.
' On startup this duplicated file is used to replace the user.config file in the 
' %localAppData%-directory.
'
' For safely importing Setting-files from other versions there is also an importfunction 
' included. This funciton will ask the user to select a file to import and restarts the 
' application in order to process the import and data-upgrade.
'
'
' To use those functions add a reference to System.Configuration and call StartupCheck() 
' in the my.Application.StartUp-Eventhandler.
' 
'===========================================================================================
Imports System.ComponentModel
Imports System.Configuration
Imports System.IO

Namespace My

    Partial Class MySettings

        ''' =========================================================================================================
        ''' <summary>
        ''' Determins the path the path to store and load the user.config-file from/to.
        ''' </summary>
        Private Shared BackupDir As String = Application.Info.DirectoryPath & "\System\Settings\"

        ''' =========================================================================================================
        ''' <summary>
        ''' Procedure to check whether to load or import a custom user.config-file. 
        ''' </summary>
        Friend Shared Sub StartupCheck()
            Dim importSettingFile As String = Application.CommandLineArgs.FirstOrDefault(Function(x) x.StartsWith("ImportSettings-"))

            If importSettingFile IsNot Nothing Then
                importConfig(importSettingFile)
            Else
                loadCustomUserConfig()
            End If
        End Sub

        ''' =========================================================================================================
        ''' <summary>
        ''' Replace the user.config-file located in %LocalAppData% with the duplicated filöe.
        ''' </summary>
        Private Shared Sub loadCustomUserConfig()
            Try
                Dim configAppData As String = GetLocalFilepath()
                Dim configAppDataDir As String = Path.GetDirectoryName(configAppData)

                Dim dupeFilePath As String = GetDuplicatePath()
                Dim dupeFileDir As String = Path.GetDirectoryName(dupeFilePath)

                ' check if there is a duplicate file 
                If Directory.Exists(dupeFileDir) AndAlso File.Exists(dupeFilePath) Then

                    'check if the directory in %LocalAppData% exits and create it if not
                    If Directory.Exists(configAppDataDir) = False Then _
                                Directory.CreateDirectory(configAppDataDir)

                    ' Copy the duplicated file to %LoaclAppData%-Dir.
                    File.Copy(dupeFilePath, configAppData, True)

                End If
            Catch ex As Exception
                '▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨
                '                                            All Errors
                '▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨
                MsgBox("Exception on loading custom-user.config." & vbCrLf & ex.Message,
                       MsgBoxStyle.Exclamation, "Load user.config")
                Log.WriteError(ex.Message, ex, "Exception on loading custom-user.config.")
            End Try
        End Sub

#Region "---------------------------------------MyBaseRelated--------------------------------------------"

        ''' =========================================================================================================
        ''' <summary>
        ''' Raises the SettingsSaving Event and copies afterwards the File to the 
        ''' specific filepath.
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Protected Overrides Sub OnSettingsSaving(sender As Object, e As CancelEventArgs)
            MyBase.OnSettingsSaving(sender, e)
            Try
                Dim configAppDataPath As String = GetLocalFilepath()
                Dim dupeFilePath As String = GetDuplicatePath()

                ' Check if Directory and file to copy exist.
                If Directory.Exists(Path.GetDirectoryName(configAppDataPath)) _
                AndAlso File.Exists(configAppDataPath) Then

                    ' Create Traget directoy if needed.
                    If Directory.Exists(BackupDir) = False Then _
                    Directory.CreateDirectory(BackupDir)

                    ' Copy File
                    My.Computer.FileSystem.CopyFile(configAppDataPath, dupeFilePath, True)
                End If
            Catch ex As Exception
                '▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨
                '                                            All Errors
                '▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨
                Log.WriteError(ex.Message, ex, "Exception while duplicating user.config.")
            End Try
        End Sub

        Shadows Sub Reset()

            Dim dupeFilePath As String = GetDuplicatePath()

            If Directory.Exists(BackupDir) = True AndAlso File.Exists(dupeFilePath) Then
                File.Delete(dupeFilePath)
            End If

            MyBase.Reset()
        End Sub


#End Region ' MyBaseRelated

#Region "------------------------------------ General Functions -----------------------------------------"

        Friend Shared Function GetDuplicatePath() As String

            ' get the Filepath 
            Dim SettingsFilePath As String = GetLocalFilepath()
            Dim TargetPath As String = BackupDir & Application.Info.Version.ToString & "." & Path.GetFileName(SettingsFilePath)

            Return TargetPath
        End Function

        ''' =========================================================================================================
        ''' <summary>
        ''' Returns the current user.config-file path.
        ''' </summary>
        ''' <returns>A String representing the path to the current user.config-file.</returns>
        Friend Shared Function GetLocalFilepath() As String

            Dim roamingConfig As Configuration =
            ConfigurationManager.OpenExeConfiguration(
            ConfigurationUserLevel.PerUserRoamingAndLocal)

            Dim filePAth As String = roamingConfig.FilePath

            Return filePAth

        End Function

#End Region ' General functions 

#Region "---------------------------------------- Import file -------------------------------------------"

        ''' =========================================================================================================
        ''' <summary>
        ''' Asks the user for an user-contig file to import. If a file is selected, the current application
        ''' is hard-stopped and restarted with new CommandLine Args. Existing CommandLineArgs are overwritten.
        ''' </summary>
        Friend Shared Sub importOnRestart()
            Try
                Dim fs As New OpenFileDialog With
                {.Filter = "config|*.config",
                .Multiselect = False,
                .Title = "Select config file to import.",
                .InitialDirectory = Reflection.Assembly.GetExecutingAssembly().CodeBase}

                If fs.ShowDialog = DialogResult.OK And fs.FileName <> "" And File.Exists(fs.FileName) Then
                    ' Restart the application with new Start-parameters
                    Dim startInfo As New ProcessStartInfo()
                    startInfo.FileName = Reflection.Assembly.GetExecutingAssembly().CodeBase
                    startInfo.Arguments = "ImportSettings-""" & fs.FileName & """"

                    If MessageBox.Show("Do you want to import the setting from file: " & vbCrLf &
                                       fs.FileName & vbCrLf &
                                        "If you continue, the program will restart and load the specified config-file. " &
                                        "This will overwrite your current settings. " & vbCrLf & vbCrLf &
                                        "Would you like to continue?",
                                       "Import Settings",
                                       MessageBoxButtons.OKCancel,
                                       MessageBoxIcon.Question,
                                       MessageBoxDefaultButton.Button2) = DialogResult.OK Then

                        Process.Start(startInfo)
                        Process.GetCurrentProcess().Kill()
                    End If
                End If
            Catch ex As Win32Exception When ex.ErrorCode = -2147467259
                '▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨
                '                                      Process.Start() cancelled
                '▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨
                MsgBox("Import has benn cancelled.", MsgBoxStyle.Information)
            Catch ex As Exception
                '▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨
                '                                            All Errors
                '▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨
                Log.WriteError(ex.Message, ex, "Exception while determining the import file.")
                MsgBox("An exception occured while determining the import file: " &
                       ex.Message, MsgBoxStyle.Exclamation, "Exception occured")
            End Try
        End Sub

        ''' =========================================================================================================
        ''' <summary>
        ''' Performs an Upgrade from a given UserConfig.file.
        ''' </summary>
        ''' <param name="filepath">The File to import.</param>
        ''' <remarks>In order to perform and upgrade, a pseudo-Version-number is calculated. This
        ''' Version-number is 1 step smaller as the current version. This calculated Version-number is
        ''' used to create a new version-folder in the %LocalAppData%-Folder. If there is already
        ''' another folder with this Version Number the user has to confirm overwriting. </remarks>
        Private Shared Sub importConfig(ByVal filepath As String)
            Try
                If filepath = Nothing Then Exit Sub

                Dim fileToLoad As String = filepath.Replace("ImportSettings-", "")

                If File.Exists(fileToLoad) Then
                    '▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
                    '                     Calculate-Previous-Version-Start 
                    '▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
                    Dim splitVersion() As String = Application.Info.Version.ToString.Split(".")
                    Dim calcVersion As New List(Of Integer)
                    Dim copyRest As Boolean = True

                    For i = splitVersion.Count - 1 To 0 Step -1
                        Dim currNumber As Integer = Integer.Parse(splitVersion(i))
                        Dim prevNumber As Integer = currNumber - 1

                        If prevNumber = -1 And i <> 0 Then
                            ' Number was 0 => convert to 9999 if not Major Number
                            calcVersion.Insert(0, 9999)
                        ElseIf prevNumber < currNumber
                            ' Number is smaller than the current Number => Copy rest of Numbers
                            copyRest = True
                            calcVersion.Insert(0, prevNumber)
                        Else
                            Throw New ArgumentException("Unknown case while calculation previous Version.")
                        End If
                    Next

                    Dim prevVersion As New Version(String.Join(".", calcVersion))

                    If prevVersion >= Application.Info.Version Then
                        Throw New ArithmeticException("The calculated version number is not smaller than the current.")
                    End If
                    '▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
                    '    Calculate-Previous-Version-END 
                    '▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

                    ' Get directory path for the current user.config-file.
                    Dim ImportDir As String = Path.GetDirectoryName(GetLocalFilepath())

                    ' Determine the destination directory 
                    ImportDir = Path.GetDirectoryName(ImportDir) & "\" & prevVersion.ToString & "\"

                    ' Extract the filename, if something changes over time
                    Dim targetFile As String = Path.GetFileName(GetLocalFilepath)

                    ' Ask for confirmation if there is already a directory.
                    If Directory.Exists(ImportDir) Then
                        If MsgBox("There is already a directory '""" & ImportDir &
                                """ if you continue, the content in this directory " &
                                " will be overriden. " & vbCrLf &
                                "Do you wish to overwrite the content?",
                                  MsgBoxStyle.YesNo, "Overwrite Content") = MsgBoxResult.No Then
                            MsgBox("Import has been cancelled.", MsgBoxStyle.Information, "Import")
                            Exit Sub
                        End If
                    End If

                    ' Create target directory 
                    Directory.CreateDirectory(ImportDir)

                    ' Copy the file to import
                    File.Copy(fileToLoad, ImportDir & targetFile, True)

                    ' Perform a Settings-Upgrade.
                    My.Settings.Upgrade()

                    ' Delete directory and all content
                    Directory.Delete(ImportDir, True)
                End If

            Catch ex As Exception
                '▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨
                '                                            All Errors
                '▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨▨
                Log.WriteError(ex.Message, ex, "Import settings")
                MsgBox("An exception occured while importing settings: " & vbCrLf & ex.Message,
                       MsgBoxStyle.Exclamation, "Import settings")
            End Try
        End Sub

#End Region  ' Import file

    End Class
End Namespace