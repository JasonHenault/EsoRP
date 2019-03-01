Imports System
Imports System.IO
Imports System.Data.SqlClient

Public Class Form1
    Dim EsoPath = "C:\Program Files (x86)\Zenimax Online\Launcher\Bethesda.net_Launcher.exe"

    Function ReadAddonDescriptor() As Object
        Dim version As String
        Dim addonFiles() As String = {}

        Try
            Dim descriptor As String = My.Computer.FileSystem.ReadAllText(".\EsoRP.txt")
            Dim descriptorLines() As String = Split(descriptor, Environment.NewLine)

            For i As Integer = 0 To descriptorLines.Length - 1
                If descriptorLines(i).IndexOf("## Version") <> -1 Then 'Looking for "version" indicator line
                    version = descriptorLines(i).Substring(12)
                ElseIf descriptorLines(i) = "" Then 'Looking for addon files enumeration beginning
                    For j As Integer = i + 1 To descriptorLines.Length - 1
                        If descriptorLines(j) <> "core\database.lua" Then
                            addonFiles = addonFiles.Concat({descriptorLines(j)}).ToArray 'Add each lines of path in addonFilePath
                        End If
                    Next
                    Exit For
                End If
            Next

        Catch ex As Exception
            MsgBox("Error : Check your internet connection. If the problem persists, please report the following error :" + vbCrLf + vbCrLf + "In ReadFileDescriptor() : " + ex.Message)
        End Try

        Return {version, addonFiles}
    End Function

    Private Sub AutoKill()
        Process.GetCurrentProcess().Kill()
    End Sub

    Private Sub UpdateOrKill()
        Try
            Dim eso64 = Process.GetProcessesByName("eso64")
            Dim eso = Process.GetProcessesByName("eso")
            Dim updater = Process.GetProcessesByName("Install EsoRP")
            Dim updaterNew = Process.GetProcessesByName("Install EsoRP_newVersion")

            For i = 0 To updater.Count() - 1
                If updater(i).Id <> Process.GetCurrentProcess().Id Then
                    updater(i).Kill()
                End If
            Next

            For i = 0 To updaterNew.Count() - 1
                If updaterNew(i).Id <> Process.GetCurrentProcess().Id Then
                    updaterNew(i).Kill()
                End If
            Next

            If eso64.Count() > 0 Or eso.Count() > 0 Then
                DatabaseUpdate()
                MsgBox("Database downloaded and user profil uploaded successfully.", Title:="EsoRP")
                Close()
            End If
        Catch ex As Exception
            MsgBox("Error. If the problem persists, please report the following error :" + vbCrLf + vbCrLf + "In UpdateOrKill() : " + ex.Message)
        End Try
    End Sub

    Private Sub AutoUpdate()
        Try
            If File.Exists("updating") Then 'If a new version is downloaded
                Dim esoRP = Process.GetProcessesByName("Install EsoRP")
                While esoRP.Count > 0 'wait for Install EsoRP.exe close
                    System.Threading.Thread.Sleep(100)
                    esoRP = Process.GetProcessesByName("Install EsoRP")
                End While
                File.Copy("Install EsoRP_newVersion.exe", "Install EsoRP.exe", True) 'Copy current updater
                File.Delete("updating")
                Process.Start("Install EsoRP.exe")
                AutoKill()
            Else
                If File.Exists("Install EsoRP_newVersion.exe") Then
                    Dim esoRPnew = Process.GetProcessesByName("Install EsoRP_newVersion")
                    While esoRPnew.Count > 0 'wait for Install EsoRP_newVersion.exe close
                        System.Threading.Thread.Sleep(100)
                        esoRPnew = Process.GetProcessesByName("Install EsoRP_newVersion")
                    End While
                    File.Delete("Install EsoRP_newVersion.exe")
                End If
                Dim currentVersion As String = ReadAddonDescriptor()(0)
                My.Computer.Network.DownloadFile("https://esorp.cnw.scot/addon/EsoRP.txt", ".\EsoRP.txt", "", "", True, 1000, True)
                Dim newDescriptor As Object = ReadAddonDescriptor()
                Dim newVersion As String = newDescriptor(0)
                Dim newAddonFiles() As String = newDescriptor(1)

                If newVersion <> currentVersion Then 'An update is avalaible
                    For i As Integer = 0 To newAddonFiles.Length - 1 'Update all files of the addon
                        My.Computer.Network.DownloadFile("https://esorp.cnw.scot/addon/" + newAddonFiles(i).Replace("\", "/"), ".\" + newAddonFiles(i), "", "", True, 1000, True)
                    Next

                    My.Computer.Network.DownloadFile("https://esorp.cnw.scot/addon/Install EsoRP.exe", ".\Install EsoRP_newVersion.exe", "", "", True, 1000, True) 'Download the new updater on another name
                    File.Create("updating")
                    Process.Start(".\Install EsoRP_newVersion.exe")
                    AutoKill()
                End If
            End If
        Catch ex As Exception
            MsgBox("Error : Check your internet connection. If the problem persists, please report the following error :" + vbCrLf + vbCrLf + "In AutoUpdate() : " + ex.Message)
        End Try
    End Sub

    Private Sub CreateShortcut()
        Dim oShell As Object
        Dim oLink As Object

        Try
            oShell = CreateObject("WScript.Shell")
            oLink = oShell.CreateShortcut(My.Computer.FileSystem.SpecialDirectories.Desktop & "\EsoRP.lnk")

            oLink.TargetPath = Directory.GetCurrentDirectory() + "\Install EsoRP.exe"
            oLink.WorkingDirectory = Directory.GetCurrentDirectory()
            oLink.Save()
        Catch ex As Exception
            MsgBox("Error. If the problem persists, please report the following error :" + vbCrLf + vbCrLf + "In CreateShortcut() : " + ex.Message)
        End Try
    End Sub

    Private Sub CheckSystem()
        UpdateOrKill()
        Directory.CreateDirectory("core")
        AutoUpdate()
        CreateShortcut()
    End Sub

    Private Sub LoadConfigFile()
        If Not File.Exists(".\config.ini") Then
            If Not File.Exists(EsoPath) And Not File.Exists("C:\Program Files (x86)\Steam\SteamApps\common\Zenimax Online\Launcher\Bethesda.net_Launcher.exe") Then
                Dim steamUsed = MessageBox.Show("No default ESO installation detected." + vbCrLf + "Do you launch ESO via Steam?", "EsoRP Updater - Use Steam?", MessageBoxButtons.YesNoCancel)
                If steamUsed = Global.System.Windows.Forms.DialogResult.Cancel Then
                    AutoKill()
                ElseIf steamUsed = Global.System.Windows.Forms.DialogResult.Yes Then
                    EsoPath = "steam://rungameid/306130"
                    My.Computer.FileSystem.WriteAllText(".\config.ini", EsoPath, False)
                    Exit Sub
                End If
                Dim folderBrowser = FolderBrowserDialog
                folderBrowser.Description = "Please select your ""Zenimax Online"" folder."
                If folderBrowser.ShowDialog() = Global.System.Windows.Forms.DialogResult.Cancel Then
                    AutoKill()
                End If
                EsoPath = folderBrowser.SelectedPath() + "\Launcher\Bethesda.net_Launcher.exe"
                If EsoPath = "\Launcher\Bethesda.net_Launcher.exe" Then
                    MsgBox("Vous devez choisir le dossier ""Zenimax Online"".", MsgBoxStyle.Critical)
                    LoadConfigFile()
                End If
                My.Computer.FileSystem.WriteAllText(".\config.ini", EsoPath, False)
            ElseIf File.Exists("C:\Program Files (x86)\Steam\SteamApps\common\Zenimax Online\Launcher\Bethesda.net_Launcher.exe") Then
                EsoPath = "steam://rungameid/306130"
            End If
        Else
            EsoPath = My.Computer.FileSystem.ReadAllText(".\config.ini")
            If Not File.Exists(EsoPath) And EsoPath <> "steam://rungameid/306130" Then
                My.Computer.FileSystem.DeleteFile(".\config.ini")
                LoadConfigFile()
            End If
        End If
    End Sub

    Private Sub UploadProfils()
        Try
            Dim forbidden = {"#", "%", "{", "}", "|", """", "^", ",", "~", "[", "]", "`", ",", "/", "?", ":", "@", "=", "&"}
            Dim allowed = {"%23", "%25", "%7B", "%7D", "%7C", "%22", "%5E", "%2C", "%7E", "%5B", "%5D", "%60", "%3B", "%2F", "%3F", "%3A", "%40", "%3D", "%26"}

            If File.Exists("../../SavedVariables/EsoRP.lua") Then
                Dim savedVars As String = My.Computer.FileSystem.ReadAllText("../../SavedVariables/EsoRP.lua")

                savedVars = savedVars.Substring(savedVars.IndexOf("{") + 3, savedVars.LastIndexOf("}") - (savedVars.IndexOf("{") + 3) - 2)
                savedVars = savedVars.Substring(savedVars.IndexOf("{") + 3, savedVars.LastIndexOf("}") - (savedVars.IndexOf("{") + 3) - 2)

                Dim userID = savedVars.Substring(savedVars.IndexOf("[""") + 3, savedVars.IndexOf("""]", savedVars.IndexOf("[""") + 3) - (savedVars.IndexOf("[""") + 3))

                savedVars = savedVars.Substring(savedVars.IndexOf("{") + 3, savedVars.LastIndexOf("}") - (savedVars.IndexOf("{") + 3) - 2)

                Dim eof = False
                While Not eof
                    Dim name = savedVars.Substring(savedVars.IndexOf("[""") + 2, savedVars.IndexOf("""]") - (savedVars.IndexOf("[""") + 2))

                    Dim counter = 1, position = savedVars.IndexOf("{"), last

                    While counter > 0
                        last = position
                        position = Math.Min(savedVars.IndexOf("{", position + 1), savedVars.IndexOf("}", position + 1))

                        If position = -1 Then
                            position = savedVars.Length
                            counter = -1
                            eof = True
                        Else
                            If savedVars(position) = "{" Then
                                counter = counter + 1
                            Else
                                counter = counter - 1
                            End If
                        End If
                    End While

                    Dim subSet As String = savedVars.Substring(savedVars.IndexOf("{"), position - savedVars.IndexOf("{"))

                    Dim server = subSet.Substring(subSet.IndexOf("[""server""] = """) + 14, (subSet.IndexOf(""",", subSet.IndexOf("[""server""] = """) + 14) - (subSet.IndexOf("[""server""] = """) + 14)))
                    Dim language = subSet.Substring(subSet.IndexOf("[""language""] = """) + 16, (subSet.IndexOf(""",", subSet.IndexOf("[""language""] = """) + 16) - (subSet.IndexOf("[""language""] = """) + 16)))
                    Dim race = subSet.Substring(subSet.IndexOf("[""race""] = """) + 12, (subSet.IndexOf(""",", subSet.IndexOf("[""race""] = """) + 12) - (subSet.IndexOf("[""race""] = """) + 12)))
                    Dim classe = subSet.Substring(subSet.IndexOf("[""class""] = """) + 13, (subSet.IndexOf(""",", subSet.IndexOf("[""class""] = """) + 13) - (subSet.IndexOf("[""class""] = """) + 13)))
                    Dim lore = subSet.Substring(subSet.IndexOf("[""lore""] = """) + 12, (subSet.IndexOf(""",", subSet.IndexOf("[""lore""] = """) + 12) - (subSet.IndexOf("[""lore""] = """) + 12)))

                    savedVars = savedVars.Substring(position)

                    For i = 0 To forbidden.Length() - 1
                        name = name.Replace(forbidden(i), allowed(i))
                        lore = lore.Replace(forbidden(i), allowed(i))
                    Next

                    name = name.Replace("'", "''") 'echapment caracter -> double it
                    lore = lore.Replace("'", "''") 'echapment caracter -> double it

                    name = name.Replace("\r", "") 'echapment caracter -> double it
                    lore = lore.Replace("\r", "") 'echapment caracter -> double it

                    If lore = "        [" Then
                        lore = ""
                    End If

                    Dim request As String = "https://esorp.cnw.scot/pushDB.php?name=" + name + "&server=" + server + "&userID=" + userID + "&language=" + language + "&race=" + race + "&class=" + classe + "&lore=" + lore
                    Dim webClient As New System.Net.WebClient
                    Dim result As String = webClient.DownloadString(request)
                End While
            End If
        Catch ex As Exception
            MsgBox("Error. If the problem persists, please report the following error :" + vbCrLf + vbCrLf + "In UploadProfils() : " + ex.Message)
            Close()
        End Try
    End Sub

    Private Sub DownloadDB()
        Try
            My.Computer.Network.DownloadFile("https://esorp.cnw.scot/getDB.php", ".\core\database.lua", "", "", True, 1000, True)
        Catch ex As Exception
            MsgBox("Error. If the problem persists, please report the following error :" + vbCrLf + vbCrLf + "In DownloadDB() : " + ex.Message)
            Close()
        End Try
    End Sub

    Private Sub DatabaseUpdate()
        UploadProfils()
        DownloadDB()
    End Sub

    Private Sub LaunchESO()
        If Process.GetProcessesByName("Bethesda.net_Launcher").Count() <= 0 Then
            Dim esoProcessInfo As New System.Diagnostics.ProcessStartInfo(EsoPath)
            If EsoPath <> "steam://rungameid/306130" Then
                esoProcessInfo.WorkingDirectory = EsoPath.ToString().Substring(0, EsoPath.ToString().Length() - 25) 'path - "Bethesda.net_Launcher.exe"
            End If
            Process.Start(esoProcessInfo)
            End If

            Dim launcher = Process.GetProcessesByName("Bethesda.net_Launcher")
        Dim eso = Process.GetProcessesByName("eso64")
        Dim update = False

        While launcher.Count <= 0 'wait for launcher open
            launcher = Process.GetProcessesByName("Bethesda.net_Launcher")
        End While

        While True
            While eso.Count <= 0 And launcher.Count > 0 'wait for eso open or launcher close
                System.Threading.Thread.Sleep(500)
                eso = Process.GetProcessesByName("eso64")
                launcher = Process.GetProcessesByName("Bethesda.net_Launcher")
            End While

            If launcher.Count <= 0 Then
                Return
            End If

            While eso.Count > 0 'wait for eso close
                System.Threading.Thread.Sleep(500)
                eso = Process.GetProcessesByName("eso64")
                update = True
            End While

            If update = True Then
                update = False
                DatabaseUpdate()
            End If

            System.Threading.Thread.Sleep(500)
        End While
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
#If DEBUG Then
#Else
        CheckSystem()
#End If
        LoadConfigFile()
        DatabaseUpdate()
        LaunchESO()
        Close()
    End Sub
End Class
