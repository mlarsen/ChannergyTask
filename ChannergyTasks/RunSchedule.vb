Imports System.Data
Imports System.Data.Odbc
Imports System.IO
Imports System.Collections
Imports System.Text.RegularExpressions
Imports System.Text
Imports System.Timers
Imports System.Threading

Module RunSchedule
    'Public stPath As String = "E:\Core\HeltonTools\2020"
    Public stPath As String = System.AppDomain.CurrentDomain.BaseDirectory()
    Public stImportTableName As String
    Public stImportFolderPath As String
    Public stImportFileFilter As String
    Public stExeFilePath As String
    Public stScriptName As String
    Public stExportTableName As String
    Public stExportFolderPath As String
    Public bolIsTimeStamp As Boolean


    Sub Main()
        Dim stDay As String
        Dim stDayofMonth As String
        Dim tSendTime As TimeSpan
        Dim tTime As TimeSpan
        Dim tTimeDiff As TimeSpan
        Dim tLastRun As TimeSpan
        Dim cmd As OdbcCommandBuilder
        Dim stFrequency As String
        Dim stNextScheduled As String
        Dim stLastRun As String
        Dim tSendDate As Date
        Dim tDate As Date
        Dim iMCMVersion As Integer
        Dim iTotalMinutes As Integer
        Dim iTotalHours As Integer
        Dim stInstalledVersion As String = Reflection.Assembly.GetExecutingAssembly.GetName.Version.Major.ToString + "." + Reflection.Assembly.GetExecutingAssembly.GetName.Version.Minor.ToString + "." + Reflection.Assembly.GetExecutingAssembly.GetName.Version.Build.ToString + "." + Reflection.Assembly.GetExecutingAssembly.GetName.Version.Revision.ToString
        Dim clArgs() As String = Environment.GetCommandLineArgs
        Dim bolDebug As Boolean = False

        If clArgs.Length = 2 Then
            If Strings.UCase(clArgs(1)) = "DEBUG" Then
                bolDebug = True
            End If

        End If

        'Make sure that the stPath is formatted correctly
        If Right(stPath, 1) <> "\" Then
            stPath = stPath + "\"
        End If


        'For testing purposes

        Dim stLogFile As String = stPath + "ChannergyTasks.txt"
            Dim file As System.IO.StreamWriter



        iMCMVersion = GetMCMVersion(stPath)
        If iMCMVersion < 1909 Then
            Console.WriteLine("You must use MCM verion 1909 or higher.")
            Exit Sub
        End If

        'Get the connection string
        stODBCString = CoreFunctions.GetDSN(stPath)

        'Code added 05/25/2021: Updated the create table sub to update all of the related tables if necessary.
        CreateTables()

        Do While True
            If bolDebug = True Then
                file = My.Computer.FileSystem.OpenTextFileWriter(stLogFile, True)
            End If


            Dim con As New OdbcConnection(stODBCString)
            Dim da As New OdbcDataAdapter("SELECT *,CAST((CURRENT_TIMESTAMP-IF(LastRun=NULL,CURRENT_TIMESTAMP,LastRun))/(1000*60*60) AS INTEGER) AS Hours  FROM ChannergyScripts WHERE IsActive=True;", con)
            Dim ds As New DataSet()
            da.Fill(ds, "RunScripts")

            'Get the values for stDay,iDay and tTime
            stDay = Today.ToString("MM/dd/yyyy")
            stDayofMonth = DateAndTime.Day(Today)
            tTime = TimeSpan.Parse(DateTime.Now.ToString("HH:mm"))
            tDate = Today.ToString("MM/dd/yyyy")

            'Code added 08/11/2020: Check to see if the service has been checked for updates today
            If CheckLast("ChannergyTasks", tDate) = False Then 'See if there are any updates
                If IsNewVersion("ChannergyTasks", stInstalledVersion) = True Then
                    'If My.Computer.FileSystem.FileExists(stPath + "UpdateService.exe") = False Then
                    DownloadApplicationFiles("UpdateService", stPath)
                    'End If

                    Dim startInfo = New ProcessStartInfo(stPath + "UpdateService.exe")
                    startInfo.WindowStyle = ProcessWindowStyle.Normal
                    startInfo.Arguments = Chr(34) + "ChannergyTasks" + Chr(34)
                    startInfo.UseShellExecute = False
                    System.Diagnostics.Process.Start(startInfo)

                End If
            End If


            'Loop through the tables and see if we have any to send during this time
            For Each dr In ds.Tables("RunScripts").Rows
                stScriptName = dr.item("ScriptName").ToString
                stFrequency = dr.item("Frequency").ToString
                stLastRun = dr.item("LastRun").ToString
                stNextScheduled = dr.item("NextScheduled").ToString
                stImportTableName = dr.item("ImportTableName").ToString
                stImportFolderPath = dr.item("ImportFolderPath").ToString
                stImportFileFilter = dr.item("ImportFileExtension").ToString
                stExeFilePath = dr.item("ExeFilePath").ToString
                stExportFolderPath = dr.item("ExportFilePath").ToString
                stExportTableName = dr.item("ExportTableName").ToString
                bolIsTimeStamp = dr.item("IsTimeStamp")
                iTotalHours = CInt(dr.item("Hours").ToString)

                'If stLastRun <> "Never" Then
                '    tLastrun = TimeSpan.Parse(stLastRun)
                '    tTimeDiff = (tTime - tLastrun)
                'End If

                'If stNextScheduled <> "" Then
                '    'Get the send time
                '    tSendTime = TimeSpan.Parse(dr.item("NextScheduled").ToString("MM/dd/yyyy HH:mm:ss tt"))
                'End If

                Console.WriteLine("Current Time:" + stDay + " " + tTime.ToString + " " + stScriptName + "," + stFrequency + ", NextScheduled:" + stNextScheduled)

                'For testing
                If bolDebug = True Then
                    file.WriteLine("Current Time:" + stDay + " " + tTime.ToString + " " + stScriptName + "," + stFrequency + ", NextScheduled:" + stNextScheduled)
                End If


                If stFrequency = "Daily" Then
                    tSendTime = convTime(stScriptName, "Time")
                    tTimeDiff = (tTime - tSendTime)


                    If stNextScheduled <> "" Then
                        tSendDate = convDateStamp(stScriptName, "NextScheduled")
                    End If

                    If stNextScheduled = "" And tTime >= tSendTime Then 'run the script now
                        If bolDebug = True Then
                            file.WriteLine("Running " + stScriptName)
                        End If

                        RunScript(stScriptName, stImportTableName, stExeFilePath)
                    ElseIf (tDate = tSendDate And tTime >= tSendTime) Or (iTotalHours >= 24 And iTotalHours <= 25) Then
                        If bolDebug = True Then
                            file.WriteLine("Running " + stScriptName)
                        End If

                        RunScript(stScriptName, stImportTableName, stExeFilePath)
                    End If

                ElseIf stFrequency = "Hourly" Or stFrequency = "Every 30 Min" Then
                    If stNextScheduled <> "" Then
                        tSendTime = convTimeStamp(stScriptName, "NextScheduled")
                        tSendDate = convDateStamp(stScriptName, "NextScheduled")
                        tLastRun = convTimeStamp(stScriptName, "LastRun")
                        tTimeDiff = (tTime - tSendTime)
                        iTotalMinutes = Math.Abs(tTimeDiff.TotalMinutes)

                        'If (tTime >= tSendTime And tDate >= tSendDate And tTimeDiff.TotalHours >= 1) Or (tDate > tSendDate) Then
                        '    RunScript(stScriptName, stImportTableName)
                        'End If

                        If iTotalMinutes <= 5 Or iTotalMinutes > 60 Or (tDate > tSendDate) Then
                            If bolDebug = True Then
                                file.WriteLine("Running " + stScriptName)
                            End If

                            RunScript(stScriptName, stImportTableName, stExeFilePath)
                        End If
                    Else
                        If bolDebug = True Then
                            file.WriteLine("Running " + stScriptName)
                        End If

                        RunScript(stScriptName, stImportTableName, stExeFilePath)
                    End If
                ElseIf stFrequency = "Every 2 Hours" Then
                    If stNextScheduled <> "" Then
                        tSendTime = convTimeStamp(stScriptName, "NextScheduled")
                        tSendDate = convDateStamp(stScriptName, "NextScheduled")
                        tLastRun = convTimeStamp(stScriptName, "LastRun")
                        tTimeDiff = (tTime - tSendTime)
                        iTotalMinutes = Math.Abs(tTimeDiff.TotalMinutes)

                        'If (tTime >= tSendTime And tDate >= tSendDate And tTimeDiff.TotalHours >= 1) Or (tDate > tSendDate) Then
                        '    RunScript(stScriptName, stImportTableName)
                        'End If

                        If iTotalMinutes <= 5 Or iTotalMinutes > 120 Or (tDate > tSendDate) Then
                            If bolDebug = True Then
                                file.WriteLine("Running " + stScriptName)
                            End If

                            RunScript(stScriptName, stImportTableName, stExeFilePath)
                        End If
                    Else
                        If bolDebug = True Then
                            file.WriteLine("Running " + stScriptName)
                        End If

                        RunScript(stScriptName, stImportTableName)
                    End If
                ElseIf stFrequency = "Every 4 Hours" Then
                    If stNextScheduled <> "" Then
                        tSendTime = convTimeStamp(stScriptName, "NextScheduled")
                        tSendDate = convDateStamp(stScriptName, "NextScheduled")
                        tLastRun = convTimeStamp(stScriptName, "LastRun")
                        tTimeDiff = (tTime - tSendTime)
                        iTotalMinutes = Math.Abs(tTimeDiff.TotalMinutes)

                        'If (tTime >= tSendTime And tDate >= tSendDate And tTimeDiff.TotalHours >= 1) Or (tDate > tSendDate) Then
                        '    RunScript(stScriptName, stImportTableName)
                        'End If

                        If iTotalMinutes <= 5 Or iTotalMinutes > 240 Or (tDate > tSendDate) Then
                            If bolDebug = True Then
                                file.WriteLine("Running " + stScriptName)
                            End If

                            RunScript(stScriptName, stImportTableName, stExeFilePath)
                        End If
                    Else
                        If bolDebug = True Then
                            file.WriteLine("Running " + stScriptName)
                        End If

                        RunScript(stScriptName, stImportTableName, stExeFilePath)
                    End If
                End If

            Next
            'AppendScriptLog("Waiting...", "Waiting")
            con.Close()
            Console.WriteLine("Waiting...")

            If bolDebug = True Then
                file.WriteLine("Waiting...")
                file.Close()
            End If

            Thread.Sleep(60 * 1000)

        Loop
    End Sub
    Function GetMCMVersion(ByRef stPath As String) As Integer
        Dim stFileExePath As String = stPath + "MCM.exe"
        Dim stFileVersion As String = FileVersionInfo.GetVersionInfo(stFileExePath).FileVersion

        Return CInt(Right(stFileVersion, 4))

    End Function
    Function convTimeStamp(ByRef stScriptName As String, ByRef stFieldName As String, Optional ByRef isDate As Boolean = False) As TimeSpan
        Dim stSQL As String
        Dim tTime As TimeSpan

        If isDate = False Then
            stSQL = "SELECT CAST(EXTRACT(HOUR FROM " + stFieldName + ") AS VARCHAR(2))+':'+CAST(EXTRACT(MINUTE FROM " + stFieldName + ") AS VARCHAR(2))+':'+CAST(EXTRACT(SECOND FROM " + stFieldName + ") AS VARCHAR(2)) FROM ChannergyScripts WHERE ScriptName='" + stScriptName + "';"
            SQLTextQuery("S", stSQL, stODBCString, 1)
            If sqlArray(0) <> "::" Then
                tTime = TimeSpan.Parse(sqlArray(0))
            End If
        Else
            stSQL = "SELECT CAST(EXTRACT(MONTH FROM " + stFieldName + ") AS VARCHAR(2))+'/'+CAST(EXTRACT(DAY FROM " + stFieldName + ") AS VARCHAR(2))+'/'+CAST(EXTRACT(YEAR FROM " + stFieldName + ") AS VARCHAR(4)) FROM ChannergyScripts WHERE ScriptName='" + stScriptName + "';"
            If sqlarray(0) <> "" Then
                tTime = TimeSpan.Parse(sqlarray(0))
            End If
        End If


        Return tTime
    End Function
    Function convTime(ByRef stScriptName As String, ByRef stFieldName As String) As TimeSpan
        Dim stSQL As String
        Dim tTime As TimeSpan


        stSQL = "SELECT Time FROM ChannergyScripts WHERE ScriptName='" + stScriptName + "';"
        SQLTextQuery("S", stSQL, stODBCString, 1)
        If sqlArray(0) <> ":" Then
            tTime = TimeSpan.Parse(sqlArray(0))
        End If



        Return tTime
    End Function
    Function convDateStamp(ByRef stSearchValue As String, ByRef stDateFieldName As String, Optional ByRef stTableName As String = "ChannergyScripts", Optional ByRef stFieldName As String = "ScriptName") As Date
        Dim stSQL As String
        Dim tDate As Date

        stSQL = "SELECT CAST(EXTRACT(MONTH FROM " + stDateFieldName + ") AS VARCHAR(2))+'/'+CAST(EXTRACT(DAY FROM " + stDateFieldName + ") AS VARCHAR(2))+'/'+CAST(EXTRACT(YEAR FROM " + stDateFieldName + ") AS VARCHAR(4)) FROM " + stTableName + " WHERE " + stFieldName + "='" + stSearchValue + "';"
        SQLTextQuery("S", stSQL, stODBCString, 1)

        If sqlArray(0) <> "" Then
            tDate = Date.Parse(sqlArray(0))
        End If



        Return tDate
    End Function
    Sub RunScript(ByRef stScriptName As String, Optional stImportTableName As String = "", Optional stExeFilePath As String = "")

        'Reset the error flag
        bolIsError = False
        Console.WriteLine("Running: " + stScriptName)
        ChannergyStopMCM.Main()

        If bolIsError = False Then


            'If the stImportTableName is not blank then run the ImportTable function
            If stImportTableName <> "" Then
                RunScripts.ImportTable(stScriptName, stImportTableName, stImportFileFilter, stImportFolderPath)

                If bolIsError = False Then
                    RunScripts.ProcessFile(stScriptName)
                End If
            ElseIf stExeFilePath <> "" Then
                RunScripts.ProcessExe(stScriptName)
            Else
                RunScripts.ProcessFile(stScriptName)
            End If

            'Code added 05/25/2021: Add option to export a file.
            If stExportTableName <> "" And stExportFolderPath <> "" Then
                ExportTable()
            End If

            'If the scripts were completed update the time stamp
            'If bolIsError = False Then
            UpdateTimeStamp(stScriptName)
            'End If
            ChannergyStartMCM.Main()
        End If
    End Sub
    Sub ExportTable()
        Dim stDirectoryFile As String
        Dim stSQL As String
        Dim stExportFileName As String = stExportTableName + ".csv"

        'Check to see if the stExportFolder exists, if not create it
        If My.Computer.FileSystem.DirectoryExists(stExportFolderPath) = False Then
            My.Computer.FileSystem.CreateDirectory(stExportFolderPath)
        End If

        If Right(stExportFolderPath, 1) <> "\" Then
            stExportFolderPath = stExportFolderPath + "\"
        End If

        If bolIsTimeStamp = True Then
            stExportFileName = stExportFileName.Replace(".csv", "_" & System.DateTime.Now.ToString("yyyyMMdd") & ".csv")
        End If

        stDirectoryFile = stExportFolderPath + stExportFileName

        stSQL = "EXPORT TABLE " + stExportTableName + " "
        stSQL = stSQL + "TO " + Chr(34) + stDirectoryFile + Chr(34) + " "
        stSQL = stSQL + "WITH HEADERS;"

        SQLTextQuery("U", stSQL, stODBCString, 0)

        If bolIsError = False Then
            AppendScriptLog(stScriptName, "Table Exported", stExportFileName)
        End If

    End Sub
    Sub CreateTables()
        Dim stSQL As String

        stSQL = "CREATE TABLE IF NOT EXISTS ChannergyServiceUpdate("
        stSQL = stSQL + "ServiceName VARCHAR(30),"
        stSQL = stSQL + "LastChecked DATE DEFAULT CURRENT_DATE);"
        SQLTextQuery("U", stSQL, stODBCString, 0)

        'Create the ChannergyScripts table if it doesn't exist
        stSQL = "CREATE TABLE IF NOT EXISTS ChannergyScripts ("
        stSQL = stSQL + "ScriptNo AUTOINC,"
        stSQL = stSQL + "ScriptName VARCHAR(30),"
        stSQL = stSQL + "ImportTableName VARCHAR(30),"
        stSQL = stSQL + "CreateTables MEMO,"
        stSQL = stSQL + "Process MEMO,"
        stSQL = stSQL + "IsActive BOOLEAN DEFAULT False, "
        stSQL = stSQL + "PRIMARY KEY (ScriptName));"
        SQLTextQuery("U", stSQL, stODBCString, 0) '

        stSQL = "ALTER TABLE ChannergyScripts "
        stSQL = stSQL + "ADD COLUMN IF NOT EXISTS ImportFolderPath VARCHAR(255), "
        stSQL = stSQL + "ADD COLUMN IF NOT EXISTS ImportFileExtension VARCHAR(5) DEFAULT '*.*', "
        stSQL = stSQL + "ADD COLUMN IF Not EXISTS Frequency VARCHAR(20), "
        stSQL = stSQL + "ADD COLUMN IF Not EXISTS Day VARCHAR(20), "
        stSQL = stSQL + "ADD COLUMN IF Not EXISTS Time VARCHAR(10), "
        stSQL = stSQL + "ADD COLUMN IF Not EXISTS LastRun TIMESTAMP, "
        stSQL = stSQL + "ADD COLUMN IF Not EXISTS NextScheduled TIMESTAMP,"
        stSQL = stSQL + "ADD COLUMN IF NOT EXISTS ExeFilePath VARCHAR(255),"
        stSQL = stSQL + "ADD COLUMN IF NOT EXISTS ExportFilePath VARCHAR(255),"
        stSQL = stSQL + "ADD COLUMN IF NOT EXISTS IsTimeStamp BOOLEAN,"
        stSQL = stSQL + "ADD COLUMN IF NOT EXISTS ExportTableName VARCHAR(30);"
        SQLTextQuery("U", stSQL, stODBCString, 0)

        'Create the ScriptLog table if it does not exist
        stSQL = "CREATE TABLE IF Not EXISTS ScriptLog ("
        stSQL = stSQL + "LogNo AUTOINC,"
        stSQL = stSQL + "Date DATE DEFAULT CURRENT_DATE,"
        stSQL = stSQL + "Time TIME DEFAULT CURRENT_TIME,"
        stSQL = stSQL + "ScriptName VARCHAR(50),"
        stSQL = stSQL + "Status VARCHAR(50),"
        stSQL = stSQL + "IsError BOOLEAN DEFAULT False,"
        stSQL = stSQL + "ErrorMessage MEMO);"
        SQLTextQuery("U", stSQL, stODBCString, 0)

        'Update the ScriptLog table
        stSQL = "ALTER TABLE ScriptLog "
        stSQL = stSQL + "REDEFINE COLUMN Status VARCHAR(50);"
        SQLTextQuery("U", stSQL, stODBCString, 0)

        'Update 12/29/2019:Redefine the IsActive in the ChannergyScripts table to default to False
        stSQL = "ALTER TABLE ChannergyScripts "
        stSQL = stSQL + "REDEFINE COLUMN IsActive BOOLEAN DEFAULT False;"
        SQLTextQuery("U", stSQL, stODBCString, 0)

        'Update 03/16/2020:Add a Service field to the table
        stSQL = "ALTER TABLE ScriptLog "
        stSQL = stSQL + "ADD COLUMN IF Not EXISTS Service VARCHAR(50) AT 5;"
        SQLTextQuery("U", stSQL, stODBCString, 0)
    End Sub
    Function CheckLast(ByRef stApplication As String, ByRef tDate As Date) As Boolean
        Dim stSQL As String = "SELECT ServiceName,LastChecked FROM ChannergyServiceUpdate WHERE ServiceName='" + stApplication + "';"

        SQLTextQuery("S", stSQL, stODBCString, 2)

        If sqlArray(0) = "NoData" Or sqlArray(0) = "" Then 'The table has not been populated yet
            stSQL = "INSERT INTO ChannergyServiceUpdate(ServiceName) VALUES('" + stApplication + "');"
            SQLTextQuery("U", stSQL, stODBCString, 0)
            Return False
        ElseIf convDateStamp(stApplication, "LastChecked", "ChannergyServiceUpdate", "ServiceName") = tDate Then
            Return True
        ElseIf convDateStamp(stApplication, "LastChecked", "ChannergyServiceUpdate", "ServiceName") < tDate Then
            'Update the LastChecked field
            stSQL = "UPDATE ChannergyServiceUpdate SET LastChecked=CURRENT_DATE WHERE ServiceName='" + stApplication + "';"
            SQLTextQuery("U", stSQL, stODBCString, 0)

            Return False
        Else
            Return True
        End If


    End Function
End Module
