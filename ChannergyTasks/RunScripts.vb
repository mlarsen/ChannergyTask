Imports System.Data.Odbc
Module RunScripts

    Public stOutputfile As String
    Sub ImportTable(ByRef stScriptName As String, ByRef stImportTableName As String, ByRef stFileFilter As String, ByRef stImportFolder As String)
        'First fix the stImportFolder
        If Right(stImportFolder, 1) <> "\" Then
            stImportFolder = stImportFolder + "\"
        End If

        Dim stSQL As String
        Dim stExportFilePath = stImportFolder + "\Processed\"
        Dim stProcessFilePath = stImportFolder + "\Prep\"
        Dim fi As IO.FileInfo
        Dim diP As New IO.DirectoryInfo(stImportFolder)
        Dim aryFiP As IO.FileInfo() = diP.GetFiles(stFileFilter)
        Dim FileContents As String
        Dim FileExt As String
        Dim stFileName As String

        'Reset the error flag   
        bolIsError = False

        'Make sure that the stProcessFilePath exists
        If My.Computer.FileSystem.DirectoryExists(stProcessFilePath) = False Then
            My.Computer.FileSystem.CreateDirectory(stProcessFilePath)
        End If

        AppendScriptLog(stScriptName, "Begin", False)
        'UpdateDialog(stScriptName + ":Start", frmTestScript.TextBox1)

        'Move the files into ProcessFilePath
        For Each fi In aryFiP
            My.Computer.FileSystem.MoveFile(fi.FullName, stProcessFilePath + fi.Name, True)
        Next

        'Create an array of files in the ProcessFilePath
        Dim di As New IO.DirectoryInfo(stProcessFilePath)
        Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")

        'Create Order tables
        CreateOrdertables(stScriptName)

        Try
            For Each fi In aryFi
                'Reset bolIsError
                bolIsError = False

                stFileName = fi.FullName

                If CheckIfDOS(stFileName) = False Then
                    ConvUnix2Dos(stFileName)
                Else
                    'Read source text file
                    FileContents = FileGet(stFileName)
                    FileExt = Right(stFileName, 4)
                    stOutputFile = Left(stFileName, Len(stFileName) - 4) + "_MW" + FileExt
                    FileWrite(stOutputFile, FileContents)
                End If

                If bolIsError = False Then
                    stSQL = "IMPORT TABLE " + stImportTableName + " "
                    stSQL = stSQL + "FROM " + Chr(34) + stOutputFile + Chr(34)
                    stSQL = stSQL + " WITH HEADERS;"

                    CoreFunctions.SQLTextQuery("U", stSQL, stODBCString, 0)

                    If bolIsError = False Then
                        AppendScriptLog(stScriptName, fi.Name, False)
                        ProcessFile(stScriptName)


                        If bolIsError = False Then
                            'If there Then are no errors, move the files To the processed folder
                            My.Computer.FileSystem.MoveFile(fi.FullName, stExportFilePath + fi.Name, True)

                            My.Computer.FileSystem.DeleteFile(stOutputFile)
                        End If

                    End If
                End If
            Next

        Catch ex As Exception
            AppendScriptLog(stScriptName, "Error", True, ex.Message)

        Finally
            AppendScriptLog(stScriptName, "Complete")

        End Try

    End Sub
    Sub ProcessFile(ByRef stScriptName As String)
        Dim stSQL As String
        Dim stCommand As String


        AppendScriptLog(stScriptName, "Begin")
        'Get the SQL from the ScriptEntry table
        stSQL = "SELECT Process FROM ChannergyScripts WHERE ScriptName='" + stScriptName + "';"
        CoreFunctions.SQLTextQuery("S", stSQL, stODBCString, 1)
        stCommand = sqlArray(0)

        CoreFunctions.SQLTextQuery("U", stCommand, stODBCString, 0)

        AppendScriptLog(stScriptName, "Complete")
    End Sub
    Sub ProcessExe(ByRef stScriptName As String)
        Dim stCommand As String
        Dim stSQL As String
        Dim procExecuting As New Process

        'Get the ExeFilePath from the ScriptEntry table
        stSQL = "SELECT ExeFilePath FROM ChannergyScripts WHERE ScriptName='" + stScriptName + "';"

        SQLTextQuery("S", stSQL, stODBCString, 1)
        stCommand = sqlArray(0)

        procExecuting = Process.Start(stCommand)
        procExecuting.WaitForExit()

    End Sub
    Sub UpdateTimeStamp(ByRef stScriptName As String)
        Dim stSQL As String
        Dim stFrequency As String
        Dim stScheduledTime As String

        'Update the last run timestamp
        stSQL = "UPDATE ChannergyScripts SET LastRun=CURRENT_TIMESTAMP WHERE ScriptName='" + stScriptName + "';"
        SQLTextQuery("U", stSQL, stODBCString)

        'Get the frequency and time
        stSQL = "SELECT Frequency,Time FROM ChannergyScripts WHERE ScriptName='" + stScriptName + "';"
        SQLTextQuery("S", stSQL, stODBCString, 2)
        stFrequency = sqlArray(0)
        stScheduledTime = sqlArray(1)

        If stFrequency = "Daily" Then
            'Code added 10/01/2020: Update the LastRun to be the current date and the scheduled time
            stSQL = "UPDATE ChannergyScripts SET LastRun=CAST((CAST(CURRENT_DATE AS VARCHAR(10))+' '+Time) AS TIMESTAMP) WHERE ScriptName='" + stScriptName + "';"
            SQLTextQuery("U", stSQL, stODBCString)

            stSQL = "UPDATE ChannergyScripts SET NextScheduled=LastRun+24*3600*1000 WHERE ScriptName='" + stScriptName + "';"
        ElseIf stFrequency = "Hourly" Then
            stSQL = "UPDATE ChannergyScripts SET NextScheduled=LastRun+3600*1000 WHERE ScriptName='" + stScriptName + "';"
        ElseIf stFrequency = "Every 2 Hours" Then
            stSQL = "UPDATE ChannergyScripts SET NextScheduled=LastRun+2*3600*1000 WHERE ScriptName='" + stScriptName + "';"
        ElseIf stFrequency = "Every 4 Hours" Then
            stSQL = "UPDATE ChannergyScripts SET NextScheduled=LastRun+4*3600*1000 WHERE ScriptName='" + stScriptName + "';"
        End If
        SQLTextQuery("U", stSQL, stODBCString)
    End Sub
    Function CheckIfDOS(ByVal stFileName As String) As Boolean
        Dim strFind As String = Chr(13) + Chr(10)
        Dim FileContents As String

        'Read source text file
        FileContents = FileGet(stFileName)

        'Find /r
        Dim FirstCharacter As Integer = FileContents.IndexOf(strFind)

        'If the <CR> character is found we can assume that the file has been converted to DOS format
        If FirstCharacter > 0 Then
            Return True
        Else
            Return False
        End If


    End Function
    Sub ConvUnix2Dos(ByVal stFileName As String)
        Dim strFind As String
        Dim strReplace As String
        Dim FileContents As String
        Dim dFileContents As String
        Dim FileExt As String

        FileExt = Right(stFileName, 4)

        stOutputFile = Left(stFileName, Len(stFileName) - 4) + "_MW" + FileExt

        strFind = Chr(10)
        strReplace = Chr(13) + Chr(10)
        'Read source text file
        FileContents = FileGet(stFileName)

        'replace all string In the source file
        dFileContents = Replace(FileContents, strFind, strReplace, 1, -1, 1)

        'Compare source And result
        If dFileContents <> FileContents Then
            'write result If different
            FileWrite(stOutputFile, dFileContents)

            'MsgBox("New file saved as " + OutputFile)
        End If

    End Sub

    Function FileGet(ByRef Filename As String)
        If Len(Filename) > 0 Then
            Dim FileStream As String
            FileStream = My.Computer.FileSystem.ReadAllText(Filename)
            Return FileStream
        End If
    End Function

    'Write string As a text file.
    Sub FileWrite(ByVal FileName, ByVal Contents)

        My.Computer.FileSystem.WriteAllText(FileName, Contents, False, System.Text.Encoding.ASCII)

    End Sub
    Sub CreateOrdertables(ByRef stScriptName As String)
        Dim stSQL As String = "SELECT CreateTables FROM ChannergyScripts WHERE ScriptName='" + stScriptName + "';"
        Dim stCommand As String

        'Get the SQL from the ChanenrgyScripts table
        SQLTextQuery("S", stSQL, stODBCString, 1)

        stCommand = sqlArray(0)

        SQLTextQuery("U", stCommand, stODBCString, 0)

    End Sub
    Sub AppendScriptLog(ByVal stScriptName As String, ByVal stStatus As String, Optional ByVal stService As String = "", Optional ByVal isError As Boolean = False, Optional ByVal stErrorMsg As String = "")
        Dim CRLF As String = Chr(13) + Chr(10)
        Dim con As New OdbcConnection(stODBCString)
        Dim da As New OdbcDataAdapter("SELECT ScriptName,Service,Status,IsError,ErrorMessage FROM ScriptLog;", con)
        Dim ds As New DataSet
        Dim dt As DataTable
        Dim cmd As OdbcCommandBuilder = New OdbcCommandBuilder(da)


        da.Fill(ds, "ScriptLog")

        dt = ds.Tables("ScriptLog")

        Dim newChannelLogRow As DataRow = dt.NewRow

        Try
            newChannelLogRow("ScriptName") = stScriptName
            newChannelLogRow("Status") = stStatus
            newChannelLogRow("Service") = stService

            If isError = False Then
                newChannelLogRow("IsError") = "False"
            Else
                newChannelLogRow("IsError") = "True"
            End If
            newChannelLogRow("ErrorMessage") = stErrorMsg

            dt.Rows.Add(newChannelLogRow)
            da.Update(ds, "ScriptLog")

        Catch ex As Exception
            Console.WriteLine("Exception Message: " & ex.Message)
        End Try
    End Sub
End Module
