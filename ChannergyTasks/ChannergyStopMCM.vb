Imports System.Threading
Imports System.ServiceProcess
Imports System.Data.Odbc
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Module ChannergyStopMCM

    Public stServiceName As String
    Sub Main()
        Dim bolIsComplete As Boolean = False
        Dim stSQL As String
        Dim stODBCString As String

        'Check to see if the path is valid
        If Right(stPath, 1) <> "\" Then
            stPath = stPath + "\"
        End If

        stODBCString = GetDSN(stPath)
        If stODBCString <> "" Then
            ' First check to see if the MCM is installed, if not there is nothing to do
            If CheckServiceExits("MCM") = True Then

                Try
                    Dim theMCM As System.ServiceProcess.ServiceController
                    theMCM = New System.ServiceProcess.ServiceController(stServiceName)

                    If theMCM.Status = ServiceControllerStatus.Running Then
                        'Check the status of the MCM
                        If GetStatus(stODBCString) <> "Paused" Then

                            'Delete any records from the ChannelMessage table
                            stSQL = "DELETE FROM ChannelMessage;"
                            SQLTextQuery("U", stSQL, stODBCString, 0)

                            'Add the Pause Channels to the ChannelMessage table
                            stSQL = "INSERT INTO ChannelMessage(ChannelNo,ChannelAccountNo,Request) "
                            stSQL = stSQL + "VALUES(0,0,'Pause Channels');"
                            SQLTextQuery("U", stSQL, stODBCString, 0)

                            Do While bolIsComplete = False
                                'Now check for the answer field
                                stSQL = "SELECT Request,Answer FROM ChannelMessage;"
                                SQLTextQuery("S", stSQL, stODBCString, 2)

                                If sqlarray(1) = "Command Received" Then 'The MCM has received the command


                                    If GetStatus(stODBCString) = "Paused" Then
                                        bolIsComplete = True
                                    Else
                                        bolIsComplete = False
                                    End If


                                Else
                                    bolIsComplete = False
                                    Thread.Sleep(5000)
                                End If


                            Loop
                            AppendScriptLog(stScriptName, "Paused", stServiceName)

                            'Delete any records from the ChannelMessage table
                            stSQL = "DELETE FROM ChannelMessage;"
                            SQLTextQuery("U", stSQL, stODBCString, 0)


                        End If
                    End If

                    'Stop other services if they are running
                    If CheckServiceExits("Mailware FTP Download Service") = True Then
                        Dim theFTPDownload As System.ServiceProcess.ServiceController
                        theFTPDownload = New System.ServiceProcess.ServiceController("Mailware FTP Download Service")

                        If theFTPDownload.Status = ServiceControllerStatus.Running Then
                            theFTPDownload.Pause()
                            theFTPDownload.Refresh()
                            AppendScriptLog(stScriptName, "Paused", stServiceName)
                        End If

                    End If

                    If CheckServiceExits("Mailware FTP Upload Service") = True Then
                        Dim theFTPUpload As System.ServiceProcess.ServiceController
                        theFTPUpload = New System.ServiceProcess.ServiceController("Mailware FTP Upload Service")

                        If theFTPUpload.Status = ServiceControllerStatus.Running Then
                            theFTPUpload.Pause()
                            theFTPUpload.Refresh()
                            AppendScriptLog(stScriptName, "Paused", stServiceName)
                        End If

                    End If
                Catch ex As Exception
                    Console.WriteLine(ex.Message)
                    AppendScriptLog(stScriptName, "Error", stServiceName, True, ex.Message)
                End Try
            End If
        End If
    End Sub
    Function CheckServiceExits(ByRef stService As String) As Boolean
        Dim services() As ServiceController = ServiceController.GetServices()
        Dim bolServiceInstalled As Boolean = False
        Dim iLenghth As Integer = Len(stService)
        For Each service As ServiceController In services
            If Left(service.DisplayName, iLenghth) = stService Then
                stServiceName = service.DisplayName
                bolServiceInstalled = True
            End If
        Next

        Return bolServiceInstalled
    End Function

    Function GetStatus(ByRef stODBCstring As String) As String
        'Now loop through the active MCM accounts and see what the last record is.
        Dim stSQL As String = "SELECT DISTINCT ChannelNo FROM ChannelAccounts WHERE IsProductUploadActive=True OR IsShippingConfirmationActive=True OR IsOrderDownloadActive=True;"
        Dim dt As New DataTable
        Dim da As New OdbcDataAdapter
        Dim con As New OdbcConnection(stODBCstring)
        Dim i As Integer
        Dim stChannelNo As String
        Dim LastRecord As String
        Dim stStatus As String
        Dim stTempStatus As String

        da = New OdbcDataAdapter(stSQL, con)
        da.Fill(dt)

        If dt.Rows.Count <> 0 Then


            For i = 0 To dt.Rows.Count - 1
                stChannelNo = dt.Rows(i).Item("ChannelNo").ToString
                'Check the last record for each channel
                stSQL = "SELECT MAX(ChannelLogNo) FROM ChannelLog WHERE ChannelNo=" + stChannelNo + ";"
                CoreFunctions.SQLTextQuery("S", stSQL, stODBCstring, 1)
                LastRecord = CoreFunctions.sqlArray(0)

                If LastRecord = "" Then 'There was no record for the channel account.
                    stTempStatus = "Paused"
                Else
                    'Get the status of the last record
                    stSQL = "SELECT Status FROM ChannelLog WHERE ChannelLogNo=" + LastRecord + ";"
                    CoreFunctions.SQLTextQuery("S", stSQL, stODBCstring, 1)
                    stTempStatus = CoreFunctions.sqlArray(0)
                End If


                If i = 0 Then
                    stStatus = stTempStatus
                Else
                    If stTempStatus <> stStatus Then
                        stStatus = stTempStatus
                    End If
                End If

            Next
            con.Close()
        Else
            stStatus = "Paused"
        End If

        Return stStatus

    End Function
End Module
