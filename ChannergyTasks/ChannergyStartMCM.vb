Imports System.Threading
Imports System.ServiceProcess
Imports System.Data.Odbc
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Module ChannergyStartMCM

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
                        'Check to see if the MCM is in Pause mode
                        If GetStatus(stODBCString) = "Paused" Then
                            'Delete any records from the ChannelMessage table
                            stSQL = "DELETE FROM ChannelMessage;"
                            SQLTextQuery("U", stSQL, stODBCString, 0)

                            'Add the Resume Channels to the ChannelMessage table
                            stSQL = "INSERT INTO ChannelMessage(ChannelNo,ChannelAccountNo,Request) "
                            stSQL = stSQL + "VALUES(0,0,'Resume Channels');"
                            SQLTextQuery("U", stSQL, stODBCString, 0)

                            Do While bolIsComplete = False
                                'Now check for the answer field
                                stSQL = "SELECT Request,Answer FROM ChannelMessage;"
                                SQLTextQuery("S", stSQL, stODBCString, 2)

                                If sqlarray(1) = "Command Received" Then 'The MCM has received the command
                                    bolIsComplete = True
                                Else
                                    bolIsComplete = False
                                    Thread.Sleep(5000)
                                End If

                            Loop

                            AppendScriptLog(stScriptName, "Resumed", stServiceName)
                            'Delete any records from the ChannelMessage table
                            stSQL = "DELETE FROM ChannelMessage;"
                            SQLTextQuery("U", stSQL, stODBCString, 0)

                        End If

                        'Start other services if they are stopped
                        If CheckServiceExits("Mailware FTP Download Service") = True Then
                            Dim theFTPDownload As System.ServiceProcess.ServiceController
                            theFTPDownload = New System.ServiceProcess.ServiceController("Mailware FTP Download Service")

                            If theFTPDownload.Status = ServiceControllerStatus.Paused Then
                                theFTPDownload.Continue()
                                theFTPDownload.Refresh()
                                AppendScriptLog(stScriptName, "Resumed", stServiceName)
                            End If

                        End If

                        If CheckServiceExits("Mailware FTP Upload Service") = True Then
                            Dim theFTPUpload As System.ServiceProcess.ServiceController
                            theFTPUpload = New System.ServiceProcess.ServiceController("Mailware FTP Upload Service")

                            If theFTPUpload.Status = ServiceControllerStatus.Paused Then
                                theFTPUpload.Continue()
                                theFTPUpload.Refresh()
                                AppendScriptLog(stScriptName, "Resumed", stServiceName)
                            End If

                        End If
                    End If
                Catch ex As Exception
                    AppendScriptLog(stScriptName, "Error", stServiceName, True, ex.Message)
                    Console.WriteLine(ex.Message)
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
End Module
