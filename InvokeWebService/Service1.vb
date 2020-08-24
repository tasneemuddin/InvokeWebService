Imports System.Timers
Imports System.Configuration
Imports System.Collections.Specialized
Imports System.Net
Imports System.IO
Imports System.Text
Imports System.Runtime.Serialization.Json


Public Class Service1
    WithEvents timerObj As New Timer
    ''' Install Web Service
    ''' Step 1 : Open command prompt in Administrator mode
    ''' Step 2 : Go to the the path, where serice exe file exist
    ''' Step 3 : Install Service   "C:\Windows\Microsoft.NET\Framework\v4.0.30319\installutil.exe" "InvokeWebService.exe"    
    ''' Step 4 : Uninstall Service "C:\Windows\Microsoft.NET\Framework\v4.0.30319\installutil.exe" -u "InvokeWebService.exe"
    Protected Overrides Sub OnStart(ByVal args() As String)
        Try
            ' Add code here to start your service. This method should set things
            ' in motion so your service can do its work.
            PrintServiceLog("Service Started")
            ScheduleTimer()
        Catch ex As Exception
            PrintServiceLog("OnStart()  Exception : " + ex.Message)
        End Try
    End Sub

    Private Sub ScheduleTimer()
        Try
            Dim NotificationTime As DateTime = Convert.ToDateTime(GetConnectionString("NotificationTime"))
            Dim now As DateTime = DateTime.Now
            Dim firstRun As DateTime = New DateTime(now.Year, now.Month, now.Day, NotificationTime.Hour, NotificationTime.Minute, NotificationTime.Second, NotificationTime.Millisecond)
            PrintServiceLog("Current Time : " + now.ToString)

            If now > firstRun Then
                firstRun = firstRun.AddDays(1)
                PrintServiceLog("First Run : " + firstRun.ToString)
            Else
                PrintServiceLog("First Run : " + firstRun.ToString)
            End If

            Dim timeToGo As TimeSpan = firstRun - now
            PrintServiceLog("Time to go : " + timeToGo.ToString)
            If timeToGo <= TimeSpan.Zero Then
                timeToGo = TimeSpan.Zero
            End If
            Me.timerObj.Interval = timeToGo.TotalMilliseconds
            Me.timerObj.Start()
        Catch ex As Exception
            Throw New Exception("ScheduleTimer() Exception : " + ex.Message)
        End Try
    End Sub

    Private Sub TimerElapsed(sender As System.Object, e As System.EventArgs) Handles timerObj.Elapsed
        Try
            'PrintServiceLog("Calling Web Service at : " + DateTime.Now.ToString)
            Me.timerObj.Interval = GetConnectionString("Intervel")

            '''Service call
            Dim token As String = GetAuthorizationToken()
            If String.IsNullOrWhiteSpace(token) Then Throw New Exception("token is null or empty")

            PrintServiceLog(SendRequest(New Uri(GetConnectionString("SubscriptioinExpireNotificationURL")), Nothing, token, "application/json", "GET"))
            PrintServiceLog(SendRequest(New Uri(GetConnectionString("Signout")), Nothing, token, "application/json", "DELETE"))

        Catch ex As Exception
            PrintServiceLog("TimerElapsed() Exception : " + ex.Message)
        End Try
    End Sub

    Protected Overrides Sub OnStop()
        Try
            ' Add code here to perform any tear-down necessary to stop your service.
            PrintServiceLog("Service Stopped")
            Me.timerObj.Stop()
            Me.timerObj.Dispose()
        Catch ex As Exception
            PrintServiceLog("OnStop()   Exception : " + ex.Message)
        End Try
    End Sub

    Private Sub PrintServiceLog(ByVal dataString As String)
        Try
            Dim DateString As String = DateTime.Today.Day.ToString + "-" + DateTime.Today.Month.ToString + "-" + DateTime.Today.Year.ToString
            Dim LogPath As String = AppDomain.CurrentDomain.BaseDirectory + "\Logs\"
            If (Not System.IO.Directory.Exists(LogPath)) Then
                System.IO.Directory.CreateDirectory(LogPath)
            End If
            Dim logFile As String = LogPath & "ServiceLog_" & DateString & ".Log"
            Using Owriter As IO.StreamWriter = New System.IO.StreamWriter(logFile, True)
                Owriter.WriteLine("TIME :" + DateTime.Now & " ---- :" & dataString & vbCrLf)
                Owriter.Close()
            End Using
        Catch ex As Exception
        End Try
    End Sub

    Public Function GetConnectionString(ByVal Key As String) As String
        Try
            Return ConfigurationManager.ConnectionStrings(Key).ConnectionString
        Catch ex As Exception
            Throw New Exception("GetConnectionString(" & Key & ") Exception : " + ex.Message)
        End Try
    End Function

    Public Sub SetConnectionString(ByVal Key As String, ByVal Value As String)
        Try
            ConfigurationManager.ConnectionStrings(Key).ConnectionString = Value
        Catch ex As Exception
            Throw New Exception("SetConnectionString(" & Key & ", " & Value & ") Exception : " + ex.Message)
        End Try
    End Sub

    Public Function GetAuthorizationToken() As String
        Try
            Dim jsonSring As String = "{""checkbox"":"""",""Password"":""" & GetConnectionString("Password") & """,""Email"":""" & GetConnectionString("Username") & """}"

            Dim data = Encoding.UTF8.GetBytes(jsonSring)

            Dim responceJSON = SendRequest(New Uri(GetConnectionString("LoginURL")), data, String.Empty, "application/json", "POST")

            Dim ResponseObject As ResponseObject = ConvertJsonToObject(responceJSON, GetType(ResponseObject))

            If IsNothing(ResponseObject.token) Then
                responceJSON = SendRequest(New Uri(GetConnectionString("ReplaceExistingUserTokenURL")), data, String.Empty, "application/json", "POST")
                ResponseObject = ConvertJsonToObject(responceJSON, GetType(ResponseObject))
            End If
            Return ResponseObject.token
        Catch ex As Exception
            Throw New Exception("CallWebService() Exception : " + ex.Message)
        End Try
    End Function

    Private Function SendRequest(ByVal uri As Uri, ByVal jsonDataBytes As Byte(), ByVal token As String, ByVal contentType As String, ByVal method As String) As String
        Try
            Dim response As String = String.Empty
            Dim request As WebRequest = WebRequest.Create(uri)
            request.ContentType = contentType
            request.Method = method

            If Not String.IsNullOrWhiteSpace(token) Then request.Headers.Add("Authorization", "Bearer " + token)

            If IsNothing(jsonDataBytes) Then
                Using responseStream = request.GetResponse.GetResponseStream
                    Using reader As New StreamReader(responseStream)
                        response = reader.ReadToEnd()
                    End Using
                End Using
            Else
                request.ContentLength = jsonDataBytes.Length
                Using requestStream = request.GetRequestStream
                    requestStream.Write(jsonDataBytes, 0, jsonDataBytes.Length)
                    requestStream.Close()

                    Using responseStream = request.GetResponse.GetResponseStream
                        Using reader As New StreamReader(responseStream)
                            response = reader.ReadToEnd()
                        End Using
                    End Using
                End Using
            End If
            Return response
        Catch ex As Exception
            Throw New Exception("SendRequest() Exception : " + ex.Message)
        End Try
    End Function

    Public Function ConvertJsonToObject(ByVal jsonString As String, ByVal type As System.Type) As Object
        Dim vbObj As Object
        Try
            Dim serializer As DataContractJsonSerializer = New DataContractJsonSerializer(type)
            Dim oMemoryStream As MemoryStream = New MemoryStream()
            Dim enc As New UTF8Encoding
            Dim arrBytData() As Byte = enc.GetBytes(jsonString)
            oMemoryStream.Write(arrBytData, 0, arrBytData.Length)
            oMemoryStream.Position = 0
            vbObj = serializer.ReadObject(oMemoryStream)
        Catch ex As Exception
            vbObj = Nothing
            Throw New Exception("ConvertJsonToObject() Exception : " + ex.Message)
        End Try
        Return vbObj
    End Function

    Public Class ResponseObject
        Public message As String
        Public token As String
    End Class

End Class