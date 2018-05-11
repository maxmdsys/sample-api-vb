
Imports System.IO
Imports System.ServiceModel

Public Class RetrieveDirectMessagesSample
    Public Property Auth As ServiceReference1.authentication = Nothing
    Public Property Username As String
    Public Property Password As String
    Public Property OutputFolder As String
    Public Property MailFolder As String
    Public Property EndpointURL As String

    Private Function getProxy() As ServiceReference1.DirectMessageServiceClient
        Dim proxy As New ServiceReference1.DirectMessageServiceClient
        If EndpointURL.StartsWith("https://", StringComparison.CurrentCultureIgnoreCase) Then
            Dim binding As New BasicHttpsBinding With {
                .MaxReceivedMessageSize = 10000000,
                .MaxBufferSize = 10000000,
                .MessageEncoding = WSMessageEncoding.Mtom
            }
            proxy.Endpoint.Binding = binding
        Else
            Dim binding As New BasicHttpBinding With {
                .MaxReceivedMessageSize = 10000000,
                .MaxBufferSize = 10000000,
                .MessageEncoding = WSMessageEncoding.Mtom
            }
            proxy.Endpoint.Binding = binding
        End If
        Dim endpointAddress As New EndpointAddress(EndpointURL)
        proxy.Endpoint.Address = endpointAddress
        Return proxy
    End Function

    Private Function getAuthentication() As ServiceReference1.authentication
        If Auth Is Nothing Then
            Dim authentication As New ServiceReference1.authentication With {
                .username = Username,
                .password = Password
            }
            Auth = authentication
        End If
        Return Auth
    End Function

    Private Sub ProcessErrorResponse(ByVal response As ServiceReference1.apiResponse)
        Select Case response.code
            Case 1
                Console.WriteLine("authentication failed")
            Case 2
                Console.WriteLine("Connection to IMAP server failed")
            Case 3
                Console.WriteLine("Unknown IMAP server host")
            Case 4
            Case 5
                Console.WriteLine("Other error")
        End Select
        Console.WriteLine("Information of the SOAP request: " + response.message)
    End Sub

    Private Sub ProcessCreateFolder(ByVal response As ServiceReference1.apiResponse, ByVal newFolder As String)
        If response.success Then
            Console.WriteLine("Create folder: {0}", newFolder)
        Else
            ProcessErrorResponse(response)
        End If
    End Sub

    Private Sub ProcessMessage(ByVal msg As ServiceReference1.messageDetail)
        Dim msgFolder As String = Path.Combine(OutputFolder, msg.msgnum.ToString)
        If Directory.Exists(msgFolder) Then
            Directory.Delete(msgFolder, True)
        End If
        Directory.CreateDirectory(msgFolder)
        Console.WriteLine("Subject: " + msg.subject)
        If msg.htmlBody Then
            File.WriteAllText(Path.Combine(msgFolder, "body.html"), msg.body)
        Else
            File.WriteAllText(Path.Combine(msgFolder, "body.txt"), msg.body)
        End If
        File.WriteAllText(Path.Combine(msgFolder, "subject.txt"), msg.subject)
        If msg.headers IsNot Nothing And msg.headers.Length > 0 Then
            Dim attachmentFolder As String = Path.Combine(msgFolder, "attached_documents")
            Directory.CreateDirectory(attachmentFolder)
            If msg.attachmentList IsNot Nothing Then
                For Each attachment As ServiceReference1.attachment In msg.attachmentList
                    Try
                        Dim filepath As String = Path.Combine(attachmentFolder, attachment.filename)
                        File.WriteAllBytes(filepath, attachment.content)
                    Catch e As UnauthorizedAccessException
                        Console.WriteLine(e)
                    End Try
                Next
            End If
        End If
        Console.WriteLine("Message detail saved in " + msgFolder)
        Console.WriteLine()
    End Sub

    Private Sub ProcessSetFolderSubscribed(ByVal response As ServiceReference1.apiResponse, ByVal folderName As String())
        If response.success Then
            For Each i As String In folderName
                Console.WriteLine("Set folder {0} subscribed", i)
            Next
        Else
            ProcessErrorResponse(response)
        End If
    End Sub

    Private Sub ProcessMessageResponse(ByVal response As ServiceReference1.getMessageResponse)
        If response.success Then
            If response.messages Is Nothing Then
                Console.WriteLine("No messages found")
            Else
                Console.WriteLine("Retrieved {0} message(s) from server", response.messages.Length)
                For i As Integer = 0 To response.messages.Length - 1
                    ProcessMessage(response.messages(i))
                Next
            End If
        Else
            ProcessErrorResponse(response)
        End If
    End Sub

    Private Function GetUnixTime(ByVal time As DateTime) As Long
        Dim unixTime As Long = time.Subtract(New Date(1970, 1, 1, 0, 0, 0, 0).ToLocalTime()).TotalMilliseconds
        Return unixTime
    End Function

    Private Sub ProcessCopyMessagesByUID(ByVal response As ServiceReference1.apiResponse, ByVal uids As Long())
        If response.success Then
            For i As Integer = 0 To uids.Length - 1
                Console.WriteLine("Copy messages by uid {0}", uids(i))
            Next
        Else
            ProcessErrorResponse(response)
        End If
    End Sub

    Private Sub ProcessSendMessageResponseType(ByVal response As ServiceReference1.sendMessageResponseType)
        If response.success Then
            Console.WriteLine("Send message response, smtp: {0}", response.smtpId)
        Else
            ProcessErrorResponse(response)
        End If
    End Sub

    Private Sub ProcessDirectMessageStatusReportResponseType(ByVal response As ServiceReference1.directMessageStatusReportResponseType)
        If response.hasCvData Then
            Console.WriteLine("csvData is {0}", response.csvData)
        Else
            If response.messageLogs IsNot Nothing Then
                Console.WriteLine("Retrieved {0} message(s) from server", response.messageLogs.Length)
                For i As Integer = 0 To response.messageLogs.Length - 1
                    ProcessMessageLog(response.messageLogs(i))
                Next
            Else
                Console.WriteLine("message log is null")
            End If
        End If
    End Sub

    Private Sub ProcessMessageLog(ByVal response As ServiceReference1.directMessageLog)
        Console.WriteLine("Message Log:")
        Console.WriteLine("from: {0}", response.from)
        Console.WriteLine("to: {0}", response.to)
        Console.WriteLine("status: {0}", response.status)
        Console.WriteLine("size: {0}", response.messageSize)
        Console.WriteLine("status detail: {0}", response.statusDetails)
    End Sub

    Private Sub ProcessMarkMessagesAsReadByUID(ByVal response As ServiceReference1.apiResponse, ByVal uids As Long())
        If response.success Then
            For Each uid As Long In uids
                Console.WriteLine("Mark messages as read by uid {0}", uid)
            Next
        Else
            ProcessErrorResponse(response)
        End If
    End Sub

    'get the number of messages
    Public Sub GetMessageCount()
        Dim proxy As ServiceReference1.DirectMessageServiceClient = getProxy()
        Dim response As ServiceReference1.getCountResponse = proxy.GetMessageCount(getAuthentication(), MailFolder)
        If response.success Then
            Console.WriteLine("Messages Count: " + response.count.ToString)
        Else
            ProcessErrorResponse(response)
        End If
    End Sub

    'get the number of unread messages
    Public Sub GetUnreadMessageCount()
        Dim proxy As ServiceReference1.DirectMessageServiceClient = getProxy()
        Dim response As ServiceReference1.getCountResponse = proxy.GetUnreadMessageCount(getAuthentication(), MailFolder)
        If response.success Then
            Console.WriteLine("Unread Messages Count: " + response.count.ToString)
        Else
            ProcessErrorResponse(response)
        End If
    End Sub

    'get the number of deleted messages
    Public Sub GetDeletedMessageCount()
        Dim proxy As ServiceReference1.DirectMessageServiceClient = getProxy()
        Dim response As ServiceReference1.getCountResponse = proxy.GetDeletedMessageCount(getAuthentication(), MailFolder)
        If response.success Then
            Console.WriteLine("Unread Messages Count: " + response.count.ToString)
        Else
            ProcessErrorResponse(response)
        End If
    End Sub

    'get the name of folders
    Public Sub GetFolders(ByVal subscribedFolderOnly As Boolean)
        Dim proxy As ServiceReference1.DirectMessageServiceClient = getProxy()
        Dim response As ServiceReference1.getFoldersResponseType = proxy.GetFolders(getAuthentication(), MailFolder, subscribedFolderOnly)
        Console.WriteLine("Get folder: {0}", response.folders)
    End Sub

    'rename a un-reserved folder
    Public Sub MoveFolder(ByVal folderName As String, ByVal newFolderName As String)
        Dim proxy As ServiceReference1.DirectMessageServiceClient = getProxy()
        Dim response As ServiceReference1.apiResponse = proxy.MoveFolder(getAuthentication(), folderName, newFolderName)
        Console.WriteLine("Move messages from {0} to {1}", folderName, newFolderName)
    End Sub

    'Create a new folder. All folder names should start with "INBOX." And below reserved folders can not be re-created or deleted:
    'INBOX
    'INBOX.Sent 
    'INBOX.Templates    
    'INBOX.Drafts
    'INBOX.Spam
    Public Sub CreateFolder(ByVal newFolderName As String)
        Dim proxy As ServiceReference1.DirectMessageServiceClient = getProxy()
        Dim response As ServiceReference1.apiResponse = proxy.CreateFolder(getAuthentication(), newFolderName)
        ProcessCreateFolder(response, newFolderName)
    End Sub

    'Set a list of folders to subscribed/unsubscribed
    Public Sub SetFolderSubscribed(ByVal folderName As String())
        Dim proxy As ServiceReference1.DirectMessageServiceClient = getProxy()
        Dim response As ServiceReference1.apiResponse = proxy.SetFolderSubscribed(getAuthentication(), folderName, True)
        ProcessSetFolderSubscribed(response, folderName)
    End Sub

    'Get all UIDs in a specified folder. This function is useful to synchronize the folder's messages with server
    Public Sub GetUIDs()
        Dim proxy As ServiceReference1.DirectMessageServiceClient = getProxy()
        Dim response As ServiceReference1.getUIDResponseType = proxy.GetUIDs(getAuthentication(), MailFolder)
        For i As Integer = 0 To response.uids.Length - 1
            Console.WriteLine("Get {0}th UIDs: {1}", i, response.uids(i))
        Next
    End Sub

    'Get all messages from a specified folder.
    Public Sub GetMessages()
        Dim proxy As ServiceReference1.DirectMessageServiceClient = getProxy()
        Dim response As ServiceReference1.getMessageResponse = proxy.GetMessages(getAuthentication(), MailFolder)
        ProcessMessageResponse(response)
    End Sub

    'get message by index
    Public Sub GetMessagebyIndex(ByVal index As Integer)
        Dim proxy As ServiceReference1.DirectMessageServiceClient = getProxy()
        Dim response As ServiceReference1.getMessageResponse = proxy.GetMessageByIndex(getAuthentication(), MailFolder, index)
        ProcessMessageResponse(response)
    End Sub

    'get message by indexes
    Public Sub GetMessagesbyIndexes(ByVal index As Integer())
        Dim proxy As ServiceReference1.DirectMessageServiceClient = getProxy()
        Dim response As ServiceReference1.getMessageResponse = proxy.GetMessagesByIndexes(getAuthentication(), MailFolder, index)
        ProcessMessageResponse(response)
    End Sub

    'get unead messages
    Public Sub GetUnreadMessages()
        Dim proxy As ServiceReference1.DirectMessageServiceClient = getProxy()
        Dim response As ServiceReference1.getMessageResponse = proxy.GetUnreadMessages(getAuthentication(), MailFolder)
        ProcessMessageResponse(response)
    End Sub

    'Get messages filtered by received Date
    'if @beginTime Is null, search messages from the beginning to @entTime
    'if @endTime Is null, search messages from @beginTime to now
    'if woth @beginTime And @endTime are null, get all messages
    '<param name="beginTime">search messages greater than or equals beginTime</param>
    '<param name="endTime">search messages less than or equals endTime </param>
    Public Sub GetMessagesByReceivedDate(ByVal beginTime As DateTime, ByVal endTime As DateTime)
        Try
            Dim proxy As ServiceReference1.DirectMessageServiceClient = getProxy()
            Dim response As ServiceReference1.getMessageResponse = proxy.GetMessagesByReceivedDate(getAuthentication(), MailFolder, GetUnixTime(beginTime), GetUnixTime(endTime))
            ProcessMessageResponse(response)
        Catch ex As CommunicationException
            Console.WriteLine(ex)
        End Try
    End Sub

    'get messages filtered by received date
    'if @beginTime Is null, search messages from the beginning to @entTime
    'beginTime And endTime in UTC. Format: yyyyMMddHHmmss, ex: 20160705174715 == Jul 05,2016 17:47:15 UTC Or 13:47:15 EDT. 
    '<param name="beginTime"> Format: yyyyMMddHHmmss, ex: 20160705174715 == Jul 05,2016 17:47:15 UTC Or 13:47:15 EDT.</param>
    '<param name="endTime"> Format: yyyyMMddHHmmss, ex: 20160705174715 == Jul 05,2016 17:47:15 UTC Or 13:47:15 EDT. </param>
    Public Sub GetMessagesByReceivedDateUTC(ByVal beginTime As String, ByVal endTime As String)
        Try
            Dim proxy As ServiceReference1.DirectMessageServiceClient = getProxy()
            Dim response As ServiceReference1.getMessageResponse = proxy.GetMessagesByReceivedDateUTC(getAuthentication(), MailFolder, beginTime, endTime)
            ProcessMessageResponse(response)
        Catch ex As CommunicationException
            Console.WriteLine(ex)
        End Try
    End Sub

    'get messages filtered by sender pattern
    Public Sub GetMessagesBySender(ByVal senderPattern As String)
        Dim proxy As ServiceReference1.DirectMessageServiceClient = getProxy()
        Dim response As ServiceReference1.getMessageResponse = proxy.GetMessagesBySender(getAuthentication(), MailFolder, senderPattern)
        ProcessMessageResponse(response)
    End Sub

    'get messages filtered by uids
    Public Sub GetMessagesByUID(ByVal uids As Long())
        Dim proxy As ServiceReference1.DirectMessageServiceClient = getProxy()
        Dim response As ServiceReference1.getMessageResponse = proxy.GetMessagesByUID(getAuthentication(), MailFolder, uids)
        ProcessMessageResponse(response)
    End Sub

    'copy messages by UIDs
    Public Sub CopyMessagesByUID(ByVal folderName As String, ByVal newFolderName As String, ByVal uids As Long())
        Dim proxy As ServiceReference1.DirectMessageServiceClient = getProxy()
        Dim response As ServiceReference1.getMessageResponse = proxy.CopyMessagesByUID(getAuthentication(), folderName, uids, newFolderName)
        ProcessCopyMessagesByUID(response, uids)
    End Sub

    'Send a FHIR Query for the patient. The query will be saved into INBOX.Sent folder as an unread message automatically
    Public Sub PatientFHIRQuery(ByVal query As ServiceReference1.fhirQueryType)
        Dim proxy As ServiceReference1.DirectMessageServiceClient = getProxy()
        Dim response As ServiceReference1.sendMessageResponseType = proxy.PatientFHIRQuery(getAuthentication(), query)
        ProcessSendMessageResponseType(response)
    End Sub

    'Forword your messages to recipient
    Public Sub ForwardMessagesByUID(ByVal uids As Long(), ByVal rcpts As ServiceReference1.recipient())
        Dim proxy As ServiceReference1.DirectMessageServiceClient = getProxy()
        Dim response As ServiceReference1.sendMessageResponseType = proxy.ForwardMessagesByUID(getAuthentication(), MailFolder, uids, rcpts)
        ProcessSendMessageResponseType(response)
    End Sub

    'Get information of multiple messages
    Public Sub GetMessagesStatusBySmtpIds(ByVal smtpIds As String())
        Dim proxy As ServiceReference1.DirectMessageServiceClient = getProxy()
        Dim response As ServiceReference1.directMessageStatusReportResponseType = proxy.GetMessagesStatusBySmtpIds(getAuthentication(), smtpIds)
        ProcessDirectMessageStatusReportResponseType(response)
    End Sub

    'Mark messages as READ by UIDs
    Public Sub MarkMessagesAsReadByUID(ByVal uids As Long())
        Dim proxy As ServiceReference1.DirectMessageServiceClient = getProxy()
        Dim response As ServiceReference1.apiResponse = proxy.MarkMessagesAsReadByUID(getAuthentication(), MailFolder, uids)
        ProcessMarkMessagesAsReadByUID(response, uids)
    End Sub

    'Mark messages as UNREAD by UIDs
    Public Sub MarkMessagesAsUnreadByUID(ByVal uids As Long())
        Dim proxy As ServiceReference1.DirectMessageServiceClient = getProxy()
        Dim response As ServiceReference1.apiResponse = proxy.MarkMessagesAsUnReadByUID(getAuthentication(), MailFolder, uids)
        ProcessMarkMessagesAsReadByUID(response, uids)
    End Sub

    'get messages filtered by subject pattern
    Public Sub GetMessagesBySubject(ByVal subjectPattern As String)
        Dim proxy As ServiceReference1.DirectMessageServiceClient = getProxy()
        Dim response As ServiceReference1.getMessageResponse = proxy.GetMessagesBySubject(getAuthentication(), MailFolder, subjectPattern)
        ProcessMessageResponse(response)
    End Sub

    'Get all metadata in the folder.
    Public Sub GetMessageMetadata()
        Dim proxy As ServiceReference1.DirectMessageServiceClient = getProxy()
        Dim response As ServiceReference1.getMessageResponse = proxy.GetMessageMetadata(getAuthentication(), MailFolder)
        ProcessMessageResponse(response)
    End Sub

    'Get all metadata in the folder that filtered by the received date 
    'beginTime Is Unix Time stamp: milliseconds since the Unix Epoch (1970-01-01T00:00:00Z ISO-8601) endTime Is Unix Time stamp: milliseconds since the Unix Epoch (1970-01-01T00:00:00Z ISO-8601) 
    '<param name="beginTime">search messages greater than or equals beginTime</param>
    '<param name="endTime">search messages less than Or equals endTime </param>
    Public Sub SearchMessagesMetadataByReceivedDate(ByVal beginTime As DateTime, ByVal endTime As DateTime)
        Try
            Dim proxy As ServiceReference1.DirectMessageServiceClient = getProxy()
            Dim response As ServiceReference1.getMessageResponse = proxy.SearchMessagesMetadataByReceivedDate(getAuthentication(), MailFolder, GetUnixTime(beginTime), GetUnixTime(endTime))
            ProcessMessageResponse(response)
        Catch ex As CommunicationException
            Console.WriteLine(ex)
        End Try
    End Sub

    'Get all metadata in the folder that filtered by the received date 
    'beginTime and endTime in UTC. Format: yyyyMMddHHmmss, ex: 20160705174715 == Jul 05,2016 17:47:15 UTC or 13:47:15 EDT.
    '<param name="beginTime">search messages greater than or equals beginTime</param>
    '<param name="endTime">search messages less than or equals endTime </param>

    Public Sub SearchMessagesMetadataByReceivedDateUTC(ByVal beginTime As String, ByVal endTime As String)
        Try
            Dim proxy As ServiceReference1.DirectMessageServiceClient = getProxy()
            Dim response As ServiceReference1.getMessageResponse = proxy.SearchMessagesMetadataByReceivedDateUTC(getAuthentication(), MailFolder, beginTime, endTime)
            ProcessMessageResponse(response)
        Catch ex As CommunicationException
            Console.WriteLine(ex)
        End Try
    End Sub
End Class

