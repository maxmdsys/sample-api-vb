Module Main
    Dim productionEndpointURL As String = "https://rs4b.max.md:8445/message/services/DirectMessageService"
    Dim evaluationEndpointURL As String = "https://rs5c.max.md:8445/message/services/DirectMessageService"
    Dim smtpId As String = ""
    Sub Main()
        Dim sendSample As New SendDirectMessagesSample With {
            .EndpointURL = productionEndpointURL,'Change it to evaluation endpoint url if you use service reference on evaluation server
            .Username = "username@directDomain",'Replace with your direct address
            .Password = "",'Replace with your password
            .Recipients = {
                    New ServiceReference1.recipient With {
                    .email = "recipient1@directDomain",'Replace with recipient direct address
                    .type = ServiceReference1.recipientType.TO,
                    .typeSpecified = True
                }
            }
        }

        ''uncomment line below to send a test direct email.
        'sendSample.Send()

        Dim retrieveSample As New RetrieveDirectMessagesSample With {
            .EndpointURL = productionEndpointURL,'Change it to evaluation endpoint url if you use service reference on evaluation server
            .Username = "username@directDomain",'Replace with your direct address
            .Password = "",'Replace with your password
            .OutputFolder = "C:\test\messages",'Replace with your output folder
            .MailFolder = "INBOX"'Replace with your mail folder name
        }

        Dim rcpts As ServiceReference1.recipient() = {
                New ServiceReference1.recipient With {
                .email = "recipient1@directDomain",'Replace with recipient direct address
                .type = ServiceReference1.recipientType.TO,
                .typeSpecified = True
            }
        }

        Dim fhirqpt As ServiceReference1.fhirQueryParameterType() = {
                New ServiceReference1.fhirQueryParameterType With {
                .name = "",
                .value = ""
            }
        }

        Dim fhirrt As ServiceReference1.fhirResourceType() = {
                New ServiceReference1.fhirResourceType With {
                .queryParameters = fhirqpt,
                .resourceType = ""
            }
        }

        Dim fhirqt As New ServiceReference1.fhirQueryType With {
            .identifier = "",
            .recipients = rcpts,
            .resources = fhirrt,
            .subject = ""
        }

        ''uncomment below lines to get results.
        'retrieveSample.GetMessageCount()
        'retrieveSample.GetUnreadMessageCount()
        'retrieveSample.GetDeletedMessageCount()
        'retrieveSample.GetMessagebyIndex(1)
        'retrieveSample.GetMessagesByIndexes(New Integer() {1, 2})
        'retrieveSample.GetMessages()
        'retrieveSample.GetMessagesByReceivedDate(New DateTime(2018, 4, 20), New DateTime(2018, 4, 30))
        'retrieveSample.GetMessagesBySender("recipient1@directDomain")
        'retrieveSample.GetMessagesBySubject("test")
        'retrieveSample.MarkMessagesAsReadByUID(New Long() {1})
        'retrieveSample.MarkMessagesAsUnreadByUID(New Long() {1})
        'retrieveSample.GetMessagesStatusBySmtpIds(New String() {smtpId})
        'retrieveSample.ForwardMessagesByUID(New Long() {1}, rcpts)
        'retrieveSample.GetFolders(False)
        'retrieveSample.CreateFolder("INBOX.hello")
        'retrieveSample.SetFolderSubscribed(New String() {"INBOX.hello"})
        'retrieveSample.CreateFolder("INBOX.newFolder")
        'retrieveSample.SetFolderSubscribed(New String() {"INBOX.newFolder"})
        'retrieveSample.MoveFolder("INBOX.hello", "INBOX.newFolder")
        'retrieveSample.getMessagesByUID(New Long() {1})
        'retrieveSample.CopyMessagesByUID("INBOX", "INBOX.hello", New Long() {1})
        'retrieveSample.GetUIDs()
        'retrieveSample.GetMessagesByReceivedDateUTC("20170623000000", "20170623235959")
        'retrieveSample.GetMessageMetadata()
        'retrieveSample.SearchMessagesMetadataByReceivedDate(New DateTime(2017, 6, 23), New DateTime(2017, 6, 23))
        'retrieveSample.SearchMessagesMetadataByReceivedDateUTC("20170623000000", "20170623235959")
        'retrieveSample.PatientFHIRQuery(fhirqt)
    End Sub

End Module
