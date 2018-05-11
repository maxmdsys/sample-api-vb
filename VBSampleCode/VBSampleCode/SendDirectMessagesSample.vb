Imports System.IO
Imports System.ServiceModel

Public Class SendDirectMessagesSample
    Public Property Username As String
    Public Property Password As String
    Public Property EndpointURL As String
    Public Property Recipients As ServiceReference1.recipient()

    Private Function GetProxy() As ServiceReference1.DirectMessageServiceClient
        Dim proxy As New ServiceReference1.DirectMessageServiceClient
        If EndpointURL.StartsWith("https://", StringComparison.CurrentCultureIgnoreCase) Then
            'binding.MaxBufferSize = 10000000
            Dim binding As New BasicHttpsBinding With {
                .MaxReceivedMessageSize = 10000000,
                .MessageEncoding = WSMessageEncoding.Mtom
            }
            proxy.Endpoint.Binding = binding
        Else
            'binding.MaxBufferSize = 10000000
            Dim binding As New BasicHttpBinding With {
                .MaxReceivedMessageSize = 10000000,
                .MessageEncoding = WSMessageEncoding.Mtom
            }
            proxy.Endpoint.Binding = binding
        End If
        Dim endpointAddress As New EndpointAddress(EndpointURL)
        proxy.Endpoint.Address = endpointAddress
        Return proxy
    End Function

    Private Function GetRequest() As ServiceReference1.sendRequest
        Dim authentication As New ServiceReference1.authentication With {
            .username = Username,
            .password = Password
        }
        Dim att1 As New ServiceReference1.attachment With {
            .content = File.ReadAllBytes("C:\test\a.txt"),'Replace with your attachment file path here
            .contentType = "text/plain",'if the attachment is a text file, type should be text/plain
            .filename = "a.txt"'Replace it with your file name
        }
        'Dim att2 As New ServiceReference1.attachment With {
        '    .content = File.ReadAllBytes("C:\test\a.pdf"),
        '    .contentType = "application/pdf",
        '    .filename = "a.pdf"
        '}

        'Dim attachments As ServiceReference1.attachment() = {att1, att2}
        Dim attachments As ServiceReference1.attachment() = {att1}

        Dim message As New ServiceReference1.message With {
            .sender = Username,
            .subject = "Test Direct Message",'Replace with your subject
            .recipients = Recipients,
            .body = "Test <strong>Direct</strong> message Body",'Replace with your message body
            .htmlBody = True,
            .attachmentList = attachments
        }

        Dim request As New ServiceReference1.sendRequest With {
            .authentication = authentication,
            .message = message
        }
        Return request
    End Function

    Public Sub Send()
        Dim proxy As ServiceReference1.DirectMessageServiceClient = GetProxy()
        Dim response As ServiceReference1.apiResponse = proxy.Send(GetRequest())
        If response.success Then
            Console.WriteLine("Sent Direct message completed.")
        Else
            Select Case response.code
                Case 1
                    Console.WriteLine("authentication failed")
                Case 2
                    Console.WriteLine("Incorrect addresses")
                Case 3
                Case 4
                    Console.WriteLine("Other error")
            End Select
            Console.WriteLine("Information of the SOAP request: " + response.message)
        End If
    End Sub
End Class
