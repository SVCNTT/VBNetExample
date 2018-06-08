Imports System.IO
Imports System.Net
Imports System.Text

Module MoMoExample

    Sub Main()
        Dim PartnerCode = "MOMOIQA420180417" 'Lấy từ config 
        Dim BillId = "30F8777D-8E69-4C37-80E2-86570C319C38" 'Lấy từ config 
        Dim MoMoSecretKey = "PPuDXq1KowPT1ftR8DvlQTHhC03aul17" 'Lấy từ config 
        'Dim MoMoServiceUrl = "http://testing.momo.vn:18099/pay/query-status" 'Lấy từ config
        Dim MoMoServiceUrl = "http://172.16.13.25:18099/pay/query-status" 'Lấy từ config
        Dim RequestId = System.Guid.NewGuid.ToString()
        Dim DataBeforeHash = New StringBuilder("partnerCode=").Append(PartnerCode).Append("&billId=").Append(BillId).Append("&requestId=").Append(RequestId).ToString()
        Dim Signature = HashString(DataBeforeHash, MoMoSecretKey)
        Dim DataSendRequest As String
        DataSendRequest = New StringBuilder("{""partnerCode"" : """).Append(PartnerCode).Append(""", ").Append("""billId"" : """).Append(BillId).Append(""",").Append("""requestId"" : """).Append(RequestId).Append(""",").Append("""signature"" : """).Append(Signature).Append("""}").ToString()
        Console.WriteLine("Call Service Response data : {0}", DataSendRequest)
        Dim DataMoMoResponse = HttpSendPost(DataSendRequest, MoMoServiceUrl)
        Dim MoMoResponseData As MoMoResponseData = Newtonsoft.Json.JsonConvert.DeserializeObject(Of MoMoResponseData)(DataMoMoResponse)
        Console.WriteLine("Call Service Response Status : {0}", MoMoResponseData.status)
        Console.WriteLine("Call Service Response Message : {0}", MoMoResponseData.message)
        If MoMoResponseData.status = 0 Then
            Console.WriteLine("Order Transaction get Status in MoMo Service : {0}", MoMoResponseData.data.status)
            Console.WriteLine("Order Transaction get Message in MoMo Service : {0}", MoMoResponseData.data.message)
            'Todo check signature
            Try
                Dim SignatureField = ",""signature"":""" + MoMoResponseData.signature + """"
                Console.WriteLine("My Service Get MoMo SignatureField : {0}", SignatureField)
                Dim MoMoResponseDataRaw = DataMoMoResponse.Replace(SignatureField, "")
                Console.WriteLine("My Service Parse MoMo MoMoResponseDataRaw : {0}", MoMoResponseDataRaw)
                Dim MySiganture = HashString(MoMoResponseDataRaw, MoMoSecretKey)
                Console.WriteLine("MySiganture : {0}", MySiganture)
                If MySiganture.Equals(MoMoResponseData.signature) Then
                    Console.WriteLine("My Service Check Signature Return : PASS")
                    ' ToDo Update Data And Close Bill Success .... 
                Else
                    Console.WriteLine("My Service Check Signature Return : FAIL")
                End If
            Catch ex As Exception
                ' Check signature Fail Recall Query Status Or Close Bill With Another Payment Flow
            End Try
        End If

        Console.ReadLine()
    End Sub

    Private Function HttpSendPost(postdata As String, momoUrl As String)
        Dim webRequest As HttpWebRequest
        Dim enc As UTF8Encoding
        Dim postdatabytes As Byte()
        Dim serviceResponse As String
        Dim encode As Encoding = System.Text.Encoding.GetEncoding("utf-8")
        webRequest = HttpWebRequest.Create(momoUrl)
        enc = New System.Text.UTF8Encoding()

        postdatabytes = enc.GetBytes(postdata)
        webRequest.Method = "POST"
        webRequest.ContentType = "application/json"
        webRequest.ContentLength = postdatabytes.Length
        webRequest.Timeout = 30000
        Try
            Using stream = webRequest.GetRequestStream()

                stream.Write(postdatabytes, 0, postdatabytes.Length)
                Dim result = webRequest.GetResponse().GetResponseStream()
                Dim reader = New StreamReader(result, encode)
                serviceResponse = reader.ReadToEnd()
                reader.Close()
            End Using
        Catch ex As Exception
            Console.WriteLine("Get request catch exception : {0}", ex)
            serviceResponse = "{""status"": -1, ""message"":""HttpRequest error"", ""data"" : {""status"" : -1, ""message"" : ""HttpRequest error""}}"
        End Try
        Console.WriteLine("Call Service Response data : {0}", serviceResponse)
        Return serviceResponse

    End Function

    Private Function HashString(StringToHash As String, Key As String) As String
        Dim myEncoder As New System.Text.UTF8Encoding
        Dim dataByte() As Byte = myEncoder.GetBytes(StringToHash)
        Dim myHMACSHA256 As New System.Security.Cryptography.HMACSHA256(myEncoder.GetBytes(Key))
        Dim HashCode As Byte() = myHMACSHA256.ComputeHash(dataByte)
        Dim hash As String = Replace(BitConverter.ToString(HashCode), "-", "")
        Return hash.ToLower
    End Function

    Public Class MoMoResponseData
        Public Property status As Int32
        Public Property message As String
        Public Property data As TransactionBodyData
        Public Property signature As String
    End Class

    Public Class TransactionBodyData
        Public Property status As Int32
        Public Property message As String
        Public Property partnerCode As String
        Public Property billId As String
        Public Property transId As String
        Public Property amount As String
        Public Property discountAmount As String
        Public Property fee As String
        Public Property phoneNumber As String
        Public Property storeId As String
        Public Property requestDate As String
        Public Property responseDate As String
        Public Property customerName As String
    End Class
End Module