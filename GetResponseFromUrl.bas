Function getResponse(params As String)

Dim hReq As Object

Dim strUrl As String
    strUrl = "htts://www.url.com.br?params=" & params

Set hReq = CreateObject("MSXML2.XMLHTTP")
    With hReq
        .Open "GET", strUrl, False
        .Send
    End With

getResponse = hReq.ResponseText

End Function
