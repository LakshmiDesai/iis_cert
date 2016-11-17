dim xHttp: Set xHttp = createobject("MSXML2.ServerXMLHTTP")

xHttp.Open "GET", "https://yourhost.example.com/foo", False

' 2 stands for SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS
' 13056 means ignore all server side cert error
xHttp.setOption 2, 13056
xHttp.Send

' read response body
WScript.Echo xHttp.responseBody