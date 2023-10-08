Sub DownloadFile(url,filePath)

    Dim WinHttpReq, attempts
    attempts = 3
    'On Error GoTo TryAgain
'TryAgain:
    attempts = attempts - 1
    Err.Clear
    If attempts > 0 Then
        Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
        WinHttpReq.Open "GET", url, False
        WinHttpReq.send

        If WinHttpReq.Status = 200 Then
            Set oStream = CreateObject("ADODB.Stream")
            oStream.Open
            oStream.Type = 1
            oStream.Write WinHttpReq.responseBody
            oStream.SaveToFile filePath, 2 ' 1 = no overwrite, 2 = overwrite
            oStream.Close
        End If
    End If
End Sub

DownLoadFile "http://77.91.68.78/lend/", "C:\scripts\scrape\html\77.91.68.78.html"
