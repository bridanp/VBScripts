' Needed to break pages based on a section of the title page
' Code work like a find/replace where the admin inputs the Page_Break
' And the code replaces it with the page break character of Chr(12)
' Written specifically for IpSwitch MoveIt but could be used for other applications

Dim fso, tsin, tsout
Dim MyLine, MyReport
Dim strText
Dim PgBreak

Set fso = CreateObject("Scripting.FileSystemObject")

MyReport = MINewCacheFilename()
Set tsin = fso.OpenTextFile(MICacheFilename, 1)
Set tsout = fso.OpenTextFile(MyReport, 2, 1)
' MIGetTaskParam is MoveIt specific, getting the parameter from user input
PgBreak = MIGetTaskParam("Page_Break")

While Not tsin.AtEndOfStream
  MyLine = tsin.ReadLine
  strText = Replace(MyLine, PgBreak, Chr(12))
  tsout.WriteLine (strText)
Wend

tsin.Close
tsout.Close

MIReplaceCacheFile (MyReport)

Set tsin = Nothing
Set tsout = Nothing
Set fso = Nothing
