' This VBS script is used to lengthen to a specific record length
' NACHA files specifically need to be a certain length for mainframe usage
' Code is written for IpSwitch MoveIt input but could be used for other applications

' DECLARE VARIABLES
Dim fso, tsin, tsout
Dim MyText
Dim WrapText
Dim FileLength
Dim TextOut
Dim Record_Length

' OPEN AND READ THE INPUT FILE
Set fso = CreateObject("Scripting.FileSystemObject")

' 1=ForReading, 2=ForWriting, 8=ForAppending
Set tsin = fso.OpenTextFile(MICacheFilename, 1)
MyText = tsin.ReadAll
tsin.Close
Set tsin = Nothing
p = 1
Record_Length = MIGetTaskParam("Record_Length")
FileLength = Len(MyText)

' PROCESS FILE
' Process entire block of text here

' OPEN AND WRITE THE OUTPUT FILE
Set tsout = fso.OpenTextFile(MICacheFilename, 2, 1)
Do while p <= FileLength
   WrapText = (mid(MyText,p,Record_Length)) & vbCrLf
   TextOut = TextOut + WrapText
   p = p + Record_Length				  
Loop

tsout.Write TextOut
tsout.Close
Set tsout = Nothing
Set fso = Nothing
