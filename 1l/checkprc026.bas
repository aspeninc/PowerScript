' ASPEN PowerScript Sample Program
'
' CHECKPRC026.BAS
'
' Run CHECKRELAYOPERATIONPRC026 command
'
' Version: 1.0
' Category: OneLiner
'
' sInput must include a XML node with name CHECKRELAYOPERATIONPRC026  and with attributes 
' in the list below (* denotes required entries, [] denote default value).
'
' Cmd name CHECKRELAYOPERATIONPRC026
'
' Attributes (*: required; []: default)
'
'  REPORTPATHNAME= (*) full valid pathname of report file
'  REPORTCOMMENT= Report comment string. 255 char or shorter
'  SELECTEDOBJ=
'       PICKED Check devices in selected relaygroup
'       BNO1;'BNAME1';KV1;BNO2;'BNAME2';KV2;'CKT';BTYP;  location string of line to check(Help section 10.2)
'  TIERS= check relaygroups in vicinity within this tier number
'  AREAS= Check all relaygroups in area range
'  ZONES= Check all relaygroups in zone range
'  KVS=   Additional KV filter
'  TAGS=  Additional tag filter
'  DEVICETYPE= [OCP DSP] Devide type to check. Space delimited
'  APPENDREPORT=	Append report file: 0-False; [1]-True
'  SEPARATIONANGLE=	[120] System separation angle for stable power swing calculation
'  ULOSSRVRATIO=	[1.43] Upper loss-of-synchronism circle sending-end to receiving-end voltage ratio
'  LLOSSRVRATIO=	[0.7] Lower loss-of-synchronism circle sending-end to receiving-end voltage ratio
'  DELAYLIMIT= [15] Report violation if relay trips faster than this limit (in cycles)
'  CURRMULT= [1.0] Current multiplier to apply in relay trip checking
'

Sub main
  sReportFile$ = GetOlrFileName()
  For ii = Len(sReportFile) to 1 step -1
    If Mid( sReportFile, ii, 1 ) = "\" Then exit For
  Next
  sReportFile = Left(sReportFile, ii) & "prc026.xml"  
  sInput$ = "<CHECKRELAYOPERATIONPRC026 " & _
            "REPORTPATHNAME=""" & sReportFile  & """ " & _
            "APPENDREPORT=""0"" " & _
            "DEVICETYPE=""DSP"" " & _
            "KVS=""300-9999"" " & _
            "SEPARATIONANGLE=""120"" " & _
            "ULOSSRVRATIO=""1.43"" " & _
            "LLOSSRVRATIO=""0.7"" " & _
            "CURRMULT=""1.1"" " & _
            " />"
  Print sInput
  If Run1LPFCommand( sInput ) Then 
    Print "Success Report in " & sReportFile
  Else 
    Print "Error: " & ErrorString()
  End If
End Sub

Function AdditionalInputStringSamples
  sInput$ = "<CHECKRELAYOPERATIONPRC026 " & _
            "REPORTPATHNAME=""" & sReportFile  & """ " & _
            "DEVICETYPE=""DSP"" " & _
            "KVS=""200-9999"" " & _
            "TAGS=""prc026"" " & _
            "SEPARATIONANGLE=""115"" " & _
            "CURRMULT=""1.1"" " & _
            " />"
  sInput$ = "<CHECKRELAYOPERATIONPRC026 " & _
            "REPORTPATHNAME=""" & sReportFile  & """ " & _
            "KVS=""200-500"" " & _
            "APPENDREPORT=""1"" " & _
            " />"
            
  sInput$ = "<CHECKRELAYOPERATIONPRC026 " & _
            "REPORTPATHNAME=""" & sReportFile  & """ " & _
            "SELECTEDOBJ=""0; 'CLAYTOR'; 132; 0; 'NEVADA'; 132; '1'; 1;"" " & _
            "APPENDREPORT=""1"" " & _
            " />"

  sInput$ = "<CHECKRELAYOPERATIONPRC026 " & _
            "REPORTPATHNAME=""" & sReportFile  & """ " & _
            "SELECTEDOBJ=""PICKED"" " & _
            "APPENDREPORT=""1"" " & _
            " />"

End Function
