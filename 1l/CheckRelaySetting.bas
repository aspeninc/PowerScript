' ASPEN PowerScript Sample Program
'
' CHECKRELAYSETTING.BAS
'
' Run OneLiner command: Check relay setting
'
' Version: 1.0
' Category: OneLiner
'
' sInput must include a XML node with name CHECKPRIBACKCOORD  and with attributes 
' in the list below (* denotes required entries, [] denote default value).
'
' Cmd name CHECKRELAYSETTINGS
'
' Attributes (*: required; []: default)
'
'  SELECTEDOBJ=
'       PICKED Check line with selected relaygroup
'       BNO1;'BNAME1';KV1;BNO2;'BNAME2';KV2;'CKT';BTYP;  location string of line to check(Help section 10.2)
'  TIERS= check lines in vicinity within this tier number
'  AREAS= Check all lines in area range
'  ZONES= Check all lines in zone range
'  KVS=   Additional KV filter
'  TAGS=  Additional tag filter
'  REPORTPATHNAME= (*) full valid path to report folder with write access
'  REPORTCOMMENT= Report comment string. 255 char or shorter
'  FAULTTYPE= 1LG, 3LG. Fault type to check. Space delimited
'  DEVICETYPE= OCG, OCP, DSG, DSP, LOGIC, VOLTAGE, DIFF Devide type to check. Space delimited
'  OUTAGELINES	Run Line outage contingency: 0-False; 1-True
'  OUTAGEXFMRS	Run transformer outage contingency: 0-False; 1-True
'  OUTAGE3SOURCES= 1 or 0 Outage only 3 strongest sources
'  OUTAGEMULINES= 1 or 0 Outage mutually coupled lines
'  OUTAGEMULINESGND= 1 or 0 Outage and ground ends of mutually coupled lines
'  OUTAGE2LINES= 1 or 0 Double outage lines
'  OUTAGE1LINE1XFMR= 1 or 0 Double outage line and transformer
'  OUTAGE2XFMR= 1 or 0 Double outage transformers
'

Sub main
  sInput$ = "<CHECKRELAYSETTINGS " & _
            "SELECTEDOBJ=""PICKED"" " & _
            "REPORTPATHNAME=""c:\000tmp\checksettings.csv"" " & _
            " />"
  Print sInput
  If Run1LPFCommand( sInput ) Then 
    Print "Success"
  Else 
    Print ErrorString()
  End If
End Sub

Function AdditionalInputStringSamples
  sInput$ = "<CHECKRELAYSETTINGS " & _
            "REPORTPATHNAME=""c:\000tmp\checksettings.csv"" " & _
            "DEVICETYPE=""OCP OCG"" " & _
            "KVS=""0-9999"" " & _
            "TAGS="""" " & _
            "OUTAGELINES=""1"" " & _ 
            "OUTAGE2LINES=""1"" " & _ 
            " />"
  sInput$ = "<CHECKRELAYSETTINGS " & _
            "SELECTEDOBJ=""PICKED"" " & _
            "REPORTPATHNAME=""c:\000tmp\checksettings.csv"" " & _
            " />"
  sInput$ = "<CHECKRELAYSETTINGS " & _
            "REPORTPATHNAME=""c:\000tmp\"" " & _
            " />"

  sInput$ = "<CHECKRELAYSETTINGS " & _
            "REPORTPATHNAME=""c:\000tmp\checksettings.csv"" " & _
            "SELECTEDOBJ=""6; 'NEVADA'; 132.; 8; 'REUSENS'; 132.; '1'; 1;"" " & _
            " />"

End Function
