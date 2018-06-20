' ASPEN PowerScrip sample program
'
' CHECKRELAYCOORD.BAS
'
' Run OneLiner command: Check primary/backup relay coordination
'
' Version: 1.0
' Category: OneLiner
'
'
' sInput must include a XML node with name CHECKPRIBACKCOORD  and with attributes 
' in the list below (* denotes required entries, [] denote default value).
'
' Cmd name CHECKPRIBACKCOORD
'
' Attributes
'  REPFILENAME= (*) pathname of report file
'  OUTFILETYPE 	[2] Output file type 1- TXT; 2- CSV
'     SELECTEDOBJ	Relay group to check against backups. Must  have one of following values
'        PICKED  	Check the highlighted relaygroup on the 1-line diagram
'        BNO1;'BNAME1';KV1;BNO2;'BNAME2';KV2;'CKT';BTYP;  location string of relay to check per format in help section 10.2.
'  TIERS	[0] Number of tiers around selected object. This attribute is ignored if SELECTEDOBJ is not found.
'  AREAS	[0-9999] Comma delimited list of area numbers and ranges to check against backups.
'  ZONES	[0-9999] Comma delimited list of zone numbers and ranges to check against backups. This attribute is ignored if AREAS is found.
'  KVS	0-999] Comma delimited list of KV leves and ranges to check against backups. This attribute is ignored if SELECTEDOBJ is found.
'  TAGS	Comma delimited list of tags to check against backups. This attribute is ignored If SELECTEDOBJ is found.
'  COORDTYPE	Coordination Type to check. Must  have one of following values
'      0	OC backup/OC primary (Classical)
'      1	OC backup/OC primary (Multi-point)
'      2	DS backup/OC primary
'      3	OC backup/DS primary
'      4	DS backup/DS primary
'      5	OC backup/Recloser primary
'      6	All types/All types
'  LINEPERCENT	Percent interval For sliding intermediate faults. This attribute is ignored If COORDTYPE is 0 Or 5.
'  RUNINTERMEOP	1-true; 0-false. Check  intermediate faults With End-opened. This attribute is ignored If COORDTYPE is 0 Or 5.
'  RUNCLOSEIN	1-true; 0-false. Check Close-in fault. This attribute is ignored If COORDTYPE is 0 Or 5.
'  RUNCLOSEINEOP	1-true; 0-false. Check Close-in fault With End-opened. This attribute is ignored If COORDTYPE is 0 Or 5.
'  RUNLINEEND	1-true; 0-false. Check Line-End fault. This attribute is ignored If COORDTYPE is 0 Or 5.
'  RUNREMOTEBUS	1-true; 0-false. Check remote bus fault. This attribute is ignored If COORDTYPE is 0 Or 5.
'  RELAYTYPE	Relay types to check: 1-Ground; 2-Phase; 3-Both.
'  FAULTTYPE	Fault  types to check: 1-3LG; 2-2LG; 4-1LF; 8-LL; Or sum of values For desired selection
'  OUTPUTALL	1- Include all cases in report; 0- Include only flagged cases in report
'  MINCTI	Lower limit of acceptable CTI range
'  MAXCTI	Upper limit of acceptable CTI range
'  OUTRLYPARAMS	Include relay settings in report: 0-None; 1-OC;2-DS;3-Both
'  OUTAGELINES	Run Line outage contingency: 0-False; 1-True
'  OUTAGEXFMRS	Run transformer outage contingency: 0-False; 1-True
'  OUTAGEMULINES	Run mutual Line outage contingency: 0-False; 1-True
'  OUTAGEMULINESGND 	Run mutual Line outage And grounded contingency: 0-False; 1-True
'  OUTAGE2LINES	Run double Line outage contingency: 0-False; 1-True
'  OUTAGE1LINE1XFMR	Run double Line And transformer outage contingency: 0-False; 1-True
'  OUTAGE2XFMRS	Run double And transformer outage contingency: 0-False; 1-True
'  OUTAGE3SOURCES	Outage only  3 strongest sources: 0-False; 1-True
'

Sub main()

  sInput$ = "<CHECKPRIBACKCOORD " & _
            "REPFILENAME=""c:\\000tmp\\checkcoord.csv"" " & _
            "OUTFILETYPE=""1"" " & _
            "SELECTEDOBJ=""6; 'NEVADA'; 132.; 8; 'REUSENS'; 132.; '1'; 1;"" " & _
            "COORDTYPE=""6"" " & _
            "OUTPUTALL=""1"" " & _
            "MINCTI=""0.05"" " & _ 
            "MAXCTI=""99"" " & _ 
            "LINEPERCENT=""15"" " & _ 
            "RELAYTYPE=""3"" " & _ 
            "FAULTTYPE=""5"" " & _ 
            " />"
  Print sInput
  If Run1LPFCommand( sInput ) Then 
    Print "Success"
  Else 
    Print ErrorString()
  End If
End Sub

