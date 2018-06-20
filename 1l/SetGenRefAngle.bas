' ASPEN PowerScript Sample Program
'
' SETGENREFANGLE.BAS
'
' Run Set Generator Reference Angle command 
'
' Version: 1.0
' Category: OneLiner
'
'
' Cmd name SETGENREFANGLE
'
' Attributes (Entries with * are required)
'
'  REPORTPATHNAME= full valid pathname to report file
'        Default: None. Output to TTY window only.
'  REFERENCEGEN= Bus name and kV of reference generator in format: 'BNAME', KV
'        Default: None. Use existing reference generator
'  EQUSOURCEOPTION= Option for calculating equivalent sources ref angle
'       ROTATE (default) apply delta angle of existing reference gen
'       SKIP   Leave unchanged. This option will be in effect
'              automatically when old reference is not valid
'       ASGEN  Use angle computed for regular generator
'
Sub main

  sInput$ = "<SETGENREFANGLE " & _
            " />"

  Print sInput
  If Run1LPFCommand( sInput ) Then 
    Print "Success"
  Else 
    Print ErrorString()
  End If
End Sub

Function AdditionalInputSamples
  sInput$ = "<SETGENREFANGLE " & _
            " />"
  sInput$ = "<SETGENREFANGLE " & _
            "REPORTPATHNAME=""c:\000tmp\setrefangle.txt"" " & _
            "EQUSOURCEOPTION=""SKIP"" " & _
            "REFERENCEGEN=""'ARIZONA' 13.8"" " & _
            " />"

End Function
