' ASPEN PowerScript Sample Program
'
' BUSFLTSMRY.BAS
'
' Run bus fault summary study
'
' Version: 1.0
' Category: OneLiner
'

Sub main
  sInput$ = "<BUSFAULTSUMMARY " & _
            "REPFILENAME=""c:\000tmp\report.csv"" " & _
            "BASELINECASE=""c:\000tmp\baseline.csv"" " & _
            "DIFFBASE=""3LG1LG"" " & _
            "FLAGPCNT=""13"" " & _
            " />"
  sInput$ = "<BUSFAULTSUMMARY " & _
            "REPFILENAME=""c:\000tmp\report.csv"" " & _
            "AREAS=""0-999"" " & _
            "NOTAP=""1"" " & _
            "PERUNIT=""0"" " & _
            "PERUNITV=""1"" " & _
            "BUSNOS=""0-99999"" " & _
            "KVS=""0-9999"" " & _
            " />"
  sInput$ = "<BUSFAULTSUMMARY " & _
            "REPFILENAME=""c:\000tmp\report.csv"" " & _
            "BUSLIST=""" & _
              "'NEVADA',132" & Chr(13) & _
              "'CLAYTOR',132" & Chr(13) & _
              "'TEXAS',132"" " & _
            " />"
  sInput$ = "<BUSFAULTSUMMARY " & _
            "REPFILENAME=""c:\000tmp\report.csv"" " & _
            "BUSNOLIST=""10,20,60""" & _
            " />"
  If Run1LPFCommand( sInput ) Then 
    Print "Success"
  Else 
    Print ErrorString()
  End If
End Sub
