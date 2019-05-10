' ASPEN PowerScript Sample Program
'
' GETAREAZONENAME.BAS
'
' Demo PowerScript functions GetAreaName and GetZoneName
'
' Version: 1.0
'
Sub main

  Do
    nNo = Val(InputBox( "Area no" ))
    If nNo = 0 Then Stop
    Print "Area " & nNo & " name=" & GetAreaName( nNo )
    Print "Zone " & nNo & " name=" & GetZoneName( nNo )
  Loop

End Sub
